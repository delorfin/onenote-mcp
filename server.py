"""
OneNote MCP Server (Local Files)
=================================
An MCP (Model Context Protocol) server that reads local OneNote (.one) files
directly from disk and writes to OneNote via the COM API.
No Azure registration or authentication needed.

Reading: parses backup files at:
    C:\\Users\\<user>\\AppData\\Local\\Microsoft\\OneNote\\16.0\\Backup\\

Writing: uses the OneNote COM API via PowerShell (requires OneNote desktop app).

It exposes tools for Claude Code to:
    - List all notebooks and sections
    - Read page text content
    - Search across all pages (semantic + exact match)
    - Create new pages in any notebook/section
    - Append content to existing pages

Prerequisites:
    pip install "mcp[cli]" pyOneNote
    (or: uv add "mcp[cli]" pyOneNote)
    + OneNote desktop app (for write features)

Usage with Claude Code:
    claude mcp add --transport stdio onenote -- uv --directory "path/to/this/project" run server.py
"""

import logging
import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

from pyOneNote.OneDocument import OneDocment
from mcp.server.fastmcp import FastMCP
from vector_index import EmbeddingIndex
from ocr import ocr_image

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Where OneNote stores local backup files.
# Override with ONENOTE_BACKUP_DIR environment variable if yours is elsewhere.
if sys.platform == "darwin":
    _container = Path.home() / "Library" / "Containers" / "com.microsoft.onenote.mac"
    _base = _container / "Data" / "Library" / "Application Support" / "Microsoft User Data" / "OneNote" / "15.0"
    # "Sicherung" is the German-locale name for "Backup"; check both
    DEFAULT_BACKUP_DIRS = [_base / "Sicherung", _base / "Backup"]
else:
    _win_base = Path(os.environ.get("APPDATA", "")).parent / "Local" / "Microsoft" / "OneNote" / "16.0"
    DEFAULT_BACKUP_DIRS = [_win_base / "Backup"]

ONENOTE_DIRS: list[Path] = []
if os.environ.get("ONENOTE_BACKUP_DIR"):
    ONENOTE_DIRS = [Path(os.environ["ONENOTE_BACKUP_DIR"])]
else:
    ONENOTE_DIRS = [d for d in DEFAULT_BACKUP_DIRS if d.exists()]

# ---------------------------------------------------------------------------
# Logging (to stderr so it doesn't break stdio MCP transport)
# ---------------------------------------------------------------------------

LOG_FILE = os.path.join(tempfile.gettempdir(), "onenote_mcp.log")

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stderr),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger("onenote-mcp")
log.info("Log file: %s", LOG_FILE)

# ---------------------------------------------------------------------------
# OneNote file parsing helpers
# ---------------------------------------------------------------------------


def _discover_notebooks() -> dict[str, dict]:
    """
    Scan the OneNote backup directory and build a notebook -> section -> files map.

    Returns a dict like:
    {
        "My Notebook": {
            "path": Path(...),
            "sections": {
                "Algorithm": {
                    "files": [Path("Algorithm (On 1-4-2026).one"), ...],
                    "latest": Path(...)   # most recently modified
                },
                ...
            }
        },
        ...
    }
    """
    if not ONENOTE_DIRS:
        log.error("No OneNote backup directories found")
        return {}

    notebooks = {}
    for onenote_dir in ONENOTE_DIRS:
        if not onenote_dir.exists():
            continue
        for notebook_dir in onenote_dir.iterdir():
            if not notebook_dir.is_dir():
                continue

            notebook_name = notebook_dir.name

            # Merge into existing entry if notebook already found in another dir
            if notebook_name not in notebooks:
                notebooks[notebook_name] = {
                    "path": notebook_dir,
                    "sections": {},
                }
            sections = notebooks[notebook_name]["sections"]

            # Walk all .one files in this notebook (including subdirectories)
            for one_file in notebook_dir.rglob("*.one"):
                # Skip recycle bin
                if "RecycleBin" in str(one_file):
                    continue

                # Extract the base section name (strip the date suffix)
                # e.g. "Algorithm (On 1-4-2026).one" -> "Algorithm"
                # e.g. "Daily (On 02.02.26).one" -> "Daily"
                fname = one_file.name
                # Remove .one extension(s) and date suffixes
                section_name = re.sub(r"\.one$", "", fname)
                section_name = re.sub(r"\s*\(On [\d.\-]+\)$", "", section_name)
                section_name = re.sub(r"\.one$", "", section_name)  # handle double .one
                section_name = section_name.strip()

                if not section_name:
                    section_name = "(unnamed)"

                # Build relative path for context (subfolder within notebook)
                rel_parts = one_file.parent.relative_to(notebook_dir).parts
                if rel_parts:
                    section_key = "/".join(rel_parts) + "/" + section_name
                else:
                    section_key = section_name

                if section_key not in sections:
                    sections[section_key] = {"files": [], "latest": None}

                sections[section_key]["files"].append(one_file)

            # For each section, determine the latest (most recently modified) file
            for sec_info in sections.values():
                sec_info["files"].sort(key=lambda p: p.stat().st_mtime, reverse=True)
                sec_info["latest"] = sec_info["files"][0]

    return notebooks


def _parse_pages(filepath: Path) -> list[dict]:
    """
    Parse a .one file and extract pages with titles and text content.

    Returns a list of dicts: [{"title": "...", "texts": ["block1", ...]}, ...]
    Uses jcidPageMetaData (CachedTitleString) as page boundaries and
    jcidRichTextOENode (RichEditTextUnicode) for text content.
    Also extracts embedded images via jcidImageNode and runs OCR on them.
    """
    pages = []
    try:
        with open(filepath, "rb") as f:
            doc = OneDocment(f)

        props = doc.get_properties()
        current_page = None
        seen_titles: set[str] = set()

        # Build a lookup from identity string -> file info for image cross-referencing
        files = {}
        try:
            files = doc.get_files()
        except Exception as e:
            log.debug("Could not get files from %s: %s", filepath, e)
        file_by_identity: dict[str, dict] = {}
        for guid, finfo in files.items():
            identity = finfo.get("identity", "")
            if identity:
                file_by_identity[identity] = {**finfo, "guid": guid}

        # Collect image references per page, then OCR after loop
        image_refs_per_page: list[list[str]] = []  # parallel to pages
        current_image_refs: list[str] = []

        for prop in props:
            ptype = prop.get("type", "")
            val = prop.get("val", {})
            if not isinstance(val, dict):
                continue

            if ptype == "jcidPageMetaData":
                title = val.get("CachedTitleString", "").replace("\x00", "").strip()
                if title and title not in seen_titles:
                    seen_titles.add(title)
                    if current_page:
                        pages.append(current_page)
                        image_refs_per_page.append(current_image_refs)
                    current_page = {"title": title, "texts": []}
                    current_image_refs = []

            if current_page and ptype == "jcidRichTextOENode":
                text = val.get("RichEditTextUnicode", "")
                if text and isinstance(text, str) and text.strip():
                    current_page["texts"].append(text.strip())

            if current_page and ptype == "jcidImageNode":
                pic_ref = val.get("PictureContainer")
                if isinstance(pic_ref, list) and pic_ref:
                    current_image_refs.append(pic_ref[0])

        if current_page:
            pages.append(current_page)
            image_refs_per_page.append(current_image_refs)

        # OCR images and append text to respective pages
        if file_by_identity:
            _ocr_page_images(pages, image_refs_per_page, file_by_identity)

    except Exception as e:
        log.warning("Failed to parse pages from %s: %s", filepath, e)

    return pages


_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif"}


def _ocr_page_images(
    pages: list[dict],
    image_refs_per_page: list[list[str]],
    file_by_identity: dict[str, dict],
) -> None:
    """Run OCR on images referenced by each page and append text to page texts."""
    for page, refs in zip(pages, image_refs_per_page):
        for ref_str in refs:
            finfo = file_by_identity.get(ref_str)
            if not finfo:
                continue
            ext = finfo.get("extension", "")
            if not ext:
                continue
            if not ext.startswith("."):
                ext = "." + ext
            if ext.lower() not in _IMAGE_EXTENSIONS:
                continue
            content = finfo.get("content")
            if not content or not isinstance(content, bytes):
                continue
            try:
                text = ocr_image(content)
                if text and text.strip():
                    page["texts"].append(f"[OCR from image]: {text.strip()}")
                    log.debug("OCR extracted %d chars for page '%s'", len(text), page["title"])
            except Exception as e:
                log.debug("OCR failed for image in page '%s': %s", page["title"], e)


def _parse_one_file(filepath: Path) -> list[str]:
    """
    Parse a .one file and extract all text content.

    Returns a flat list of text strings found in the file.
    """
    pages = _parse_pages(filepath)
    texts = []
    for page in pages:
        texts.extend(page["texts"])
    return texts


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

mcp = FastMCP("onenote")

# Global search index (initialized at startup)
_search_index: EmbeddingIndex | None = None


@mcp.tool()
async def list_notebooks() -> str:
    """List all locally available OneNote notebooks.

    Shows notebook names and how many sections each one has.
    """
    notebooks = _discover_notebooks()
    if not notebooks:
        return f"No notebooks found in {ONENOTE_DIRS}"

    lines = []
    for name, info in sorted(notebooks.items()):
        section_count = len(info["sections"])
        lines.append(f"- {name}  ({section_count} sections)")
    return "\n".join(lines)


@mcp.tool()
async def list_sections(notebook_name: str) -> str:
    """List all sections in a specific notebook.

    Args:
        notebook_name: The name of the notebook (from list_notebooks).
    """
    notebooks = _discover_notebooks()
    if notebook_name not in notebooks:
        # Try case-insensitive match
        for key in notebooks:
            if key.lower() == notebook_name.lower():
                notebook_name = key
                break
        else:
            available = ", ".join(sorted(notebooks.keys()))
            return f"Notebook '{notebook_name}' not found. Available: {available}"

    sections = notebooks[notebook_name]["sections"]
    lines = []
    for sec_name, sec_info in sorted(sections.items()):
        latest = sec_info["latest"]
        size_kb = latest.stat().st_size / 1024
        lines.append(f"- {sec_name}  ({size_kb:.0f} KB)")
    return "\n".join(lines)


@mcp.tool()
async def read_section(notebook_name: str, section_name: str) -> str:
    """Read all text content from a specific section of a notebook.

    Shows content organized by page with titles.

    Args:
        notebook_name: The name of the notebook.
        section_name: The name of the section (from list_sections).
    """
    notebooks = _discover_notebooks()

    # Case-insensitive notebook match
    nb = None
    for key, val in notebooks.items():
        if key.lower() == notebook_name.lower():
            nb = val
            break
    if nb is None:
        available = ", ".join(sorted(notebooks.keys()))
        return f"Notebook '{notebook_name}' not found. Available: {available}"

    # Case-insensitive section match
    sec_info = None
    for key, val in nb["sections"].items():
        if key.lower() == section_name.lower():
            sec_info = val
            break
    if sec_info is None:
        available = ", ".join(sorted(nb["sections"].keys()))
        return f"Section '{section_name}' not found. Available: {available}"

    filepath = sec_info["latest"]
    pages = _parse_pages(filepath)

    if not pages:
        return f"No text content found in section '{section_name}'."

    lines = []
    for page in pages:
        lines.append(f"## {page['title']}")
        if page["texts"]:
            lines.append("\n\n".join(page["texts"]))
        else:
            lines.append("(no text content)")
        lines.append("")

    return "\n\n".join(lines)


@mcp.tool()
async def read_page(notebook_name: str, section_name: str, page_title: str) -> str:
    """Read a specific page by title from a notebook section.

    Args:
        notebook_name: The name of the notebook.
        section_name: The name of the section.
        page_title: The title of the page (from read_section output).
    """
    notebooks = _discover_notebooks()

    nb = None
    for key, val in notebooks.items():
        if key.lower() == notebook_name.lower():
            nb = val
            break
    if nb is None:
        available = ", ".join(sorted(notebooks.keys()))
        return f"Notebook '{notebook_name}' not found. Available: {available}"

    sec_info = None
    for key, val in nb["sections"].items():
        if key.lower() == section_name.lower():
            sec_info = val
            break
    if sec_info is None:
        available = ", ".join(sorted(nb["sections"].keys()))
        return f"Section '{section_name}' not found. Available: {available}"

    filepath = sec_info["latest"]
    pages = _parse_pages(filepath)

    for page in pages:
        if page["title"].lower() == page_title.lower():
            if not page["texts"]:
                return f"Page '{page['title']}' exists but has no text content."
            return f"# {page['title']}\n\n" + "\n\n".join(page["texts"])

    available_pages = [p["title"] for p in pages]
    return f"Page '{page_title}' not found. Available pages: {', '.join(available_pages)}"


@mcp.tool()
async def search_notes(query: str, exact_match: bool = False) -> str:
    """Search for text across ALL notebooks and sections.

    By default uses semantic search (finds conceptually related content).
    Set exact_match=True for literal substring matching.

    Args:
        query: The text to search for.
        exact_match: If True, use exact substring matching (case-insensitive).
                     If False (default), use semantic similarity search.
    """
    if not exact_match and _search_index is not None:
        matches = _search_index.search(query, top_k=20)
        if not matches:
            return f"No results found for '{query}'."
        lines = [f"Found {len(matches)} match(es) for '{query}' (semantic search):\n"]
        for m in matches:
            score = m["score"]
            nb = m["notebook"]
            sec = m["section"]
            title = m["page_title"]
            text = m["text"]
            snippet = text[:200] + "..." if len(text) > 200 else text
            lines.append(f'[{nb} / {sec} / "{title}"]  (score: {score:.2f})\n  {snippet}')
        return "\n\n".join(lines)

    # Exact match fallback
    query_lower = query.lower()
    notebooks = _discover_notebooks()
    results = []

    for nb_name, nb_info in sorted(notebooks.items()):
        for sec_name, sec_info in sorted(nb_info["sections"].items()):
            filepath = sec_info["latest"]
            pages = _parse_pages(filepath)

            for page in pages:
                for text in page["texts"]:
                    if query_lower in text.lower():
                        idx = text.lower().index(query_lower)
                        start = max(0, idx - 80)
                        end = min(len(text), idx + len(query) + 80)
                        snippet = text[start:end].strip()
                        if start > 0:
                            snippet = "..." + snippet
                        if end < len(text):
                            snippet = snippet + "..."

                        results.append(
                            f'[{nb_name} / {sec_name} / "{page["title"]}"]'
                            f"\n  {snippet}"
                        )

    if not results:
        return f"No results found for '{query}'."

    header = f"Found {len(results)} match(es) for '{query}' (exact match):\n\n"
    return header + "\n\n".join(results[:30])


@mcp.tool()
async def rebuild_search_index() -> str:
    """Rebuild the semantic search index from all notebooks.

    Use this if you've added new content and want it to be searchable,
    or if the index seems stale.
    """
    global _search_index
    notebooks = _discover_notebooks()
    if not notebooks:
        return "No notebooks found -- nothing to index."

    _search_index = EmbeddingIndex()
    count = _search_index.build(notebooks, _parse_pages)
    return f"Search index rebuilt: {count} pages indexed."


@mcp.tool()
async def list_all_sections() -> str:
    """List ALL sections across ALL notebooks.

    Useful for getting a complete overview of everything in your OneNote.
    """
    notebooks = _discover_notebooks()
    if not notebooks:
        return f"No notebooks found in {ONENOTE_DIRS}"

    lines = []
    for nb_name, nb_info in sorted(notebooks.items()):
        lines.append(f"\n## {nb_name}")
        for sec_name, sec_info in sorted(nb_info["sections"].items()):
            latest = sec_info["latest"]
            size_kb = latest.stat().st_size / 1024
            lines.append(f"  - {sec_name}  ({size_kb:.0f} KB)")

    return "\n".join(lines)


@mcp.tool()
async def get_notebook_summary(notebook_name: str) -> str:
    """Get a summary of a notebook: its sections and a preview of each section's content.

    Args:
        notebook_name: The name of the notebook.
    """
    notebooks = _discover_notebooks()

    nb = None
    for key, val in notebooks.items():
        if key.lower() == notebook_name.lower():
            nb = val
            notebook_name = key
            break
    if nb is None:
        available = ", ".join(sorted(notebooks.keys()))
        return f"Notebook '{notebook_name}' not found. Available: {available}"

    lines = [f"# {notebook_name}\n"]

    for sec_name, sec_info in sorted(nb["sections"].items()):
        filepath = sec_info["latest"]
        texts = _parse_one_file(filepath)

        lines.append(f"## {sec_name}")
        if texts:
            # Show first ~200 chars as preview
            preview = " | ".join(texts)
            if len(preview) > 200:
                preview = preview[:200] + "..."
            lines.append(f"  Preview: {preview}")
        else:
            lines.append("  (no text content)")
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# OneNote COM API helpers (for writing)
# ---------------------------------------------------------------------------

ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _sanitize_html_for_onenote(html: str) -> str:
    """
    Convert HTML to OneNote-compatible inline HTML.

    OneNote's <one:T> element only supports inline HTML (b, i, span, br, etc.).
    Block-level elements (h1-h6, p, ul, ol, li, div, table, etc.) cause
    UpdatePageContent to silently fail.

    Also escapes ]]> which would break the CDATA wrapper.
    """
    # Escape ]]> so it doesn't break CDATA sections
    html = html.replace("]]>", "]]&gt;")

    # Convert block-level closing tags to <br/>
    html = re.sub(r"</(?:p|div|h[1-6]|li|tr|blockquote|pre|code|section|article|header|footer|nav|aside|details|summary|figure|figcaption|dl|dt|dd)>", "<br/>", html, flags=re.IGNORECASE)

    # Remove block-level opening tags (keep their content)
    html = re.sub(r"<(?:p|div|h[1-6]|li|tr|td|th|blockquote|ul|ol|table|thead|tbody|pre|code|section|article|header|footer|nav|aside|details|summary|figure|figcaption|dl|dt|dd)(?:\s[^>]*)?>", "", html, flags=re.IGNORECASE)

    # Remove remaining closing tags for container elements
    html = re.sub(r"</(?:ul|ol|table|thead|tbody|td|th)>", "", html, flags=re.IGNORECASE)

    # Clean up multiple consecutive <br/> tags
    html = re.sub(r"(<br\s*/?>){3,}", "<br/><br/>", html, flags=re.IGNORECASE)

    # Normalize br tags
    html = re.sub(r"<br\s*/?>", "<br/>", html, flags=re.IGNORECASE)

    # Strip leading/trailing <br/>
    html = re.sub(r"^(<br/>)+", "", html)
    html = re.sub(r"(<br/>)+$", "", html)

    return html.strip()


def _run_powershell(script: str) -> tuple[bool, str]:
    """Run a PowerShell script and return (success, output)."""
    try:
        result = subprocess.run(
            ["powershell.exe", "-Command", script],
            capture_output=True, text=True, timeout=30,
        )
        output = result.stdout.strip()
        if result.returncode != 0:
            return False, result.stderr.strip() or output
        return True, output
    except subprocess.TimeoutExpired:
        return False, "PowerShell command timed out"
    except FileNotFoundError:
        return False, "PowerShell not found (write features require Windows)"


def _com_get_hierarchy(level: int = 3) -> ET.Element | None:
    """
    Get OneNote hierarchy via COM API.
    Levels: 0=Notebooks, 1=SectionGroups, 2=Sections, 3=Sections(full), 4=Pages
    """
    tmpfile = os.path.join(tempfile.gettempdir(), "onenote_hierarchy.xml")
    # Escape backslashes for PowerShell string
    tmpfile_ps = tmpfile.replace("\\", "\\\\")
    script = (
        f'$onenote = New-Object -ComObject OneNote.Application; '
        f'$h = ""; '
        f'$onenote.GetHierarchy("", {level}, [ref]$h); '
        f'$h | Out-File -FilePath "{tmpfile_ps}" -Encoding UTF8; '
        f'Write-Output "OK"'
    )
    ok, msg = _run_powershell(script)
    if not ok:
        log.warning("COM GetHierarchy failed: %s", msg)
        return None
    try:
        with open(tmpfile, "r", encoding="utf-8-sig") as f:
            xml_content = f.read()
        return ET.fromstring(xml_content)
    except Exception as e:
        log.warning("Failed to parse hierarchy XML: %s", e)
        return None
    finally:
        try:
            os.remove(tmpfile)
        except OSError:
            pass


def _com_find_section_id(notebook_name: str, section_name: str) -> str | None:
    """Find a section ID by notebook and section name (case-insensitive)."""
    root = _com_get_hierarchy(3)
    if root is None:
        return None

    for nb in root.iter(f"{{{ONE_NS}}}Notebook"):
        if nb.get("name", "").lower() != notebook_name.lower():
            continue
        for sec in nb.iter(f"{{{ONE_NS}}}Section"):
            if sec.get("isInRecycleBin") == "true":
                continue
            if sec.get("name", "").lower() == section_name.lower():
                return sec.get("ID")
    return None


def _run_powershell_file(script: str) -> tuple[bool, str]:
    """Write a PowerShell script to a temp file and execute it."""
    ps_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_cmd.ps1")
    try:
        with open(ps_file, "w", encoding="utf-8") as f:
            f.write(script)
        log.debug("Running PowerShell script (%d chars): %s", len(script), ps_file)
        result = subprocess.run(
            ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True, text=True, timeout=30,
        )
        output = result.stdout.strip()
        stderr = result.stderr.strip()
        log.debug("PowerShell exit=%d stdout=%s stderr=%s",
                  result.returncode, output[:500] if output else "(empty)",
                  stderr[:500] if stderr else "(empty)")
        if result.returncode != 0:
            return False, stderr or output
        return True, output
    except subprocess.TimeoutExpired:
        log.error("PowerShell timed out")
        return False, "PowerShell command timed out"
    except FileNotFoundError:
        log.error("PowerShell not found")
        return False, "PowerShell not found (write features require Windows)"
    finally:
        try:
            os.remove(ps_file)
        except OSError:
            pass


def _com_create_page(section_id: str, title: str, body_html: str) -> tuple[bool, str]:
    """Create a new page in a section using the OneNote COM API."""
    log.info("create_page: title=%r, body_len=%d, section=%s", title, len(body_html), section_id)
    log.debug("create_page: raw body=%r", body_html[:500])
    # Sanitize HTML to OneNote-compatible inline format
    body_html = _sanitize_html_for_onenote(body_html)
    log.debug("create_page: sanitized body=%r", body_html[:500])

    # Write title and body to temp files to avoid all escaping issues
    title_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_title.txt")
    body_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_body.txt")
    with open(title_file, "w", encoding="utf-8") as f:
        f.write(title)
    with open(body_file, "w", encoding="utf-8") as f:
        f.write(body_html)

    section_id_esc = section_id.replace("'", "''")

    script = f"""
$titleContent = Get-Content -Path '{title_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
$bodyContent = Get-Content -Path '{body_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
if ($titleContent) {{ $titleContent = $titleContent.Trim() }}
if ($bodyContent) {{ $bodyContent = $bodyContent.Trim() }}

$onenote = New-Object -ComObject OneNote.Application
$pageId = ""
$onenote.CreateNewPage('{section_id_esc}', [ref]$pageId, 0)

# Get the new page's XML
$pageXml = ""
$onenote.GetPageContent($pageId, [ref]$pageXml, 0)
$xml = [xml]$pageXml

# Set title
$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$titleNode = $xml.SelectSingleNode("//one:Title/one:OE/one:T", $nsMgr)
if ($titleNode) {{
    $titleNode.InnerXml = "<![CDATA[" + $titleContent + "]]>"
}}

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Title UpdatePageContent failed: $_"
    exit 1
}}

# Re-fetch to add body
$pageXml2 = ""
$onenote.GetPageContent($pageId, [ref]$pageXml2, 0)
$xml2 = [xml]$pageXml2

# Add body outline
$outline = $xml2.CreateElement("one", "Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeChildren = $xml2.CreateElement("one", "OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oe = $xml2.CreateElement("one", "OE", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$t = $xml2.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$cdata = $xml2.CreateCDataSection($bodyContent)
$t.AppendChild($cdata) | Out-Null
$oe.AppendChild($t) | Out-Null
$oeChildren.AppendChild($oe) | Out-Null
$outline.AppendChild($oeChildren) | Out-Null
$xml2.DocumentElement.AppendChild($outline) | Out-Null

try {{
    $onenote.UpdatePageContent($xml2.OuterXml)
}} catch {{
    Write-Error "Body UpdatePageContent failed: $_"
    exit 1
}}
Write-Output $pageId
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("create_page result: ok=%s output=%r", ok, output[:200] if output else "(empty)")
        if ok and output:
            return True, f"Page '{title}' created successfully (ID: {output})"
        return False, f"Failed to create page: {output}"
    finally:
        for f in (title_file, body_file):
            try:
                os.remove(f)
            except OSError:
                pass


def _com_append_to_page(page_id: str, body_html: str) -> tuple[bool, str]:
    """Append content to an existing page using the OneNote COM API."""
    log.info("append_to_page: page=%s, body_len=%d", page_id, len(body_html))
    log.debug("append_to_page: raw body=%r", body_html[:500])
    body_html = _sanitize_html_for_onenote(body_html)
    log.debug("append_to_page: sanitized body=%r", body_html[:500])

    # Write body to temp file to avoid escaping issues
    body_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_body.txt")
    with open(body_file, "w", encoding="utf-8") as f:
        f.write(body_html)

    page_id_esc = page_id.replace("'", "''")

    script = f"""
$bodyContent = Get-Content -Path '{body_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
if ($bodyContent) {{ $bodyContent = $bodyContent.Trim() }}

$onenote = New-Object -ComObject OneNote.Application
$pageXml = ""
$onenote.GetPageContent('{page_id_esc}', [ref]$pageXml, 0)
$xml = [xml]$pageXml

$outline = $xml.CreateElement("one", "Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeChildren = $xml.CreateElement("one", "OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oe = $xml.CreateElement("one", "OE", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$t = $xml.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$cdata = $xml.CreateCDataSection($bodyContent)
$t.AppendChild($cdata) | Out-Null
$oe.AppendChild($t) | Out-Null
$oeChildren.AppendChild($oe) | Out-Null
$outline.AppendChild($oeChildren) | Out-Null
$xml.DocumentElement.AppendChild($outline) | Out-Null

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Append UpdatePageContent failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("append_to_page result: ok=%s output=%r", ok, output[:200] if output else "(empty)")
        if ok:
            return True, "Content appended successfully."
        return False, f"Failed to append content: {output}"
    finally:
        try:
            os.remove(body_file)
        except OSError:
            pass


def _com_list_pages(section_id: str) -> list[dict]:
    """List pages in a section via COM API. Returns list of {id, name}."""
    root = _com_get_hierarchy(4)
    if root is None:
        return []

    pages = []
    for sec in root.iter(f"{{{ONE_NS}}}Section"):
        if sec.get("ID") == section_id:
            for page in sec.iter(f"{{{ONE_NS}}}Page"):
                if page.get("isInRecycleBin") == "true":
                    continue
                pages.append({
                    "id": page.get("ID", ""),
                    "name": page.get("name", "(untitled)"),
                })
            break
    return pages


# ---------------------------------------------------------------------------
# MCP Write Tools (Windows only -- requires OneNote COM API via PowerShell)
# ---------------------------------------------------------------------------

if sys.platform == "win32":

    @mcp.tool()
    async def list_live_notebooks() -> str:
        """List notebooks from the running OneNote app (live, not backup files).

        This uses the OneNote COM API and shows the notebooks currently open in
        the OneNote desktop app, including their sections. Use this to find
        where to create new pages.
        """
        root = _com_get_hierarchy(3)
        if root is None:
            return "Could not connect to OneNote. Make sure the OneNote desktop app is installed."

        lines = []
        for nb in root.findall(f"{{{ONE_NS}}}Notebook"):
            nb_name = nb.get("name", "?")
            lines.append(f"\n## {nb_name}")
            for sec in nb.iter(f"{{{ONE_NS}}}Section"):
                if sec.get("isInRecycleBin") == "true":
                    continue
                sec_name = sec.get("name", "?")
                locked = " (locked)" if sec.get("locked") == "true" else ""
                lines.append(f"  - {sec_name}{locked}")

        if not lines:
            return "No notebooks found in OneNote."
        return "\n".join(lines)

    @mcp.tool()
    async def create_page(notebook_name: str, section_name: str, title: str, content: str) -> str:
        """Create a new page in a OneNote notebook section.

        The content is written as HTML. You can use basic HTML tags like
        <b>, <i>, <br>, <ul>, <li>, <h1>-<h6>, etc.

        Requires the OneNote desktop app to be installed.

        Args:
            notebook_name: Name of the notebook (from list_live_notebooks).
            section_name: Name of the section within the notebook.
            title: Title for the new page.
            content: The page content (plain text or HTML).
        """
        section_id = _com_find_section_id(notebook_name, section_name)
        if section_id is None:
            return (
                f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
                f"Use list_live_notebooks to see available notebooks and sections."
            )

        ok, msg = _com_create_page(section_id, title, content)
        return msg

    @mcp.tool()
    async def list_live_pages(notebook_name: str, section_name: str) -> str:
        """List pages in a section from the running OneNote app.

        Use this to find page IDs for appending content to existing pages.

        Args:
            notebook_name: Name of the notebook.
            section_name: Name of the section.
        """
        section_id = _com_find_section_id(notebook_name, section_name)
        if section_id is None:
            return (
                f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
                f"Use list_live_notebooks to see available notebooks and sections."
            )

        pages = _com_list_pages(section_id)
        if not pages:
            return "No pages found in this section."

        lines = []
        for p in pages:
            lines.append(f"- {p['name']}  (id: {p['id']})")
        return "\n".join(lines)

    @mcp.tool()
    async def append_to_page(page_id: str, content: str) -> str:
        """Append content to an existing OneNote page.

        The content is added as a new outline block at the bottom of the page.
        Supports HTML formatting (<b>, <i>, <br>, <ul>, <li>, etc.).

        Args:
            page_id: The page ID (from list_live_pages).
            content: The content to append (plain text or HTML).
        """
        ok, msg = _com_append_to_page(page_id, content)
        return msg


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _build_search_index():
    """Build or update the semantic search index."""
    global _search_index
    notebooks = _discover_notebooks()
    if not notebooks:
        log.warning("No notebooks found, skipping search index build")
        return

    try:
        _search_index = EmbeddingIndex()
        count = _search_index.build(notebooks, _parse_pages)
        log.info("Search index ready: %d pages indexed", count)
    except Exception as e:
        log.warning("Failed to build search index: %s (semantic search disabled)", e)
        _search_index = None


def main():
    if not ONENOTE_DIRS:
        log.error(
            "No OneNote backup directories found. Checked: %s\n"
            "Set the ONENOTE_BACKUP_DIR environment variable to the correct path.",
            DEFAULT_BACKUP_DIRS,
        )
        sys.exit(1)

    log.info("Starting OneNote MCP server (local files)...")
    log.info("Reading from: %s", ONENOTE_DIRS)

    _build_search_index()

    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
