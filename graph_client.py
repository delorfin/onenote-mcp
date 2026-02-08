"""
Microsoft Graph API client for OneNote operations.

Provides read and write access to OneNote notebooks, sections, and pages
via the Microsoft Graph REST API (v1.0).
"""

import logging
import re
from html import escape as html_escape

import requests

from graph_auth import get_access_token

log = logging.getLogger("onenote-mcp")

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# In-memory cache for notebook/section hierarchy (avoids repeated API calls)
_hierarchy_cache: dict | None = None


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------


def _get_headers() -> dict:
    """Get authorization headers for Graph API calls."""
    token = get_access_token()
    if not token:
        raise RuntimeError("No access token available. Run the 'authenticate' tool first.")
    return {"Authorization": f"Bearer {token}"}


def _graph_get(path: str, **kwargs) -> requests.Response:
    """Perform an authenticated GET request to Graph API."""
    url = f"{_GRAPH_BASE}{path}"
    log.debug("Graph GET %s", url)
    resp = requests.get(url, headers=_get_headers(), timeout=60, **kwargs)
    _handle_error(resp)
    return resp


def _graph_post(path: str, **kwargs) -> requests.Response:
    """Perform an authenticated POST request to Graph API."""
    url = f"{_GRAPH_BASE}{path}"
    log.debug("Graph POST %s", url)
    resp = requests.post(url, headers=_get_headers(), timeout=60, **kwargs)
    _handle_error(resp)
    return resp


def _graph_patch(path: str, **kwargs) -> requests.Response:
    """Perform an authenticated PATCH request to Graph API."""
    url = f"{_GRAPH_BASE}{path}"
    log.debug("Graph PATCH %s", url)
    resp = requests.patch(url, headers=_get_headers(), timeout=60, **kwargs)
    _handle_error(resp)
    return resp


def _handle_error(resp: requests.Response) -> None:
    """Raise descriptive errors for common Graph API failures."""
    if resp.ok:
        return
    status = resp.status_code
    try:
        body = resp.json()
        msg = body.get("error", {}).get("message", resp.text[:300])
    except Exception:
        msg = resp.text[:300]

    if status == 401:
        raise RuntimeError(f"Authentication expired or invalid. Run the 'authenticate' tool to re-authenticate. ({msg})")
    if status == 429:
        raise RuntimeError(f"Rate limited by Microsoft Graph API. Please wait a moment and try again. ({msg})")
    if status == 404:
        raise RuntimeError(f"Resource not found: {msg}")
    raise RuntimeError(f"Graph API error {status}: {msg}")


# ---------------------------------------------------------------------------
# Hierarchy cache
# ---------------------------------------------------------------------------


def invalidate_cache() -> None:
    """Clear the in-memory hierarchy cache."""
    global _hierarchy_cache
    _hierarchy_cache = None


def _ensure_hierarchy() -> dict:
    """Fetch and cache the notebook → section hierarchy.

    Returns:
        {
            "notebooks": [
                {
                    "id": "...",
                    "displayName": "...",
                    "sections": [
                        {"id": "...", "displayName": "...", ...},
                        ...
                    ]
                },
                ...
            ]
        }
    """
    global _hierarchy_cache
    if _hierarchy_cache is not None:
        return _hierarchy_cache

    resp = _graph_get("/me/onenote/notebooks?$expand=sections($select=id,displayName)&$select=id,displayName")
    data = resp.json()
    notebooks = []
    for nb in data.get("value", []):
        notebooks.append({
            "id": nb["id"],
            "displayName": nb.get("displayName", ""),
            "sections": nb.get("sections", []),
        })
    _hierarchy_cache = {"notebooks": notebooks}
    log.debug("Hierarchy cache loaded: %d notebooks", len(notebooks))
    return _hierarchy_cache


def _find_notebook(notebook_name: str) -> dict | None:
    """Find a notebook by name (case-insensitive)."""
    hierarchy = _ensure_hierarchy()
    for nb in hierarchy["notebooks"]:
        if nb["displayName"].lower() == notebook_name.lower():
            return nb
    return None


def _find_section(notebook_name: str, section_name: str) -> dict | None:
    """Find a section by notebook and section name (case-insensitive)."""
    nb = _find_notebook(notebook_name)
    if nb is None:
        return None
    for sec in nb.get("sections", []):
        if sec["displayName"].lower() == section_name.lower():
            return sec
    return None


# ---------------------------------------------------------------------------
# Read operations
# ---------------------------------------------------------------------------


def list_notebooks_graph() -> list[dict]:
    """List all notebooks via Graph API.

    Returns list of {"id": ..., "displayName": ...}.
    """
    hierarchy = _ensure_hierarchy()
    return [
        {"id": nb["id"], "displayName": nb["displayName"],
         "section_count": len(nb.get("sections", []))}
        for nb in hierarchy["notebooks"]
    ]


def list_sections_graph(notebook_name: str) -> list[dict] | None:
    """List sections in a notebook by name.

    Returns list of {"id": ..., "displayName": ...} or None if notebook not found.
    """
    nb = _find_notebook(notebook_name)
    if nb is None:
        return None
    return [
        {"id": sec["id"], "displayName": sec.get("displayName", "")}
        for sec in nb.get("sections", [])
    ]


def list_pages_graph(notebook_name: str, section_name: str) -> list[dict] | None:
    """List pages in a section by notebook and section name.

    Returns list of {"id": ..., "title": ...} or None if section not found.
    """
    sec = _find_section(notebook_name, section_name)
    if sec is None:
        return None

    resp = _graph_get(f"/me/onenote/sections/{sec['id']}/pages?$select=id,title,createdDateTime,lastModifiedDateTime&$orderby=createdDateTime")
    data = resp.json()
    return [
        {
            "id": p["id"],
            "title": p.get("title", "(untitled)"),
            "created": p.get("createdDateTime", ""),
            "modified": p.get("lastModifiedDateTime", ""),
        }
        for p in data.get("value", [])
    ]


def get_page_content_graph(page_id: str) -> str:
    """Get page content as HTML.

    Returns the raw HTML body of the page.
    """
    resp = _graph_get(f"/me/onenote/pages/{page_id}/content")
    return resp.text


def _extract_text_from_html(html: str) -> str:
    """Extract readable plain text from HTML (simple regex-based approach)."""
    # Remove script/style blocks
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html, flags=re.DOTALL | re.IGNORECASE)
    # Replace block-level closing tags with newlines
    text = re.sub(r"</(?:p|div|h[1-6]|li|tr|br|hr)[^>]*>", "\n", text, flags=re.IGNORECASE)
    # Replace <br> tags
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    # Strip all remaining tags
    text = re.sub(r"<[^>]+>", "", text)
    # Decode common entities
    text = text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    text = text.replace("&nbsp;", " ").replace("&quot;", '"')
    # Collapse whitespace
    lines = [line.strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


def read_page_graph(notebook_name: str, section_name: str, page_title: str) -> str | None:
    """Read a specific page by title, returning plain text content.

    Returns the page text or None if not found.
    """
    pages = list_pages_graph(notebook_name, section_name)
    if pages is None:
        return None

    for page in pages:
        if page["title"].lower() == page_title.lower():
            html = get_page_content_graph(page["id"])
            return _extract_text_from_html(html)
    return None


def read_section_graph(notebook_name: str, section_name: str) -> list[dict] | None:
    """Read all pages in a section, returning list of {"title": ..., "text": ...}.

    Returns None if section not found.
    """
    pages = list_pages_graph(notebook_name, section_name)
    if pages is None:
        return None

    result = []
    for page in pages:
        html = get_page_content_graph(page["id"])
        text = _extract_text_from_html(html)
        result.append({"title": page["title"], "text": text})
    return result


def search_pages_graph(query: str) -> list[dict]:
    """Search pages by title via Graph API.

    Uses OData filter on title. Returns list of matching pages.
    """
    # Escape single quotes in the query for OData
    safe_query = query.replace("'", "''")
    path = f"/me/onenote/pages?$filter=contains(tolower(title), '{safe_query.lower()}')&$select=id,title,createdDateTime,lastModifiedDateTime,parentSection&$expand=parentSection($select=displayName)&$orderby=lastModifiedDateTime desc&$top=20"
    resp = _graph_get(path)
    data = resp.json()
    results = []
    for p in data.get("value", []):
        section_name = ""
        if p.get("parentSection"):
            section_name = p["parentSection"].get("displayName", "")
        results.append({
            "id": p["id"],
            "title": p.get("title", "(untitled)"),
            "section": section_name,
            "modified": p.get("lastModifiedDateTime", ""),
        })
    return results


# ---------------------------------------------------------------------------
# Write operations
# ---------------------------------------------------------------------------


def _format_graph_html_document(title: str, body_html: str) -> str:
    """Wrap content in a full HTML document as required by Graph API for page creation."""
    return (
        "<!DOCTYPE html>\n"
        "<html>\n"
        "<head>\n"
        f"  <title>{html_escape(title)}</title>\n"
        "</head>\n"
        "<body>\n"
        f"  {body_html}\n"
        "</body>\n"
        "</html>"
    )


def create_page_graph(notebook_name: str, section_name: str, title: str, html_content: str) -> tuple[bool, str]:
    """Create a new page in a section via Graph API.

    Args:
        notebook_name: Name of the notebook.
        section_name: Name of the section.
        title: Page title.
        html_content: Body HTML content.

    Returns:
        (success, message) tuple.
    """
    sec = _find_section(notebook_name, section_name)
    if sec is None:
        nb = _find_notebook(notebook_name)
        if nb is None:
            available = [n["displayName"] for n in _ensure_hierarchy()["notebooks"]]
            return False, f"Notebook '{notebook_name}' not found. Available: {', '.join(available)}"
        available = [s["displayName"] for s in nb.get("sections", [])]
        return False, f"Section '{section_name}' not found in '{notebook_name}'. Available: {', '.join(available)}"

    # If content looks like plain text, wrap in <p> tags
    if not re.search(r"<[a-zA-Z]", html_content):
        html_content = "<p>" + html_escape(html_content).replace("\n", "</p>\n<p>") + "</p>"

    page_html = _format_graph_html_document(title, html_content)

    headers = _get_headers()
    headers["Content-Type"] = "application/xhtml+xml"
    url = f"{_GRAPH_BASE}/me/onenote/sections/{sec['id']}/pages"
    log.debug("Creating page in section %s: title=%r", sec["id"], title)

    resp = requests.post(url, headers=headers, data=page_html.encode("utf-8"), timeout=60)
    _handle_error(resp)

    data = resp.json()
    page_id = data.get("id", "unknown")
    return True, f"Page '{title}' created successfully (ID: {page_id})"


def append_to_page_graph(page_id: str, html_content: str) -> tuple[bool, str]:
    """Append content to an existing page via Graph API.

    Args:
        page_id: The page ID.
        html_content: HTML content to append.

    Returns:
        (success, message) tuple.
    """
    # If content looks like plain text, wrap in tags
    if not re.search(r"<[a-zA-Z]", html_content):
        html_content = "<p>" + html_escape(html_content).replace("\n", "</p>\n<p>") + "</p>"

    patch_body = [
        {
            "target": "body",
            "action": "append",
            "content": f"<div>{html_content}</div>",
        }
    ]

    headers = _get_headers()
    headers["Content-Type"] = "application/json"
    url = f"{_GRAPH_BASE}/me/onenote/pages/{page_id}/content"
    log.debug("Appending to page %s", page_id)

    resp = requests.patch(url, headers=headers, json=patch_body, timeout=60)
    _handle_error(resp)

    return True, "Content appended successfully."


def update_page_content_graph(page_id: str, html_content: str) -> tuple[bool, str]:
    """Replace the body content of an existing page via Graph API.

    Args:
        page_id: The page ID.
        html_content: New HTML body content.

    Returns:
        (success, message) tuple.
    """
    # If content looks like plain text, wrap in tags
    if not re.search(r"<[a-zA-Z]", html_content):
        html_content = "<p>" + html_escape(html_content).replace("\n", "</p>\n<p>") + "</p>"

    patch_body = [
        {
            "target": "body",
            "action": "replace",
            "content": f"<div>{html_content}</div>",
        }
    ]

    headers = _get_headers()
    headers["Content-Type"] = "application/json"
    url = f"{_GRAPH_BASE}/me/onenote/pages/{page_id}/content"
    log.debug("Replacing content of page %s", page_id)

    resp = requests.patch(url, headers=headers, json=patch_body, timeout=60)
    _handle_error(resp)

    return True, "Page content updated successfully."


def prepend_to_page_graph(page_id: str, html_content: str) -> tuple[bool, str]:
    """Prepend content to the top of an existing page via Graph API.

    Args:
        page_id: The page ID.
        html_content: HTML content to prepend.

    Returns:
        (success, message) tuple.
    """
    if not re.search(r"<[a-zA-Z]", html_content):
        html_content = "<p>" + html_escape(html_content).replace("\n", "</p>\n<p>") + "</p>"

    patch_body = [
        {
            "target": "body",
            "action": "prepend",
            "content": f"<div>{html_content}</div>",
        }
    ]

    headers = _get_headers()
    headers["Content-Type"] = "application/json"
    url = f"{_GRAPH_BASE}/me/onenote/pages/{page_id}/content"
    log.debug("Prepending to page %s", page_id)

    resp = requests.patch(url, headers=headers, json=patch_body, timeout=60)
    _handle_error(resp)

    return True, "Content prepended successfully."


def replace_text_in_page_graph(page_id: str, find_text: str, replace_text: str, case_sensitive: bool = False) -> tuple[bool, str]:
    """Find and replace text in a page via Graph API.

    Fetches the page HTML, performs the replacement, and patches back.

    Args:
        page_id: The page ID.
        find_text: Text to find.
        replace_text: Text to replace with.
        case_sensitive: Whether the search is case-sensitive.

    Returns:
        (success, message) tuple.
    """
    # Fetch current content
    html = get_page_content_graph(page_id)

    # Count matches
    flags = 0 if case_sensitive else re.IGNORECASE
    escaped = re.escape(find_text)
    matches = re.findall(escaped, html, flags=flags)
    if not matches:
        return False, f"No matches found for '{find_text}' in this page."

    # Perform replacement
    updated_html = re.sub(escaped, replace_text, html, flags=flags)

    # Extract just the body content (between <body> tags)
    body_match = re.search(r"<body[^>]*>(.*)</body>", updated_html, flags=re.DOTALL | re.IGNORECASE)
    body_content = body_match.group(1).strip() if body_match else updated_html

    patch_body = [
        {
            "target": "body",
            "action": "replace",
            "content": body_content,
        }
    ]

    headers = _get_headers()
    headers["Content-Type"] = "application/json"
    url = f"{_GRAPH_BASE}/me/onenote/pages/{page_id}/content"
    log.debug("Replacing text in page %s: %d occurrences", page_id, len(matches))

    resp = requests.patch(url, headers=headers, json=patch_body, timeout=60)
    _handle_error(resp)

    return True, f"Replaced {len(matches)} occurrence(s) of '{find_text}' with '{replace_text}'."
