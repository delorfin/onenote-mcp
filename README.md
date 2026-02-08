# OneNote MCP Server

A fork of [mhzarem/onenote-mcp](https://github.com/mhzarem/onenote-mcp) with macOS support, semantic search, page-level access, and image OCR.

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that gives Claude access to your local Microsoft OneNote notebooks. It reads `.one` files directly from disk — no Azure registration, no API keys, no authentication required. On Windows, it can also write to OneNote via the COM API.

## Changes from upstream

- **macOS support** — auto-discovers OneNote backup directories inside the macOS app container (handles locale variants like "Backup" / "Sicherung")
- **Page-level access** — `read_page` tool to read a single page by title instead of an entire section
- **Semantic search** — uses [sentence-transformers](https://www.sbert.net/) (`paraphrase-multilingual-MiniLM-L12-v2`) to embed pages and find conceptually related content; index is persisted to `~/.cache/onenote-mcp/` and incrementally updated
- **Image OCR (macOS)** — extracts text from images embedded in OneNote pages using the macOS Vision framework; results are cached on disk
- **Content-hash dedup** — when OneNote rotates backup files, the embedding index reuses existing vectors for pages whose content hasn't changed, avoiding redundant re-embedding
- **Improved backup file matching** — handles more date suffix formats in backup filenames (dots, dashes)

## Tools

### Reading

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all locally available OneNote notebooks (from backup files) |
| `list_sections` | List all sections in a specific notebook |
| `list_all_sections` | List every section across every notebook |
| `get_notebook_summary` | Get a notebook overview with content previews |
| `read_section` | Read all pages in a section |
| `read_page` | Read a single page by title |
| `search_notes` | Semantic search across all pages (or exact match with `exact_match=True`) |
| `rebuild_search_index` | Rebuild the semantic search index manually |

### Writing (Windows only)

| Tool | Description |
|------|-------------|
| `list_live_notebooks` | List notebooks/sections from the running OneNote app |
| `create_page` | Create a new page in any notebook section |
| `list_live_pages` | List pages in a section (with IDs for appending) |
| `append_to_page` | Append content to an existing page |

Writing tools use the OneNote COM API and require the OneNote desktop app on Windows.

## Prerequisites

- Python 3.12+
- [uv](https://docs.astral.sh/uv/) (recommended) or pip
- Microsoft OneNote desktop app (with local backup files)

## Installation

```bash
git clone https://github.com/delorfin/onenote-mcp.git
cd onenote-mcp
uv sync
```

## Setup

### Claude Code

```bash
claude mcp add --transport stdio onenote -- uv --directory /path/to/onenote-mcp run server.py
```

Verify it's connected:

```bash
claude mcp list
```

### Claude Desktop

Add this to your `claude_desktop_config.json`:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "onenote": {
      "command": "uv",
      "args": [
        "--directory",
        "/absolute/path/to/onenote-mcp",
        "run",
        "server.py"
      ]
    }
  }
}
```

On Windows, use the full path to `uv.exe` and double backslashes:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "C:\\Users\\YOUR_USER\\.local\\bin\\uv.exe",
      "args": [
        "--directory",
        "C:\\path\\to\\onenote-mcp",
        "run",
        "server.py"
      ]
    }
  }
}
```

Restart Claude Desktop after saving.

## Where It Reads Files From

The server auto-detects the OneNote backup directory:

- **macOS**: `~/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support/Microsoft User Data/OneNote/15.0/Backup/`
- **Windows**: `C:\Users\<user>\AppData\Local\Microsoft\OneNote\16.0\Backup\`

To override, set the `ONENOTE_BACKUP_DIR` environment variable:

```bash
# Claude Code
claude mcp add --transport stdio --env ONENOTE_BACKUP_DIR=/path/to/notes onenote -- uv --directory /path/to/onenote-mcp run server.py

# Or export it
export ONENOTE_BACKUP_DIR=/path/to/your/onenote/files
```

## Usage Examples

Once connected, you can ask Claude:

- "List my OneNote notebooks"
- "Show me the sections in my Machine Learning notebook"
- "Read the Algorithm page from my CS notebook"
- "Search my notes for transformers"
- "Give me a summary of my Programming notebook"
- "Create a new page in My Notebook / Quick Notes titled 'Meeting Notes'"
- "Append today's summary to my existing page"

## How It Works

**Reading:**
1. Scans the OneNote backup directory for `.one` files
2. Organizes them by notebook and section (grouping backup versions together)
3. Uses [pyOneNote](https://github.com/delorfin/pyOneNote) to parse the binary `.one` format
4. Extracts page titles and text content, plus OCR text from embedded images (macOS)
5. Builds a semantic search index at startup and incrementally updates it before each search (cached to disk)

**Writing (Windows):**
1. Connects to the running OneNote desktop app via the COM API
2. Uses PowerShell subprocess calls to create pages and update content
3. Supports HTML formatting in page content

## Limitations

- **Reading**: Uses OneNote desktop backup files — not OneDrive-only notebooks without local backup
- **Writing**: Requires Windows with the OneNote desktop app installed
- **OCR**: macOS only (uses the Vision framework); images are skipped on other platforms

## License

MIT
