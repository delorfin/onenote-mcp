# OneNote MCP Server

A fork of [mhzarem/onenote-mcp](https://github.com/mhzarem/onenote-mcp) with macOS support, cross-platform writing, semantic search, page-level access, and image OCR.

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that gives Claude access to your Microsoft OneNote notebooks. It supports two data sources:

- **Local backup files** (default) — fast, offline, no authentication needed
- **Microsoft Graph API** — live/current data, write support on all platforms, requires one-time device code authentication

## Changes from upstream

- **Cross-platform write support** — create pages and append content on macOS/Linux/Windows via Microsoft Graph API (one-time device code auth)
- **Dual data source** — read from local backup files (default, fast, offline) or live Graph API (`use_api=True`); switch default with `set_data_source`
- **macOS support** — auto-discovers OneNote backup directories inside the macOS app container (handles locale variants like "Backup" / "Sicherung")
- **Page-level access** — `read_page` tool to read a single page by title instead of an entire section
- **Semantic search** — uses [sentence-transformers](https://www.sbert.net/) (`paraphrase-multilingual-MiniLM-L12-v2`) to embed pages and find conceptually related content; index is persisted to `~/.cache/onenote-mcp/` and incrementally updated
- **Image OCR (macOS)** — extracts text from images embedded in OneNote pages using the macOS Vision framework; results are cached on disk
- **Content-hash dedup** — when OneNote rotates backup files, the embedding index reuses existing vectors for pages whose content hasn't changed, avoiding redundant re-embedding
- **Improved backup file matching** — handles more date suffix formats in backup filenames (dots, dashes)

## Tools

### Reading

All read tools accept an optional `use_api` parameter. When `use_api=True`, data is fetched live from Microsoft Graph API instead of local backup files.

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all available OneNote notebooks |
| `list_sections` | List all sections in a specific notebook |
| `list_all_sections` | List every section across every notebook |
| `get_notebook_summary` | Get a notebook overview with content previews |
| `read_section` | Read all pages in a section |
| `read_page` | Read a single page by title |
| `list_pages` | List pages in a section (with IDs for writing) |
| `search_notes` | Semantic search across all pages (or exact match with `exact_match=True`) |
| `rebuild_search_index` | Rebuild the semantic search index manually |

### Writing

Write tools use Graph API by default (all platforms). On Windows, set `use_com=True` to use the COM API instead.

| Tool | Description |
|------|-------------|
| `create_page` | Create a new page in any notebook section |
| `append_to_page` | Append content to an existing page |

### Authentication & Configuration

| Tool | Description |
|------|-------------|
| `authenticate` | Start device code flow for Graph API access |
| `set_data_source` | Set default source: `'local'` or `'api'` |
| `clear_auth` | Remove cached auth tokens (for re-authentication) |

## Prerequisites

- Python 3.12+
- [uv](https://docs.astral.sh/uv/) (recommended) or pip
- Microsoft OneNote desktop app (for local backup files)
- Microsoft account (for Graph API write/live features — free, no Azure app registration)

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

## Authentication Setup

Authentication is only needed for **write operations** and **live API reads**. Local reads from backup files work without any authentication.

1. Ask Claude to "authenticate with OneNote" (or call the `authenticate` tool)
2. Open the displayed URL in a browser
3. Enter the device code and sign in with your Microsoft account
4. Authentication is cached in `~/.cache/onenote-mcp/graph-token.json` and auto-refreshes

To use a custom Azure app registration, set `AZURE_CLIENT_ID`:

```bash
claude mcp add --transport stdio --env AZURE_CLIENT_ID=your-app-id onenote -- uv --directory /path/to/onenote-mcp run server.py
```

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
- "Authenticate with OneNote so I can create pages"
- "Create a new page in My Notebook / Quick Notes titled 'Meeting Notes'"
- "List pages in my Quick Notes section and append to the latest one"
- "Show me live notebooks from the API" (use_api=True)

## How It Works

**Reading (local, default):**
1. Scans the OneNote backup directory for `.one` files
2. Organizes them by notebook and section (grouping backup versions together)
3. Uses [pyOneNote](https://github.com/delorfin/pyOneNote) to parse the binary `.one` format
4. Extracts page titles and text content, plus OCR text from embedded images (macOS)
5. Builds a semantic search index at startup and incrementally updates it before each search (cached to disk)

**Reading (API):**
1. Authenticates via MSAL device code flow (one-time)
2. Queries Microsoft Graph REST API for notebooks, sections, and pages
3. Fetches page content as HTML and extracts text

**Writing (Graph API, all platforms):**
1. Authenticates via MSAL device code flow (one-time)
2. Creates pages via POST to Graph API with XHTML content
3. Appends to pages via PATCH with JSON content operations

**Writing (COM API, Windows only):**
1. Connects to the running OneNote desktop app via the COM API
2. Uses PowerShell subprocess calls to create pages and update content
3. Supports HTML formatting in page content

## Limitations

- **Local reading**: Uses OneNote desktop backup files — not OneDrive-only notebooks without local backup
- **Writing**: Requires Microsoft account authentication (one-time device code flow)
- **API search**: Title-matching only (local semantic search is more powerful for content search)
- **OCR**: macOS only (uses the Vision framework); images are skipped on other platforms

## License

MIT
