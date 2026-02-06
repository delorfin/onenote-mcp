# OneNote MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that gives Claude access to your local Microsoft OneNote notebooks. It reads `.one` files directly from disk and writes to OneNote via the COM API — no Azure registration, no API keys, no authentication required.

## What It Does

This server parses the OneNote backup files that the desktop app stores locally and exposes them as tools that Claude can use to browse, read, and write to your notes.

### Reading Tools

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all locally available OneNote notebooks (from backup files) |
| `list_sections` | List all sections in a specific notebook |
| `read_section` | Read the full text content of a section |
| `search_notes` | Search for text across all notebooks and sections |
| `list_all_sections` | List every section across every notebook |
| `get_notebook_summary` | Get a notebook overview with content previews |

### Writing Tools

| Tool | Description |
|------|-------------|
| `list_live_notebooks` | List notebooks/sections from the running OneNote app |
| `create_page` | Create a new page in any notebook section |
| `list_live_pages` | List pages in a section (with IDs for appending) |
| `append_to_page` | Append content to an existing page |

Writing tools use the OneNote COM API and require the OneNote desktop app to be running on Windows.

## Prerequisites

- Python 3.12+
- [uv](https://docs.astral.sh/uv/) (recommended) or pip
- Microsoft OneNote desktop app (with local backup files)

## Installation

```bash
git clone https://github.com/mhzarem/onenote-mcp.git
cd onenote-mcp
uv sync
```

Or with pip:

```bash
git clone https://github.com/mhzarem/onenote-mcp.git
cd onenote-mcp
pip install "mcp[cli]" pyOneNote
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

By default, the server reads from the OneNote desktop app's local backup directory:

```
C:\Users\<user>\AppData\Local\Microsoft\OneNote\16.0\Backup\
```

To use a different location, set the `ONENOTE_BACKUP_DIR` environment variable:

```bash
# Claude Code
claude mcp add --transport stdio --env ONENOTE_BACKUP_DIR=/path/to/notes onenote -- uv --directory /path/to/onenote-mcp run server.py

# Or export it
export ONENOTE_BACKUP_DIR=/path/to/your/onenote/files
```

## Usage Examples

Once connected, you can ask Claude:

**Reading:**
- "List my OneNote notebooks"
- "Show me the sections in my Machine Learning notebook"
- "Read my Algorithm notes"
- "Search my notes for transformers"
- "Give me a summary of my Programming notebook"

**Writing:**
- "Create a new page in My Notebook / Quick Notes titled 'Meeting Notes'"
- "Add my interview prep notes to OneNote"
- "Append today's summary to my existing page"

## How It Works

**Reading:**
1. Scans the OneNote backup directory for `.one` files
2. Organizes them by notebook and section (grouping backup versions together)
3. Uses [pyOneNote](https://github.com/DissectMalware/pyOneNote) to parse the binary `.one` format
4. Extracts `RichEditTextUnicode` text content from each section

**Writing:**
1. Connects to the running OneNote desktop app via the COM API
2. Uses PowerShell subprocess calls to create pages and update content
3. Supports HTML formatting in page content

## Limitations

- **Reading**: Uses OneNote desktop backup files — not OneDrive-only notebooks without local backup
- **Reading**: Extracts text content only; images and embedded files are not included
- **Writing**: Requires Windows with the OneNote desktop app installed
- **Writing**: The OneNote app must be installed (it doesn't need to be open — the COM API will start it)

## License

MIT
