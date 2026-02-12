# Google Workspace MCP Server

An MCP server that provides tools for interacting with Google Docs, Sheets, and Drive.

## Setup

### 1. Install dependencies

```bash
cd /Users/goncaloreis/Projects/google-workspace-mcp
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. First-time authentication

Run the server once to authenticate:

```bash
python3 server.py
```

This will open a browser window for Google OAuth. Sign in and grant permissions.
A `token.json` file will be created to store your credentials.

### 3. Configure Claude Code

Add to your Claude Code MCP settings (`~/.config/claude/claude_desktop_config.json` or similar):

```json
{
  "mcpServers": {
    "google-workspace": {
      "command": "/Users/goncaloreis/Projects/google-workspace-mcp/venv/bin/python",
      "args": ["/Users/goncaloreis/Projects/google-workspace-mcp/server.py"]
    }
  }
}
```

## Available Tools

### Google Docs
- `google_docs_create` - Create a new Google Doc
- `google_docs_read` - Read content from a Google Doc
- `google_docs_append` - Append content to a Google Doc

### Google Sheets
- `google_sheets_read` - Read data from a Google Sheet
- `google_sheets_write` - Write data to a Google Sheet
- `google_sheets_append` - Append rows to a Google Sheet

### Google Drive
- `google_drive_create_folder` - Create a new folder
- `google_drive_move` - Move a file to a different folder
- `google_drive_list` - List files in a folder

## Usage Examples

### Create a Google Doc
```
Tool: google_docs_create
Arguments: {"title": "Meeting Notes", "content": "# Meeting Notes\n\nDate: 2026-01-22"}
```

### Read a Google Sheet
```
Tool: google_sheets_read
Arguments: {"spreadsheet_id": "1abc123...", "range": "Sheet1!A1:D10"}
```

### Create a folder
```
Tool: google_drive_create_folder
Arguments: {"name": "New Project", "parent_id": "folder_id_here"}
```
