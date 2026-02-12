# Google Workspace MCP Server

This server gives Claude direct access to your Google Workspace — Gmail, Google Drive, Sheets, Docs, Slides, Calendar, and Tasks. Once installed, you can ask Claude things like "search my email for messages from João", "create a spreadsheet with this data", or "what's on my calendar this week" and it will do it directly.

**98 tools** across 7 Google services. Built by Gonçalo.

---

## What you'll need before starting

- A Mac (these instructions are for macOS)
- Python 3.10 or higher (check with `python3 --version` in Terminal)
- Claude Desktop app installed ([download here](https://claude.ai/download) if you don't have it)
- The `credentials.json` file that Gonçalo sent you (should be attached to the email)
- About 15 minutes

---

## Installation — step by step

### Step 1: Open Terminal

Press `Cmd + Space`, type "Terminal", and hit Enter. A black/white window will appear. This is where you'll type all the commands below.

### Step 2: Check that Python is installed

Type this and press Enter:

```bash
python3 --version
```

You should see something like `Python 3.11.4`. If you get "command not found", you need to install Python first:
- Go to https://www.python.org/downloads/
- Download the latest version for macOS
- Run the installer
- Close and reopen Terminal, then try `python3 --version` again

### Step 3: Create a folder for MCP servers

```bash
mkdir -p ~/Projects
```

This creates a `Projects` folder inside your home directory (if it doesn't exist already).

### Step 4: Download the code

```bash
cd ~/Projects
git clone https://github.com/goncaloreis/google-workspace-mcp.git
cd google-workspace-mcp
```

If you get "git: command not found", run `xcode-select --install` first, then try again.

### Step 5: Set up the Python environment

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

You should see a bunch of packages being installed. Wait for it to finish.

### Step 6: Add the credentials file

Take the `credentials.json` file that Gonçalo sent you and put it in the project folder. You can do this by dragging it into Finder at `~/Projects/google-workspace-mcp/`, or with Terminal:

```bash
cp ~/Downloads/credentials.json ~/Projects/google-workspace-mcp/credentials.json
```

(Adjust the path if you saved it somewhere other than Downloads.)

### Step 7: Authenticate with your Google account

```bash
python3 server.py
```

This will open your web browser and ask you to sign in with your Google account. **Sign in with your Stake Capital Google account** (your @stake.capital email).

You'll see a warning saying "Google hasn't verified this app" — this is normal for internal tools. Click **"Advanced"** → **"Go to [app name] (unsafe)"** → then **"Allow"** on each permission screen.

Once you've granted permissions, go back to Terminal. You should see the server running. Press `Ctrl + C` to stop it.

A `token.json` file has been created — this stores your personal access. **Never share this file.**

### Step 8: Connect it to Claude Desktop

Open this file in a text editor:

```bash
open -a TextEdit ~/Library/Application\ Support/Claude/claude_desktop_config.json
```

If the file doesn't exist yet, create it:

```bash
mkdir -p ~/Library/Application\ Support/Claude
echo '{}' > ~/Library/Application\ Support/Claude/claude_desktop_config.json
open -a TextEdit ~/Library/Application\ Support/Claude/claude_desktop_config.json
```

Replace the contents with this (or add to the existing `mcpServers` section):

```json
{
  "mcpServers": {
    "google-workspace": {
      "command": "/bin/bash",
      "args": [
        "-c",
        "cd $HOME/Projects/google-workspace-mcp && source venv/bin/activate && python server.py"
      ]
    }
  }
}
```

Save and close the file.

### Step 9: Restart Claude Desktop

Quit Claude Desktop completely (`Cmd + Q`) and reopen it.

### Step 10: Verify it works

In Claude Desktop, you should see a small hammer/tools icon (🔨) at the bottom of the chat input. Click it and you should see tools like `google_docs_create`, `gmail_search`, `google_sheets_read`, etc.

Try asking Claude: **"Search my email for the most recent message from Gonçalo"**

If it works, you're done! 🎉

---

## Troubleshooting

**"I don't see the tools icon in Claude"**
- Make sure you saved the config file correctly (check for typos, especially in the path)
- Make sure you fully quit and restarted Claude Desktop
- Check the file at `~/Library/Application Support/Claude/claude_desktop_config.json` is valid JSON

**"Google sign-in failed"**
- Make sure you're signing in with your @stake.capital Google account
- If you see "access blocked", contact Gonçalo — he may need to add you as a test user in the Google Cloud project

**"Python not found" or "venv not found"**
- Make sure Python 3.10+ is installed
- Try using the full path: replace `python` with `python3` everywhere

**"Permission denied"**
- Run: `chmod +x ~/Projects/google-workspace-mcp/server.py`

---

## What can it do?

Once installed, you can ask Claude things like:

- "Search my email for messages about the Stellar project"
- "Create a Google Doc called Meeting Notes with today's date"
- "Read the data from this spreadsheet: [paste Google Sheets URL]"
- "What's on my calendar this week?"
- "Create a new folder in Google Drive called Q1 Reports"
- "Draft an email to [person] about [topic]"
- "List my Google Tasks"
- "Add a slide to my presentation about [topic]"

Claude will use the tools automatically — you just ask in plain language.
