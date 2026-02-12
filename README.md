# Google Workspace MCP Server

This server gives Claude direct access to your Google Workspace — Gmail, Google Drive, Sheets, Docs, Slides, Calendar, and Tasks. Once installed, you can ask Claude things like "search my email for messages from João", "create a spreadsheet with this data", or "what's on my calendar this week" and it will do it directly.

**98 tools** across 7 Google services. Built by Gonçalo.

---

## What you'll need before starting

- Python 3.10 or higher
- Claude Desktop app installed ([download here](https://claude.ai/download))
- Claude Code installed (see below)
- The `credentials.json` file that Gonçalo sent you (attached to the email)
- About 15 minutes

---

## Installation

### Step 1: Install Claude Code

If you don't have Claude Code yet, open your terminal (Terminal on Mac, Command Prompt or PowerShell on Windows) and run:

```bash
npm install -g @anthropic-ai/claude-code
```

If you don't have `npm`, install Node.js first from https://nodejs.org (download the LTS version, run the installer, then try the command above again).

### Step 2: Save the credentials.json file

Download the `credentials.json` file from the email attachment and save it to your Downloads folder. Remember where it is — you'll tell Claude Code about it in the next step.

### Step 3: Let Claude Code do the rest

Open your terminal and run:

```bash
claude
```

Then paste this prompt:

---

**Prompt to paste into Claude Code:**

```
I need you to install a Google Workspace MCP server on my machine. Here's what to do:

1. Check that Python 3.10+ is installed (run python3 --version on Mac/Linux, or python --version on Windows)

2. Create a Projects folder in my home directory if it doesn't exist

3. Clone the repo:
   cd ~/Projects (or %USERPROFILE%\Projects on Windows)
   git clone https://github.com/goncaloreis/google-workspace-mcp.git
   cd google-workspace-mcp

4. Set up Python virtual environment:
   python3 -m venv venv (or python -m venv venv on Windows)
   Activate it and install requirements: pip install -r requirements.txt

5. Copy credentials.json from my Downloads folder into the google-workspace-mcp folder

6. Run "python server.py" once to trigger Google OAuth — this will open my browser. I'll sign in with my Google account and grant permissions. Wait for me to confirm this is done.

7. After I confirm OAuth is complete, configure Claude Desktop by editing the config file:
   - Mac: ~/Library/Application Support/Claude/claude_desktop_config.json
   - Windows: %APPDATA%\Claude\claude_desktop_config.json
   
   Add this to the mcpServers section (adapt paths for my OS):
   
   On Mac:
   {
     "mcpServers": {
       "google-workspace": {
         "command": "/bin/bash",
         "args": ["-c", "cd $HOME/Projects/google-workspace-mcp && source venv/bin/activate && python server.py"]
       }
     }
   }
   
   On Windows:
   {
     "mcpServers": {
       "google-workspace": {
         "command": "cmd",
         "args": ["/c", "cd /d %USERPROFILE%\\Projects\\google-workspace-mcp && venv\\Scripts\\activate && python server.py"]
       }
     }
   }

   If the file already has other MCP servers, merge this into the existing mcpServers object.

8. Tell me to restart Claude Desktop (Cmd+Q on Mac, or close fully on Windows) and reopen it.

Start by checking my OS and Python version, then proceed step by step. Ask me if anything is unclear.
```

---

### Step 4: Verify it works

After restarting Claude Desktop, you should see a tools icon (🔨) at the bottom of the chat input. Click it and you should see tools like `google_docs_create`, `gmail_search`, `google_sheets_read`, etc.

Try asking Claude: **"Search my email for the most recent message from Gonçalo"**

If it works, you're done! 🎉

---

## Troubleshooting

**"Google hasn't verified this app" warning during sign-in**
This is normal for internal tools. Click **"Advanced"** → **"Go to [app name] (unsafe)"** → **"Allow"**.

**"access_denied" or "blocked" during Google sign-in**
Contact Gonçalo — he may need to add you as a test user in the Google Cloud project.

**Tools don't show up in Claude Desktop**
- Check the config file for typos
- Make sure you fully quit and restarted Claude Desktop
- Ask Claude Code to check that the config file is valid JSON

**Any other error**
Paste the error message to Gonçalo on Slack.

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

---

## Security notes

- `credentials.json` identifies the app, not any person's account. It's shared and safe to have.
- `token.json` is YOUR personal access token, created when you sign in. **Never share this file.**
- The `.gitignore` is configured to exclude both `token.json` and any other sensitive files from git.
