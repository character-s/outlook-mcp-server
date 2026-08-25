# 🤖 Outlook MCP Server

**AI-powered email management for Microsoft Outlook** - Search, compose, organize, and batch forward emails with natural language commands.

<div align="center">

⭐ **This saved you time? [Star us](https://github.com/marlonluo2018/outlook-mcp-server/stargazers) - takes 2 seconds, helps thousands of Outlook users find this AI email assistant!** ⭐

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6)](https://www.microsoft.com/windows)

</div>

## 🤔 Project Choice: Local vs Cloud Approach

**This project uses win32COM which works entirely locally with your Outlook data (more secure), while our new [Microsoft Graph MCP Server](https://github.com/marlonluo2018/microsoft_graph_mcp_server) provides online access to the full Microsoft 365 ecosystem (more features). Choose based on your security and functionality needs.**

### Why Choose This Project (win32COM)?
- **✅ Local Data Processing**: All email operations happen locally on your computer - no data leaves your machine
- **✅ Enhanced Security**: No cloud authentication or API tokens required - works with your existing Outlook setup
- **✅ Offline Capability**: Access and manage emails even without internet connection
- **✅ Privacy-First**: Your email data never touches external servers or cloud services

### Why Switch to Microsoft Graph API?
- **✅ Full Microsoft 365 Ecosystem**: Access Mail, Calendar, Teams, OneDrive, SharePoint
- **✅ Cross-Platform Support**: Works on Windows, macOS, Linux, and mobile devices
- **✅ Modern Authentication**: OAuth 2.0 with device code flow
- **✅ Cloud-Scale Performance**: No local application dependencies
- **✅ Online Access**: Works directly with cloud-based Microsoft 365 services

---

## 🚀 Quick Start (2 Minutes)

### What You'll Get
- **Smart Email Search**: "Find emails about budget approval from last week"
- **AI Email Writing**: Draft replies with context-aware suggestions  
- **Easy Organization**: Create folders and move emails with simple commands
- **Batch Forwarding**: Send emails to 100s of recipients in minutes, not hours

### 🤖 AI Behavior & Workflow

The [`agent_prompt_template.md`](agent_prompt_template.md) defines how the AI assistant behaves when managing emails:
- **Purpose**: Guide the AI's workflow for email search, summarization, and drafting
- **Key Rules**: AND logic for searches, 5-by-5 email display, confirmation before sending
- **Safety**: Built-in constraints ensure user control and prevent unauthorized actions
- **How to Use**: Copy the template content and use it to configure your AI assistant

*See the [template](agent_prompt_template.md) for complete behavior guidelines and workflow definitions*

### Why Choose Outlook MCP Server?

| Feature | Outlook MCP Server | Traditional Outlook | Outlook Add-ins |
|---------|-------------------|-------------------|-----------------|
| **AI Email Search** | ✅ "Find urgent emails from my boss" | ❌ Manual folder browsing | ⚠️ Basic search only |
| **Natural Language** | ✅ "Show me budget emails from last week" | ❌ Complex filters needed | ⚠️ Limited keywords |
| **Batch Forward 100+ Emails** | ✅ 2 minutes with CSV | ❌ 50+ minutes manual | ⚠️ 20+ minutes |
| **AI Email Writing** | ✅ Context-aware replies | ❌ Manual composition | ⚠️ Basic templates |
| **Setup Time** | ✅ 2 minutes | ✅ Already installed | ❌ 10-30 minutes |
| **Privacy** | ✅ 100% local processing | ✅ Local only | ⚠️ Cloud-dependent |
| **Cost** | ✅ Completely free | ✅ Included | 💰 $5-50/month |
| **Learning Curve** | ✅ Natural language | ✅ Familiar interface | ⚠️ New interface |

### Real-World Impact

**Before**: "I need to forward this email to 150 team members..."
- Manual: Click Forward → Type each email → Send → **2+ hours wasted**
- Traditional Outlook: Create distribution list → Add members → Forward → **30+ minutes**

**After**: "Forward this email to everyone in team.csv"
- Outlook MCP Server: Load CSV → AI forwards to all → **2 minutes total**
- **Time Saved**: 48+ minutes per batch operation!

### Requirements
- ✅ Python 3.8+
- ✅ Microsoft Outlook 2016+ (must be running)
- ✅ Windows 10+

### Installation & Setup

**Method 1: UVX (Recommended - Auto Dependencies)**
```bash
# 1. Install
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server
uvx --with "pywin32>=226" --with-editable "." outlook-mcp-server

# 2. Configure your AI assistant
# Use this in your MCP client settings:
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--with", "pywin32>=226", "--with-editable", ".", "outlook-mcp-server"]
    }
  }
}
```

**Method 2: Standard Python**
```bash
# 1. Install
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server
pip install -r requirements.txt
python -m outlook_mcp_server

# 2. Configure your AI assistant
# Use this in your MCP client settings:
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["-m", "outlook_mcp_server"]
    }
  }
}
```

**Method 3: Direct Source (Development)**
```bash
# 1. Install
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server
pip install -r requirements.txt

# 2. Configure your AI assistant
# Use this in your MCP client settings:
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\Project\\outlook-mcp-server\\outlook_mcp_server\\__main__.py"]
    }
  }
}
```

### Test Your Setup ✅
Ask your AI assistant: "Show me my recent emails" - if it works, you're ready!

### Configuration Troubleshooting 🛠️

**Common Issues & Solutions:**

| Problem | Check This | Solution |
|---------|------------|----------|
| **"uvx not found"** | Is UV installed? | `pip install uv` then retry |
| **"python not found"** | Python in PATH? | Use full path like `C:\Python39\python.exe` |
| **"Outlook not running"** | Outlook window open? | Start Outlook first, then restart MCP |
| **"Permission denied"** | Admin rights? | Run terminal as administrator |
| **"Module not found"** | Dependencies installed? | `pip install -r requirements.txt` |

**Configuration Verification:**
```bash
# Test your setup before connecting to AI
python -c "import outlook_mcp_server; print('✅ Server module loaded')"

# Test Outlook connection (Windows only)
python -c "import win32com.client; outlook = win32com.client.Dispatch('Outlook.Application'); print('✅ Outlook connected')"
```

**MCP Client-Specific Setup:**

**Claude Desktop:**
1. Open Claude Desktop settings
2. Find "MCP Servers" section
3. Click "Add Server" → "Custom"
4. Paste the JSON configuration
5. Restart Claude Desktop

**Other MCP Clients:**
- Look for "MCP Configuration" or "Server Settings"
- Add the JSON to your client's config file
- Usually located at: `~/.config/[client]/mcp.json`

**Still Stuck?** [Report an issue](https://github.com/marlonluo2018/outlook-mcp-server/issues) with your error message and setup details.

## 🎯 Core Features

### Email Management
- **Search**: Find emails by subject, sender, content, or date range
- **Compose**: Write new emails with AI assistance
- **Reply**: Smart replies that understand conversation context
- **Batch Forward**: Send emails to 100s of recipients from CSV files (saves hours!)

### Folder Management  
- **List**: See all your Outlook folders
- **Create**: Make new folders with simple commands
- **Move**: Organize emails between folders
- **Delete**: Remove folders (careful - this is permanent!)

## 📧 Batch Forwarding: Save Hours on Email Distribution

### Real-World Use Cases

**🎯 Team Updates**
- Forward weekly reports to your entire team
- Send important updates to all team members
- Distribute project updates to stakeholders

**📊 Marketing Campaigns**  
- Send newsletters to subscriber lists
- Forward promotional emails to customer segments
- Distribute event invitations to contact groups

**🏢 Corporate Communications**
- Send policy updates to all employees
- Forward training materials to departments
- Distribute announcements to company distribution lists

## 🔄 How It Works

### Simple Workflow
1. **Load emails**: "Show me recent emails" → Emails appear in cache
2. **Browse results**: View 5 emails per page with clear formatting
3. **Take action**: Reply, move, delete, or get AI summary
4. **Confirm before sending**: AI always asks before sending emails

### AI Assistant Behavior
- **Understands natural language**: "Find urgent emails from my boss"
- **Shows email summaries**: One-line overview + key action items
- **Drafts with context**: Replies understand the conversation
- **Never sends without permission**: Always confirms before sending

### Batch Forwarding Workflow
**1. Prepare Your CSV File**
```csv
email
john@company.com
jane@company.com
team@company.com
```

**2. Use Natural Language**
```
"Forward this email to everyone in my contacts.csv"
"Send this project update to my team list"
"Distribute this newsletter to subscribers.csv"
```

**3. AI Handles the Rest**
- Automatically splits large lists (max 500 per batch)
- Sends via BCC to protect recipient privacy
- Adds your custom message before original email
- Provides delivery confirmation

### Time Savings
**Manual forwarding**: 100 emails × 30 seconds = 50 minutes
**Batch forwarding**: 30 seconds setup + 2 minutes processing = 2.5 minutes
**You save**: 47.5 minutes per batch!

## 🔧 Common Commands

Try these with your AI assistant:

```
"Show me emails from last 3 days"
"Find emails about project updates" 
"Draft a reply to John about rescheduling"
"Create a folder called 'Work Projects'"
"Move email #3 to the Archive folder"
```

## 🛠️ Essential Tools

### Email Search & Loading
- `list_recent_emails_tool(days=7)` - Load recent emails (max: 30 days)
- `search_email_by_subject_tool("search term")` - Search email subjects
- `search_email_by_sender_name_tool("sender name")` - Search by sender
- `search_email_by_body_tool("search term")` - Search email content (slower)

### Email Actions
- `view_email_cache_tool(page=1)` - Browse loaded emails (5 per page)
- `get_email_by_number_tool(email_number)` - Get full email details
- `reply_to_email_by_number_tool(email_number, "reply text")` - Draft a reply
- `compose_email_tool("recipient@email.com", "subject", "body")` - Draft a new email

Both accept an `attachments` argument - a local file path, or a list of paths:

```python
reply_to_email_by_number_tool(1, "See the corrected sheet attached.",
                              attachments=r"C:\Users\me\Documents\report.xlsx")
compose_email_tool("team@example.com", "Monthly report", "Attached.",
                   attachments=[r"C:\a.pdf", r"C:\b.xlsx"])
```

Paths must exist on the machine running the server (`~` and environment variables
are expanded). Limits are 20 files and 25 MB in total, configurable via
`AttachmentConfig` in `backend/config.py`. If any file cannot be attached, no
draft is saved at all - you never get a draft that is silently missing a file.

### Automatic CC

Set `OUTLOOK_MCP_DEFAULT_CC` to CC someone on every draft this server creates -
one or more addresses separated by `;` or `,`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "...",
      "args": ["-m", "outlook_mcp_server"],
      "env": { "OUTLOOK_MCP_DEFAULT_CC": "boss@example.com" }
    }
  }
}
```

It is read from the environment, not committed to config, so a mailbox-specific
address stays out of version control. The addresses are added on top of whatever
CC list a reply already has, skipping anyone already there and never CCing the
person being replied to. Because everything lands in Drafts, you see the CC field
and can remove it before sending.

### Folder Operations
- `get_folder_list_tool()` - **Always use first** to see available folders
- `create_folder_tool("folder name")` - Create new folder
- `move_email_tool(email_number, "target folder")` - Move email between folders
- `move_folder_tool("source", "target")` - Move folders

### Safety Notes
- **Always check folder list first** before moving/deleting
- **Emails are cached** - use numbers from cache for operations
- **Nothing is ever sent** - composed messages, replies and batch forwards are
  saved to Drafts; you review and press Send yourself in Outlook

## ⚡ Performance

- Loads 100 emails in ~2 seconds
- Searches complete in real-time
- All processing happens locally (your data stays private)

## ✅ Quality & Reliability

### Robust Validation System
- **Input Validation**: All user inputs are validated before processing
- **Custom Error Messages**: Clear, actionable error messages for common issues
- **Safety Checks**: Prevents invalid operations before they cause problems

### Comprehensive Testing
- **145+ Unit Tests**: Every feature thoroughly tested
- **Configuration Coverage**: All settings and constants validated
- **Edge Case Handling**: Tested against unusual inputs and scenarios
- **Continuous Quality**: Tests run on every change to ensure reliability

### What This Means for You
- **Fewer Errors**: Validation catches mistakes before they cause problems
- **Better Error Messages**: Know exactly what went wrong and how to fix it
- **Reliable Operation**: Comprehensive testing ensures consistent performance
- **Safe Operations**: Built-in safeguards prevent accidental data loss

## 🛡️ Privacy & Security

- **100% Local Processing**: No data leaves your computer
- **No Cloud Services**: Works entirely offline
- **Secure by Design**: Uses your existing Outlook installation



## 📚 Need Help?

- Check the [agent prompt template](agent_prompt_template.md) for AI assistant setup
- See [configuration examples](mcp-config-uvx.json) for different installation methods
- Read the [Win32COM vs Graph API comparison](docs/win32com_vs_graph_api_comparison.md) for technical details
- Review [Win32COM API implementation guide](docs/win32com_api_implementation.md) for performance optimization
- Report issues on the [GitHub repository](https://github.com/marlonluo2018/outlook-mcp-server/issues)

---

<div align="center">

**⭐ Love saving time on email? [Star this repo](https://github.com/marlonluo2018/outlook-mcp-server/stargazers) - helps 10,000+ Outlook users discover AI email management! ⭐**

**💡 Quick star tip: Click the ⭐ button above - it takes 2 seconds and supports open-source email AI!**

</div>