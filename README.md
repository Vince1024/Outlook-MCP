# MCP Outlook

[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io) 

<img src="https://api.iconify.design/material-symbols/mail-outline.svg?color=%23004080" width="60"> <img src="https://api.iconify.design/material-symbols/calendar-month-outline.svg?color=%23004080" width="60"> <img src="https://api.iconify.design/material-symbols/perm-contact-calendar.svg?color=%23004080" width="60"> <img src="https://api.iconify.design/material-symbols/folder-outline.svg?color=%23004080" width="60"> <img src="https://api.iconify.design/material-symbols/settings-account-box.svg?color=%23004080" width="60">

A Model Context Protocol (MCP) server for Microsoft Outlook integration.

**Version**: 1.0.3 | **Documentation**: [DOCUMENTATION.md](DOCUMENTATION.md) | **Contributing**: [CONTRIBUTING.md](CONTRIBUTING.md) | **Examples**: [EXAMPLES.md](EXAMPLES.md)

## Overview

This MCP server provides AI assistants with the ability to interact with Microsoft Outlook, including:

- <img src="https://api.iconify.design/material-symbols/mail-outline.svg?color=%23004080" width="20"> **Email Management**: Read, search, send, and draft emails with HTML support and Outlook signatures
- <img src="https://api.iconify.design/material-symbols/attach-file.svg?color=%23004080" width="20"> **Attachment Management**: List, download, and send email attachments with robust detection
- <img src="https://api.iconify.design/material-symbols/calendar-month-outline.svg?color=%23004080" width="20"> **Calendar Management**: View, create, and search calendar events
- <img src="https://api.iconify.design/material-symbols/perm-contact-calendar.svg?color=%23004080" width="20"> **Contact Management**: View, create, and search contacts
- <img src="https://api.iconify.design/material-symbols/folder-outline.svg?color=%23004080" width="20"> **Folder Management**: List folders, search in custom folders, and view Outlook rules
- <img src="https://api.iconify.design/material-symbols/settings-account-box.svg?color=%23004080" width="20"> **Auto-Learning Style**: Automatically learns your email formatting preferences (font, size, color) from sent emails

## Features

> **Detailed documentation for each feature**: [DOCUMENTATION.md](DOCUMENTATION.md)

### <img src="https://api.iconify.design/material-symbols/mail-outline.svg?color=%23004080" width="30"> 5 Email Tools ([Documentation](DOCUMENTATION.md#email-tools)) 

- `get_inbox_emails` - Retrieve emails from inbox with filtering options
- `get_sent_emails` - Retrieve sent emails
- `search_emails` - Search emails across folders by subject, body, or sender
- `send_email` - Send emails with CC/BCC support, HTML content, and **auto-learned style**
- `create_draft_email` - Create draft emails without sending, with HTML and **auto-learned style**

> **Auto-Learning Style**: When `OUTLOOK_AUTO_LEARN_STYLE=true`, the server dynamically learns your email formatting style (font-family, font-size, color) from your most recent sent email each time you send or create a draft. This learned style is automatically applied to plain-text emails. No caching - learning happens live on every send/draft operation.

### <img src="https://api.iconify.design/material-symbols/attach-file.svg?color=%23004080" width="30"> 3 Attachment Tools ([Documentation](DOCUMENTATION.md#attachment-tools))

- `get_email_attachments` - List all attachments from an email (with robust detection)
- `download_email_attachment` - Download a specific attachment to disk
- `send_email_with_attachments` - Send emails with file attachments

> **Robust Detection**: Attachment detection intelligently filters out inline images, email signatures, and embedded items to show only real file attachments. Uses ContentID detection, file size filtering, and type checking.

### <img src="https://api.iconify.design/material-symbols/calendar-month-outline.svg?color=%23004080" width="30"> 3 Calendar Tools ([Documentation](DOCUMENTATION.md#calendar-tools))

- `get_calendar_events` - Get upcoming calendar events
- `create_calendar_event` - Create new calendar events with attendees
- `search_calendar_events` - Search events by subject or location

### <img src="https://api.iconify.design/material-symbols/perm-contact-calendar.svg?color=%23004080" width="30"> 3 Contact Tools ([Documentation](DOCUMENTATION.md#contact-tools))

- `get_contacts` - Retrieve contacts with optional name filtering
- `create_contact` - Create new contacts
- `search_contacts` - Search contacts by name, email, or company

### <img src="https://api.iconify.design/material-symbols/folder-outline.svg?color=%23004080" width="30"> 3 Folder Tools ([Documentation](DOCUMENTATION.md#folder-tools))

- `list_outlook_folders` - List all Outlook folders (ultra-fast, no item counts)
- `search_emails_in_custom_folder` - Search in specific custom folders with date filtering
- `list_outlook_rules` - List all Outlook rules with conditions and actions

## Performance Optimizations

This MCP has been heavily optimized for **large mailboxes** and to **minimize Outlook freezing**:

- **Folder caching** - 45x faster on repeated searches
- **Date filtering** - Search only recent emails (default: 2 days)
- **Direct indexing** - Faster iteration without `items.Count`
- **Reduced limits** - Prevents long freezes (max 50 emails)
- **Smart defaults** - Optimized for daily usage
- **Silent logging** - Minimal log output for cleaner integration

**See [DOCUMENTATION.md - Performances](DOCUMENTATION.md#performances) for detailed performance information.**

## Installation

> **Quick Start Guide**: [QUICK_START.md](QUICK_START.md) | **Full Installation Guide**: [DOCUMENTATION.md - Installation](DOCUMENTATION.md#installation)

### Prerequisites

- **Windows OS** (required for COM automation)
- **Microsoft Outlook** installed and configured
- **Python 3.10+**

### Setup

1. Clone or download this repository

2. Install dependencies:

```bash
pip install -r requirements.txt
```

Or using the project file:

```bash
pip install -e .
```

3. Verify Outlook is running and configured with an account

4. Test the installation:

```bash
python tests/test_connection.py
```

## Usage

### Running the Server

Run the MCP server directly:

```bash
python src/outlook_mcp.py
```

Or using FastMCP's built-in CLI:

```bash
fastmcp run src/outlook_mcp.py
```

### Configuration for Cursor/Claude Desktop

Add this configuration to your MCP settings file:

**For Cursor** (`~/.cursor/mcp.json` or workspace settings):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "C:/Users/YOUR_USERNAME/source/repos/MCP/src/outlook_mcp.py"
      ],
      "env": {}
    }
  }
}
```

**For Claude Desktop** (`%APPDATA%/Claude/claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "C:/Users/YOUR_USERNAME/source/repos/MCP/src/outlook_mcp.py"
      ]
    }
  }
}
```

**Important**: Replace `YOUR_USERNAME` with your actual Windows username.

### Testing the Server

You can test the server using FastMCP's interactive mode:

```bash
fastmcp dev src/outlook_mcp.py
```

This will open an interactive prompt where you can test the tools.

## Tool Examples

> **More Examples**: [EXAMPLES.md](EXAMPLES.md) - Real-world use cases and workflows

### Reading Emails

```python
# Get last 10 unread emails
get_inbox_emails(limit=10, unread_only=True)

# Search for emails about "meeting"
search_emails(query="meeting", folder="inbox", limit=20)
```

### Sending Emails

```python
# Send a simple email
send_email(
    to="colleague@company.com",
    subject="Meeting Follow-up",
    body="Hi, following up on our meeting...",
    importance="high"
)

# Send email with HTML content and Outlook signature
send_email(
    to="colleague@company.com",
    subject="Project Update",
    html_body="<h1>Update</h1><p>Here are the details...</p>",
    signature_name="My Signature"
)

# Create a draft with multiple recipients and signature
create_draft_email(
    to="team@company.com",
    subject="Project Update",
    body="Here's the latest update...",
    cc="manager@company.com",
    signature_name="My Signature"
)
```

### Calendar Management

```python
# Get next 7 days of events
get_calendar_events(days_ahead=7)

# Create a meeting
create_calendar_event(
    subject="Team Standup",
    start_time="2025-01-15 09:00",
    end_time="2025-01-15 09:30",
    location="Conference Room A",
    required_attendees="team@company.com",
    reminder_minutes=15
)

# Search for meetings
search_calendar_events(query="standup", days_range=30)
```

### Contact Management

```python
# Get all contacts
get_contacts(limit=50)

# Search for a contact
search_contacts(query="John Smith")

# Create a new contact
create_contact(
    full_name="Jane Doe",
    email="jane.doe@company.com",
    company="Acme Corp",
    job_title="Product Manager",
    mobile_phone="+1-555-1234"
)
```


## Security & Permissions

- This server requires access to your Outlook data
- It uses Windows COM automation (no credentials stored)
- All operations are performed with your current Outlook profile's permissions
- Make sure Outlook is running and configured before starting the server

## Troubleshooting

### "Unable to connect to Outlook"

- Ensure Microsoft Outlook is installed and running
- Verify Outlook is configured with at least one email account
- Try restarting Outlook

### "ImportError: No module named 'win32com'"

- Install pywin32: `pip install pywin32`
- After installation, run: `python Scripts/pywin32_postinstall.py -install` (if needed)

### Permission Errors

- Run your terminal/IDE as Administrator (may be required for COM automation)
- Check that Outlook is not blocked by security policies

### Date Parsing Issues

- Use ISO format for dates: `2025-01-15 14:00`
- Supported formats: "YYYY-MM-DD HH:MM", "tomorrow 2pm", "next Monday 10am"

## Development

### Project Structure

```
mcp-outlook/
├── src/
│   ├── __init__.py
│   └── outlook_mcp.py       # Main MCP server
├── pyproject.toml           # Project configuration
├── requirements.txt         # Dependencies
├── .gitignore
└── README.md
```

### Adding New Tools

To add a new tool, use the `@mcp.tool()` decorator:

```python
@mcp.tool()
def my_new_tool(param1: str, param2: int = 10) -> str:
    """
    Tool description.
    
    Args:
        param1: Description of param1
        param2: Description of param2 (default: 10)
    
    Returns:
        JSON string with results
    """
    # Implementation
    return json.dumps({"success": True, "data": "..."})
```

### Running Tests

```bash
pytest
```

### Code Formatting

```bash
black src/
ruff check src/
```

## Limitations

- **Windows Only**: Uses COM automation which is Windows-specific
- **Outlook Required**: Microsoft Outlook must be installed and running
- **Single Account**: Works with the default Outlook profile only
- **Performance**: Large mailboxes may have slower search performance

## Roadmap

### Recently Added
- [x] HTML email support
- [x] Outlook signature integration
- [x] Silent logging for cleaner integration
- [x] Outlook rules listing
- [x] Auto-learning email style from sent emails
- [x] Recursive folder search

### Planned Features
- [ ] Attachment download/upload support
- [ ] Meeting response handling (accept/decline/tentative)
- [ ] Out-of-office settings (get/set/disable)
- [ ] Task management integration
- [ ] Folder management (create, move, delete)
- [ ] Advanced filtering (flags, categories, custom properties)
- [ ] Email rules creation and modification
- [ ] Cross-platform support (investigate MAPI alternatives)

## License

MIT License - See LICENSE file for details.

## Contributing

Contributions are welcome! Please read our [Contributing Guide](CONTRIBUTING.md) for details on:

- How to set up your development environment
- Code style guidelines (Black + Ruff)
- How to submit pull requests
- Roadmap and planned features

Quick Start for Contributors:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes and test thoroughly
4. Commit with conventional commits (`feat:`, `fix:`, `docs:`, etc.)
5. Push and create a Pull Request

See [CONTRIBUTING.md](CONTRIBUTING.md) for detailed instructions.

## Documentation

- **[README.md](README.md)** (this file) - Overview and quick start
- **[DOCUMENTATION.md](DOCUMENTATION.md)** - Complete technical documentation
- **[EXAMPLES.md](EXAMPLES.md)** - Real-world examples and use cases
- **[QUICK_START.md](QUICK_START.md)** - 5-minute setup guide
- **[CONTRIBUTING.md](CONTRIBUTING.md)** - How to contribute
- **[CHANGELOG.md](CHANGELOG.md)** - Version history

## Support

For issues or questions:
- **Create an issue**: [GitHub Issues](https://github.com/YOUR_USERNAME/mcp-outlook/issues)
- **Check existing issues** for similar problems
- **Provide details**: Windows version, Outlook version, Python version, error logs

Before creating an issue:
1. Run `python tests/test_connection.py` and include the output
2. Check the [Troubleshooting](DOCUMENTATION.md#gestion-des-erreurs) section
3. Review [existing issues](https://github.com/YOUR_USERNAME/mcp-outlook/issues)

## Acknowledgments

- Built with [FastMCP](https://github.com/jlowin/fastmcp)
- Uses [pywin32](https://github.com/mhammond/pywin32) for COM automation
- Inspired by the MCP Atlassian server architecture

---

**Note**: This tool accesses your local Outlook data. Ensure you follow your organization's security policies when handling email and calendar data.

