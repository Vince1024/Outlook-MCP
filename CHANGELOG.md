# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-12-31

### Initial Public Release

MCP Outlook is a Model Context Protocol server that provides AI assistants with comprehensive Microsoft Outlook integration capabilities.

#### Features

**Email Management** (9 tools)
- `get_inbox_emails` - Retrieve emails from inbox with unread filtering
- `get_sent_emails` - Retrieve sent emails
- `search_emails` - Search emails across standard folders
- `send_email` - Send emails with CC/BCC, importance levels, HTML content, and Outlook signature integration
- `create_draft_email` - Create draft emails without sending, with HTML and signature support
- `get_email_attachments` - Get list of attachments from a specific email with details
- `download_email_attachment` - Download specific attachment from an email to disk
- `send_email_with_attachments` - Send emails with file attachments
- `learn_user_email_preferences` - Automatically learn email formatting preferences from sent emails

**Folder Management** (3 tools)
- `list_outlook_folders` - List all Outlook folders (ultra-fast, no item counts)
- `search_emails_in_custom_folder` - Search in specific custom folders with date filtering and recursive support
- `list_outlook_rules` - List all Outlook mail rules

**Calendar Management** (5 tools)
- `get_calendar_events` - Get upcoming calendar events
- `create_calendar_event` - Create new calendar events with attendees
- `search_calendar_events` - Search events by subject or location
- `get_meeting_requests` - Get pending meeting invitations that need a response
- `respond_to_meeting` - Accept, decline, or tentatively respond to meeting invitations

**Contact Management** (3 tools)
- `get_contacts` - Retrieve contacts with optional name filtering
- `create_contact` - Create new contacts
- `search_contacts` - Search contacts by name, email, or company

**Out-of-Office Management** (3 tools)
- `get_out_of_office_settings` - Get current automatic reply configuration
- `set_out_of_office` - Configure automatic replies (immediate or scheduled)
- `disable_out_of_office` - Disable automatic replies

**MCP Resources** (3 resources)
- `outlook://inbox/unread-count` - Real-time unread email count monitoring
- `outlook://inbox/recent` - 5 most recent emails snapshot
- `outlook://calendar/today` - Today's calendar events overview

#### Architecture

**Ultra-Modular Design**
- 14 separate tool files (1 file per tool, < 100 lines each)
- 4 utility modules (connection, formatters, signature, folder helpers)
- Centralized configuration in `config.py`
- Comprehensive unit test suite (14 test files)
- Archive of monolithic version preserved in `/V1/` for reference

#### Performance Optimizations

- **Folder caching** - 45x faster on repeated searches
- **Date filtering** - Search only recent emails (default: 2 days)
- **Direct indexing** - Faster iteration without items.Count
- **Reduced limits** - Prevents long freezes (max 50 emails)
- **Smart defaults** - Optimized for daily usage
- **Recursive folder search** - With safety limits (max depth, cycle detection)

#### Documentation

- Comprehensive README with installation and usage
- QUICK_START guide for rapid setup
- EXAMPLES with real-world use cases
- ARCHITECTURE documentation with ultra-modular design details
- CONTRIBUTING guide for developers
- Complete technical documentation (1700+ lines)

#### Security & Privacy

- **GDPR-compliant logging** - Only counts, flags, and hashes logged (no PII)
- **Email content truncation** - Prevents data leakage in responses
- **No credential storage** - Uses Windows COM automation only
- **Local-only operation** - All data stays on user's machine

#### Technical Details

- **Framework**: FastMCP
- **COM Automation**: pywin32
- **Python Version**: 3.10+
- **Platform**: Windows 10/11
- **Outlook Version**: 2010+
- **Testing**: pytest with comprehensive unit tests
- **Code Quality**: Full docstring coverage, structured logging, robust error handling

---

## Future Roadmap

### Planned Features
- [ ] Task management integration
- [ ] Folder management (create, move, delete)
- [ ] Advanced filtering (flags, categories, custom properties)
- [ ] Email rules management (create, modify, delete)
- [ ] Cross-platform support exploration
- [ ] Attachment content preview/thumbnails
- [ ] Bulk attachment operations

### Performance Improvements
- [ ] Async operations for better responsiveness
- [ ] Batch operations for multiple items
- [ ] Enhanced caching strategies
- [ ] Background sync capabilities

---

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for information on how to contribute to this project.

---

[1.0.0]: https://github.com/YOUR_USERNAME/mcp-outlook/releases/tag/v1.0.0
