# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.4] - 2026-01-15

### Fixed
- Attachment tools module structure: `__init__.py` now follows consistent pattern with other tool packages
- Resolved AttributeError when registering attachment tools with MCP server
- Server now starts successfully without import errors

---

## [1.0.3] - 2026-01-12

### Added
- **3 new Attachment Tools**: `get_email_attachments`, `download_email_attachment`, `send_email_with_attachments`
- Robust attachment detection in `formatters.py` via new `count_real_attachments()` function
- Smart filtering of inline images, email signatures, and embedded items based on ContentID, file type, and size

### Fixed
- Date format in `search_emails_in_custom_folder` now uses DD/MM/YYYY for European Windows locales (was MM/DD/YYYY)
- Email attachment count now accurately reflects real file attachments, not inline images

### Changed
- All email-related tools now use robust attachment detection
- `has_attachments` and `attachment_count` fields now filter out inline images and signature logos

---

## [1.0.2] - 2026-01-06

### Changed
- Auto-learning now happens dynamically on every `send_email` / `create_draft_email` operation (no longer at startup)
- Style learning analyzes only the most recent sent email (configurable via `STYLE_LEARNING_EMAIL_COUNT`)
- Removed caching mechanism - learning is performed live when `OUTLOOK_AUTO_LEARN_STYLE=true`

### Removed
- `config/user_email_style.json` cache file (no longer needed)
- `insert_styled_content_into_html()` function from `style_learner.py` (unused)

### Fixed
- Circular reference detection log level changed from WARNING to DEBUG in `folder_helper.py`
- Documentation updated to reflect actual auto-learning behavior

### Documentation
- Updated `signature_name` parameter documentation to clarify current limitations
- README.md auto-learning section now accurately describes live learning behavior

---

## [1.0.1] - 2025-12-31

### Changed
- `search_emails_in_custom_folder` now searches recursively by default (recursive: bool = True)
- This ensures subfolders are automatically included in searches, preventing missed emails

### Fixed
- Search in custom folders no longer misses emails stored in subfolders

---

## [1.0.0] - 2025-12-31

### Initial Public Release

MCP Outlook is a Model Context Protocol server that provides AI assistants with comprehensive Microsoft Outlook integration capabilities.

#### Features

**Email Management** (5 tools)
- `get_inbox_emails` - Retrieve emails from inbox with unread filtering
- `get_sent_emails` - Retrieve sent emails
- `search_emails` - Search emails across standard folders
- `send_email` - Send emails with CC/BCC, importance levels, HTML content, and Outlook signature integration
- `create_draft_email` - Create draft emails without sending, with HTML and signature support

**Folder Management** (3 tools)
- `list_outlook_folders` - List all Outlook folders (ultra-fast, no item counts)
- `search_emails_in_custom_folder` - Search in specific custom folders with date filtering and recursive support
- `list_outlook_rules` - List all Outlook mail rules

**Calendar Management** (3 tools)
- `get_calendar_events` - Get upcoming calendar events
- `create_calendar_event` - Create new calendar events with attendees
- `search_calendar_events` - Search events by subject or location

**Contact Management** (3 tools)
- `get_contacts` - Retrieve contacts with optional name filtering
- `create_contact` - Create new contacts
- `search_contacts` - Search contacts by name, email, or company

#### Architecture

**Ultra-Modular Design**
- 14 separate tool files (1 file per tool)
- Utility modules (connection, formatters, signature, folder helpers, style learner)
- Centralized configuration in `config.py`
- Auto-learning email style from sent emails

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

[1.0.1]: https://github.com/YOUR_USERNAME/mcp-outlook/releases/tag/v1.0.1
[1.0.0]: https://github.com/YOUR_USERNAME/mcp-outlook/releases/tag/v1.0.0
