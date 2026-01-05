# MCP-OUTLOOK - Architecture (Ultra-Modular)

## Overview

MCP-OUTLOOK uses an **ultra-modular** architecture where **each MCP tool is in its own file**. This approach maximizes granularity, maintainability, and testability of the code.

The old monolithic version (1990 lines) is preserved in `/V1/` for reference.

## Project Structure

```
MCP-OUTLOOK/
├── V1/
│   └── outlook_mcp.py              # Old monolithic version (1990 lines, archived)
│
├── src/
│   ├── __init__.py
│   ├── outlook_mcp.py              # Current entry point (193 lines, modular)
│   │
│   ├── config.py                   # Centralized configuration and constants
│   │
│   ├── utils/                      # Reusable utilities
│   │   ├── __init__.py
│   │   ├── outlook_connection.py  # Outlook COM connection
│   │   ├── formatters.py          # Outlook object to JSON formatting
│   │   ├── signature_helper.py    # Email signature management
│   │   └── folder_helper.py       # Folder operations
│   │
│   └── tools/                      # MCP Tools - 1 FILE PER TOOL
│       ├── __init__.py
│       │
│       ├── email/                  # 5 email tools
│       │   ├── __init__.py
│       │   ├── get_inbox_emails.py
│       │   ├── get_sent_emails.py
│       │   ├── search_emails.py
│       │   ├── send_email.py
│       │   └── create_draft_email.py
│       │
│       ├── folder/                 # 3 folder tools
│       │   ├── __init__.py
│       │   ├── list_outlook_folders.py
│       │   ├── search_emails_in_custom_folder.py
│       │   └── list_outlook_rules.py
│       │
│       ├── calendar/               # 3 calendar tools
│       │   ├── __init__.py
│       │   ├── get_calendar_events.py
│       │   ├── create_calendar_event.py
│       │   └── search_calendar_events.py
│       │
│       └── contact/                # 3 contact tools
│           ├── __init__.py
│           ├── get_contacts.py
│           ├── create_contact.py
│           └── search_contacts.py
│
├── tests/                          # Unit tests (to be created)
│   ├── test_utils/
│   │   ├── test_formatters.py
│   │   └── test_folder_helper.py
│   └── test_tools/
│       ├── test_email/
│       │   ├── test_get_inbox_emails.py
│       │   ├── test_send_email.py
│       │   └── ...
│       ├── test_folder/
│       ├── test_calendar/
│       └── test_contact/
│
├── ARCHITECTURE.md                 # This file
├── MIGRATION_GUIDE.md              # Migration guide
└── start_mcp_server.bat            # Startup script
```

## Ultra-Modular Architecture Principles

### 1. One File Per Tool

Each MCP tool has its own file:

- **get_inbox_emails.py**: 1 file for 1 tool
- **send_email.py**: 1 file for 1 tool
- **search_emails.py**: 1 file for 1 tool
- etc.

**Benefits**:
- Easy to find code for a specific tool
- Isolated modifications = no risk of regression on other tools
- Ultra-targeted unit tests
- Very clear Git diffs

### 2. Organization by Category

Tools are organized in subfolders by domain:

```
tools/
├── email/      # All email-related tools
├── folder/     # All folder-related tools
├── calendar/   # All calendar-related tools
└── contact/    # All contact-related tools
├── calendar/   # All calendar-related tools
└── contact/    # All contact-related tools
```

### 3. Uniform Registration Pattern

Each tool file follows the same pattern:

```python
# In tools/email/get_inbox_emails.py
def register(mcp):
    @mcp.tool()
    def get_inbox_emails(limit: int = 5, unread_only: bool = False) -> str:
        """Detailed docstring"""
        # Implementation
        pass
```

### 4. Relative Imports

Tools use relative imports to access utils and config:

```python
from ...config import MAX_EMAIL_LIMIT
from ...utils import get_outlook_application, format_email
```

## Complete Tool Inventory

### Email Tools (5 files)

| File | Tool | Description |
|------|------|-------------|
| `get_inbox_emails.py` | `get_inbox_emails()` | Retrieves emails from inbox |
| `get_sent_emails.py` | `get_sent_emails()` | Retrieves sent emails |
| `search_emails.py` | `search_emails()` | Searches emails by keyword |
| `send_email.py` | `send_email()` | Sends an email |
| `create_draft_email.py` | `create_draft_email()` | Creates a draft email |

### Folder Tools (3 files)

| File | Tool | Description |
|------|------|-------------|
| `list_outlook_folders.py` | `list_outlook_folders()` | Lists all Outlook folders |
| `search_emails_in_custom_folder.py` | `search_emails_in_custom_folder()` | Searches in a custom folder |
| `list_outlook_rules.py` | `list_outlook_rules()` | Lists Outlook rules |

### Calendar Tools (3 files)

| File | Tool | Description |
|------|------|-------------|
| `get_calendar_events.py` | `get_calendar_events()` | Retrieves calendar events |
| `create_calendar_event.py` | `create_calendar_event()` | Creates an event |
| `search_calendar_events.py` | `search_calendar_events()` | Searches events |

### Contact Tools (3 files)

| File | Tool | Description |
|------|------|-------------|
| `get_contacts.py` | `get_contacts()` | Retrieves contacts |
| `create_contact.py` | `create_contact()` | Creates a contact |
| `search_contacts.py` | `search_contacts()` | Searches contacts |

**Total: 14 files for 14 tools**

## Architecture Evolution

| Aspect | Monolithic (archived in /V1/) | Current (Ultra-Modular) |
|--------|---------------------|----------------------|
| **Files** | 1 monolith (1990 lines) | 14 files (< 100 lines each) |
| **Organization** | Everything mixed | 1 file = 1 tool |
| **Maintainability** | Difficult | Very easy |
| **Testability** | Difficult | Very easy |
| **Scalability** | Risk of regression | No risk |
| **Finding code** | Search in 1990 lines | Open the right file |
| **Modifying a tool** | Risk of affecting others | Isolated modification |
| **Unit tests** | 1 large test file | 1 test per tool |

## Usage

### Starting the Server

```bash
python src/outlook_mcp.py
```

or

```bash
start_mcp_server.bat
```

### Expected Output

```
============================================================
MCP Outlook Server - Ultra-Modular Architecture
============================================================
✓ 14 tools loaded (each in its own file)
  - 5 Email tools
  - 3 Folder tools
  - 3 Calendar tools
  - 3 Contact tools

Server ready! Press Ctrl+C to stop.
============================================================
```

### Importing into Python Code

```python
from src.outlook_mcp import mcp
from src.utils import get_outlook_application, format_email
from src.config import MAX_EMAIL_LIMIT
```

## Benefits of Ultra-Modular Architecture

### 1. Ultra-Simplified Maintenance

- Need to modify `send_email`? Open `tools/email/send_email.py`
- No need to search through 2000 lines
- No risk of affecting `get_inbox_emails` or other tools

### 2. Ultra-Targeted Tests

```python
# tests/test_tools/test_email/test_send_email.py
from src.tools.email.send_email import register

def test_send_email_success():
    # Test only send_email
    pass

def test_send_email_with_signature():
    # Test signature
    pass
```

### 3. Facilitated Collaboration

- Dev A works on `send_email.py`
- Dev B works on `get_calendar_events.py`
- **Zero Git conflicts!**

### 4. Trivial Addition of New Tools

To add a new tool:

1. Create `tools/email/forward_email.py`
2. Copy the pattern from `send_email.py`
3. Implement the logic
4. Add the import in `outlook_mcp.py`

That's all! No other file to modify.

### 5. Ultra-Fast Code Review

Pull request to modify `send_email`:
- **1 modified file**: `tools/email/send_email.py`
- Clear and easy to review diff
- No pollution with other tools

## Recommended Unit Tests

### Structure

```
tests/
├── test_utils/
│   ├── test_formatters.py
│   ├── test_folder_helper.py
│   └── test_signature_helper.py
└── test_tools/
    ├── test_email/
    │   ├── test_get_inbox_emails.py
    │   ├── test_get_sent_emails.py
    │   ├── test_search_emails.py
    │   ├── test_send_email.py
    │   └── test_create_draft_email.py
    ├── test_folder/
    │   ├── test_list_outlook_folders.py
    │   ├── test_search_emails_in_custom_folder.py
    │   └── test_list_outlook_rules.py
    ├── test_calendar/
    │   └── ...
    └── test_contact/
        └── ...
```

### Example Test

```python
# tests/test_tools/test_email/test_send_email.py
import pytest
from unittest.mock import Mock, patch
from src.tools.email.send_email import register

def test_send_email_basic():
    mcp_mock = Mock()
    register(mcp_mock)
    
    # Get the registered function
    send_email_func = mcp_mock.tool.call_args[0][0]
    
    # Test with mocks
    with patch('src.tools.email.send_email.get_outlook_application'):
        result = send_email_func(
            to="test@example.com",
            subject="Test",
            body="Test message"
        )
        # Assertions...
```

## Technical Notes

### Folder Cache

The cache is in `utils/folder_helper.py`:

```python
from src.utils import FOLDER_CACHE

# Clear cache if needed
FOLDER_CACHE.clear()
```

### Logging

Centralized configuration in `config.py`:

```python
LOG_LEVEL = logging.CRITICAL
SILENT_LOGGERS = ['mcp', 'FastMCP', ...]
```

### Imports in outlook_mcp.py

```python
# Import each individual tool
from tools.email import (
    get_inbox_emails, get_sent_emails, search_emails,
    send_email, create_draft_email
)

# Register each tool
get_inbox_emails.register(mcp)
send_email.register(mcp)
# etc.
```

## FAQ

### Q: Why 1 file per tool and not per category?

**A:** 
- **Now**: 1 file of 100 lines easy to understand
- **Before**: 1 file of 400 lines with 5 tools mixed
- Result: 5x easier maintenance!

### Q: Isn't that too many files?

**A:** No! That's exactly the advantage:
- Need `send_email`? Open `send_email.py`
- No need to open and search in `email_tools.py`
- Modern IDE: Ctrl+P → `send_email.py` → found in 1 second

### Q: Is performance impacted?

**A:** Absolutely not! Python imports are cached. No performance impact.

### Q: How to add a tool?

**A:**
1. Create `tools/email/my_new_tool.py`
2. Copy the pattern from an existing tool
3. Implement
4. Add in `outlook_mcp.py`:
   ```python
   from tools.email import my_new_tool
   my_new_tool.register(mcp)
   ```

### Q: What is the V1/ directory?

**A:** It contains the archived monolithic version (1990 lines). It's kept for historical reference and to understand the evolution of the project.

## Conclusion

The ultra-modular architecture brings a drastic improvement:

- **14 files** instead of 1
- **< 100 lines** per file instead of 1990
- **1 tool = 1 file** = trivial maintenance
- **No risk** of regression between tools
- **Ultra-targeted** and easy tests

This is the maximum level of modularity for this project!

**Last Update**: December 31, 2025  
**Version**: 1.0.2  
**Architecture**: Ultra-Modular (1 file per tool)  
**Author**: Claude + VP
