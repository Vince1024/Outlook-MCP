"""
Configuration and Constants for MCP Outlook Server

This module contains all configuration constants used across the MCP Outlook server.
Centralizing configuration makes it easier to maintain and modify settings.

Version: 2.0.0
"""

import logging

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================

# Logging level - controls verbosity of logs
# Options: 
#   - logging.DEBUG   : Verbose (shows operation start, intermediate steps, success)
#   - logging.INFO    : Normal (shows success/errors only)
#   - logging.WARNING : Warnings and above
#   - logging.ERROR   : Only errors
#   - logging.CRITICAL: Silent (no logs)
LOG_LEVEL = logging.DEBUG

# Enable/disable file logging
# True  = Logs written to file logs/outlook_mcp_YYYY-MM-DD.log
# False = No logs (silent mode)
LOG_FILE = True

# Log message format (timestamp, level, module name, message)
# Padding: levelname=8 chars (left-aligned), name=40 chars (left-aligned)
LOG_FORMAT = '%(asctime)s | %(levelname)-8s | %(name)-32s | %(message)s'

# Loggers to silence (third-party noise)
# NOTE: Do NOT include '__main__' here - that's OUR logger!
SILENT_LOGGERS = [
    'mcp', 'FastMCP', 'fastmcp', 'mcp.server', 'fastmcp.server',
    'mcp.client', 'fastmcp.client', 'asyncio'
]

# ============================================================================
# OUTLOOK FOLDER CONSTANTS
# ============================================================================
# From Microsoft Outlook Object Model documentation
# These constants represent the default folder IDs in the Outlook namespace

OUTLOOK_FOLDER_INBOX = 6      # Inbox folder for incoming emails
OUTLOOK_FOLDER_SENT = 5        # Sent items folder
OUTLOOK_FOLDER_DRAFTS = 16     # Drafts folder for unsent emails
OUTLOOK_FOLDER_DELETED = 3     # Deleted items (trash)
OUTLOOK_FOLDER_OUTBOX = 4      # Outbox for emails pending send
OUTLOOK_FOLDER_JUNK = 23       # Junk/spam folder
OUTLOOK_FOLDER_CALENDAR = 9    # Calendar folder for appointments
OUTLOOK_FOLDER_CONTACTS = 10   # Contacts folder

# ============================================================================
# OUTLOOK ITEM TYPE CONSTANTS
# ============================================================================
# For CreateItem method

OUTLOOK_ITEM_MAIL = 0          # Mail item type
OUTLOOK_ITEM_APPOINTMENT = 1   # Calendar appointment type
OUTLOOK_ITEM_CONTACT = 2       # Contact item type

# ============================================================================
# EMAIL IMPORTANCE LEVEL CONSTANTS
# ============================================================================

IMPORTANCE_LOW = 0
IMPORTANCE_NORMAL = 1
IMPORTANCE_HIGH = 2

# ============================================================================
# PERFORMANCE AND LIMITS
# ============================================================================
# Default limits to prevent performance issues with large mailboxes

DEFAULT_EMAIL_LIMIT = 5            # Reduced from 10 to minimize Outlook freezing
MAX_EMAIL_LIMIT = 50               # Reduced from 100 to prevent long freezes
DEFAULT_CONTACT_LIMIT = 50
MAX_CONTACT_LIMIT = 200            # Reasonable limit for contact queries
EMAIL_BODY_PREVIEW_LENGTH = 500    # Truncate email bodies to prevent excessive data transfer
DEFAULT_DAYS_BACK = 2              # Only search emails from last 2 days by default (ultra-fast!)

# ============================================================================
# EXCLUDED STORES/FOLDERS
# ============================================================================
# These will be skipped when listing folders or searching
# Add your specific folders to exclude here if needed

EXCLUDED_STORES = [
    # Example: "Team Mailbox Name",
]

