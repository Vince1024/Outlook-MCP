"""
Configuration and Constants for MCP Outlook Server

This module contains all configuration constants used across the MCP Outlook server.
Centralizing configuration makes it easier to maintain and modify settings.

Version: 1.0.2
"""

import logging
import os

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
# USER CONFIGURATION (from environment variables)
# ============================================================================

# User email addresses
USER_EMAIL = os.getenv('OUTLOOK_USER_EMAIL', '')
PERSO_EMAIL = os.getenv('OUTLOOK_PERSO_EMAIL', '')
TEAM_EMAIL = os.getenv('OUTLOOK_TEAM_EMAIL', '')

# Primary folder for searches (where user's emails arrive via rules)
PRIMARY_FOLDER = os.getenv('OUTLOOK_PRIMARY_FOLDER', '')

# Default signature name for sent emails
DEFAULT_SIGNATURE = os.getenv('OUTLOOK_DEFAULT_SIGNATURE', '')

# Default limit for email searches (can be overridden per tool)
USER_DEFAULT_LIMIT = int(os.getenv('OUTLOOK_DEFAULT_LIMIT', DEFAULT_EMAIL_LIMIT))

# Auto-learn email style at startup (true/false)
AUTO_LEARN_STYLE = os.getenv('OUTLOOK_AUTO_LEARN_STYLE', 'false').lower() == 'true'

# Number of sent emails to analyze for style learning
STYLE_LEARNING_EMAIL_COUNT = 1

# ============================================================================
# EXCLUDED STORES/FOLDERS
# ============================================================================
# These will be skipped when listing folders or searching
# Can be configured via OUTLOOK_EXCLUDED_FOLDERS env var (semicolon-separated)

EXCLUDED_FOLDERS_ENV = os.getenv('OUTLOOK_EXCLUDED_FOLDERS', '')
EXCLUDED_STORES = [folder.strip() for folder in EXCLUDED_FOLDERS_ENV.split(';') if folder.strip()]

# Add default exclusions if not already present
DEFAULT_EXCLUSIONS = ['All Public Folders', 'Dossiers publics']
for exclusion in DEFAULT_EXCLUSIONS:
    if exclusion not in EXCLUDED_STORES:
        EXCLUDED_STORES.append(exclusion)

