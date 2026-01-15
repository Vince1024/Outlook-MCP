"""
MCP Server for Microsoft Outlook (Ultra-Modular Architecture)

This module provides a FastMCP server that interfaces with Microsoft Outlook via COM automation.
Each MCP tool is in its own individual file for maximum granularity and maintainability.

Features:
    - Ultra-modular architecture: 1 file per tool
    - Improved maintainability and testability
    - Better code organization by feature
    - Reusable utility functions
    - Easy to extend with new tools

Tools:
    - Email management: Read, send, search, and create drafts
    - Calendar management: Read, create, and search calendar events
    - Contact management: Read, create, and search contacts
    - Folder management: List folders, search in custom folders, list rules

Requirements:
    - Microsoft Outlook installed and configured on Windows
    - Python packages: win32com, fastmcp, python-dateutil

Usage:
    Run as a standalone MCP server:
        python outlook_mcp.py
        
Security Notes:
    - This module accesses local Outlook data via COM
    - No credentials are logged or transmitted
    - Email body content is truncated in responses to prevent data leakage
    
Version: 1.0.2
Architecture: Ultra-Modular (1 file per tool)
"""

import logging
import sys
from pathlib import Path
from datetime import datetime

# Add src directory to path to enable absolute imports
src_path = Path(__file__).parent
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

# ============================================================================
# LOGGING CONFIGURATION (BEFORE FastMCP IMPORT!)
# ============================================================================
# This MUST be done before importing FastMCP to prevent it from overriding our config
#
# IMPORTANT: No console output (no stderr handler) to keep console clean.
# MCP protocol uses stdout for JSON-RPC communication, and stderr would 
# display as [error] in red in Cursor console, which is confusing.
# All logs go to file only: logs/outlook_mcp_YYYY-MM-DD.log

# Import configuration first
from config import LOG_LEVEL, LOG_FORMAT, LOG_FILE, SILENT_LOGGERS

# Get the root logger and configure it directly (bypasses basicConfig issues)
root_logger = logging.getLogger()
root_logger.setLevel(LOG_LEVEL)

# Remove any existing handlers from root logger (clean slate)
for handler in root_logger.handlers[:]:
    root_logger.removeHandler(handler)

# Create formatter
formatter = logging.Formatter(LOG_FORMAT)

# Create logs directory if it doesn't exist
log_dir = Path(__file__).parent.parent / "logs"
log_dir.mkdir(exist_ok=True)

# Create and add file handler (ONLY if LOG_FILE is enabled)
if LOG_FILE:
    # Generate filename with today's date: outlook_mcp_2025-12-19.log
    today = datetime.now().strftime('%Y-%m-%d')
    log_filename = f"outlook_mcp_{today}.log"
    
    file_handler = logging.FileHandler(log_dir / log_filename, encoding='utf-8')
    file_handler.setLevel(LOG_LEVEL)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)
else:
    # No file logging: add NullHandler to prevent "No handlers" warnings
    root_logger.addHandler(logging.NullHandler())

# Get module logger
logger = logging.getLogger(__name__)
logger.setLevel(LOG_LEVEL)

# Silence noisy third-party loggers (keep them at WARNING or higher)
for logger_name in SILENT_LOGGERS:
    third_party_logger = logging.getLogger(logger_name)
    third_party_logger.setLevel(logging.WARNING)
    third_party_logger.propagate = False

# Log successful initialization (BEFORE FastMCP import)
if LOG_FILE:
    logger.info("MCP Outlook Server - Starting (Log Level: %s)", logging.getLevelName(LOG_LEVEL))
    logger.debug("Log File: %s", log_dir / log_filename)

# ============================================================================
# MCP SERVER INITIALIZATION
# ============================================================================
# Import FastMCP AFTER logging is configured

from fastmcp import FastMCP

# Import individual tool modules
from tools.email import get_inbox_emails, get_sent_emails, search_emails, send_email, create_draft_email
from tools.folder import list_outlook_folders, search_emails_in_custom_folder, list_outlook_rules
from tools.calendar import get_calendar_events, create_calendar_event, search_calendar_events
from tools.contact import get_contacts, create_contact, search_contacts
from tools.attachment import get_email_attachments, download_email_attachment, send_email_with_attachments

mcp = FastMCP("outlook")
logger.debug("FastMCP server instance created")

# ============================================================================
# EMAIL STYLE LEARNING
# ============================================================================
# Style learning is now dynamic - no startup learning required.
# Styles are learned on-the-fly when sending/creating drafts if AUTO_LEARN_STYLE=true.

# ============================================================================
# TOOL REGISTRATION
# ============================================================================

# Register all tools - each tool is in its own file for maximum granularity
logger.debug("Registering MCP tools...")

# Email tools (5 tools)
get_inbox_emails.register(mcp)
get_sent_emails.register(mcp)
search_emails.register(mcp)
send_email.register(mcp)
create_draft_email.register(mcp)
logger.debug("  [OK] 5 Email tools registered")

# Folder tools (3 tools)
list_outlook_folders.register(mcp)
search_emails_in_custom_folder.register(mcp)
list_outlook_rules.register(mcp)
logger.debug("  [OK] 3 Folder tools registered")

# Calendar tools (3 tools)
get_calendar_events.register(mcp)
create_calendar_event.register(mcp)
search_calendar_events.register(mcp)
logger.debug("  [OK] 3 Calendar tools registered")

# Contact tools (3 tools)
get_contacts.register(mcp)
create_contact.register(mcp)
search_contacts.register(mcp)
logger.debug("  [OK] 3 Contact tools registered")

# Attachment tools (3 tools)
get_email_attachments.register(mcp)
download_email_attachment.register(mcp)
send_email_with_attachments.register(mcp)
logger.debug("  [OK] 3 Attachment tools registered")

# Total: 17 individual tools, each in its own file
logger.info("MCP Server ready - 17 tools registered (5 Email, 3 Folder, 3 Calendar, 3 Contact, 3 Attachment)")

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    """
    Main entry point for the MCP Outlook server.
    
    Starts the FastMCP server which listens for requests from MCP clients
    (such as Claude Desktop, Cursor, or other AI assistants).
    
    The server exposes all registered tools as callable functions
    that clients can invoke to interact with Outlook.
    
    Usage:
        python outlook_mcp.py
        
    Notes:
        - Server runs indefinitely until interrupted (Ctrl+C)
        - Requires Microsoft Outlook to be installed and configured
        - All operations use the currently logged-in Outlook profile
        - Ultra-modular: 14 tools, each in its own file
    """
    try:
        # NOTE: NO print() statements here! MCP protocol uses stdout for JSON communication.
        # All output must go to the log file via logger.info/debug/error
        
        # Run the MCP server (blocks until interrupted)
        logger.debug("MCP server starting (blocking mode)...")
        mcp.run()
        logger.info("MCP server stopped")
    except KeyboardInterrupt:
        logger.info("Server stopped by user (KeyboardInterrupt)")
        pass
    except Exception as e:
        logger.critical("Server crashed unexpectedly", exc_info=True)
        raise
