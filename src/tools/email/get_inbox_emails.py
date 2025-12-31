"""
Get Inbox Emails Tool

Retrieves emails from the Outlook Inbox folder.
"""

import json
import logging
from config import DEFAULT_EMAIL_LIMIT, MAX_EMAIL_LIMIT, OUTLOOK_FOLDER_INBOX
from utils import get_outlook_application, format_email

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def get_inbox_emails(limit: int = DEFAULT_EMAIL_LIMIT, unread_only: bool = False) -> str:
        """
        Get emails from the Outlook Inbox folder.
        
        This function retrieves emails from the user's Inbox, sorted by received time
        (most recent first). It can optionally filter to only show unread emails.
        
        Args:
            limit: Maximum number of emails to return (default: 10, max: 100)
            unread_only: If True, only return unread emails (default: False)
        
        Returns:
            JSON string with structure:
            {
                "success": bool,
                "count": int,
                "emails": [list of email dictionaries]
            }
            
        Examples:
            >>> get_inbox_emails(limit=5)
            {"success": true, "count": 5, "emails": [...]}
            
            >>> get_inbox_emails(limit=10, unread_only=True)
            {"success": true, "count": 3, "emails": [...]}
            
        Notes:
            - Limited to MAX_EMAIL_LIMIT (50) to prevent performance issues
            - When unread_only=True, we fetch up to limit*2 items to ensure enough results
        """
        # Log operation start
        logger.debug("Starting get_inbox_emails operation", extra={
            "operation": "get_inbox_emails",
            "limit": limit,
            "unread_only": unread_only
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(OUTLOOK_FOLDER_INBOX)
            
            limit = min(limit, MAX_EMAIL_LIMIT)
            
            emails = []
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            
            if unread_only:
                items = items.Restrict("[Unread] = True")
            
            mail = items.GetFirst()
            count = 0
            
            while mail is not None and count < limit:
                try:
                    emails.append(format_email(mail))
                    count += 1
                except Exception:
                    pass
                
                mail = items.GetNext()
            
            # Log success
            logger.info("Retrieved inbox emails successfully", extra={
                "count": len(emails),
                "limit": limit,
                "unread_only": unread_only
            })
            
            return json.dumps({
                "success": True,
                "count": len(emails),
                "emails": emails
            }, indent=2)
            
        except Exception as e:
            logger.error("Failed to get inbox emails", exc_info=True, extra={
                "limit": limit,
                "unread_only": unread_only
            })
            return json.dumps({"success": False, "error": str(e)})

