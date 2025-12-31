"""
Get Sent Emails Tool

Retrieves emails from the Outlook Sent Items folder.
"""

import json
import logging
from config import DEFAULT_EMAIL_LIMIT, MAX_EMAIL_LIMIT, OUTLOOK_FOLDER_SENT
from utils import get_outlook_application, format_email

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def get_sent_emails(limit: int = DEFAULT_EMAIL_LIMIT) -> str:
        """
        Get emails from the Outlook Sent Items folder.
        
        Retrieves emails that the user has sent, sorted by send time (most recent first).
        Useful for reviewing sent correspondence or finding previously sent information.
        
        Args:
            limit: Maximum number of emails to return (default: 10, max: 100)
        
        Returns:
            JSON string with structure:
            {
                "success": bool,
                "count": int,
                "emails": [list of email dictionaries]
            }
            
        Examples:
            >>> get_sent_emails(limit=5)
            {"success": true, "count": 5, "emails": [...]}
            
        Notes:
            - Limited to MAX_EMAIL_LIMIT (50) for performance
            - Sorted by SentOn date in descending order
        """
        # Log operation start
        logger.debug("Starting get_sent_emails operation", extra={
            "operation": "get_sent_emails",
            "limit": limit
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            sent_folder = namespace.GetDefaultFolder(OUTLOOK_FOLDER_SENT)
            
            limit = min(limit, MAX_EMAIL_LIMIT)
            
            emails = []
            items = sent_folder.Items
            items.Sort("[SentOn]", True)
            
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
            logger.info("Retrieved sent emails successfully", extra={
                "count": len(emails),
                "limit": limit
            })
            
            return json.dumps({
                "success": True,
                "count": len(emails),
                "emails": emails
            }, indent=2)
            
        except Exception as e:
            logger.error("Failed to get sent emails", exc_info=True, extra={
                "limit": limit
            })
            return json.dumps({"success": False, "error": str(e)})

