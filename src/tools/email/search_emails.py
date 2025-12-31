"""
Search Emails Tool

Searches for emails in Outlook folders using keyword matching.
"""

import hashlib
import json
import logging
from config import (
    MAX_EMAIL_LIMIT,
    OUTLOOK_FOLDER_INBOX, OUTLOOK_FOLDER_SENT, OUTLOOK_FOLDER_DRAFTS, OUTLOOK_FOLDER_DELETED
)
from utils import get_outlook_application, format_email

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def search_emails(query: str, folder: str = "inbox", limit: int = 20) -> str:
        """
        Search for emails in Outlook folders using keyword matching.
        
        Searches across subject, body, and sender name fields. Can search in a specific
        folder or across all mail folders.
        
        Args:
            query: Search query (searches in subject, body, sender)
            folder: Folder to search in (inbox, sent, drafts, deleted, all) (default: inbox)
            limit: Maximum number of results (default: 20, max: 100)
        
        Returns:
            JSON string with structure:
            {
                "success": bool,
                "query": str,
                "count": int,
                "emails": [list of matching email dictionaries]
            }
            
        Examples:
            >>> search_emails("payment", folder="inbox", limit=10)
            {"success": true, "query": "payment", "count": 5, "emails": [...]}
            
            >>> search_emails("project update", folder="all", limit=20)
            {"success": true, "query": "project update", "count": 15, "emails": [...]}
            
        Notes:
            - Uses Outlook's SQL-like filter syntax for efficient searching
            - Limited to MAX_EMAIL_LIMIT (50) for performance
            - When folder="all", searches inbox, sent, and drafts folders
        """
        # Log operation start
        logger.debug("Starting search_emails operation", extra={
            "operation": "search_emails",
            "query_length": len(query),
            "folder": folder,
            "limit": limit
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            
            folder_map = {
                "inbox": OUTLOOK_FOLDER_INBOX,
                "sent": OUTLOOK_FOLDER_SENT,
                "drafts": OUTLOOK_FOLDER_DRAFTS,
                "deleted": OUTLOOK_FOLDER_DELETED,
            }
            
            limit = min(limit, MAX_EMAIL_LIMIT)
            emails = []
            
            if folder == "all":
                folders_to_search = [OUTLOOK_FOLDER_INBOX, OUTLOOK_FOLDER_SENT, OUTLOOK_FOLDER_DRAFTS]
            else:
                folder_id = folder_map.get(folder.lower(), OUTLOOK_FOLDER_INBOX)
                folders_to_search = [folder_id]
            
            for folder_id in folders_to_search:
                search_folder = namespace.GetDefaultFolder(folder_id)
                
                filter_str = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{query}%' OR " \
                            f"\"urn:schemas:httpmail:textdescription\" LIKE '%{query}%' OR " \
                            f"\"urn:schemas:httpmail:fromname\" LIKE '%{query}%'\""
                
                items = search_folder.Items.Restrict(filter_str)
                items.Sort("[ReceivedTime]", True)
                
                remaining_limit = limit - len(emails)
                mail = items.GetFirst()
                count = 0
                
                while mail is not None and count < remaining_limit:
                    try:
                        emails.append(format_email(mail))
                        count += 1
                        
                        if len(emails) >= limit:
                            break
                    except Exception:
                        pass
                    
                    mail = items.GetNext()
                
                if len(emails) >= limit:
                    break
            
            # Log success (query length only for privacy)
            logger.info("Email search completed successfully", extra={
                "result_count": len(emails),
                "query_length": len(query),
                "folder": folder,
                "limit": limit
            })
            
            return json.dumps({
                "success": True,
                "query": query,
                "count": len(emails),
                "emails": emails
            }, indent=2)
            
        except Exception as e:
            # Hash query for privacy (only first 8 chars of hash)
            query_hash = hashlib.sha256(query.encode()).hexdigest()[:8]
            logger.error("Failed to search emails", exc_info=True, extra={
                "query_hash": query_hash,
                "query_length": len(query),
                "folder": folder,
                "limit": limit
            })
            return json.dumps({"success": False, "error": str(e)})

