"""
Get Email Attachments Tool

Lists all attachments from a specific email.
"""

import json
import logging
from utils import get_outlook_application

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def get_email_attachments(entry_id: str) -> str:
        """
        Get list of attachments from a specific email.
        
        Retrieves all attachments from an email, including their filenames, sizes,
        types, and indices. This information is needed to download specific attachments
        using the download_email_attachment tool.
        
        Args:
            entry_id: The EntryID of the email (obtained from get_inbox_emails, search_emails, etc.)
        
        Returns:
            JSON string with structure:
            {
              "success": true,
              "count": 2,
              "attachments": [
                {
                  "filename": "report.pdf",
                  "size": 245678,
                  "type": 1,
                  "index": 1
                }
              ]
            }
            
        Attachment Types:
            - type: 1 - Standard file (olByValue)
            - type: 5 - Embedded Outlook item (olEmbeddeditem)
            - type: 6 - OLE object (olOLE)
            
        Examples:
            >>> # Get entry_id from get_inbox_emails
            >>> emails = get_inbox_emails(limit=1)
            >>> entry_id = emails["emails"][0]["entry_id"]
            >>> 
            >>> # List attachments
            >>> attachments = get_email_attachments(entry_id)
        
        Security Notes:
            - Only returns metadata (no file content)
            - File sizes and names are logged for GDPR compliance
            - No sensitive data is exposed
        """
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            
            # Get email by EntryID
            try:
                mail = namespace.GetItemFromID(entry_id)
            except Exception as e:
                logger.error(f"Failed to get item from EntryID: {str(e)}")
                return json.dumps({
                    "success": False,
                    "error": "Failed to get item from EntryID. The email may have been deleted or the ID is invalid."
                })
            
            # Get attachments collection
            attachments_collection = mail.Attachments
            attachment_count = attachments_collection.Count
            
            logger.debug(f"Email has {attachment_count} attachment(s)")
            
            # Build attachment list
            attachments = []
            for i in range(1, attachment_count + 1):  # COM indices start at 1
                try:
                    attachment = attachments_collection.Item(i)
                    
                    att_info = {
                        "filename": attachment.FileName,
                        "size": attachment.Size,
                        "type": attachment.Type,
                        "index": i
                    }
                    
                    attachments.append(att_info)
                    logger.debug(f"Attachment {i}: {attachment.FileName} ({attachment.Size} bytes)")
                    
                except Exception as e:
                    logger.warning(f"Failed to get attachment {i}: {str(e)}")
                    # Continue with other attachments
                    continue
            
            return json.dumps({
                "success": True,
                "count": len(attachments),
                "attachments": attachments
            }, indent=2)
            
        except Exception as e:
            logger.error(f"Error getting email attachments: {str(e)}", exc_info=True)
            return json.dumps({
                "success": False,
                "error": f"Error getting email attachments: {str(e)}"
            })
