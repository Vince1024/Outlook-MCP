"""
Download Email Attachment Tool

Downloads a specific attachment from an email to disk.
"""

import json
import logging
from pathlib import Path
from utils import get_outlook_application

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def download_email_attachment(entry_id: str, attachment_index: int, save_path: str) -> str:
        """
        Download a specific attachment from an email to disk.
        
        Downloads an attachment from an email and saves it to the specified location.
        Parent directories are created automatically if they don't exist. Existing
        files are overwritten without confirmation.
        
        Args:
            entry_id: The EntryID of the email (from get_inbox_emails, search_emails, etc.)
            attachment_index: Index of the attachment to download (1-based, from get_email_attachments)
            save_path: Full path where to save the file (e.g., "C:/Users/user/Downloads/report.pdf")
        
        Returns:
            JSON string with structure:
            {
              "success": true,
              "message": "Attachment 'report.pdf' downloaded successfully",
              "saved_path": "C:/Downloads/report.pdf",
              "filename": "report.pdf",
              "size": 245678
            }
            
        Examples:
            >>> # Download first attachment
            >>> download_email_attachment(
            ...     entry_id="00000000...",
            ...     attachment_index=1,
            ...     save_path="C:/Users/user/Downloads/report.pdf"
            ... )
            
            >>> # Download to relative path
            >>> download_email_attachment(
            ...     entry_id="00000000...",
            ...     attachment_index=2,
            ...     save_path="./downloads/data.xlsx"
            ... )
        
        Notes:
            - Parent directories are created automatically
            - Existing files are overwritten without confirmation
            - Index starts at 1 (not 0) - use get_email_attachments to get valid indices
            - Large files may take time to download
            
        Security Notes:
            - File content is not logged
            - Only file path and size are logged (GDPR-compliant)
            - User should verify file safety before opening
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
            
            # Validate attachment index
            if attachment_index < 1 or attachment_index > attachment_count:
                return json.dumps({
                    "success": False,
                    "error": f"Invalid attachment index {attachment_index}. Email has {attachment_count} attachment(s). Index must be between 1 and {attachment_count}."
                })
            
            # Get the specific attachment
            try:
                attachment = attachments_collection.Item(attachment_index)
            except Exception as e:
                logger.error(f"Failed to get attachment {attachment_index}: {str(e)}")
                return json.dumps({
                    "success": False,
                    "error": f"Failed to get attachment at index {attachment_index}: {str(e)}"
                })
            
            # Get attachment details
            filename = attachment.FileName
            size = attachment.Size
            
            # Convert to Path object and resolve to absolute path
            save_path_obj = Path(save_path).resolve()
            
            # Create parent directories if they don't exist
            save_path_obj.parent.mkdir(parents=True, exist_ok=True)
            
            # Save attachment
            try:
                attachment.SaveAsFile(str(save_path_obj))
                logger.info(f"Attachment '{filename}' saved to: {save_path_obj}")
            except Exception as e:
                logger.error(f"Failed to save attachment: {str(e)}")
                return json.dumps({
                    "success": False,
                    "error": f"Failed to save attachment '{filename}': {str(e)}"
                })
            
            # Verify file was created
            if not save_path_obj.exists():
                return json.dumps({
                    "success": False,
                    "error": f"Attachment '{filename}' was not saved to {save_path_obj}. Check file path and permissions."
                })
            
            return json.dumps({
                "success": True,
                "message": f"Attachment '{filename}' downloaded successfully",
                "saved_path": str(save_path_obj),
                "filename": filename,
                "size": size
            }, indent=2)
            
        except Exception as e:
            logger.error(f"Error downloading attachment: {str(e)}", exc_info=True)
            return json.dumps({
                "success": False,
                "error": f"Error downloading attachment: {str(e)}"
            })
