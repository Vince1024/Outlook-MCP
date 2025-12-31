"""
List Outlook Folders Tool

Lists all available Outlook folders with their paths.
"""

import json
import logging
from config import EXCLUDED_STORES
from utils import get_outlook_application, get_all_folders

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def list_outlook_folders() -> str:
        """
        List all available Outlook folders with their paths (FAST - no item counts).
        
        Retrieves a hierarchical list of all mail folders in Outlook, including
        default folders (Inbox, Sent, etc.) and custom user-created folders.
        Useful for discovering folder names before searching.
        
        Returns:
            JSON string with structure:
            {
                "success": bool,
                "count": int,
                "folders": [
                    {
                        "name": str,
                        "path": str
                    }
                ]
            }
            
        Examples:
            >>> list_outlook_folders()
            {
                "success": true,
                "count": 25,
                "folders": [
                    {"name": "Inbox", "path": "Inbox"},
                    {"name": "Archive", "path": "Inbox/Archive"},
                    {"name": "Personal", "path": "Personal"},
                    {"name": "My Mails", "path": "Personal/My Mails"}
                ]
            }
            
        Notes:
            - Performance optimization: Does NOT include item_count/unread_count by default
            - These counts can take several minutes on large mailboxes
            - Includes all folders recursively (nested folders)
            - System folders that can't be accessed are skipped
            - Useful to find the exact folder name/path for searching
            - Returns in seconds instead of minutes!
        """
        # Log operation start
        logger.debug("Starting list_outlook_folders operation", extra={
            "operation": "list_outlook_folders"
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            
            all_folders = []
            
            for store in namespace.Stores:
                try:
                    if store.DisplayName in EXCLUDED_STORES:
                        continue
                    
                    root_folder = store.GetRootFolder()
                    store_folders = get_all_folders(root_folder, include_counts=False)
                    all_folders.extend(store_folders)
                except Exception:
                    pass
            
            # Log success
            logger.info("Listed Outlook folders successfully", extra={
                "folder_count": len(all_folders),
                "stores_scanned": len(namespace.Stores)
            })
            
            return json.dumps({
                "success": True,
                "count": len(all_folders),
                "folders": all_folders
            }, indent=2)
            
        except Exception as e:
            logger.error("Failed to list Outlook folders", exc_info=True)
            return json.dumps({"success": False, "error": str(e)})

