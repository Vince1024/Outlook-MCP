"""Search Emails in Custom Folder Tool"""
import hashlib, json, logging
from datetime import datetime, timedelta
from typing import Optional
from config import DEFAULT_DAYS_BACK, MAX_EMAIL_LIMIT
from utils import get_outlook_application, format_email, get_folder_by_path

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def search_emails_in_custom_folder(folder_path: str, query: Optional[str] = None, limit: int = 20, days_back: int = DEFAULT_DAYS_BACK, recursive: bool = False) -> str:
        """
        Search for emails in a specific custom Outlook folder.
        
        Args:
            folder_path: Path to the folder (e.g., "Vincent PAPUCHON (PERSO)/My Mails")
            query: Search query (optional, searches in subject/body/sender)
            limit: Maximum number of results (default: 20, max: 50)
            days_back: Number of days to search back (default: 2, 0 for all)
            recursive: If True, searches in all subfolders recursively (default: False)
        
        Returns:
            JSON with success status, folder path, email count, and matching emails
        """
        # Log operation start
        logger.debug("Starting search_emails_in_custom_folder operation", extra={
            "operation": "search_emails_in_custom_folder",
            "has_query": bool(query),
            "query_length": len(query) if query else 0,
            "limit": limit,
            "days_back": days_back,
            "recursive": recursive
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            limit = min(limit, MAX_EMAIL_LIMIT)
            target_folder = get_folder_by_path(namespace, folder_path, use_cache=True)
            if target_folder is None:
                return json.dumps({"success": False, "error": f"Folder '{folder_path}' not found. Use list_outlook_folders() to see available folders."})
            
            # Build list of folders to search
            folders_to_search = [target_folder]
            if recursive:
                # Recursively collect all subfolder objects
                def collect_subfolders(folder, max_depth=50, current_depth=0):
                    """Recursively collect all subfolder objects"""
                    subfolders = []
                    if current_depth >= max_depth:
                        return subfolders
                    try:
                        for subfolder in folder.Folders:
                            try:
                                subfolders.append(subfolder)
                                # Recursively add nested subfolders
                                subfolders.extend(collect_subfolders(subfolder, max_depth, current_depth + 1))
                            except Exception:
                                continue
                    except Exception:
                        pass
                    return subfolders
                
                folders_to_search.extend(collect_subfolders(target_folder))
            
            emails = []
            folders_searched = 0
            
            # Search in each folder
            for folder in folders_to_search:
                if len(emails) >= limit:
                    break
                
                folders_searched += 1
                try:
                    items = folder.Items
                    if days_back > 0:
                        start_date = datetime.now() - timedelta(days=days_back)
                        filter_str = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y')}'"
                        items = items.Restrict(filter_str)
                    items.Sort("[ReceivedTime]", True)
                    
                    max_index = limit * 5 if query else limit
                    if query:
                        query_lower = query.lower()
                        for i in range(max_index):
                            if len(emails) >= limit:
                                break
                            try:
                                mail = items[i + 1]
                                subject = mail.Subject.lower() if mail.Subject else ""
                                body = mail.Body.lower() if mail.Body else ""
                                sender = mail.SenderName.lower() if mail.SenderName else ""
                                if query_lower in subject or query_lower in body or query_lower in sender:
                                    emails.append(format_email(mail))
                            except Exception:
                                break
                    else:
                        for i in range(limit - len(emails)):
                            try:
                                mail = items[i + 1]
                                emails.append(format_email(mail))
                            except Exception:
                                break
                except Exception:
                    continue
            result = {"success": True, "folder": folder_path, "count": len(emails), "emails": emails}
            if query:
                result["query"] = query
            if days_back > 0:
                result["days_back"] = days_back
            if recursive:
                result["folders_searched"] = folders_searched
                info_parts = []
                if days_back > 0:
                    info_parts.append(f"emails from last {days_back} days")
                info_parts.append(f"{folders_searched} folder(s) searched recursively")
                result["info"] = "Searched " + " in ".join(info_parts)
            elif days_back > 0:
                result["info"] = f"Searched emails from last {days_back} days only"
            
            # Log success (folder path only, not query content)
            logger.info("Searched custom folder successfully", extra={
                "result_count": len(emails),
                "has_query": bool(query),
                "query_length": len(query) if query else 0,
                "days_back": days_back,
                "limit": limit,
                "recursive": recursive,
                "folders_searched": folders_searched if recursive else 1
            })
            
            return json.dumps(result, indent=2)
        except Exception as e:
            # Hash query for privacy (only first 8 chars of hash)
            query_hash = hashlib.sha256(query.encode()).hexdigest()[:8] if query else None
            logger.error("Failed to search in custom folder", exc_info=True, extra={
                "query_hash": query_hash,
                "query_length": len(query) if query else 0,
                "limit": limit,
                "days_back": days_back
            })
            return json.dumps({"success": False, "error": str(e)})

