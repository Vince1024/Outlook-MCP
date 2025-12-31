"""Get Contacts Tool"""
import hashlib, json, logging
from typing import Optional
from config import DEFAULT_CONTACT_LIMIT, MAX_CONTACT_LIMIT, OUTLOOK_FOLDER_CONTACTS
from utils import get_outlook_application, format_contact

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def get_contacts(limit: int = DEFAULT_CONTACT_LIMIT, search_name: Optional[str] = None) -> str:
        """Get contacts from the Outlook Contacts folder."""
        # Log operation start
        logger.debug("Starting get_contacts operation", extra={
            "operation": "get_contacts",
            "limit": limit,
            "has_search_filter": bool(search_name),
            "search_name_length": len(search_name) if search_name else 0
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            contacts_folder = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CONTACTS)
            limit = min(limit, MAX_CONTACT_LIMIT)
            items = contacts_folder.Items
            items.Sort("[FullName]")
            contacts = []
            max_scan = limit * 3 if search_name else limit
            contact = items.GetFirst()
            scanned = 0
            while contact is not None and len(contacts) < limit and scanned < max_scan:
                try:
                    scanned += 1
                    if search_name:
                        full_name = contact.FullName.lower() if contact.FullName else ""
                        if search_name.lower() not in full_name:
                            contact = items.GetNext()
                            continue
                    contacts.append(format_contact(contact))
                except Exception:
                    pass
                contact = items.GetNext()
            
            # Log success (search name length only for privacy)
            logger.info("Retrieved contacts successfully", extra={
                "contact_count": len(contacts),
                "limit": limit,
                "has_search_filter": bool(search_name),
                "search_name_length": len(search_name) if search_name else 0
            })
            
            return json.dumps({"success": True, "count": len(contacts), "contacts": contacts}, indent=2)
        except Exception as e:
            # Hash search_name for privacy (only first 8 chars of hash)
            search_hash = hashlib.sha256(search_name.encode()).hexdigest()[:8] if search_name else None
            logger.error("Failed to get contacts", exc_info=True, extra={
                "limit": limit,
                "search_hash": search_hash,
                "search_name_length": len(search_name) if search_name else 0
            })
            return json.dumps({"success": False, "error": str(e)})

