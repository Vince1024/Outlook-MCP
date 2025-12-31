"""Search Contacts Tool"""
import hashlib, json, logging
from config import OUTLOOK_FOLDER_CONTACTS
from utils import get_outlook_application, format_contact

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def search_contacts(query: str) -> str:
        """Search for contacts by keyword in name, email, or company."""
        # Log operation start
        logger.debug("Starting search_contacts operation", extra={
            "operation": "search_contacts",
            "query_length": len(query)
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            contacts_folder = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CONTACTS)
            items = contacts_folder.Items
            contacts = []
            query_lower = query.lower()
            for contact in items:
                try:
                    full_name = contact.FullName.lower() if contact.FullName else ""
                except Exception:
                    full_name = ""
                try:
                    email = contact.Email1Address.lower() if contact.Email1Address else ""
                except Exception:
                    email = ""
                try:
                    company = contact.CompanyName.lower() if contact.CompanyName else ""
                except Exception:
                    company = ""
                if query_lower in full_name or query_lower in email or query_lower in company:
                    formatted = format_contact(contact)
                    if "error" not in formatted:
                        contacts.append(formatted)
            
            # Log success (query length only for privacy)
            logger.info("Contact search completed successfully", extra={
                "result_count": len(contacts),
                "query_length": len(query)
            })
            
            return json.dumps({"success": True, "query": query, "count": len(contacts), "contacts": contacts}, indent=2)
        except Exception as e:
            # Hash query for privacy (only first 8 chars of hash)
            query_hash = hashlib.sha256(query.encode()).hexdigest()[:8]
            logger.error("Failed to search contacts", exc_info=True, extra={
                "query_hash": query_hash,
                "query_length": len(query)
            })
            return json.dumps({"success": False, "error": str(e)})

