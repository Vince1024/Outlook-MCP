"""Create Contact Tool"""
import json, logging
from typing import Optional
from config import OUTLOOK_ITEM_CONTACT
from utils import get_outlook_application

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def create_contact(full_name: str, email: str, company: Optional[str] = None, job_title: Optional[str] = None, business_phone: Optional[str] = None, mobile_phone: Optional[str] = None, home_phone: Optional[str] = None) -> str:
        """Create a new contact in the Outlook Contacts folder."""
        # Log operation start (no personal data)
        logger.debug("Starting create_contact operation", extra={
            "operation": "create_contact",
            "has_company": bool(company),
            "has_phone": bool(business_phone or mobile_phone or home_phone)
        })
        
        try:
            outlook = get_outlook_application()
            contact = outlook.CreateItem(OUTLOOK_ITEM_CONTACT)
            contact.FullName = full_name
            contact.Email1Address = email
            if company:
                contact.CompanyName = company
            if job_title:
                contact.JobTitle = job_title
            if business_phone:
                contact.BusinessTelephoneNumber = business_phone
            if mobile_phone:
                contact.MobileTelephoneNumber = mobile_phone
            if home_phone:
                contact.HomeTelephoneNumber = home_phone
            contact.Save()
            
            # Log success (no sensitive data)
            logger.info("Contact created successfully", extra={
                "has_name": bool(full_name),
                "has_email": bool(email),
                "email_domain": email.split("@")[1] if email and "@" in email else "unknown",
                "has_company": bool(company),
                "has_phone": bool(business_phone or mobile_phone or home_phone)
            })
            
            return json.dumps({"success": True, "message": f"Contact '{full_name}' created"}, indent=2)
        except Exception as e:
            # GDPR compliant: no personal data in logs
            logger.error("Failed to create contact", exc_info=True, extra={
                "has_name": bool(full_name),
                "has_email": bool(email),
                "email_domain": email.split("@")[1] if email and "@" in email else "unknown",
                "has_company": bool(company)
            })
            return json.dumps({"success": False, "error": str(e)})

