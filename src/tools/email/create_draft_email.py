"""
Create Draft Email Tool

Creates a draft email in Outlook without sending it.
"""

import json
import logging
from typing import Optional
from config import OUTLOOK_ITEM_MAIL
from utils import get_outlook_application, get_outlook_signature

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def create_draft_email(
        to: str,
        subject: str,
        body: str,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        html_body: Optional[str] = None,
        signature_name: Optional[str] = None
    ) -> str:
        """
        Create a draft email in Outlook without sending it.
        
        Creates an email and saves it to the Drafts folder where the user can
        review, edit, and send it later. Useful for preparing emails that need
        review before sending.
        
        Args:
            to: Recipient email address(es), semicolon-separated for multiple
            subject: Email subject line
            body: Email body content (plain text format, used if html_body and signature_name are not provided)
            cc: CC recipients (optional), semicolon-separated
            bcc: BCC recipients (optional), semicolon-separated
            html_body: HTML body content (optional). If provided, this will be used instead of body
            signature_name: Name of Outlook signature to use (optional). If provided, loads signature from Outlook
                Example: "Work Signature" will load "Work Signature (user@company.com).htm" from signatures folder
        
        Returns:
            JSON string with structure:
            {
                "success": bool,
                "message": str
            }
            
        Examples:
            >>> create_draft_email("manager@company.com", "Report", "Draft report...")
            {"success": true, "message": "Draft email created"}
            
            >>> create_draft_email("user@company.com", "Hello", "Message", signature_name="VP DXT")
            {"success": true, "message": "Draft email created"}
            
        Notes:
            - Draft is saved in the user's Drafts folder
            - User can find and edit the draft in Outlook
            - No email is sent until the user manually sends it
        """
        # Log operation start
        logger.debug("Starting create_draft_email operation", extra={
            "operation": "create_draft_email",
            "recipient_count": len(to.split(";")),
            "has_cc": bool(cc),
            "has_bcc": bool(bcc)
        })
        
        try:
            outlook = get_outlook_application()
            mail = outlook.CreateItem(OUTLOOK_ITEM_MAIL)
            
            mail.To = to
            mail.Subject = subject
            
            # Determine the email body
            if signature_name:
                try:
                    mail.Display(False)
                    signature_html = mail.HTMLBody
                    
                    if html_body:
                        mail.HTMLBody = html_body + signature_html
                    else:
                        body_html = body.replace('\n', '<br>')
                        mail.HTMLBody = f"<html><body><p>{body_html}</p>{signature_html}</body></html>"
                        
                except Exception:
                    signature_html = get_outlook_signature(signature_name)
                    if signature_html:
                        if html_body:
                            mail.HTMLBody = html_body + "<br>" + signature_html
                        else:
                            body_html = body.replace('\n', '<br>')
                            mail.HTMLBody = f"<html><body><p>{body_html}</p><br>{signature_html}</body></html>"
                    else:
                        if html_body:
                            mail.HTMLBody = html_body
                        else:
                            mail.Body = body
            elif html_body:
                mail.HTMLBody = html_body
            else:
                mail.Body = body
            
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            
            mail.Save()
            
            # Log success (no sensitive data)
            logger.info("Draft email created successfully", extra={
                "recipient_count": len(to.split(";")),
                "has_cc": bool(cc),
                "has_bcc": bool(bcc),
                "has_signature": bool(signature_name)
            })
            
            return json.dumps({
                "success": True,
                "message": "Draft email created"
            }, indent=2)
            
        except Exception as e:
            # GDPR compliant: no personal data in logs
            logger.error("Failed to create draft email", exc_info=True, extra={
                "recipient_count": len(to.split(";")) if to else 0,
                "has_subject": bool(subject),
                "has_signature": bool(signature_name)
            })
            return json.dumps({"success": False, "error": str(e)})

