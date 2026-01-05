"""
Send Email Tool

Sends an email via Outlook.
"""

import json
import logging
from typing import Optional
from config import OUTLOOK_ITEM_MAIL, IMPORTANCE_LOW, IMPORTANCE_NORMAL, IMPORTANCE_HIGH, DEFAULT_SIGNATURE, AUTO_LEARN_STYLE
from utils import get_outlook_application, get_outlook_signature
from utils.style_learner import apply_style_to_html, learn_user_email_style

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def send_email(
        to: str,
        subject: str,
        body: str,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        importance: str = "normal",
        html_body: Optional[str] = None,
        signature_name: Optional[str] = None
    ) -> str:
        """
        Send an email via Outlook.
        
        Creates and sends a new email through the user's Outlook account. The email
        is sent immediately and a copy is saved in the Sent Items folder.
        
        Args:
            to: Recipient email address(es), semicolon-separated for multiple
                Example: "user1@example.com" or "user1@example.com; user2@example.com"
            subject: Email subject line
            body: Email body content (plain text format, used if html_body and signature_name are not provided)
            cc: CC recipients (optional), semicolon-separated
            bcc: BCC recipients (optional), semicolon-separated
            importance: Email importance level (low, normal, high) (default: normal)
            html_body: HTML body content (optional). If provided, this will be used instead of body
            signature_name: CURRENTLY UNUSED - Signature is inserted via mail.Display() which always uses 
                Outlook's default signature. This parameter is kept for API compatibility but has no effect.
                To use a specific signature, change Outlook's default signature settings.
        
        Returns:
            JSON string with structure:
            {
                "success": bool,
                "message": str
            }
            
        Examples:
            >>> send_email("colleague@company.com", "Meeting", "See you at 2pm")
            {"success": true, "message": "Email sent to colleague@company.com"}
            
            >>> send_email("team@company.com", "Urgent", "...", importance="high")
            {"success": true, "message": "Email sent to team@company.com"}
            
            >>> send_email("user@company.com", "Hello", "Message", signature_name="VP DXT")
            {"success": true, "message": "Email sent to user@company.com"}
            
        Security Notes:
            - No sensitive data should be included in logs
            - Only counts and boolean flags are logged (GDPR-compliant)
            - Email addresses and content are never logged
        """
        # Log operation start
        logger.debug("Starting send_email operation", extra={
            "operation": "send_email",
            "recipient_count": len(to.split(";")),
            "has_cc": bool(cc),
            "has_bcc": bool(bcc),
            "importance": importance
        })
        
        try:
            outlook = get_outlook_application()
            mail = outlook.CreateItem(OUTLOOK_ITEM_MAIL)
            
            mail.To = to
            mail.Subject = subject
            
            # Learn style from recent sent emails if AUTO_LEARN_STYLE enabled (no caching)
            learned_style = None
            if AUTO_LEARN_STYLE:
                try:
                    learned_style = learn_user_email_style()
                    if learned_style:
                        logger.debug("Style learned from recent sent emails", extra={"style": learned_style})
                except Exception as e:
                    logger.debug("Failed to learn style: %s", e)
            
            # Determine signature to use
            # Priority: signature_name param > DEFAULT_SIGNATURE config > None
            sig_name = signature_name if signature_name is not None else (DEFAULT_SIGNATURE if DEFAULT_SIGNATURE else None)
            
            # Build email body
            if html_body:
                # User provided HTML body - use as is
                final_body = html_body
            else:
                # Plain text body - apply learned style if available
                if learned_style:
                    final_body = apply_style_to_html(body, learned_style)
                else:
                    # No style - plain text
                    final_body = body
            
            # Add signature if requested
            if sig_name:
                try:
                    # Get signature via Display (inserts Outlook's default signature)
                    # Note: This always uses the default signature, not sig_name
                    mail.Display(False)
                    signature_html = mail.HTMLBody
                    
                    # Simple prepend: put our content before the entire signature HTML
                    if isinstance(final_body, str) and '<' in final_body:
                        mail.HTMLBody = final_body + signature_html
                    else:
                        # Plain text to HTML
                        body_html = final_body.replace('\n', '<br>')
                        mail.HTMLBody = f"<p>{body_html}</p>" + signature_html
                        
                except Exception as e:
                    logger.error(f"Failed to add signature: {e}")
                    # Fallback: just use body
                    if isinstance(final_body, str) and '<' in final_body:
                        mail.HTMLBody = final_body
                    else:
                        mail.Body = final_body
            else:
                # No signature requested
                if isinstance(final_body, str) and '<' in final_body:
                    mail.HTMLBody = final_body
                else:
                    mail.Body = final_body
            
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            
            importance_map = {
                "low": IMPORTANCE_LOW,
                "normal": IMPORTANCE_NORMAL,
                "high": IMPORTANCE_HIGH
            }
            mail.Importance = importance_map.get(importance.lower(), IMPORTANCE_NORMAL)
            
            mail.Send()
            
            # Log success (no sensitive data)
            logger.info("Email sent successfully", extra={
                "recipient_count": len(to.split(";")),
                "has_cc": bool(cc),
                "has_bcc": bool(bcc),
                "importance": importance,
                "has_signature": bool(signature_name)
            })
            
            return json.dumps({
                "success": True,
                "message": f"Email sent to {to}"
            }, indent=2)
            
        except Exception as e:
            # GDPR compliant: no personal data in logs
            logger.error("Failed to send email", exc_info=True, extra={
                "recipient_count": len(to.split(";")) if to else 0,
                "has_subject": bool(subject),
                "importance": importance,
                "has_signature": bool(signature_name)
            })
            return json.dumps({"success": False, "error": str(e)})

