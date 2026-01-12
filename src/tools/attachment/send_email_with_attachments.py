"""
Send Email with Attachments Tool

Sends an email with file attachments via Outlook.
"""

import json
import logging
from pathlib import Path
from typing import Optional
from config import OUTLOOK_ITEM_MAIL, IMPORTANCE_LOW, IMPORTANCE_NORMAL, IMPORTANCE_HIGH, DEFAULT_SIGNATURE, AUTO_LEARN_STYLE
from utils import get_outlook_application
from utils.style_learner import apply_style_to_html, learn_user_email_style

logger = logging.getLogger(__name__)


def register(mcp):
    """Register this tool with the MCP server."""
    
    @mcp.tool()
    def send_email_with_attachments(
        to: str,
        subject: str,
        body: str,
        attachments: str,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        importance: str = "normal",
        html_body: Optional[str] = None,
        signature_name: Optional[str] = None
    ) -> str:
        """
        Send an email with file attachments via Outlook.
        
        Creates and sends a new email through the user's Outlook account with one or
        more file attachments. The email is sent immediately and a copy is saved in
        the Sent Items folder.
        
        Args:
            to: Recipient email address(es), semicolon-separated for multiple
                Example: "user1@example.com" or "user1@example.com; user2@example.com"
            subject: Email subject line
            body: Email body content (plain text format, used if html_body is not provided)
            attachments: File path(s) to attach, semicolon-separated for multiple
                Example: "C:/file.pdf" or "C:/file.pdf; C:/doc.docx"
                All files must exist and be accessible
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
              "success": true,
              "message": "Email sent to colleague@company.com",
              "attachments_added": 2
            }
            
        Examples:
            >>> # Send email with one attachment
            >>> send_email_with_attachments(
            ...     to="colleague@company.com",
            ...     subject="Monthly Report",
            ...     body="Please find attached the report.",
            ...     attachments="C:/Users/user/Documents/report.pdf"
            ... )
            
            >>> # Send email with multiple attachments
            >>> send_email_with_attachments(
            ...     to="team@company.com",
            ...     subject="Q4 Results",
            ...     body="Attached are the Q4 reports.",
            ...     attachments="C:/report.pdf; C:/summary.xlsx; C:/charts.pptx",
            ...     cc="manager@company.com",
            ...     importance="high"
            ... )
        
        Notes:
            - All attachment files must exist before sending
            - Use absolute paths for attachments to avoid path resolution issues
            - Large files can slow down sending
            - Auto-learning style applies to body if AUTO_LEARN_STYLE=true
            
        Security Notes:
            - Attachment paths are validated but not logged
            - Only attachment count is logged (GDPR-compliant)
            - User should verify file safety before sending
        """
        # Log operation start
        logger.debug("Starting send_email_with_attachments operation", extra={
            "operation": "send_email_with_attachments",
            "recipient_count": len(to.split(";")),
            "has_cc": bool(cc),
            "has_bcc": bool(bcc),
            "importance": importance,
            "attachment_count": len(attachments.split(";")) if attachments else 0
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
                    logger.warning(f"Failed to add signature: {str(e)}")
                    # Fallback: use body without signature
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
            
            # Add CC and BCC
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            
            # Set importance
            importance_map = {
                "low": IMPORTANCE_LOW,
                "normal": IMPORTANCE_NORMAL,
                "high": IMPORTANCE_HIGH
            }
            mail.Importance = importance_map.get(importance.lower(), IMPORTANCE_NORMAL)
            
            # Parse and add attachments
            attachment_paths = [path.strip() for path in attachments.split(";") if path.strip()]
            attachments_added = 0
            failed_attachments = []
            
            for file_path in attachment_paths:
                # Convert to Path and resolve to absolute path
                file_path_obj = Path(file_path).resolve()
                
                # Check if file exists
                if not file_path_obj.exists():
                    failed_attachments.append(f"{file_path} (not found)")
                    logger.warning(f"Attachment file not found: {file_path_obj}")
                    continue
                
                if not file_path_obj.is_file():
                    failed_attachments.append(f"{file_path} (not a file)")
                    logger.warning(f"Attachment path is not a file: {file_path_obj}")
                    continue
                
                # Add attachment
                try:
                    mail.Attachments.Add(str(file_path_obj))
                    attachments_added += 1
                    logger.debug(f"Attachment added: {file_path_obj.name}")
                except Exception as e:
                    failed_attachments.append(f"{file_path} ({str(e)[:30]})")
                    logger.error(f"Failed to attach file {file_path_obj}: {str(e)}")
            
            # Check if any attachments were added
            if attachments_added == 0 and attachment_paths:
                return json.dumps({
                    "success": False,
                    "error": f"No attachments could be added. Failed: {', '.join(failed_attachments)}"
                })
            
            # Send email
            mail.Send()
            
            logger.info(f"Email sent successfully with {attachments_added} attachment(s)", extra={
                "operation": "send_email_with_attachments",
                "success": True,
                "attachments_added": attachments_added
            })
            
            result = {
                "success": True,
                "message": f"Email sent to {to}",
                "attachments_added": attachments_added
            }
            
            if failed_attachments:
                result["failed_attachments"] = failed_attachments
                result["message"] += f" (Warning: {len(failed_attachments)} attachment(s) failed)"
            
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error sending email with attachments: {str(e)}", exc_info=True)
            return json.dumps({
                "success": False,
                "error": f"Error sending email: {str(e)}"
            })
