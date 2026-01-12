"""
Formatters Module

Contains functions to format Outlook items (emails, appointments, contacts)
into JSON-serializable dictionaries.
"""

import logging
from typing import Dict, Any
from config import EMAIL_BODY_PREVIEW_LENGTH

logger = logging.getLogger(__name__)


def count_real_attachments(mail_item) -> int:
    """
    Count only real file attachments, excluding inline images and embedded items.
    
    Args:
        mail_item: Outlook MailItem COM object
        
    Returns:
        int: Number of real file attachments
        
    Notes:
        Filters out:
        - Embedded/inline items (Type != 1)
        - Small images likely to be signatures/logos (< 5KB and common image extensions)
        - Attachments with ContentID (inline images referenced in HTML)
    """
    try:
        real_count = 0
        
        for i in range(1, mail_item.Attachments.Count + 1):
            try:
                att = mail_item.Attachments.Item(i)
                
                # Type 1 = olByValue (regular file attachment)
                # Type 5 = olEmbeddeditem, Type 6 = olOLE
                if att.Type != 1:
                    continue
                
                filename = att.FileName.lower()
                size = att.Size
                
                # Check if it's an inline image (has ContentID)
                try:
                    # Try to access ContentID property (inline images have this)
                    prop_accessor = att.PropertyAccessor
                    # PR_ATTACH_CONTENT_ID = http://schemas.microsoft.com/mapi/proptag/0x3712001F
                    content_id = prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                    if content_id:
                        # Has ContentID = inline image
                        continue
                except:
                    # No ContentID property = likely a real attachment
                    pass
                
                # Filter small images (likely signatures/logos)
                image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp')
                if filename.endswith(image_extensions) and size < 5120:  # < 5KB
                    continue
                
                # This is likely a real attachment
                real_count += 1
                
            except Exception as e:
                # If we can't determine, count it (be conservative)
                logger.debug(f"Error checking attachment: {e}")
                real_count += 1
        
        return real_count
        
    except Exception as e:
        logger.error(f"Failed to count real attachments: {e}")
        # Fallback to total count if filtering fails
        return mail_item.Attachments.Count


def format_email(mail_item) -> Dict[str, Any]:
    """
    Format an Outlook mail item as a dictionary for JSON serialization.
    
    Args:
        mail_item: Outlook MailItem COM object
        
    Returns:
        Dict[str, Any]: Dictionary containing email properties
        
    Notes:
        - Email body is truncated to EMAIL_BODY_PREVIEW_LENGTH characters to prevent
          excessive data transfer and potential memory issues
        - Returns an error dict if formatting fails to allow graceful degradation
    """
    try:
        # Truncate body to prevent excessive data exposure
        email_body = mail_item.Body if mail_item.Body else ""
        truncated_body = email_body[:EMAIL_BODY_PREVIEW_LENGTH] + "..." \
                        if len(email_body) > EMAIL_BODY_PREVIEW_LENGTH else email_body
        
        # Get folder path
        folder_path = None
        try:
            if hasattr(mail_item, 'Parent') and mail_item.Parent:
                folder_path = mail_item.Parent.FolderPath
        except Exception:
            pass
        
        # Count real attachments (excluding inline images)
        real_attachments_count = count_real_attachments(mail_item)
        
        return {
            "subject": mail_item.Subject,
            "sender": mail_item.SenderName,
            "sender_email": mail_item.SenderEmailAddress,
            "recipients": mail_item.To,
            "cc": mail_item.CC,
            "bcc": mail_item.BCC,
            "received_time": str(mail_item.ReceivedTime) if hasattr(mail_item, 'ReceivedTime') else None,
            "sent_on": str(mail_item.SentOn) if hasattr(mail_item, 'SentOn') else None,
            "body": truncated_body,
            "body_length": len(email_body),
            "has_attachments": real_attachments_count > 0,
            "attachment_count": real_attachments_count,
            "importance": mail_item.Importance,
            "unread": mail_item.UnRead,
            "categories": mail_item.Categories,
            "folder_path": folder_path,
            "entry_id": mail_item.EntryID,
        }
    except Exception as e:
        logger.error("Failed to format email item", exc_info=True, extra={
            "error_type": type(e).__name__
        })
        return {"error": f"Failed to format email: {e}"}


def format_appointment(appointment) -> Dict[str, Any]:
    """
    Format an Outlook appointment/calendar event as a dictionary for JSON serialization.
    
    Args:
        appointment: Outlook AppointmentItem COM object
        
    Returns:
        Dict[str, Any]: Dictionary containing appointment properties
        
    Notes:
        - Body is truncated for the same security reasons as emails
        - BusyStatus codes: 0=Free, 1=Tentative, 2=Busy, 3=Out of Office
    """
    try:
        # Truncate body to prevent excessive data exposure
        appointment_body = appointment.Body if appointment.Body else ""
        truncated_body = appointment_body[:EMAIL_BODY_PREVIEW_LENGTH] + "..." \
                        if len(appointment_body) > EMAIL_BODY_PREVIEW_LENGTH else appointment_body
        
        return {
            "subject": appointment.Subject,
            "start": str(appointment.Start),
            "end": str(appointment.End),
            "location": appointment.Location,
            "organizer": appointment.Organizer if hasattr(appointment, 'Organizer') else None,
            "required_attendees": appointment.RequiredAttendees,
            "optional_attendees": appointment.OptionalAttendees,
            "body": truncated_body,
            "is_all_day_event": appointment.AllDayEvent,
            "reminder_set": appointment.ReminderSet,
            "reminder_minutes": appointment.ReminderMinutesBeforeStart if appointment.ReminderSet else None,
            "categories": appointment.Categories,
            "busy_status": appointment.BusyStatus,
        }
    except Exception as e:
        logger.error("Failed to format appointment", exc_info=True, extra={
            "error_type": type(e).__name__
        })
        return {"error": f"Failed to format appointment: {e}"}


def format_contact(contact) -> Dict[str, Any]:
    """
    Format an Outlook contact as a dictionary for JSON serialization.
    
    Args:
        contact: Outlook ContactItem COM object
        
    Returns:
        Dict[str, Any]: Dictionary containing contact properties
        
    Notes:
        - Uses safe_get helper to handle missing or null properties gracefully
        - Some Outlook contacts may have incomplete data, this ensures robust handling
    """
    try:
        # Safely get attributes with fallback to empty string
        def safe_get(obj, attr, default=""):
            """
            Safely retrieve an attribute from a COM object.
            
            Args:
                obj: COM object to retrieve attribute from
                attr: Attribute name to retrieve
                default: Default value if attribute is missing or None
                
            Returns:
                Attribute value or default
            """
            try:
                value = getattr(obj, attr, default)
                return value if value is not None else default
            except Exception:
                return default
        
        return {
            "full_name": safe_get(contact, "FullName"),
            "email1": safe_get(contact, "Email1Address"),
            "email2": safe_get(contact, "Email2Address"),
            "email3": safe_get(contact, "Email3Address"),
            "company": safe_get(contact, "CompanyName"),
            "job_title": safe_get(contact, "JobTitle"),
            "business_phone": safe_get(contact, "BusinessTelephoneNumber"),
            "mobile_phone": safe_get(contact, "MobileTelephoneNumber"),
            "home_phone": safe_get(contact, "HomeTelephoneNumber"),
            "business_address": safe_get(contact, "BusinessAddress"),
            "categories": safe_get(contact, "Categories"),
        }
    except Exception as e:
        logger.error("Failed to format contact", exc_info=True, extra={
            "error_type": type(e).__name__
        })
        return {"error": f"Failed to format contact: {e}"}

