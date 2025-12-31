"""
Formatters Module

Contains functions to format Outlook items (emails, appointments, contacts)
into JSON-serializable dictionaries.
"""

import logging
from typing import Dict, Any
from config import EMAIL_BODY_PREVIEW_LENGTH

logger = logging.getLogger(__name__)


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
            "has_attachments": mail_item.Attachments.Count > 0,
            "attachment_count": mail_item.Attachments.Count,
            "importance": mail_item.Importance,
            "unread": mail_item.UnRead,
            "categories": mail_item.Categories,
            "folder_path": folder_path,
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

