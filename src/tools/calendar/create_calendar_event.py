"""Create Calendar Event Tool"""
import json, logging
from typing import Optional
from dateutil import parser as date_parser
from config import OUTLOOK_ITEM_APPOINTMENT
from utils import get_outlook_application

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def create_calendar_event(subject: str, start_time: str, end_time: str, location: Optional[str] = None, body: Optional[str] = None, required_attendees: Optional[str] = None, optional_attendees: Optional[str] = None, reminder_minutes: int = 15, is_all_day: bool = False) -> str:
        """Create a new calendar event/appointment in Outlook."""
        # Log operation start (no subject for privacy)
        logger.debug("Starting create_calendar_event operation", extra={
            "operation": "create_calendar_event",
            "has_subject": bool(subject),
            "has_location": bool(location),
            "has_attendees": bool(required_attendees or optional_attendees),
            "is_all_day": is_all_day
        })
        
        try:
            outlook = get_outlook_application()
            appointment = outlook.CreateItem(OUTLOOK_ITEM_APPOINTMENT)
            appointment.Subject = subject
            try:
                start_dt = date_parser.parse(start_time)
                end_dt = date_parser.parse(end_time)
            except Exception as e:
                return json.dumps({"success": False, "error": f"Invalid date format: {e}. Use ISO format like '2025-01-15 14:00' or natural language like 'tomorrow 2pm'"})
            appointment.Start = start_dt
            appointment.End = end_dt
            appointment.AllDayEvent = is_all_day
            if location:
                appointment.Location = location
            if body:
                appointment.Body = body
            if required_attendees:
                appointment.RequiredAttendees = required_attendees
            if optional_attendees:
                appointment.OptionalAttendees = optional_attendees
            appointment.ReminderSet = True
            appointment.ReminderMinutesBeforeStart = reminder_minutes
            appointment.Save()
            if required_attendees or optional_attendees:
                appointment.Send()
            
            # Log success (no subject for privacy)
            logger.info("Calendar event created successfully", extra={
                "has_subject": bool(subject),
                "has_location": bool(location),
                "has_attendees": bool(required_attendees or optional_attendees),
                "is_all_day": is_all_day,
                "reminder_minutes": reminder_minutes
            })
            
            return json.dumps({"success": True, "message": f"Calendar event created for {start_time}"}, indent=2)
        except Exception as e:
            logger.error("Failed to create calendar event", exc_info=True, extra={
                "has_subject": bool(subject),
                "has_start_time": bool(start_time),
                "has_end_time": bool(end_time)
            })
            return json.dumps({"success": False, "error": str(e)})

