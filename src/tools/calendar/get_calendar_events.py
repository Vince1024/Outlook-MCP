"""Get Calendar Events Tool"""
import json, logging
from datetime import datetime, timedelta
from config import OUTLOOK_FOLDER_CALENDAR
from utils import get_outlook_application, format_appointment

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def get_calendar_events(days_ahead: int = 7, include_past: bool = False) -> str:
        """Get calendar events from Outlook."""
        # Log operation start
        logger.debug("Starting get_calendar_events operation", extra={
            "operation": "get_calendar_events",
            "days_ahead": days_ahead,
            "include_past": include_past
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            calendar = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CALENDAR)
            items = calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            start_date = datetime.now()
            if include_past:
                start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = datetime.now().replace(hour=23, minute=59, second=59)
            end_date = end_date + timedelta(days=days_ahead)
            filter_str = f"[Start] >= '{start_date.strftime('%m/%d/%Y %H:%M')}' AND [End] <= '{end_date.strftime('%m/%d/%Y %H:%M')}'"
            filtered_items = items.Restrict(filter_str)
            events = []
            for appointment in filtered_items:
                events.append(format_appointment(appointment))
            
            # Log success
            logger.info("Retrieved calendar events successfully", extra={
                "event_count": len(events),
                "days_ahead": days_ahead,
                "include_past": include_past
            })
            
            return json.dumps({"success": True, "count": len(events), "events": events}, indent=2)
        except Exception as e:
            logger.error("Failed to get calendar events", exc_info=True, extra={"days_ahead": days_ahead, "include_past": include_past})
            return json.dumps({"success": False, "error": str(e)})

