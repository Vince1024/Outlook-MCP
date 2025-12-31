"""Search Calendar Events Tool"""
import hashlib, json, logging
from datetime import datetime, timedelta
from config import OUTLOOK_FOLDER_CALENDAR
from utils import get_outlook_application, format_appointment

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def search_calendar_events(query: str, days_range: int = 30) -> str:
        """Search for calendar events by keyword in subject or location."""
        # Log operation start
        logger.debug("Starting search_calendar_events operation", extra={
            "operation": "search_calendar_events",
            "query_length": len(query),
            "days_range": days_range
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            calendar = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CALENDAR)
            items = calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            start_date = datetime.now() - timedelta(days=days_range)
            end_date = datetime.now() + timedelta(days=days_range)
            filter_str = f"[Start] >= '{start_date.strftime('%m/%d/%Y')}' AND [End] <= '{end_date.strftime('%m/%d/%Y')}'"
            filtered_items = items.Restrict(filter_str)
            events = []
            query_lower = query.lower()
            for appointment in filtered_items:
                subject = appointment.Subject.lower() if appointment.Subject else ""
                location = appointment.Location.lower() if appointment.Location else ""
                if query_lower in subject or query_lower in location:
                    events.append(format_appointment(appointment))
            
            # Log success (query length only for privacy)
            logger.info("Calendar search completed successfully", extra={
                "result_count": len(events),
                "query_length": len(query),
                "days_range": days_range
            })
            
            return json.dumps({"success": True, "query": query, "count": len(events), "events": events}, indent=2)
        except Exception as e:
            # Hash query for privacy (only first 8 chars of hash)
            query_hash = hashlib.sha256(query.encode()).hexdigest()[:8]
            logger.error("Failed to search calendar events", exc_info=True, extra={
                "query_hash": query_hash,
                "query_length": len(query),
                "days_range": days_range
            })
            return json.dumps({"success": False, "error": str(e)})

