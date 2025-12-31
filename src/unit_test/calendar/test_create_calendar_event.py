"""
Unit tests for create_calendar_event tool

This test verifies that the create_calendar_event tool works correctly without creating real events.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.calendar.create_calendar_event import register


class TestCreateCalendarEvent:
    """Tests for create_calendar_event tool"""
    
    def setup_method(self):
        """Setup before each test"""
        self.mcp_mock = Mock()
        self.tool_func = None
        
        # Capture registered function
        def capture_tool(func):
            self.tool_func = func
            return func
        
        self.mcp_mock.tool.return_value = capture_tool
        
        # Register the tool
        register(self.mcp_mock)
    
    def teardown_method(self):
        """Cleanup after each test to prevent memory leaks"""
        import gc
        self.mcp_mock = None
        self.tool_func = None
        gc.collect()  # Force garbage collection
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_basic(self, mock_outlook):
        """Test basic calendar event creation"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_appt
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func(
            subject="Team Meeting",
            start_time="2025-12-20 14:00",
            end_time="2025-12-20 15:00"
        )
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert "created" in result_dict["message"]
        
        # Verify that Outlook methods were called
        mock_app.CreateItem.assert_called_once()
        mock_appt.Save.assert_called_once()
        mock_appt.Send.assert_not_called()  # No attendees, so no send
        
        # Verify that properties were set
        assert mock_appt.Subject == "Team Meeting"
        assert mock_appt.AllDayEvent is False
        assert mock_appt.ReminderSet is True
        assert mock_appt.ReminderMinutesBeforeStart == 15
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_with_location(self, mock_outlook):
        """Test creating event with location"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_appt
        mock_outlook.return_value = mock_app
        
        # Call the tool with location
        result = self.tool_func(
            subject="Meeting",
            start_time="2025-12-20 14:00",
            end_time="2025-12-20 15:00",
            location="Conference Room A"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert mock_appt.Location == "Conference Room A"
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_with_body(self, mock_outlook):
        """Test creating event with body content"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_appt
        mock_outlook.return_value = mock_app
        
        # Call the tool with body
        result = self.tool_func(
            subject="Meeting",
            start_time="2025-12-20 14:00",
            end_time="2025-12-20 15:00",
            body="Agenda:\n1. Q1 Review\n2. Q2 Planning"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert mock_appt.Body == "Agenda:\n1. Q1 Review\n2. Q2 Planning"
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_with_attendees(self, mock_outlook):
        """Test creating event with attendees (meeting invitation)"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_appt
        mock_outlook.return_value = mock_app
        
        # Call the tool with attendees
        result = self.tool_func(
            subject="Team Meeting",
            start_time="2025-12-20 14:00",
            end_time="2025-12-20 15:00",
            required_attendees="john@example.com; jane@example.com",
            optional_attendees="bob@example.com"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that attendees are set
        assert mock_appt.RequiredAttendees == "john@example.com; jane@example.com"
        assert mock_appt.OptionalAttendees == "bob@example.com"
        
        # Verify that Send was called (meeting invitation)
        mock_appt.Send.assert_called_once()
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_all_day(self, mock_outlook):
        """Test creating all-day event"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_appt
        mock_outlook.return_value = mock_app
        
        # Call the tool with is_all_day=True
        result = self.tool_func(
            subject="Holiday",
            start_time="2025-12-25",
            end_time="2025-12-25",
            is_all_day=True
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert mock_appt.AllDayEvent is True
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_custom_reminder(self, mock_outlook):
        """Test creating event with custom reminder"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_appt
        mock_outlook.return_value = mock_app
        
        # Call the tool with custom reminder
        result = self.tool_func(
            subject="Meeting",
            start_time="2025-12-20 14:00",
            end_time="2025-12-20 15:00",
            reminder_minutes=30
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert mock_appt.ReminderMinutesBeforeStart == 30
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_invalid_date(self, mock_outlook):
        """Test error handling for invalid date format"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_appt
        mock_outlook.return_value = mock_app
        
        # Call the tool with invalid date
        result = self.tool_func(
            subject="Meeting",
            start_time="invalid-date",
            end_time="2025-12-20 15:00"
        )
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Invalid date format" in result_dict["error"]
    
    @patch('tools.calendar.create_calendar_event.get_outlook_application')
    def test_create_calendar_event_failure(self, mock_outlook):
        """Test error handling during event creation"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Calendar access failed")
        
        # Call the tool
        result = self.tool_func(
            subject="Meeting",
            start_time="2025-12-20 14:00",
            end_time="2025-12-20 15:00"
        )
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Calendar access failed" in result_dict["error"]


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

