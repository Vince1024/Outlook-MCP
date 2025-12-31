"""
Unit tests for get_calendar_events tool

This test verifies that the get_calendar_events tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.calendar.get_calendar_events import register


class TestGetCalendarEvents:
    """Tests for get_calendar_events tool"""
    
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
    
    @patch('tools.calendar.get_calendar_events.format_appointment')
    @patch('tools.calendar.get_calendar_events.get_outlook_application')
    def test_get_calendar_events_basic(self, mock_outlook, mock_format):
        """Test basic calendar event retrieval"""
        # Setup mocks
        mock_appt1 = MagicMock()
        mock_appt2 = MagicMock()
        
        mock_filtered_items = [mock_appt1, mock_appt2]
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.IncludeRecurrences = False
        mock_items.Restrict.return_value = mock_filtered_items
        
        mock_calendar = MagicMock()
        mock_calendar.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_calendar
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_appointment
        mock_format.side_effect = [
            {"subject": "Team Meeting", "start": "2025-12-20 14:00"},
            {"subject": "Project Review", "start": "2025-12-21 10:00"}
        ]
        
        # Call the tool
        result = self.tool_func(days_ahead=7)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        assert len(result_dict["events"]) == 2
        assert result_dict["events"][0]["subject"] == "Team Meeting"
        
        # Verify that Outlook methods were called
        mock_app.GetNamespace.assert_called_once_with("MAPI")
        mock_items.Sort.assert_called_once_with("[Start]")
        assert mock_items.IncludeRecurrences is True
        mock_items.Restrict.assert_called_once()
    
    @patch('tools.calendar.get_calendar_events.format_appointment')
    @patch('tools.calendar.get_calendar_events.get_outlook_application')
    def test_get_calendar_events_no_events(self, mock_outlook, mock_format):
        """Test getting events when calendar is empty"""
        # Setup mocks
        mock_filtered_items = []
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.IncludeRecurrences = False
        mock_items.Restrict.return_value = mock_filtered_items
        
        mock_calendar = MagicMock()
        mock_calendar.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_calendar
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func(days_ahead=7)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["events"]) == 0
    
    @patch('tools.calendar.get_calendar_events.format_appointment')
    @patch('tools.calendar.get_calendar_events.get_outlook_application')
    def test_get_calendar_events_with_past(self, mock_outlook, mock_format):
        """Test getting events including past events"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_filtered_items = [mock_appt]
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.IncludeRecurrences = False
        mock_items.Restrict.return_value = mock_filtered_items
        
        mock_calendar = MagicMock()
        mock_calendar.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_calendar
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.return_value = {"subject": "Past Event", "start": "2025-12-19 09:00"}
        
        # Call the tool with include_past=True
        result = self.tool_func(days_ahead=7, include_past=True)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
    
    @patch('tools.calendar.get_calendar_events.format_appointment')
    @patch('tools.calendar.get_calendar_events.get_outlook_application')
    def test_get_calendar_events_custom_days_ahead(self, mock_outlook, mock_format):
        """Test getting events with custom days_ahead parameter"""
        # Setup mocks
        mock_appt1 = MagicMock()
        mock_appt2 = MagicMock()
        mock_appt3 = MagicMock()
        
        mock_filtered_items = [mock_appt1, mock_appt2, mock_appt3]
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.IncludeRecurrences = False
        mock_items.Restrict.return_value = mock_filtered_items
        
        mock_calendar = MagicMock()
        mock_calendar.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_calendar
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.side_effect = [
            {"subject": "Event 1"},
            {"subject": "Event 2"},
            {"subject": "Event 3"}
        ]
        
        # Call the tool with days_ahead=30
        result = self.tool_func(days_ahead=30)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 3
    
    @patch('tools.calendar.get_calendar_events.get_outlook_application')
    def test_get_calendar_events_failure(self, mock_outlook):
        """Test error handling during event retrieval"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Calendar access failed")
        
        # Call the tool
        result = self.tool_func(days_ahead=7)
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Calendar access failed" in result_dict["error"]
    
    @patch('tools.calendar.get_calendar_events.format_appointment')
    @patch('tools.calendar.get_calendar_events.get_outlook_application')
    def test_get_calendar_events_with_format_error(self, mock_outlook, mock_format):
        """Test that formatting errors are handled gracefully"""
        # Setup mocks
        mock_appt1 = MagicMock()
        mock_appt2 = MagicMock()
        
        mock_filtered_items = [mock_appt1, mock_appt2]
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.IncludeRecurrences = False
        mock_items.Restrict.return_value = mock_filtered_items
        
        mock_calendar = MagicMock()
        mock_calendar.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_calendar
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # First one formats successfully, second fails
        mock_format.side_effect = [
            {"subject": "Event 1"},
            Exception("Format error")
        ]
        
        # Call the tool - should raise exception since we're iterating with format_appointment
        with pytest.raises(Exception):
            self.tool_func(days_ahead=7)


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

