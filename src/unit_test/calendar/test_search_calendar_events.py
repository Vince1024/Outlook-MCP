"""
Unit tests for search_calendar_events tool

This test verifies that the search_calendar_events tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.calendar.search_calendar_events import register


class TestSearchCalendarEvents:
    """Tests for search_calendar_events tool"""
    
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
    
    @patch('tools.calendar.search_calendar_events.format_appointment')
    @patch('tools.calendar.search_calendar_events.get_outlook_application')
    def test_search_calendar_events_by_subject(self, mock_outlook, mock_format):
        """Test searching events by subject"""
        # Setup mocks
        mock_appt1 = MagicMock()
        mock_appt1.Subject = "Team Meeting"
        mock_appt1.Location = "Room A"
        
        mock_appt2 = MagicMock()
        mock_appt2.Subject = "Project Review"
        mock_appt2.Location = "Room B"
        
        mock_appt3 = MagicMock()
        mock_appt3.Subject = "Team Lunch"
        mock_appt3.Location = "Cafeteria"
        
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
        
        # Mock format_appointment (only for matches)
        mock_format.side_effect = [
            {"subject": "Team Meeting", "location": "Room A"},
            {"subject": "Team Lunch", "location": "Cafeteria"}
        ]
        
        # Call the tool - search for "team"
        result = self.tool_func(query="team", days_range=30)
        
        # Verify result - should return 2 events with "team" in subject
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["query"] == "team"
        assert result_dict["count"] == 2
        assert result_dict["events"][0]["subject"] == "Team Meeting"
        assert result_dict["events"][1]["subject"] == "Team Lunch"
    
    @patch('tools.calendar.search_calendar_events.format_appointment')
    @patch('tools.calendar.search_calendar_events.get_outlook_application')
    def test_search_calendar_events_by_location(self, mock_outlook, mock_format):
        """Test searching events by location"""
        # Setup mocks
        mock_appt1 = MagicMock()
        mock_appt1.Subject = "Meeting"
        mock_appt1.Location = "Conference Room"
        
        mock_appt2 = MagicMock()
        mock_appt2.Subject = "Review"
        mock_appt2.Location = "Office"
        
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
        
        # Mock format_appointment (only for match)
        mock_format.return_value = {"subject": "Meeting", "location": "Conference Room"}
        
        # Call the tool - search for "conference"
        result = self.tool_func(query="conference", days_range=30)
        
        # Verify result - should return 1 event with "conference" in location
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
        assert result_dict["events"][0]["location"] == "Conference Room"
    
    @patch('tools.calendar.search_calendar_events.format_appointment')
    @patch('tools.calendar.search_calendar_events.get_outlook_application')
    def test_search_calendar_events_no_results(self, mock_outlook, mock_format):
        """Test search with no matching events"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_appt.Subject = "Team Meeting"
        mock_appt.Location = "Room A"
        
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
        
        # Call the tool - search for "nonexistent"
        result = self.tool_func(query="nonexistent", days_range=30)
        
        # Verify result - should return 0 events
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["events"]) == 0
    
    @patch('tools.calendar.search_calendar_events.format_appointment')
    @patch('tools.calendar.search_calendar_events.get_outlook_application')
    def test_search_calendar_events_custom_days_range(self, mock_outlook, mock_format):
        """Test search with custom days_range parameter"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_appt.Subject = "Meeting"
        mock_appt.Location = "Room"
        
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
        
        mock_format.return_value = {"subject": "Meeting"}
        
        # Call the tool with days_range=60
        result = self.tool_func(query="meeting", days_range=60)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
    
    @patch('tools.calendar.search_calendar_events.format_appointment')
    @patch('tools.calendar.search_calendar_events.get_outlook_application')
    def test_search_calendar_events_case_insensitive(self, mock_outlook, mock_format):
        """Test that search is case-insensitive"""
        # Setup mocks
        mock_appt = MagicMock()
        mock_appt.Subject = "TEAM MEETING"  # Uppercase
        mock_appt.Location = "Room A"
        
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
        
        mock_format.return_value = {"subject": "TEAM MEETING"}
        
        # Call the tool with lowercase query
        result = self.tool_func(query="team", days_range=30)
        
        # Verify result - should match despite case difference
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
    
    @patch('tools.calendar.search_calendar_events.get_outlook_application')
    def test_search_calendar_events_failure(self, mock_outlook):
        """Test error handling during search"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Calendar access failed")
        
        # Call the tool
        result = self.tool_func(query="meeting", days_range=30)
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Calendar access failed" in result_dict["error"]
    
    @patch('tools.calendar.search_calendar_events.format_appointment')
    @patch('tools.calendar.search_calendar_events.get_outlook_application')
    def test_search_calendar_events_with_none_fields(self, mock_outlook, mock_format):
        """Test search when Subject or Location are None"""
        # Setup mocks
        mock_appt1 = MagicMock()
        mock_appt1.Subject = None  # No subject
        mock_appt1.Location = "Room A"
        
        mock_appt2 = MagicMock()
        mock_appt2.Subject = "Meeting"
        mock_appt2.Location = None  # No location
        
        mock_appt3 = MagicMock()
        mock_appt3.Subject = "Team Meeting"
        mock_appt3.Location = "Room B"
        
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
            {"subject": "Meeting"},
            {"subject": "Team Meeting", "location": "Room B"}
        ]
        
        # Call the tool - search for "meeting"
        result = self.tool_func(query="meeting", days_range=30)
        
        # Verify result - should handle None fields gracefully
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2  # Only appt2 and appt3 match


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

