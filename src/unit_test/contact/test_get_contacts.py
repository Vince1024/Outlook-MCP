"""
Unit tests for get_contacts tool

This test verifies that the get_contacts tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.contact.get_contacts import register


class TestGetContacts:
    """Tests for get_contacts tool"""
    
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
    
    @patch('tools.contact.get_contacts.format_contact')
    @patch('tools.contact.get_contacts.get_outlook_application')
    def test_get_contacts_basic(self, mock_outlook, mock_format):
        """Test basic contact retrieval"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact2 = MagicMock()
        mock_contact3 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_contact1
        mock_items.GetNext.side_effect = [mock_contact2, mock_contact3, None]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_contact
        mock_format.side_effect = [
            {"full_name": "John Doe", "email": "john@example.com"},
            {"full_name": "Jane Smith", "email": "jane@example.com"},
            {"full_name": "Bob Johnson", "email": "bob@example.com"}
        ]
        
        # Call the tool
        result = self.tool_func(limit=10)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 3
        assert len(result_dict["contacts"]) == 3
        assert result_dict["contacts"][0]["full_name"] == "John Doe"
        
        # Verify that Outlook methods were called
        mock_app.GetNamespace.assert_called_once_with("MAPI")
        mock_namespace.GetDefaultFolder.assert_called_once()
        mock_items.Sort.assert_called_once_with("[FullName]")
    
    @patch('tools.contact.get_contacts.format_contact')
    @patch('tools.contact.get_contacts.get_outlook_application')
    def test_get_contacts_with_search_filter(self, mock_outlook, mock_format):
        """Test getting contacts with search filter"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact1.FullName = "John Doe"
        
        mock_contact2 = MagicMock()
        mock_contact2.FullName = "Jane Smith"
        
        mock_contact3 = MagicMock()
        mock_contact3.FullName = "John Wilson"
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_contact1
        mock_items.GetNext.side_effect = [mock_contact2, mock_contact3, None]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_contact (only for matches)
        mock_format.side_effect = [
            {"full_name": "John Doe", "email": "john@example.com"},
            {"full_name": "John Wilson", "email": "jwilson@example.com"}
        ]
        
        # Call the tool with search_name filter
        result = self.tool_func(limit=10, search_name="John")
        
        # Verify result - should only return contacts with "John" in name
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        assert result_dict["contacts"][0]["full_name"] == "John Doe"
        assert result_dict["contacts"][1]["full_name"] == "John Wilson"
    
    @patch('tools.contact.get_contacts.get_outlook_application')
    def test_get_contacts_empty_folder(self, mock_outlook):
        """Test getting contacts from empty folder"""
        # Setup mocks for empty contacts folder
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = None
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func(limit=10)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["contacts"]) == 0
    
    @patch('tools.contact.get_contacts.format_contact')
    @patch('tools.contact.get_contacts.get_outlook_application')
    def test_get_contacts_limit_enforced(self, mock_outlook, mock_format):
        """Test that MAX_CONTACT_LIMIT is enforced"""
        # Setup mocks - simulate many contacts
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        
        # Create 60 mock contacts (more than MAX_CONTACT_LIMIT=50)
        mock_contacts = [MagicMock() for _ in range(60)]
        mock_items.GetFirst.return_value = mock_contacts[0]
        mock_items.GetNext.side_effect = mock_contacts[1:] + [None]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.side_effect = [{"full_name": f"Contact {i}"} for i in range(60)]
        
        # Call with limit > MAX_CONTACT_LIMIT
        result = self.tool_func(limit=100)
        
        # Verify that only MAX_CONTACT_LIMIT (50) contacts are returned
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] <= 50
    
    @patch('tools.contact.get_contacts.get_outlook_application')
    def test_get_contacts_failure(self, mock_outlook):
        """Test error handling during contact retrieval"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Contacts folder access failed")
        
        # Call the tool
        result = self.tool_func(limit=10)
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Contacts folder access failed" in result_dict["error"]
    
    @patch('tools.contact.get_contacts.format_contact')
    @patch('tools.contact.get_contacts.get_outlook_application')
    def test_get_contacts_with_format_error(self, mock_outlook, mock_format):
        """Test that formatting errors are handled gracefully"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact2 = MagicMock()
        mock_contact3 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_contact1
        mock_items.GetNext.side_effect = [mock_contact2, mock_contact3, None]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_contact - second one fails
        mock_format.side_effect = [
            {"full_name": "Contact 1"},
            Exception("Format error"),  # This contact will be skipped
            {"full_name": "Contact 3"}
        ]
        
        # Call the tool
        result = self.tool_func(limit=10)
        
        # Verify result - should have 2 contacts (skipped the error one)
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        assert result_dict["contacts"][0]["full_name"] == "Contact 1"
        assert result_dict["contacts"][1]["full_name"] == "Contact 3"


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

