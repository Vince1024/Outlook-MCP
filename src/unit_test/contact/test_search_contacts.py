"""
Unit tests for search_contacts tool

This test verifies that the search_contacts tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.contact.search_contacts import register


class TestSearchContacts:
    """Tests for search_contacts tool"""
    
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
    
    @patch('tools.contact.search_contacts.format_contact')
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_by_name(self, mock_outlook, mock_format):
        """Test searching contacts by name"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact1.FullName = "John Doe"
        mock_contact1.Email1Address = "john@example.com"
        mock_contact1.CompanyName = "Acme Corp"
        
        mock_contact2 = MagicMock()
        mock_contact2.FullName = "Jane Smith"
        mock_contact2.Email1Address = "jane@example.com"
        mock_contact2.CompanyName = "Tech Inc"
        
        mock_contact3 = MagicMock()
        mock_contact3.FullName = "John Wilson"
        mock_contact3.Email1Address = "jwilson@example.com"
        mock_contact3.CompanyName = "Business LLC"
        
        mock_items = [mock_contact1, mock_contact2, mock_contact3]
        
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
        
        # Call the tool - search for "john"
        result = self.tool_func(query="john")
        
        # Verify result - should return 2 contacts with "john" in name
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["query"] == "john"
        assert result_dict["count"] == 2
        assert result_dict["contacts"][0]["full_name"] == "John Doe"
        assert result_dict["contacts"][1]["full_name"] == "John Wilson"
    
    @patch('tools.contact.search_contacts.format_contact')
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_by_email(self, mock_outlook, mock_format):
        """Test searching contacts by email"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact1.FullName = "John Doe"
        mock_contact1.Email1Address = "john@acme.com"
        mock_contact1.CompanyName = "Acme Corp"
        
        mock_contact2 = MagicMock()
        mock_contact2.FullName = "Jane Smith"
        mock_contact2.Email1Address = "jane@example.com"
        mock_contact2.CompanyName = "Example Inc"
        
        mock_items = [mock_contact1, mock_contact2]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_contact (only for match)
        mock_format.return_value = {"full_name": "John Doe", "email": "john@acme.com"}
        
        # Call the tool - search for "acme"
        result = self.tool_func(query="acme")
        
        # Verify result - should return 1 contact with "acme" in email
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
        assert result_dict["contacts"][0]["email"] == "john@acme.com"
    
    @patch('tools.contact.search_contacts.format_contact')
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_by_company(self, mock_outlook, mock_format):
        """Test searching contacts by company"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact1.FullName = "John Doe"
        mock_contact1.Email1Address = "john@example.com"
        mock_contact1.CompanyName = "Acme Corp"
        
        mock_contact2 = MagicMock()
        mock_contact2.FullName = "Jane Smith"
        mock_contact2.Email1Address = "jane@example.com"
        mock_contact2.CompanyName = "Tech Solutions"
        
        mock_items = [mock_contact1, mock_contact2]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_contact (only for match)
        mock_format.return_value = {"full_name": "Jane Smith", "company": "Tech Solutions"}
        
        # Call the tool - search for "tech"
        result = self.tool_func(query="tech")
        
        # Verify result - should return 1 contact with "tech" in company
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
        assert result_dict["contacts"][0]["company"] == "Tech Solutions"
    
    @patch('tools.contact.search_contacts.format_contact')
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_no_results(self, mock_outlook, mock_format):
        """Test search with no matching contacts"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_contact.FullName = "John Doe"
        mock_contact.Email1Address = "john@example.com"
        mock_contact.CompanyName = "Acme Corp"
        
        mock_items = [mock_contact]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool - search for "nonexistent"
        result = self.tool_func(query="nonexistent")
        
        # Verify result - should return 0 contacts
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["contacts"]) == 0
    
    @patch('tools.contact.search_contacts.format_contact')
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_case_insensitive(self, mock_outlook, mock_format):
        """Test that search is case-insensitive"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_contact.FullName = "JOHN DOE"  # Uppercase
        mock_contact.Email1Address = "john@example.com"
        mock_contact.CompanyName = "Acme"
        
        mock_items = [mock_contact]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.return_value = {"full_name": "JOHN DOE"}
        
        # Call the tool with lowercase query
        result = self.tool_func(query="john")
        
        # Verify result - should match despite case difference
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
    
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_failure(self, mock_outlook):
        """Test error handling during search"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Contacts folder access failed")
        
        # Call the tool
        result = self.tool_func(query="john")
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Contacts folder access failed" in result_dict["error"]
    
    @patch('tools.contact.search_contacts.format_contact')
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_with_none_fields(self, mock_outlook, mock_format):
        """Test search when contact fields are None"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact1.FullName = None  # No name
        mock_contact1.Email1Address = "unknown@example.com"
        mock_contact1.CompanyName = None
        
        mock_contact2 = MagicMock()
        mock_contact2.FullName = "John Doe"
        mock_contact2.Email1Address = None  # No email
        mock_contact2.CompanyName = "Acme"
        
        mock_items = [mock_contact1, mock_contact2]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.return_value = {"full_name": "John Doe", "company": "Acme"}
        
        # Call the tool - search for "acme"
        result = self.tool_func(query="acme")
        
        # Verify result - should handle None fields gracefully
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1  # Only contact2 matches
    
    @patch('tools.contact.search_contacts.format_contact')
    @patch('tools.contact.search_contacts.get_outlook_application')
    def test_search_contacts_with_format_error(self, mock_outlook, mock_format):
        """Test that contacts with format errors are skipped"""
        # Setup mocks
        mock_contact1 = MagicMock()
        mock_contact1.FullName = "John Doe"
        mock_contact1.Email1Address = "john@example.com"
        mock_contact1.CompanyName = "Acme"
        
        mock_contact2 = MagicMock()
        mock_contact2.FullName = "Jane Smith"
        mock_contact2.Email1Address = "jane@example.com"
        mock_contact2.CompanyName = "Tech"
        
        mock_items = [mock_contact1, mock_contact2]
        
        mock_contacts_folder = MagicMock()
        mock_contacts_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_contacts_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_contact - first succeeds, second returns error dict
        mock_format.side_effect = [
            {"full_name": "John Doe"},
            {"error": "Format error"}  # Error dict is skipped
        ]
        
        # Call the tool
        result = self.tool_func(query="e")  # Should match both, but second has error
        
        # Verify result - should have 1 contact (skipped the error one)
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

