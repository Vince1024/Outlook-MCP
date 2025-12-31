"""
Unit tests for create_contact tool

This test verifies that the create_contact tool works correctly without creating real contacts.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.contact.create_contact import register


class TestCreateContact:
    """Tests for create_contact tool"""
    
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
    
    @patch('tools.contact.create_contact.get_outlook_application')
    def test_create_contact_basic(self, mock_outlook):
        """Test basic contact creation"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_contact
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func(
            full_name="John Doe",
            email="john@example.com"
        )
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert "John Doe" in result_dict["message"]
        assert "created" in result_dict["message"]
        
        # Verify that Outlook methods were called
        mock_app.CreateItem.assert_called_once()
        mock_contact.Save.assert_called_once()
        
        # Verify that properties were set
        assert mock_contact.FullName == "John Doe"
        assert mock_contact.Email1Address == "john@example.com"
    
    @patch('tools.contact.create_contact.get_outlook_application')
    def test_create_contact_with_company(self, mock_outlook):
        """Test creating contact with company information"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_contact
        mock_outlook.return_value = mock_app
        
        # Call the tool with company
        result = self.tool_func(
            full_name="John Doe",
            email="john@example.com",
            company="Acme Corp",
            job_title="Software Engineer"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that company info was set
        assert mock_contact.CompanyName == "Acme Corp"
        assert mock_contact.JobTitle == "Software Engineer"
    
    @patch('tools.contact.create_contact.get_outlook_application')
    def test_create_contact_with_phones(self, mock_outlook):
        """Test creating contact with phone numbers"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_contact
        mock_outlook.return_value = mock_app
        
        # Call the tool with phone numbers
        result = self.tool_func(
            full_name="John Doe",
            email="john@example.com",
            business_phone="+1-555-0100",
            mobile_phone="+1-555-0101",
            home_phone="+1-555-0102"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that phone numbers were set
        assert mock_contact.BusinessTelephoneNumber == "+1-555-0100"
        assert mock_contact.MobileTelephoneNumber == "+1-555-0101"
        assert mock_contact.HomeTelephoneNumber == "+1-555-0102"
    
    @patch('tools.contact.create_contact.get_outlook_application')
    def test_create_contact_complete(self, mock_outlook):
        """Test creating contact with all fields"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_contact
        mock_outlook.return_value = mock_app
        
        # Call the tool with all fields
        result = self.tool_func(
            full_name="John Doe",
            email="john@example.com",
            company="Acme Corp",
            job_title="CEO",
            business_phone="+1-555-0100",
            mobile_phone="+1-555-0101",
            home_phone="+1-555-0102"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify all fields were set
        assert mock_contact.FullName == "John Doe"
        assert mock_contact.Email1Address == "john@example.com"
        assert mock_contact.CompanyName == "Acme Corp"
        assert mock_contact.JobTitle == "CEO"
        assert mock_contact.BusinessTelephoneNumber == "+1-555-0100"
        assert mock_contact.MobileTelephoneNumber == "+1-555-0101"
        assert mock_contact.HomeTelephoneNumber == "+1-555-0102"
    
    @patch('tools.contact.create_contact.get_outlook_application')
    def test_create_contact_minimal(self, mock_outlook):
        """Test creating contact with only required fields"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_contact
        mock_outlook.return_value = mock_app
        
        # Call the tool with only required fields
        result = self.tool_func(
            full_name="Jane Smith",
            email="jane@example.com"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify only required fields were set
        assert mock_contact.FullName == "Jane Smith"
        assert mock_contact.Email1Address == "jane@example.com"
    
    @patch('tools.contact.create_contact.get_outlook_application')
    def test_create_contact_failure(self, mock_outlook):
        """Test error handling during contact creation"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Contacts folder access failed")
        
        # Call the tool
        result = self.tool_func(
            full_name="John Doe",
            email="john@example.com"
        )
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Contacts folder access failed" in result_dict["error"]
    
    @patch('tools.contact.create_contact.get_outlook_application')
    def test_create_contact_with_special_characters(self, mock_outlook):
        """Test creating contact with special characters in name"""
        # Setup mocks
        mock_contact = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_contact
        mock_outlook.return_value = mock_app
        
        # Call the tool with special characters
        result = self.tool_func(
            full_name="José María García",
            email="jose@example.com"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert mock_contact.FullName == "José María García"


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

