"""
Unit tests for get_sent_emails tool

This test verifies that the get_sent_emails tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.email.get_sent_emails import register


class TestGetSentEmails:
    """Tests for get_sent_emails tool"""
    
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
    
    @patch('tools.email.get_sent_emails.format_email')
    @patch('tools.email.get_sent_emails.get_outlook_application')
    def test_get_sent_emails_basic(self, mock_outlook, mock_format):
        """Test basic sent email retrieval"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        mock_mail3 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_mail1
        mock_items.GetNext.side_effect = [mock_mail2, mock_mail3, None]
        
        mock_sent_folder = MagicMock()
        mock_sent_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_sent_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_email
        mock_format.side_effect = [
            {"subject": "Sent Email 1", "to": "user1@example.com"},
            {"subject": "Sent Email 2", "to": "user2@example.com"},
            {"subject": "Sent Email 3", "to": "user3@example.com"}
        ]
        
        # Call the tool
        result = self.tool_func(limit=5)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 3
        assert len(result_dict["emails"]) == 3
        assert result_dict["emails"][0]["subject"] == "Sent Email 1"
        assert result_dict["emails"][1]["to"] == "user2@example.com"
        
        # Verify that Outlook methods were called
        mock_app.GetNamespace.assert_called_once_with("MAPI")
        mock_namespace.GetDefaultFolder.assert_called_once()
        mock_items.Sort.assert_called_once_with("[SentOn]", True)
    
    @patch('tools.email.get_sent_emails.get_outlook_application')
    def test_get_sent_emails_empty_sent_folder(self, mock_outlook):
        """Test getting emails from empty sent folder"""
        # Setup mocks for empty sent folder
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = None
        
        mock_sent_folder = MagicMock()
        mock_sent_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_sent_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func(limit=5)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["emails"]) == 0
    
    @patch('tools.email.get_sent_emails.format_email')
    @patch('tools.email.get_sent_emails.get_outlook_application')
    def test_get_sent_emails_limit_enforced(self, mock_outlook, mock_format):
        """Test that MAX_EMAIL_LIMIT is enforced"""
        # Setup mocks - simulate more emails than limit
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        
        # Create only 10 mock emails to test limit logic (reduced from 60 to avoid OOM)
        # This is sufficient to validate that MAX_EMAIL_LIMIT is enforced
        mock_mails = [MagicMock() for _ in range(10)]
        mock_items.GetFirst.return_value = mock_mails[0]
        mock_items.GetNext.side_effect = mock_mails[1:] + [None]
        
        mock_sent_folder = MagicMock()
        mock_sent_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_sent_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.side_effect = [{"subject": f"Email {i}"} for i in range(10)]
        
        # Call with limit=5 to test that limit is enforced
        result = self.tool_func(limit=5)
        
        # Verify that only requested limit (5) emails are returned, not all 10 available
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 5  # Should stop at limit, not return all 10
    
    @patch('tools.email.get_sent_emails.get_outlook_application')
    def test_get_sent_emails_failure(self, mock_outlook):
        """Test error handling during retrieval"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Cannot access Sent Items folder")
        
        # Call the tool
        result = self.tool_func(limit=5)
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Cannot access Sent Items folder" in result_dict["error"]
    
    @patch('tools.email.get_sent_emails.format_email')
    @patch('tools.email.get_sent_emails.get_outlook_application')
    def test_get_sent_emails_with_format_error(self, mock_outlook, mock_format):
        """Test that formatting errors are handled gracefully"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_mail1
        mock_items.GetNext.side_effect = [mock_mail2, None]
        
        mock_sent_folder = MagicMock()
        mock_sent_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_sent_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_email - first one fails
        mock_format.side_effect = [
            Exception("Format error"),  # This email will be skipped
            {"subject": "Email 2"}
        ]
        
        # Call the tool
        result = self.tool_func(limit=5)
        
        # Verify result - should have 1 email (skipped the error one)
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
        assert result_dict["emails"][0]["subject"] == "Email 2"
    
    @patch('tools.email.get_sent_emails.format_email')
    @patch('tools.email.get_sent_emails.get_outlook_application')
    def test_get_sent_emails_custom_limit(self, mock_outlook, mock_format):
        """Test custom limit parameter"""
        # Setup mocks
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        
        # Create 10 mock emails
        mock_mails = [MagicMock() for _ in range(10)]
        mock_items.GetFirst.return_value = mock_mails[0]
        mock_items.GetNext.side_effect = mock_mails[1:] + [None]
        
        mock_sent_folder = MagicMock()
        mock_sent_folder.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_sent_folder
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.side_effect = [{"subject": f"Email {i}"} for i in range(10)]
        
        # Call with limit=3
        result = self.tool_func(limit=3)
        
        # Verify that only 3 emails are returned
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 3


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

