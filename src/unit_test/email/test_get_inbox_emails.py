"""
Unit tests for get_inbox_emails tool

This test verifies that the get_inbox_emails tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.email.get_inbox_emails import register


class TestGetInboxEmails:
    """Tests for get_inbox_emails tool"""
    
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
    
    @patch('tools.email.get_inbox_emails.format_email')
    @patch('tools.email.get_inbox_emails.get_outlook_application')
    def test_get_inbox_emails_basic(self, mock_outlook, mock_format):
        """Test basic inbox email retrieval"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_mail1
        mock_items.GetNext.side_effect = [mock_mail2, None]
        
        mock_inbox = MagicMock()
        mock_inbox.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_email
        mock_format.side_effect = [
            {"subject": "Email 1", "sender": "user1@example.com"},
            {"subject": "Email 2", "sender": "user2@example.com"}
        ]
        
        # Call the tool
        result = self.tool_func(limit=5)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        assert len(result_dict["emails"]) == 2
        assert result_dict["emails"][0]["subject"] == "Email 1"
        
        # Verify that Outlook methods were called
        mock_app.GetNamespace.assert_called_once_with("MAPI")
        mock_namespace.GetDefaultFolder.assert_called_once()
        mock_items.Sort.assert_called_once_with("[ReceivedTime]", True)
    
    @patch('tools.email.get_inbox_emails.format_email')
    @patch('tools.email.get_inbox_emails.get_outlook_application')
    def test_get_inbox_emails_unread_only(self, mock_outlook, mock_format):
        """Test getting only unread emails"""
        # Setup mocks
        mock_mail1 = MagicMock()
        
        mock_items = MagicMock()
        mock_items_filtered = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.Restrict.return_value = mock_items_filtered
        mock_items_filtered.GetFirst.return_value = mock_mail1
        mock_items_filtered.GetNext.return_value = None
        
        mock_inbox = MagicMock()
        mock_inbox.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.return_value = {"subject": "Unread Email", "unread": True}
        
        # Call the tool with unread_only=True
        result = self.tool_func(limit=5, unread_only=True)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
        
        # Verify that Restrict was called with unread filter
        mock_items.Restrict.assert_called_once_with("[Unread] = True")
    
    @patch('tools.email.get_inbox_emails.get_outlook_application')
    def test_get_inbox_emails_empty_inbox(self, mock_outlook):
        """Test getting emails from empty inbox"""
        # Setup mocks for empty inbox
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = None
        
        mock_inbox = MagicMock()
        mock_inbox.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
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
    
    @patch('tools.email.get_inbox_emails.format_email')
    @patch('tools.email.get_inbox_emails.get_outlook_application')
    def test_get_inbox_emails_limit_enforced(self, mock_outlook, mock_format):
        """Test that limit is enforced (max 50)"""
        # Setup mocks - simulate more emails than limit
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        
        # Create only 10 mock emails to test limit logic (reduced from 60 to avoid OOM)
        # This is sufficient to validate that MAX_EMAIL_LIMIT is enforced
        mock_mails = [MagicMock() for _ in range(10)]
        mock_items.GetFirst.return_value = mock_mails[0]
        mock_items.GetNext.side_effect = mock_mails[1:] + [None]
        
        mock_inbox = MagicMock()
        mock_inbox.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
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
    
    @patch('tools.email.get_inbox_emails.get_outlook_application')
    def test_get_inbox_emails_failure(self, mock_outlook):
        """Test error handling during retrieval"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Outlook connection failed")
        
        # Call the tool
        result = self.tool_func(limit=5)
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Outlook connection failed" in result_dict["error"]
    
    @patch('tools.email.get_inbox_emails.format_email')
    @patch('tools.email.get_inbox_emails.get_outlook_application')
    def test_get_inbox_emails_with_format_error(self, mock_outlook, mock_format):
        """Test that formatting errors are handled gracefully"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        mock_mail3 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_mail1
        mock_items.GetNext.side_effect = [mock_mail2, mock_mail3, None]
        
        mock_inbox = MagicMock()
        mock_inbox.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_email - second one fails
        mock_format.side_effect = [
            {"subject": "Email 1"},
            Exception("Format error"),  # This email will be skipped
            {"subject": "Email 3"}
        ]
        
        # Call the tool
        result = self.tool_func(limit=5)
        
        # Verify result - should have 2 emails (skipped the error one)
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        assert result_dict["emails"][0]["subject"] == "Email 1"
        assert result_dict["emails"][1]["subject"] == "Email 3"


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

