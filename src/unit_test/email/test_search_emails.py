"""
Unit tests for search_emails tool

This test verifies that the search_emails tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.email.search_emails import register


class TestSearchEmails:
    """Tests for search_emails tool"""
    
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
    
    @patch('tools.email.search_emails.format_email')
    @patch('tools.email.search_emails.get_outlook_application')
    def test_search_emails_inbox_basic(self, mock_outlook, mock_format):
        """Test basic email search in inbox"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Restrict.return_value = mock_items
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
            {"subject": "Payment received", "sender": "finance@example.com"},
            {"subject": "Invoice payment", "sender": "accounting@example.com"}
        ]
        
        # Call the tool
        result = self.tool_func(query="payment", folder="inbox", limit=10)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["query"] == "payment"
        assert result_dict["count"] == 2
        assert len(result_dict["emails"]) == 2
        
        # Verify that Restrict was called (search filter)
        mock_items.Restrict.assert_called_once()
        mock_items.Sort.assert_called_once_with("[ReceivedTime]", True)
    
    @patch('tools.email.search_emails.format_email')
    @patch('tools.email.search_emails.get_outlook_application')
    def test_search_emails_in_sent_folder(self, mock_outlook, mock_format):
        """Test searching in sent folder"""
        # Setup mocks
        mock_mail1 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Restrict.return_value = mock_items
        mock_items.Sort = MagicMock()
        mock_items.GetFirst.return_value = mock_mail1
        mock_items.GetNext.return_value = None
        
        mock_sent = MagicMock()
        mock_sent.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_sent
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.return_value = {"subject": "Report sent", "to": "manager@example.com"}
        
        # Call the tool with sent folder
        result = self.tool_func(query="report", folder="sent", limit=10)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
    
    @patch('tools.email.search_emails.format_email')
    @patch('tools.email.search_emails.get_outlook_application')
    def test_search_emails_all_folders(self, mock_outlook, mock_format):
        """Test searching across all folders (inbox, sent, drafts)"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        mock_mail3 = MagicMock()
        
        # Mock for first folder (inbox)
        mock_items1 = MagicMock()
        mock_items1.Restrict.return_value = mock_items1
        mock_items1.Sort = MagicMock()
        mock_items1.GetFirst.return_value = mock_mail1
        mock_items1.GetNext.return_value = None
        
        # Mock for second folder (sent)
        mock_items2 = MagicMock()
        mock_items2.Restrict.return_value = mock_items2
        mock_items2.Sort = MagicMock()
        mock_items2.GetFirst.return_value = mock_mail2
        mock_items2.GetNext.return_value = None
        
        # Mock for third folder (drafts)
        mock_items3 = MagicMock()
        mock_items3.Restrict.return_value = mock_items3
        mock_items3.Sort = MagicMock()
        mock_items3.GetFirst.return_value = mock_mail3
        mock_items3.GetNext.return_value = None
        
        mock_folder1 = MagicMock()
        mock_folder1.Items = mock_items1
        mock_folder2 = MagicMock()
        mock_folder2.Items = mock_items2
        mock_folder3 = MagicMock()
        mock_folder3.Items = mock_items3
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.side_effect = [mock_folder1, mock_folder2, mock_folder3]
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock format_email
        mock_format.side_effect = [
            {"subject": "Project update 1"},
            {"subject": "Project update 2"},
            {"subject": "Project update 3"}
        ]
        
        # Call the tool with folder="all"
        result = self.tool_func(query="project", folder="all", limit=10)
        
        # Verify result - should have searched 3 folders
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 3
        
        # Verify that GetDefaultFolder was called 3 times (inbox, sent, drafts)
        assert mock_namespace.GetDefaultFolder.call_count == 3
    
    @patch('tools.email.search_emails.get_outlook_application')
    def test_search_emails_no_results(self, mock_outlook):
        """Test search with no results"""
        # Setup mocks for empty search results
        mock_items = MagicMock()
        mock_items.Restrict.return_value = mock_items
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
        result = self.tool_func(query="nonexistent", folder="inbox", limit=10)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["emails"]) == 0
    
    @patch('tools.email.search_emails.format_email')
    @patch('tools.email.search_emails.get_outlook_application')
    def test_search_emails_limit_enforced(self, mock_outlook, mock_format):
        """Test that limit is enforced"""
        # Setup mocks - simulate many search results
        mock_items = MagicMock()
        mock_items.Restrict.return_value = mock_items
        mock_items.Sort = MagicMock()
        
        # Create 60 mock emails (more than MAX_EMAIL_LIMIT=50)
        mock_mails = [MagicMock() for _ in range(60)]
        mock_items.GetFirst.return_value = mock_mails[0]
        mock_items.GetNext.side_effect = mock_mails[1:] + [None]
        
        mock_inbox = MagicMock()
        mock_inbox.Items = mock_items
        
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_format.side_effect = [{"subject": f"Email {i}"} for i in range(60)]
        
        # Call with high limit
        result = self.tool_func(query="test", folder="inbox", limit=100)
        
        # Verify that only MAX_EMAIL_LIMIT (50) emails are returned
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] <= 50
    
    @patch('tools.email.search_emails.get_outlook_application')
    def test_search_emails_failure(self, mock_outlook):
        """Test error handling during search"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Search failed")
        
        # Call the tool
        result = self.tool_func(query="test", folder="inbox", limit=10)
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Search failed" in result_dict["error"]
    
    @patch('tools.email.search_emails.format_email')
    @patch('tools.email.search_emails.get_outlook_application')
    def test_search_emails_with_format_error(self, mock_outlook, mock_format):
        """Test that formatting errors are handled gracefully"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        mock_mail3 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Restrict.return_value = mock_items
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
        
        # Mock format_email - middle one fails
        mock_format.side_effect = [
            {"subject": "Email 1"},
            Exception("Format error"),  # This email will be skipped
            {"subject": "Email 3"}
        ]
        
        # Call the tool
        result = self.tool_func(query="test", folder="inbox", limit=10)
        
        # Verify result - should have 2 emails (skipped the error one)
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        assert result_dict["emails"][0]["subject"] == "Email 1"
        assert result_dict["emails"][1]["subject"] == "Email 3"


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

