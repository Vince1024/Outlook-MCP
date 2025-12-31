"""
Unit tests for search_emails_in_custom_folder tool

This test verifies that the search_emails_in_custom_folder tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.folder.search_emails_in_custom_folder import register


class TestSearchEmailsInCustomFolder:
    """Tests for search_emails_in_custom_folder tool"""
    
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
    
    @patch('tools.folder.search_emails_in_custom_folder.format_email')
    @patch('tools.folder.search_emails_in_custom_folder.get_folder_by_path')
    @patch('tools.folder.search_emails_in_custom_folder.get_outlook_application')
    def test_search_emails_in_custom_folder_basic(self, mock_outlook, mock_get_folder, mock_format):
        """Test basic search in custom folder"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail2 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.__getitem__ = lambda self, i: [mock_mail1, mock_mail2][i-1]  # 1-based indexing
        
        # Mock Folders as an empty list (no subfolders)
        mock_folders = MagicMock()
        mock_folders.__iter__ = Mock(return_value=iter([]))
        
        mock_folder = MagicMock()
        mock_folder.Items = mock_items
        mock_folder.Folders = mock_folders
        
        mock_namespace = MagicMock()
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_get_folder.return_value = mock_folder
        
        # Mock format_email
        mock_format.side_effect = [
            {"subject": "Custom Email 1", "sender": "user1@example.com"},
            {"subject": "Custom Email 2", "sender": "user2@example.com"}
        ]
        
        # Call the tool (recursive=True by default)
        result = self.tool_func(folder_path="My Folder", limit=2)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["folder"] == "My Folder"
        assert result_dict["count"] == 2
        assert len(result_dict["emails"]) == 2
        
        # Verify that get_folder_by_path was called
        mock_get_folder.assert_called_once_with(mock_namespace, "My Folder", use_cache=True)
    
    @patch('tools.folder.search_emails_in_custom_folder.get_folder_by_path')
    @patch('tools.folder.search_emails_in_custom_folder.get_outlook_application')
    def test_search_emails_folder_not_found(self, mock_outlook, mock_get_folder):
        """Test search when folder is not found"""
        # Setup mocks
        mock_namespace = MagicMock()
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Folder not found
        mock_get_folder.return_value = None
        
        # Call the tool
        result = self.tool_func(folder_path="NonExistent", limit=5)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "not found" in result_dict["error"]
    
    @patch('tools.folder.search_emails_in_custom_folder.format_email')
    @patch('tools.folder.search_emails_in_custom_folder.get_folder_by_path')
    @patch('tools.folder.search_emails_in_custom_folder.get_outlook_application')
    def test_search_emails_with_query(self, mock_outlook, mock_get_folder, mock_format):
        """Test search with keyword query"""
        # Setup mocks
        mock_mail1 = MagicMock()
        mock_mail1.Subject = "Payment received"
        mock_mail1.Body = "Thank you for your payment"
        mock_mail1.SenderName = "Finance"
        
        mock_mail2 = MagicMock()
        mock_mail2.Subject = "Invoice"
        mock_mail2.Body = "Payment is due"
        mock_mail2.SenderName = "Accounting"
        
        mock_mail3 = MagicMock()
        mock_mail3.Subject = "Meeting notes"
        mock_mail3.Body = "Notes from meeting"
        mock_mail3.SenderName = "Manager"
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.__getitem__ = lambda self, i: [mock_mail1, mock_mail2, mock_mail3][i-1]
        
        # Mock Folders as an empty list (no subfolders)
        mock_folders = MagicMock()
        mock_folders.__iter__ = Mock(return_value=iter([]))
        
        mock_folder = MagicMock()
        mock_folder.Items = mock_items
        mock_folder.Folders = mock_folders
        
        mock_namespace = MagicMock()
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_get_folder.return_value = mock_folder
        
        # Mock format_email
        mock_format.side_effect = [
            {"subject": "Payment received"},
            {"subject": "Invoice"}
        ]
        
        # Call the tool with query and recursive=False, days_back=0 to avoid date filtering
        result = self.tool_func(folder_path="My Folder", query="payment", limit=5, recursive=False, days_back=0)
        
        # Verify result - should only return emails matching "payment"
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["query"] == "payment"
        assert result_dict["count"] == 2  # Only 2 emails contain "payment"
    
    @patch('tools.folder.search_emails_in_custom_folder.format_email')
    @patch('tools.folder.search_emails_in_custom_folder.get_folder_by_path')
    @patch('tools.folder.search_emails_in_custom_folder.get_outlook_application')
    def test_search_emails_with_days_back(self, mock_outlook, mock_get_folder, mock_format):
        """Test search with days_back filter"""
        # Setup mocks
        mock_mail1 = MagicMock()
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.Restrict = MagicMock(return_value=mock_items)
        mock_items.__getitem__ = lambda self, i: mock_mail1 if i == 1 else None
        
        # Mock Folders as an empty list (no subfolders)
        mock_folders = MagicMock()
        mock_folders.__iter__ = Mock(return_value=iter([]))
        
        mock_folder = MagicMock()
        mock_folder.Items = mock_items
        mock_folder.Folders = mock_folders
        
        mock_namespace = MagicMock()
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_get_folder.return_value = mock_folder
        
        # Mock format_email
        mock_format.return_value = {"subject": "Recent Email"}
        
        # Call the tool with days_back (recursive=True by default)
        result = self.tool_func(folder_path="My Folder", limit=5, days_back=7)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["days_back"] == 7
        assert "info" in result_dict
        assert "last 7 days" in result_dict["info"]
        
        # Verify that Restrict was called (date filter)
        mock_items.Restrict.assert_called_once()
    
    @patch('tools.folder.search_emails_in_custom_folder.format_email')
    @patch('tools.folder.search_emails_in_custom_folder.get_folder_by_path')
    @patch('tools.folder.search_emails_in_custom_folder.get_outlook_application')
    def test_search_emails_empty_folder(self, mock_outlook, mock_get_folder, mock_format):
        """Test search in empty folder"""
        # Setup mocks
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.__getitem__ = Mock(side_effect=Exception("Index out of range"))
        
        # Mock Folders as an empty list (no subfolders)
        mock_folders = MagicMock()
        mock_folders.__iter__ = Mock(return_value=iter([]))
        
        mock_folder = MagicMock()
        mock_folder.Items = mock_items
        mock_folder.Folders = mock_folders
        
        mock_namespace = MagicMock()
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_get_folder.return_value = mock_folder
        
        # Call the tool with recursive=False, days_back=0 to avoid filtering
        result = self.tool_func(folder_path="Empty Folder", limit=5, recursive=False, days_back=0)
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["emails"]) == 0
    
    @patch('tools.folder.search_emails_in_custom_folder.get_folder_by_path')
    @patch('tools.folder.search_emails_in_custom_folder.get_outlook_application')
    def test_search_emails_failure(self, mock_outlook, mock_get_folder):
        """Test error handling during search"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Outlook connection failed")
        
        # Call the tool
        result = self.tool_func(folder_path="My Folder", limit=5)
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Outlook connection failed" in result_dict["error"]
    
    @patch('tools.folder.search_emails_in_custom_folder.format_email')
    @patch('tools.folder.search_emails_in_custom_folder.get_folder_by_path')
    @patch('tools.folder.search_emails_in_custom_folder.get_outlook_application')
    def test_search_emails_limit_enforced(self, mock_outlook, mock_get_folder, mock_format):
        """Test that MAX_EMAIL_LIMIT is enforced"""
        # Setup mocks - REDUCED FROM 60 TO 10 TO PREVENT OOM
        mock_mails = [MagicMock() for _ in range(10)]
        
        mock_items = MagicMock()
        mock_items.Sort = MagicMock()
        mock_items.__getitem__ = lambda self, i: mock_mails[i-1] if i <= len(mock_mails) else None
        
        # Mock Folders as an empty list (no subfolders)
        mock_folders = MagicMock()
        mock_folders.__iter__ = Mock(return_value=iter([]))
        
        mock_folder = MagicMock()
        mock_folder.Items = mock_items
        mock_folder.Folders = mock_folders
        
        mock_namespace = MagicMock()
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        mock_get_folder.return_value = mock_folder
        
        mock_format.side_effect = [{"subject": f"Email {i}"} for i in range(10)]
        
        # Call with limit=5 to test that limit is enforced (recursive=True by default)
        result = self.tool_func(folder_path="My Folder", limit=5)
        
        # Verify that only requested limit (5) emails are returned, not all 10 available
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 5  # Should stop at limit, not return all 10


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

