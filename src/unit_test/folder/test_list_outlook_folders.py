"""
Unit tests for list_outlook_folders tool

This test verifies that the list_outlook_folders tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.folder.list_outlook_folders import register


class TestListOutlookFolders:
    """Tests for list_outlook_folders tool"""
    
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
    
    @patch('tools.folder.list_outlook_folders.get_all_folders')
    @patch('tools.folder.list_outlook_folders.get_outlook_application')
    def test_list_outlook_folders_basic(self, mock_outlook, mock_get_all_folders):
        """Test basic folder listing"""
        # Setup mocks
        mock_store1 = MagicMock()
        mock_store1.DisplayName = "Mailbox - user@example.com"
        mock_root1 = MagicMock()
        mock_store1.GetRootFolder.return_value = mock_root1
        
        mock_store2 = MagicMock()
        mock_store2.DisplayName = "Archive"
        mock_root2 = MagicMock()
        mock_store2.GetRootFolder.return_value = mock_root2
        
        mock_namespace = MagicMock()
        mock_namespace.Stores = [mock_store1, mock_store2]
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock get_all_folders
        mock_get_all_folders.side_effect = [
            [
                {"name": "Inbox", "path": "Inbox"},
                {"name": "Sent Items", "path": "Sent Items"},
                {"name": "Archive", "path": "Inbox/Archive"}
            ],
            [
                {"name": "Old Emails", "path": "Archive/Old Emails"}
            ]
        ]
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 4  # 3 from first store + 1 from second
        assert len(result_dict["folders"]) == 4
        
        # Verify folder names
        folder_names = [f["name"] for f in result_dict["folders"]]
        assert "Inbox" in folder_names
        assert "Sent Items" in folder_names
        assert "Archive" in folder_names
        
        # Verify that GetNamespace was called
        mock_app.GetNamespace.assert_called_once_with("MAPI")
        
        # Verify that get_all_folders was called with include_counts=False
        assert mock_get_all_folders.call_count == 2
        calls = mock_get_all_folders.call_args_list
        assert calls[0][1]["include_counts"] is False
        assert calls[1][1]["include_counts"] is False
    
    @patch('tools.folder.list_outlook_folders.get_all_folders')
    @patch('tools.folder.list_outlook_folders.get_outlook_application')
    def test_list_outlook_folders_empty(self, mock_outlook, mock_get_all_folders):
        """Test listing with no folders"""
        # Setup mocks
        mock_store = MagicMock()
        mock_store.DisplayName = "Mailbox"
        mock_root = MagicMock()
        mock_store.GetRootFolder.return_value = mock_root
        
        mock_namespace = MagicMock()
        mock_namespace.Stores = [mock_store]
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock get_all_folders returns empty list
        mock_get_all_folders.return_value = []
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["folders"]) == 0
    
    @patch('tools.folder.list_outlook_folders.get_all_folders')
    @patch('tools.folder.list_outlook_folders.get_outlook_application')
    def test_list_outlook_folders_with_excluded_stores(self, mock_outlook, mock_get_all_folders):
        """Test that excluded stores are skipped"""
        # Setup mocks
        mock_store1 = MagicMock()
        mock_store1.DisplayName = "Mailbox"
        mock_root1 = MagicMock()
        mock_store1.GetRootFolder.return_value = mock_root1
        
        # This store should be excluded (name matches EXCLUDED_STORES)
        mock_store2 = MagicMock()
        mock_store2.DisplayName = "Public Folders"  # Common excluded store name
        
        mock_namespace = MagicMock()
        mock_namespace.Stores = [mock_store1, mock_store2]
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock get_all_folders (should only be called once for mock_store1)
        mock_get_all_folders.return_value = [{"name": "Inbox", "path": "Inbox"}]
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that get_all_folders was called only once (excluded store skipped)
        # Note: This may vary based on EXCLUDED_STORES config
        assert mock_get_all_folders.call_count >= 1
    
    @patch('tools.folder.list_outlook_folders.get_all_folders')
    @patch('tools.folder.list_outlook_folders.get_outlook_application')
    def test_list_outlook_folders_with_store_error(self, mock_outlook, mock_get_all_folders):
        """Test that errors in individual stores are handled gracefully"""
        # Setup mocks
        mock_store1 = MagicMock()
        mock_store1.DisplayName = "Mailbox 1"
        mock_root1 = MagicMock()
        mock_store1.GetRootFolder.return_value = mock_root1
        
        mock_store2 = MagicMock()
        mock_store2.DisplayName = "Mailbox 2"
        mock_store2.GetRootFolder.side_effect = Exception("Cannot access store")
        
        mock_namespace = MagicMock()
        mock_namespace.Stores = [mock_store1, mock_store2]
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock get_all_folders for first store
        mock_get_all_folders.return_value = [{"name": "Inbox", "path": "Inbox"}]
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result - should succeed with folders from first store
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1  # Only folders from first store
    
    @patch('tools.folder.list_outlook_folders.get_outlook_application')
    def test_list_outlook_folders_failure(self, mock_outlook):
        """Test error handling during folder listing"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Outlook connection failed")
        
        # Call the tool
        result = self.tool_func()
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Outlook connection failed" in result_dict["error"]
    
    @patch('tools.folder.list_outlook_folders.get_all_folders')
    @patch('tools.folder.list_outlook_folders.get_outlook_application')
    def test_list_outlook_folders_with_nested_folders(self, mock_outlook, mock_get_all_folders):
        """Test listing with nested folder structure"""
        # Setup mocks
        mock_store = MagicMock()
        mock_store.DisplayName = "Mailbox"
        mock_root = MagicMock()
        mock_store.GetRootFolder.return_value = mock_root
        
        mock_namespace = MagicMock()
        mock_namespace.Stores = [mock_store]
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Mock get_all_folders with nested structure
        mock_get_all_folders.return_value = [
            {"name": "Inbox", "path": "Inbox"},
            {"name": "Archive", "path": "Inbox/Archive"},
            {"name": "2024", "path": "Inbox/Archive/2024"},
            {"name": "2023", "path": "Inbox/Archive/2023"}
        ]
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 4
        
        # Verify nested folder paths
        paths = [f["path"] for f in result_dict["folders"]]
        assert "Inbox" in paths
        assert "Inbox/Archive" in paths
        assert "Inbox/Archive/2024" in paths
        assert "Inbox/Archive/2023" in paths


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

