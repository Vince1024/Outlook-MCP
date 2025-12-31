"""
Unit tests for list_outlook_rules tool

This test verifies that the list_outlook_rules tool works correctly without accessing real Outlook.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.folder.list_outlook_rules import register


class TestListOutlookRules:
    """Tests for list_outlook_rules tool"""
    
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
    
    @patch('tools.folder.list_outlook_rules.get_outlook_application')
    def test_list_outlook_rules_basic(self, mock_outlook):
        """Test basic rule listing"""
        # Setup mocks for rules
        mock_rule1 = MagicMock()
        mock_rule1.Name = "Move Work Emails"
        mock_rule1.Enabled = True
        
        # Mock conditions for rule1
        mock_conditions1 = MagicMock()
        mock_subject_condition = MagicMock()
        mock_subject_condition.Enabled = True
        mock_subject_condition.Text = ["work", "project"]
        mock_conditions1.Subject = mock_subject_condition
        mock_rule1.Conditions = mock_conditions1
        
        # Mock actions for rule1
        mock_actions1 = MagicMock()
        mock_move_action = MagicMock()
        mock_move_action.Enabled = True
        mock_folder = MagicMock()
        mock_folder.Name = "Work"
        mock_move_action.Folder = mock_folder
        mock_actions1.MoveToFolder = mock_move_action
        mock_rule1.Actions = mock_actions1
        
        # Second rule
        mock_rule2 = MagicMock()
        mock_rule2.Name = "Delete Spam"
        mock_rule2.Enabled = False
        
        mock_conditions2 = MagicMock()
        mock_rule2.Conditions = mock_conditions2
        
        mock_actions2 = MagicMock()
        mock_delete_action = MagicMock()
        mock_delete_action.Enabled = True
        mock_actions2.Delete = mock_delete_action
        mock_rule2.Actions = mock_actions2
        
        # Mock rules collection
        mock_rules_collection = [mock_rule1, mock_rule2]
        
        mock_default_store = MagicMock()
        mock_default_store.GetRules.return_value = mock_rules_collection
        
        mock_namespace = MagicMock()
        mock_namespace.DefaultStore = mock_default_store
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        assert len(result_dict["rules"]) == 2
        
        # Verify first rule
        rule1 = result_dict["rules"][0]
        assert rule1["name"] == "Move Work Emails"
        assert rule1["enabled"] is True
        assert len(rule1["conditions"]) > 0
        assert len(rule1["actions"]) > 0
        
        # Verify second rule
        rule2 = result_dict["rules"][1]
        assert rule2["name"] == "Delete Spam"
        assert rule2["enabled"] is False
    
    @patch('tools.folder.list_outlook_rules.get_outlook_application')
    def test_list_outlook_rules_empty(self, mock_outlook):
        """Test listing with no rules"""
        # Setup mocks
        mock_rules_collection = []
        
        mock_default_store = MagicMock()
        mock_default_store.GetRules.return_value = mock_rules_collection
        
        mock_namespace = MagicMock()
        mock_namespace.DefaultStore = mock_default_store
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 0
        assert len(result_dict["rules"]) == 0
    
    @patch('tools.folder.list_outlook_rules.get_outlook_application')
    def test_list_outlook_rules_with_from_condition(self, mock_outlook):
        """Test rule with From condition"""
        # Setup mock rule with From condition
        mock_rule = MagicMock()
        mock_rule.Name = "From Boss"
        mock_rule.Enabled = True
        
        # Mock conditions
        mock_conditions = MagicMock()
        mock_from_condition = MagicMock()
        mock_from_condition.Enabled = True
        mock_recipient1 = MagicMock()
        mock_recipient1.Name = "Boss"
        mock_from_condition.Recipients = [mock_recipient1]
        mock_conditions.From = mock_from_condition
        mock_rule.Conditions = mock_conditions
        
        # Mock actions
        mock_actions = MagicMock()
        mock_mark_read = MagicMock()
        mock_mark_read.Enabled = True
        mock_actions.MarkAsRead = mock_mark_read
        mock_rule.Actions = mock_actions
        
        mock_rules_collection = [mock_rule]
        
        mock_default_store = MagicMock()
        mock_default_store.GetRules.return_value = mock_rules_collection
        
        mock_namespace = MagicMock()
        mock_namespace.DefaultStore = mock_default_store
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 1
        
        rule = result_dict["rules"][0]
        assert rule["name"] == "From Boss"
        assert any("From: Boss" in cond for cond in rule["conditions"])
        assert any("Mark as read" in action for action in rule["actions"])
    
    @patch('tools.folder.list_outlook_rules.get_outlook_application')
    def test_list_outlook_rules_with_body_condition(self, mock_outlook):
        """Test rule with Body condition"""
        # Setup mock rule
        mock_rule = MagicMock()
        mock_rule.Name = "Body Filter"
        mock_rule.Enabled = True
        
        # Mock conditions
        mock_conditions = MagicMock()
        mock_body_condition = MagicMock()
        mock_body_condition.Enabled = True
        mock_body_condition.Text = ["urgent", "important"]
        mock_conditions.Body = mock_body_condition
        mock_rule.Conditions = mock_conditions
        
        # Mock actions
        mock_actions = MagicMock()
        mock_rule.Actions = mock_actions
        
        mock_rules_collection = [mock_rule]
        
        mock_default_store = MagicMock()
        mock_default_store.GetRules.return_value = mock_rules_collection
        
        mock_namespace = MagicMock()
        mock_namespace.DefaultStore = mock_default_store
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        rule = result_dict["rules"][0]
        assert any("Body contains" in cond for cond in rule["conditions"])
    
    @patch('tools.folder.list_outlook_rules.get_outlook_application')
    def test_list_outlook_rules_with_parse_error(self, mock_outlook):
        """Test that rule parse errors are handled gracefully"""
        # Setup mock rule that will fail to parse
        mock_rule1 = MagicMock()
        mock_rule1.Name = "Good Rule"
        mock_rule1.Enabled = True
        mock_rule1.Conditions = MagicMock()
        mock_rule1.Actions = MagicMock()
        
        # This rule will cause an error during parsing
        mock_rule2 = MagicMock()
        mock_rule2.Name = "Bad Rule"
        mock_rule2.Enabled = True
        # Accessing Conditions will raise an exception
        type(mock_rule2).Conditions = property(lambda self: (_ for _ in ()).throw(Exception("Parse error")))
        
        mock_rules_collection = [mock_rule1, mock_rule2]
        
        mock_default_store = MagicMock()
        mock_default_store.GetRules.return_value = mock_rules_collection
        
        mock_namespace = MagicMock()
        mock_namespace.DefaultStore = mock_default_store
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result - should succeed with 2 rules (second has error description)
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert result_dict["count"] == 2
        
        # Verify that bad rule has error description
        rule2 = result_dict["rules"][1]
        assert rule2["name"] == "Bad Rule"
        assert "Error parsing rule" in rule2["description"]
    
    @patch('tools.folder.list_outlook_rules.get_outlook_application')
    def test_list_outlook_rules_failure(self, mock_outlook):
        """Test error handling during rule listing"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Outlook connection failed")
        
        # Call the tool
        result = self.tool_func()
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Outlook connection failed" in result_dict["error"]
    
    @patch('tools.folder.list_outlook_rules.get_outlook_application')
    def test_list_outlook_rules_with_nested_folder_path(self, mock_outlook):
        """Test rule with nested folder path in MoveToFolder action"""
        # Setup mock rule
        mock_rule = MagicMock()
        mock_rule.Name = "Move to nested"
        mock_rule.Enabled = True
        mock_rule.Conditions = MagicMock()
        
        # Mock actions with nested folder
        mock_actions = MagicMock()
        mock_move_action = MagicMock()
        mock_move_action.Enabled = True
        
        # Create nested folder structure: Inbox/Archive/2024
        mock_folder = MagicMock()
        mock_folder.Name = "2024"
        
        mock_parent1 = MagicMock()
        mock_parent1.Name = "Archive"
        mock_folder.Parent = mock_parent1
        
        mock_parent2 = MagicMock()
        mock_parent2.Name = "Inbox"
        mock_parent1.Parent = mock_parent2
        
        # Root has no Parent attribute
        del type(mock_parent2).Parent
        
        mock_move_action.Folder = mock_folder
        mock_actions.MoveToFolder = mock_move_action
        mock_rule.Actions = mock_actions
        
        mock_rules_collection = [mock_rule]
        
        mock_default_store = MagicMock()
        mock_default_store.GetRules.return_value = mock_rules_collection
        
        mock_namespace = MagicMock()
        mock_namespace.DefaultStore = mock_default_store
        
        mock_app = MagicMock()
        mock_app.GetNamespace.return_value = mock_namespace
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func()
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        rule = result_dict["rules"][0]
        # Should show full path
        assert any("Move to folder:" in action for action in rule["actions"])


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

