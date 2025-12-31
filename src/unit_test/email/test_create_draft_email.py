"""
Unit tests for create_draft_email tool

This test verifies that the create_draft_email tool works correctly without creating real drafts.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.email.create_draft_email import register


class TestCreateDraftEmail:
    """Tests for create_draft_email tool"""
    
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
    
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_basic(self, mock_outlook):
        """Test basic draft email creation"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func(
            to="test@example.com",
            subject="Draft Subject",
            body="Draft Body"
        )
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert "Draft email created" in result_dict["message"]
        
        # Verify that Outlook methods were called
        mock_app.CreateItem.assert_called_once()
        mock_mail.Save.assert_called_once()
        
        # Verify that Send was NOT called (it's a draft)
        mock_mail.Send.assert_not_called()
        
        # Verify that properties were set
        assert mock_mail.To == "test@example.com"
        assert mock_mail.Subject == "Draft Subject"
        assert mock_mail.Body == "Draft Body"
    
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_with_cc_bcc(self, mock_outlook):
        """Test creating draft with CC and BCC"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Call the tool with CC and BCC
        result = self.tool_func(
            to="test@example.com",
            subject="Test",
            body="Body",
            cc="cc@example.com",
            bcc="bcc@example.com"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that CC and BCC are set
        assert mock_mail.CC == "cc@example.com"
        assert mock_mail.BCC == "bcc@example.com"
    
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_with_html_body(self, mock_outlook):
        """Test creating draft with HTML body"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        html_content = "<html><body><h1>Draft Test</h1></body></html>"
        
        # Call the tool with HTML
        result = self.tool_func(
            to="test@example.com",
            subject="HTML Draft",
            body="Fallback text",
            html_body=html_content
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that HTMLBody is used
        assert mock_mail.HTMLBody == html_content
    
    @patch('tools.email.create_draft_email.get_outlook_signature')
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_with_signature(self, mock_outlook, mock_get_signature):
        """Test creating draft with signature"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Mock Display to simulate failure (fallback to file)
        mock_mail.Display.side_effect = Exception("Display failed")
        
        # Mock signature
        signature_html = "<p>--<br>VP DXT<br>Title</p>"
        mock_get_signature.return_value = signature_html
        
        # Call the tool with signature
        result = self.tool_func(
            to="test@example.com",
            subject="Test",
            body="Message",
            signature_name="VP DXT"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that signature was loaded
        mock_get_signature.assert_called_once_with("VP DXT")
        
        # Verify that HTMLBody contains signature
        assert signature_html in mock_mail.HTMLBody
    
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_with_signature_display_success(self, mock_outlook):
        """Test creating draft with signature using Display()"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Mock Display to succeed
        existing_signature = "<div>Existing signature from Outlook</div>"
        mock_mail.HTMLBody = existing_signature
        
        # Call the tool with signature
        result = self.tool_func(
            to="test@example.com",
            subject="Test",
            body="Message",
            signature_name="VP DXT"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that Display was called
        mock_mail.Display.assert_called_once_with(False)
        
        # Verify that HTMLBody contains the body and signature
        assert existing_signature in mock_mail.HTMLBody
        assert "<p>Message</p>" in mock_mail.HTMLBody or "Message" in mock_mail.HTMLBody
    
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_failure(self, mock_outlook):
        """Test error handling during draft creation"""
        # Setup mock to simulate error
        mock_outlook.side_effect = Exception("Outlook connection failed")
        
        # Call the tool
        result = self.tool_func(
            to="test@example.com",
            subject="Test",
            body="Body"
        )
        
        # Verify that error is handled
        result_dict = json.loads(result)
        assert result_dict["success"] is False
        assert "error" in result_dict
        assert "Outlook connection failed" in result_dict["error"]
    
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_multiple_recipients(self, mock_outlook):
        """Test creating draft with multiple recipients"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Call with multiple recipients
        result = self.tool_func(
            to="user1@example.com; user2@example.com; user3@example.com",
            subject="Multi-recipients Draft",
            body="Draft for multiple people"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that all recipients are in To
        assert "user1@example.com" in mock_mail.To
        assert "user2@example.com" in mock_mail.To
        assert "user3@example.com" in mock_mail.To
    
    @patch('tools.email.create_draft_email.get_outlook_signature')
    @patch('tools.email.create_draft_email.get_outlook_application')
    def test_create_draft_email_html_body_with_signature(self, mock_outlook, mock_get_signature):
        """Test creating draft with both HTML body and signature"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Mock Display to simulate failure (fallback to file)
        mock_mail.Display.side_effect = Exception("Display failed")
        
        # Mock signature
        signature_html = "<p>--<br>VP DXT<br>Title</p>"
        mock_get_signature.return_value = signature_html
        
        html_body = "<html><body><h1>Important Message</h1></body></html>"
        
        # Call the tool with both HTML body and signature
        result = self.tool_func(
            to="test@example.com",
            subject="Test",
            body="Fallback",
            html_body=html_body,
            signature_name="VP DXT"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that HTMLBody contains both HTML body and signature
        assert html_body in mock_mail.HTMLBody or "<h1>Important Message</h1>" in mock_mail.HTMLBody
        assert signature_html in mock_mail.HTMLBody


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

