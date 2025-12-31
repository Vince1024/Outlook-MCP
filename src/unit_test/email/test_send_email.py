"""
Unit tests for send_email tool

This test verifies that the send_email tool works correctly without sending real emails.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path))

from tools.email.send_email import register


class TestSendEmail:
    """Tests for send_email tool"""
    
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
    
    @patch('tools.email.send_email.get_outlook_application')
    def test_send_email_basic(self, mock_outlook):
        """Test basic email sending"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Call the tool
        result = self.tool_func(
            to="test@example.com",
            subject="Test Subject",
            body="Test Body"
        )
        
        # Verify result
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        assert "test@example.com" in result_dict["message"]
        
        # Verify that Outlook methods were called
        mock_app.CreateItem.assert_called_once()
        mock_mail.Send.assert_called_once()
        
        # Verify that properties were set
        assert mock_mail.To == "test@example.com"
        assert mock_mail.Subject == "Test Subject"
        assert mock_mail.Body == "Test Body"
    
    @patch('tools.email.send_email.get_outlook_application')
    def test_send_email_with_cc_bcc(self, mock_outlook):
        """Test sending with CC and BCC"""
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
    
    @patch('tools.email.send_email.get_outlook_application')
    def test_send_email_with_importance(self, mock_outlook):
        """Test sending with high importance"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Call the tool with high importance
        result = self.tool_func(
            to="test@example.com",
            subject="Urgent",
            body="Urgent message",
            importance="high"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that importance is set to 2 (IMPORTANCE_HIGH)
        assert mock_mail.Importance == 2
    
    @patch('tools.email.send_email.get_outlook_application')
    def test_send_email_with_html_body(self, mock_outlook):
        """Test sending with HTML body"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        html_content = "<html><body><h1>Test</h1></body></html>"
        
        # Call the tool with HTML
        result = self.tool_func(
            to="test@example.com",
            subject="HTML Test",
            body="Fallback text",
            html_body=html_content
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that HTMLBody is used
        assert mock_mail.HTMLBody == html_content
    
    @patch('tools.email.send_email.get_outlook_signature')
    @patch('tools.email.send_email.get_outlook_application')
    def test_send_email_with_signature(self, mock_outlook, mock_get_signature):
        """Test sending with signature"""
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
    
    @patch('tools.email.send_email.get_outlook_application')
    def test_send_email_failure(self, mock_outlook):
        """Test error handling during sending"""
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
    
    @patch('tools.email.send_email.get_outlook_application')
    def test_send_email_multiple_recipients(self, mock_outlook):
        """Test sending to multiple recipients"""
        # Setup mocks
        mock_mail = MagicMock()
        mock_app = MagicMock()
        mock_app.CreateItem.return_value = mock_mail
        mock_outlook.return_value = mock_app
        
        # Call with multiple recipients
        result = self.tool_func(
            to="user1@example.com; user2@example.com; user3@example.com",
            subject="Multi-recipients",
            body="Message for all"
        )
        
        # Verify
        result_dict = json.loads(result)
        assert result_dict["success"] is True
        
        # Verify that all recipients are in To
        assert "user1@example.com" in mock_mail.To
        assert "user2@example.com" in mock_mail.To
        assert "user3@example.com" in mock_mail.To


# Integration tests (optional, skipped by default because they require Outlook)
class TestSendEmailIntegration:
    """
    Real integration tests with Outlook
    
    WARNING: These tests send real emails!
    Use only in development with test email addresses.
    """
    
    @pytest.mark.skip(reason="Integration test - requires Outlook and sends real emails")
    def test_real_send_email(self):
        """Real email sending test (skipped by default)"""
        from tools.email.send_email import register
        
        mcp_mock = Mock()
        tool_func = None
        
        def capture_tool(func):
            nonlocal tool_func
            tool_func = func
            return func
        
        mcp_mock.tool.return_value = capture_tool
        register(mcp_mock)
        
        # Send a real test email
        result = tool_func(
            to="test@example.com",
            subject="Unit Test - send_email",
            body="This is an automated unit test",
            signature_name="VP DXT"
        )
        
        result_dict = json.loads(result)
        assert result_dict["success"] is True


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v", "--tb=short"])

