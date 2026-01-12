"""
Attachment Tools Package

Provides tools for managing email attachments in Outlook.
"""

from .get_email_attachments import register as get_email_attachments
from .download_email_attachment import register as download_email_attachment
from .send_email_with_attachments import register as send_email_with_attachments

__all__ = [
    'get_email_attachments',
    'download_email_attachment',
    'send_email_with_attachments',
]
