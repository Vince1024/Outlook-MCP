"""
Utilities package for MCP Outlook Server

This package contains reusable utility functions and helpers.
"""

from .outlook_connection import get_outlook_application
from .formatters import format_email, format_appointment, format_contact
from .signature_helper import get_outlook_signature, get_outlook_signature_via_display
from .folder_helper import get_folder_by_path, get_all_folders, FOLDER_CACHE

__all__ = [
    'get_outlook_application',
    'format_email',
    'format_appointment',
    'format_contact',
    'get_outlook_signature',
    'get_outlook_signature_via_display',
    'get_folder_by_path',
    'get_all_folders',
    'FOLDER_CACHE',
]

