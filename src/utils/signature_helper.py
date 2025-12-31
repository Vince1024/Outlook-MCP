"""
Signature Helper Module

Handles loading and applying Outlook email signatures.
"""

import os
from typing import Optional


def get_outlook_signature_via_display(mail_item, signature_name: Optional[str] = None) -> str:
    """
    Get the signature by displaying the mail item temporarily.
    This allows Outlook to insert the default or specified signature naturally.
    
    Args:
        mail_item: The Outlook mail item object
        signature_name: Optional signature name (not used for now, uses default)
        
    Returns:
        The HTMLBody with the signature inserted by Outlook
    """
    try:
        # Display the mail item (this triggers Outlook to add the default signature)
        # But don't show the window to the user
        mail_item.Display(False)  # False = don't show window
        signature_html = mail_item.HTMLBody
        mail_item.Close(1)  # 1 = olDiscard, don't save changes
        
        return signature_html
        
    except Exception:
        return ""


def get_outlook_signature(signature_name: str) -> Optional[str]:
    """
    Load an Outlook signature by name from the signatures folder.
    
    Args:
        signature_name: Name of the signature (without extension)
        
    Returns:
        HTML content of the signature, or None if not found
    """
    # Get the user's Outlook signatures folder
    signatures_path = os.path.join(
        os.environ.get('APPDATA', ''),
        'Microsoft',
        'Signatures'
    )
    
    # Try to find the signature file (with or without email suffix)
    signature_file = os.path.join(signatures_path, f"{signature_name}.htm")
    
    if not os.path.exists(signature_file):
        # Try with common suffixes
        try:
            for file in os.listdir(signatures_path):
                if file.startswith(signature_name) and file.endswith('.htm'):
                    signature_file = os.path.join(signatures_path, file)
                    break
        except Exception:
            return None
    
    if os.path.exists(signature_file):
        try:
            with open(signature_file, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        except Exception:
            return None
    
    return None

