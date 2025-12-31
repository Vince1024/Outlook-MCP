"""
Outlook Connection Module

Handles connection to Microsoft Outlook via COM automation.
"""

import logging
import win32com.client

logger = logging.getLogger(__name__)


def get_outlook_application():
    """
    Get or create an instance of the Outlook application via COM.
    
    This function establishes a connection to the local Outlook application
    using Windows COM automation. It's used by all other functions to interact
    with Outlook data.
    
    Returns:
        win32com.client.CDispatch: Outlook Application COM object
        
    Raises:
        ValueError: If Outlook is not installed or cannot be accessed via COM
        
    Notes:
        - Requires Microsoft Outlook to be installed on the system
        - The Outlook application must be properly configured with at least one profile
        - This uses late binding (Dispatch) rather than early binding for compatibility
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        logger.error("Failed to connect to Outlook Application", exc_info=True, extra={
            "error_type": type(e).__name__,
            "error_message": str(e)
        })
        raise ValueError(
            f"Unable to connect to Outlook. Make sure Outlook is installed and properly configured. Error: {e}"
        )

