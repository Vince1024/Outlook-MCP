"""
Folder Helper Module

Handles Outlook folder operations including path resolution and traversal.
"""

import logging
from typing import Dict, Any, Optional, List
from config import EXCLUDED_STORES

logger = logging.getLogger(__name__)

# Performance cache: folder_path -> Outlook Folder object
FOLDER_CACHE: Dict[str, Any] = {}


def get_folder_by_path(namespace, folder_path: str, use_cache: bool = True):
    """
    Get an Outlook folder by its path with caching support.
    
    Performance optimization: This function caches folder objects to avoid
    the expensive traversal of all Outlook stores on every request.
    
    Args:
        namespace: Outlook MAPI namespace object
        folder_path: Full path to the folder (e.g., "Inbox/Archive" or "Personal/My Mails")
        use_cache: Whether to use the folder cache (default: True)
        
    Returns:
        Outlook Folder object if found, None otherwise
        
    Notes:
        - First access to a folder may take 20-45 seconds (store traversal)
        - Subsequent accesses use cache and take ~0.01 seconds
        - Cache is invalidated when Outlook is restarted
    """
    # Check cache first
    if use_cache and folder_path in FOLDER_CACHE:
        try:
            # Verify cached folder is still valid
            _ = FOLDER_CACHE[folder_path].Name
            return FOLDER_CACHE[folder_path]
        except Exception:
            # Cache entry is stale, remove it
            del FOLDER_CACHE[folder_path]
    
    # Search for folder
    folder_parts = folder_path.split('/')
    target_folder = None
    
    # Search through all stores to find the folder (excluding team/shared mailboxes)
    for store in namespace.Stores:
        try:
            # Skip excluded stores (team mailboxes, shared mailboxes)
            if store.DisplayName in EXCLUDED_STORES:
                continue
            
            current_folder = store.GetRootFolder()
            
            # Check if the first part matches the root folder name itself
            # This aligns with how get_all_folders() builds paths (including root name)
            if folder_parts and folder_parts[0] == current_folder.Name:
                # Path starts with root folder name - remove it and continue with remaining parts
                remaining_parts = folder_parts[1:]
            else:
                # Path doesn't start with root name - search from root
                remaining_parts = folder_parts
            
            # Navigate through the folder path
            found_all = True
            for part in remaining_parts:
                found = False
                for subfolder in current_folder.Folders:
                    if subfolder.Name == part:
                        current_folder = subfolder
                        found = True
                        break
                
                if not found:
                    found_all = False
                    break
            
            # If we found all parts (or path was just root folder name), we're done
            if found_all:
                target_folder = current_folder
                break
                
        except Exception:
            continue
    
    # Cache the result if found
    if target_folder is not None and use_cache:
        FOLDER_CACHE[folder_path] = target_folder
    
    return target_folder


def get_all_folders(folder, folder_list: Optional[List] = None, parent_path: str = "", include_counts: bool = False, max_depth: int = 50, current_depth: int = 0, visited_ids: Optional[set] = None) -> List[Dict[str, Any]]:
    """
    Recursively get all folders in Outlook with safety limits.
    
    Helper function to traverse the Outlook folder hierarchy and build
    a flat list of all folders with their full paths.
    
    Args:
        folder: Outlook folder COM object to start from
        folder_list: List to accumulate folders (used in recursion)
        parent_path: Path of parent folders (used in recursion)
        include_counts: Whether to include item/unread counts (SLOW! default: False)
        max_depth: Maximum recursion depth to prevent stack overflow (default: 50)
        current_depth: Current recursion depth (used internally)
        visited_ids: Set of visited folder IDs to detect cycles (used internally)
        
    Returns:
        List of dictionaries containing folder information
        
    Notes:
        - Uses recursion to traverse nested folder structures
        - Builds full paths like "Inbox/Archive/2024"
        - Some system folders may not be accessible (handled gracefully)
        - PERFORMANCE: include_counts=True can take minutes on large mailboxes!
        - SAFETY: max_depth prevents infinite recursion and stack overflow
        - SAFETY: visited_ids prevents circular reference loops
    """
    if folder_list is None:
        folder_list = []
    
    if visited_ids is None:
        visited_ids = set()
    
    # Safety check: Maximum recursion depth
    if current_depth >= max_depth:
        logger.warning(f"Maximum recursion depth ({max_depth}) reached at path: {parent_path}")
        return folder_list
    
    try:
        # Safety check: Detect circular references
        folder_id = id(folder)
        if folder_id in visited_ids:
            logger.warning(f"Circular reference detected at path: {parent_path}")
            return folder_list
        
        visited_ids.add(folder_id)
        
        # Build the full path for this folder
        current_path = f"{parent_path}/{folder.Name}" if parent_path else folder.Name
        
        # Build folder info (optionally without expensive counts)
        folder_info = {
            "name": folder.Name,
            "path": current_path
        }
        
        # Performance optimization: Only get counts if explicitly requested
        if include_counts:
            try:
                folder_info["item_count"] = folder.Items.Count if hasattr(folder, 'Items') else 0
                folder_info["unread_count"] = folder.UnReadItemCount if hasattr(folder, 'UnReadItemCount') else 0
            except Exception:
                folder_info["item_count"] = -1
                folder_info["unread_count"] = -1
        
        folder_list.append(folder_info)
        
        # Recursively process subfolders with incremented depth
        try:
            if hasattr(folder, 'Folders'):
                for subfolder in folder.Folders:
                    get_all_folders(subfolder, folder_list, current_path, include_counts, max_depth, current_depth + 1, visited_ids)
        except Exception:
            pass
        
    except Exception:
        pass
    
    return folder_list

