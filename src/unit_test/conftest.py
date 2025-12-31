"""
Pytest configuration for MCP-OUTLOOK unit tests
"""

import sys
from pathlib import Path

# Add src to path for all tests
src_path = Path(__file__).parent.parent
sys.path.insert(0, str(src_path))

