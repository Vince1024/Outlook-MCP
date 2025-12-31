@echo off
REM Startup script for MCP Outlook Server (Modular Architecture)

echo ========================================
echo MCP Outlook Server
echo Modular Architecture
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher
    pause
    exit /b 1
)

echo Starting MCP Outlook Server...
echo.

REM Start server
python src\outlook_mcp.py

REM If server stops, display a message
echo.
echo ========================================
echo Server stopped
echo ========================================
pause

