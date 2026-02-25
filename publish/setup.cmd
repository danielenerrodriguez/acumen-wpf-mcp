@echo off
echo WPF MCP - Setup
echo ================
echo.
echo Generating shortcuts for all macros...
echo.
"%~dp0WpfMcp.exe" --export-all
echo.
echo Shortcuts are in the Shortcuts folder next to macros.
echo You can copy them to your Desktop or pin them to Start.
echo.
pause
