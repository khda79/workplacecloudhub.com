@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT_DIR%SmartThinClient-Shell.ps1" -Gui
exit /b %ERRORLEVEL%
