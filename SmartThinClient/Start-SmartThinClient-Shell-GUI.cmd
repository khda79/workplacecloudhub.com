@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
start "" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%SCRIPT_DIR%SmartThinClient-Shell.ps1" -Gui %*
exit /b %ERRORLEVEL%
