@echo off
setlocal
pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0SmartM365-ExchangeUserNotifications-GUI.ps1"
endlocal
