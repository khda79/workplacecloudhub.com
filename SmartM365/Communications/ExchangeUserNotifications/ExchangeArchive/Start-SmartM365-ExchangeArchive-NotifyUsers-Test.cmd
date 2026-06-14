@echo off
setlocal
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartM365-ExchangeArchive-NotifyUsers.ps1" -Tenant test -WhatIf
exit /b %ERRORLEVEL%
