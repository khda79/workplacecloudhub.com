@echo off
setlocal
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartM365-ExchangeArchive-NotifyUsers.ps1" -Tenant prod -WhatIf
exit /b %ERRORLEVEL%
