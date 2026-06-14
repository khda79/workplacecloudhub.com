@echo off
setlocal
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.ps1" -Tenant test -WhatIf -SkipConfirmation
exit /b %ERRORLEVEL%
