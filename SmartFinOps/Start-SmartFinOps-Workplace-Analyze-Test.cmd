@echo off
setlocal
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartFinOps-Workplace-Analyze.ps1" -Tenant test
exit /b %ERRORLEVEL%
