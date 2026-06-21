@echo off
setlocal
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartFinOps-Workplace-Analyze.ps1" -Tenant prod
exit /b %ERRORLEVEL%
