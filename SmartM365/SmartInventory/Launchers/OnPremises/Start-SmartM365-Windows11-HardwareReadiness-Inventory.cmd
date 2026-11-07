@echo off
setlocal EnableExtensions
pwsh.exe -NoProfile -File "%~dp0..\..\EndpointInventory\Windows11Readiness\SmartM365-Endpoint-Windows11Readiness-Inventory.ps1" %*
exit /b %ERRORLEVEL%
