@echo off
call "%~dp0Start-SmartM365-Intune-DiscoveredApps-Inventory.cmd" -FreshDeviceDetails %*
exit /b %ERRORLEVEL%
