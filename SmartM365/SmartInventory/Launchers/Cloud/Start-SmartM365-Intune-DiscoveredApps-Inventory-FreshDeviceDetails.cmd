@echo off
rem Compatibility alias: the main DiscoveredApps launcher already requests fresh device details.
call "%~dp0Start-SmartM365-Intune-DiscoveredApps-Inventory.cmd" %*
exit /b %ERRORLEVEL%
