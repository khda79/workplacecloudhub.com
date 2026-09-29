@echo off
rem Compatibility alias for existing shortcuts and orchestrator configurations.
call "%~dp0Start-SmartM365-Exchange-Local-MailboxCalendarPermissions-Inventory.cmd" %*
exit /b %errorlevel%
