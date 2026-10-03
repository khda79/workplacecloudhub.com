@echo off
setlocal
rem Reuse the Windows PowerShell 5.1 launcher with real read-only collection enabled.
call "%~dp0Start-SmartM365-SharePointMigration-FarmDiagnostic.cmd" %* -Run
exit /b %ERRORLEVEL%
