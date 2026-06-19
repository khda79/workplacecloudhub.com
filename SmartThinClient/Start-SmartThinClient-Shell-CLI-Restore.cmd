@echo off
setlocal
title Smart ThinClient Shell - Restore
set "SCRIPT_DIR=%~dp0"
echo Smart ThinClient Shell - Restore
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartThinClient-Shell.ps1" -Cli -Action Restore
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Exit code: %EXITCODE%
pause
exit /b %EXITCODE%
