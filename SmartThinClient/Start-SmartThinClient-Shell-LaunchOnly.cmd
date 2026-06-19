@echo off
setlocal
title Smart ThinClient Shell - Launch Only
set "SCRIPT_DIR=%~dp0"
echo Smart ThinClient Shell - Launch Only
echo.
echo This opens the thin-client launcher without installing kiosk mode or changing Windows shell settings.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartThinClient-Shell.ps1" -Cli -Action Launch -Profile Hybrid
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Exit code: %EXITCODE%
pause
exit /b %EXITCODE%
