@echo off
setlocal
title Smart ThinClient Shell - Preview
set "SCRIPT_DIR=%~dp0"
echo Smart ThinClient Shell - Preview
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartThinClient-Shell.ps1" -Cli -Action Preview
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Exit code: %EXITCODE%
pause
exit /b %EXITCODE%
