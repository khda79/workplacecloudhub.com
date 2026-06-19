@echo off
setlocal
title Smart ThinClient Shell - Apply
set "SCRIPT_DIR=%~dp0"
echo Smart ThinClient Shell - Apply
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartThinClient-Shell.ps1" -Cli -Action Apply
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Exit code: %EXITCODE%
pause
exit /b %EXITCODE%
