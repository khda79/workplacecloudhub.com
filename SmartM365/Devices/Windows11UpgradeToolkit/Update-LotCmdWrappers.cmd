@echo off
setlocal EnableExtensions


powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\SmartM365-Windows11Upgrade-Update-LotCmdWrappers.ps1" %*
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Finished with exit code %EXITCODE%.
pause
exit /b %EXITCODE%
