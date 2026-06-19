@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "TARGET_CMD=%SCRIPT_DIR%interactive\permissions\01-Scan-Source-Permissions.cmd"

if not exist "%TARGET_CMD%" (
    echo Missing launcher: %TARGET_CMD%
pause
    exit /b 1
)

call "%TARGET_CMD%" %*
exit /b %ERRORLEVEL%
