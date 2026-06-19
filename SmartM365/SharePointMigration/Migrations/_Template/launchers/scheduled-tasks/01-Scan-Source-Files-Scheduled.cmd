@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "TARGET_CMD=%SCRIPT_DIR%files\01-Scan-Source-Files-Scheduled.cmd"

if not exist "%TARGET_CMD%" (
    echo Missing launcher: %TARGET_CMD%
    exit /b 1
)

call "%TARGET_CMD%" %*
exit /b %ERRORLEVEL%
