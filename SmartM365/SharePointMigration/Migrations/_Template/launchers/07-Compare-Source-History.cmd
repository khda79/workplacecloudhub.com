@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "TARGET_CMD=%SCRIPT_DIR%interactive\files\04-Compare-Source-History.cmd"

if not exist "%TARGET_CMD%" (
    echo Missing launcher: %TARGET_CMD%
pause
    exit /b 1
)

call "%TARGET_CMD%" %*
exit /b %ERRORLEVEL%
