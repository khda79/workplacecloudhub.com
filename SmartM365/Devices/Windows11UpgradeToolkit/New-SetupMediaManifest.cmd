@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "ROOT_DIR=%CD%\"
set "SCRIPT=%ROOT_DIR%Scripts\New-SmartM365SetupMediaManifest.ps1"
set "DEFAULT_SETUP_SOURCE=%ROOT_DIR%SetupSource"
call :PrintStartupInfo

if not exist "%SCRIPT%" (
    echo ERROR: Script not found:
    echo "%SCRIPT%"
    set "EXITCODE=1"
    goto :END
)

if "%~1"=="" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -SetupSourceRoot "%DEFAULT_SETUP_SOURCE%" -Force
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
)

set "EXITCODE=%ERRORLEVEL%"

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
popd
exit /b %EXITCODE%

:PrintStartupInfo
echo.
echo SmartM365 Windows 11 Upgrade Toolkit - setup media manifest generator
echo Started          : %DATE% %TIME%
echo Root             : %ROOT_DIR%
echo Script           : %SCRIPT%
echo Default source   : %DEFAULT_SETUP_SOURCE%
echo.
exit /b 0