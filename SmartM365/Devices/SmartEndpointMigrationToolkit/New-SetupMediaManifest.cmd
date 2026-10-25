@echo off
rem Smart Endpoint Migration Toolkit - Windows 11 Upgrade setup media manifest generator.
rem Usage without argument: hashes every media folder under Campaigns\Windows11Upgrade\SetupSource.
rem Usage with arguments  : passes them to the PowerShell script, for example
rem   New-SetupMediaManifest.cmd -MediaRoot "D:\Media\FR-fr" -Force
setlocal DisableDelayedExpansion

set "SEMT_ROOT_DIR=%~dp0"
set "SEMT_SCRIPT=%SEMT_ROOT_DIR%Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1"
set "SEMT_DEFAULT_SETUP_SOURCE=%SEMT_ROOT_DIR%Campaigns\Windows11Upgrade\SetupSource"
set "SEMT_INTERACTIVE=0"
if "%~1"=="" set "SEMT_INTERACTIVE=1"

echo.
echo Smart Endpoint Migration Toolkit - Windows 11 Upgrade setup media manifest generator
echo Started          : %DATE% %TIME%
echo Root             : %SEMT_ROOT_DIR%
echo Script           : %SEMT_SCRIPT%
echo Default source   : %SEMT_DEFAULT_SETUP_SOURCE%
echo.

if not exist "%SEMT_SCRIPT%" (
    echo ERROR: Script not found: "%SEMT_SCRIPT%"
    set "SEMT_EXIT=1"
    goto :end
)

pushd "%SEMT_ROOT_DIR%" || (
    echo ERROR: Cannot open the toolkit root: "%SEMT_ROOT_DIR%"
    set "SEMT_EXIT=3"
    goto :end
)
if "%SEMT_INTERACTIVE%"=="1" (
    "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%SEMT_SCRIPT%" -SetupSourceRoot "%SEMT_DEFAULT_SETUP_SOURCE%" -Force
) else (
    "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%SEMT_SCRIPT%" %*
)
set "SEMT_EXIT=%ERRORLEVEL%"
popd

:end
echo.
echo Finished with exit code %SEMT_EXIT%.
if "%SEMT_INTERACTIVE%"=="1" pause
exit /b %SEMT_EXIT%
