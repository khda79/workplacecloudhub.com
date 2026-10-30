@echo off
rem Smart Endpoint Migration Toolkit - Entra inventory export (delegated Microsoft Graph sign-in).
rem Writes the shared cache Inventory\DevicesEntra.csv with the full Entra device inventory.
rem A shared inventory less than 24 hours old is reused; pass -ForceRefresh to regenerate.
rem Set SEMT_NO_PAUSE=1 to skip the final pause (unattended use).
setlocal DisableDelayedExpansion

set "SEMT_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SEMT_SCRIPT=%~dp0Scripts\SmartM365-EndpointMigration-Export-EntraDevicesCsv.ps1"
set "SEMT_EXIT=1"
pushd "%~dp0" || exit /b 3

if exist "%SEMT_SCRIPT%" goto :script_ok
call :log "ERROR: script not found: %SEMT_SCRIPT%"
goto :end
:script_ok
call :log "Smart Endpoint Migration Toolkit - Entra inventory export to Inventory\DevicesEntra.csv"
"%SEMT_PS%" -NoProfile -ExecutionPolicy Bypass -File "%SEMT_SCRIPT%" -PageSize 999 %*
set "SEMT_EXIT=%ERRORLEVEL%"

:end
call :log "Finished with exit code %SEMT_EXIT%."
popd
if "%SEMT_NO_PAUSE%"=="1" goto :done
call :log "Press any key to close this window."
pause >nul
:done
exit /b %SEMT_EXIT%

:log
set "SEMT_TS="
for /f "usebackq delims=" %%T in (`powershell.exe -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"`) do set "SEMT_TS=%%T"
echo %SEMT_TS% %~1
exit /b 0
