@echo off
rem Smart Endpoint Migration Toolkit - AD inventory export.
rem Writes the shared cache Inventory\DevicesAD.csv (reused while less than 12 hours old; pass -ForceRefresh to regenerate).
rem Without -Domain, exports all domains of the current AD forest. Pass -Domain contoso.local,
rem or set SEMT_AD_DOMAIN before launching, to limit the export to one domain.
rem Set SEMT_NO_PAUSE=1 to skip the final pause (unattended use).
setlocal DisableDelayedExpansion

set "SEMT_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SEMT_SCRIPT=%~dp0Scripts\SmartM365-EndpointMigration-Export-ADDevicesCsv.ps1"
set "SEMT_EXIT=1"
pushd "%~dp0" || exit /b 3

if exist "%SEMT_SCRIPT%" goto :script_ok
call :log "ERROR: script not found: %SEMT_SCRIPT%"
goto :end
:script_ok
call :log "Smart Endpoint Migration Toolkit - AD inventory export to Inventory\DevicesAD.csv"
if "%SEMT_AD_DOMAIN%"=="" goto :run_forest
"%SEMT_PS%" -NoProfile -ExecutionPolicy Bypass -File "%SEMT_SCRIPT%" -Domain "%SEMT_AD_DOMAIN%" %*
set "SEMT_EXIT=%ERRORLEVEL%"
goto :end
:run_forest
"%SEMT_PS%" -NoProfile -ExecutionPolicy Bypass -File "%SEMT_SCRIPT%" %*
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
