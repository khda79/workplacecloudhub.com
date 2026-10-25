@echo off
rem Smart Endpoint Migration Toolkit - shared Wave launcher.
rem Usage: Run-EndpointMigrationWave.cmd "<Wave folder>" Loop|Once Default|IgnoreRunGuard
rem The Campaign is the parent folder name of the Wave (Waves\<Campaign>\<Wave>).
rem Options are resolved by the engine: SEMT_* environment variables, Wave.local.json,
rem Config\SmartEndpointMigrationToolkit.local.json and the campaign defaults.
setlocal DisableDelayedExpansion

set "SEMT_WAVE_ARG=%~1"
set "SEMT_RUN_MODE=%~2"
set "SEMT_GUARD_MODE=%~3"
if "%SEMT_WAVE_ARG%"=="" goto :usage
if /I not "%SEMT_RUN_MODE%"=="Loop" if /I not "%SEMT_RUN_MODE%"=="Once" goto :usage
if "%SEMT_GUARD_MODE%"=="" set "SEMT_GUARD_MODE=Default"

for %%I in ("%SEMT_WAVE_ARG%") do set "SEMT_WAVE_DIR=%%~fI"
for %%I in ("%SEMT_WAVE_DIR%\..") do set "SEMT_CAMPAIGN=%%~nxI"
if /I "%SEMT_CAMPAIGN%"=="HybridJoin" goto :campaign_ok
if /I "%SEMT_CAMPAIGN%"=="Windows11Upgrade" goto :campaign_ok
echo ERROR: the Wave must be under Waves\HybridJoin or Waves\Windows11Upgrade.
exit /b 2
:campaign_ok

set "SEMT_GUARD_ARG="
if /I "%SEMT_GUARD_MODE%"=="IgnoreRunGuard" set "SEMT_GUARD_ARG=-IgnoreRunGuard"

net session >nul 2>&1
if errorlevel 1 echo WARNING: this console is not elevated. Run the Wave as administrator so the shared technician state can be protected.

title SEMT %SEMT_CAMPAIGN% - %SEMT_RUN_MODE%
pushd "%~dp0" || exit /b 3
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartM365-Invoke-EndpointMigrationWithPsExec.ps1" -Campaign "%SEMT_CAMPAIGN%" -WavePath "%SEMT_WAVE_DIR%" -RunMode %SEMT_RUN_MODE% %SEMT_GUARD_ARG%
set "SEMT_EXIT=%ERRORLEVEL%"
popd
exit /b %SEMT_EXIT%

:usage
echo Usage: Run-EndpointMigrationWave.cmd "<Wave folder>" Loop^|Once Default^|IgnoreRunGuard
exit /b 1
