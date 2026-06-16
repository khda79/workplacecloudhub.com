@echo off
setlocal EnableExtensions EnableDelayedExpansion

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)


rem Shared LOT launcher. Keep LOT folders with tiny wrappers only.
rem The LOT wrapper sets EHJIR_LOT_DIR and optional execution switches.

set "LOT_DIR=%EHJIR_LOT_DIR%"
if "%LOT_DIR%"=="" set "LOT_DIR=%CD%\.."

set "IGNORE_RUN_GUARD_ARG="
if /I "%EHJIR_IGNORE_RUN_GUARD%"=="1" (
    set "IGNORE_RUN_GUARD_ARG=-IgnoreRunGuard"
)

set "RUN_ONCE_ARG="
if /I "%EHJIR_RUN_ONCE%"=="1" (
    set "RUN_ONCE_ARG=-RunOnce"
)

for %%I in ("%LOT_DIR%") do set "LOT_DIR=%%~fI\"
set "ROOT_DIR=%LOT_DIR%.."
for %%I in ("%ROOT_DIR%") do set "ROOT_DIR=%%~fI"

set "SCRIPT=%ROOT_DIR%\Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1"
set "COMPUTERS=%LOT_DIR%Computers.txt"
set "PARENT_INTUNE_CSV=%ROOT_DIR%\DevicesIntune.csv"
set "LOT_INTUNE_CSV=%LOT_DIR%DevicesIntune.csv"
set "PARENT_AD_CSV=%ROOT_DIR%\DevicesAD.csv"
set "LOT_AD_CSV=%LOT_DIR%DevicesAD.csv"
set "AD_DOMAIN_FILE=%LOT_DIR%AdDomain.txt"
set "PSEXEC_LOGS=%LOT_DIR%PsExecLogs"
set "REPORTS=%LOT_DIR%Reports"
set "CENTRAL_LOGS=%LOT_DIR%CentralLogs"
set "PSEXEC_EXE=%ROOT_DIR%\Scripts\PsExec.exe"

if not exist "%SCRIPT%" (
    echo ERROR: Script not found:
    echo "%SCRIPT%"
    set "EXITCODE=1"
    goto :END
)

if not exist "%COMPUTERS%" (
    echo ERROR: Computer list not found:
    echo "%COMPUTERS%"
    set "EXITCODE=1"
    goto :END
)

net session >nul 2>&1
if errorlevel 1 (
    echo ERROR: Run this CMD from an elevated administrator command prompt.
    set "EXITCODE=1"
    goto :END
)

set "DRY_RUN_REQUESTED=0"
for %%A in (%*) do (
    if /I "%%~A"=="-DryRun" set "DRY_RUN_REQUESTED=1"
)

if not "%DRY_RUN_REQUESTED%"=="1" (
    if exist "%PSEXEC_EXE%" (
        set "PSEXEC_ARG=-PsExecPath ""%PSEXEC_EXE%"""
    ) else (
        set "PSEXEC_FOUND="
        for /f "delims=" %%P in ('where PsExec.exe 2^>nul') do if not defined PSEXEC_FOUND set "PSEXEC_FOUND=%%P"
        if not defined PSEXEC_FOUND (
            echo ERROR: PsExec.exe not found.
            echo Place PsExec.exe here:
            echo   %PSEXEC_EXE%
            echo or add PsExec.exe to PATH, then relaunch this LOT.
            set "EXITCODE=1"
            goto :END
        )
        set "PSEXEC_ARG=-PsExecPath ""!PSEXEC_FOUND!"""
    )
)

if not defined EHJIR_THROTTLE set "EHJIR_THROTTLE=10"
if not defined EHJIR_DELAY_BETWEEN_CYCLES_MINUTES set "EHJIR_DELAY_BETWEEN_CYCLES_MINUTES=5"
if not defined EHJIR_INTUNE_RETRY_SLEEP_MINUTES set "EHJIR_INTUNE_RETRY_SLEEP_MINUTES=5"
if not defined EHJIR_INTUNE_RETRY_MAX_RETRIES set "EHJIR_INTUNE_RETRY_MAX_RETRIES=5"
if not defined EHJIR_STALE_CLEANUP_DELAY_SECONDS set "EHJIR_STALE_CLEANUP_DELAY_SECONDS=60"
if not defined EHJIR_REBOOT_DELAY_SECONDS set "EHJIR_REBOOT_DELAY_SECONDS=180"
if not defined EHJIR_PSEXEC_TIMEOUT_MINUTES set "EHJIR_PSEXEC_TIMEOUT_MINUTES=120"

if "%EHJIR_AD_DOMAIN%"=="" (
    if exist "%AD_DOMAIN_FILE%" (
        for /f "usebackq tokens=* delims=" %%D in ("%AD_DOMAIN_FILE%") do if not defined EHJIR_AD_DOMAIN set "EHJIR_AD_DOMAIN=%%D"
    )
)

set "INTUNE_ARGS="
if exist "%PARENT_INTUNE_CSV%" (
    set "INTUNE_ARGS=-IntuneInventoryCsv ""%PARENT_INTUNE_CSV%"""
) else if exist "%LOT_INTUNE_CSV%" (
    set "INTUNE_ARGS=-IntuneInventoryCsv ""%LOT_INTUNE_CSV%"""
)

set "AD_ARGS="
if exist "%PARENT_AD_CSV%" (
    set "AD_ARGS=-AdRootInventoryCsv ""%PARENT_AD_CSV%"""
)
if not "%EHJIR_AD_DOMAIN%"=="" (
    set "AD_ARGS=%AD_ARGS% -AdInventoryCsv ""%LOT_AD_CSV%"" -AdDomain ""%EHJIR_AD_DOMAIN%"""
) else (
    set "AD_ARGS=%AD_ARGS% -AdInventoryCsv ""%PARENT_AD_CSV%"""
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -ComputerListPath "%COMPUTERS%" ^
  %INTUNE_ARGS% ^
  %AD_ARGS% ^
  -LogRoot "%PSEXEC_LOGS%" ^
  -ReportRoot "%REPORTS%" ^
  -CentralLogRoot "%CENTRAL_LOGS%" ^
  -AllowDsregLeave ^
  -AllowRemoveStaleIntuneEnrollment ^
  -AllowRebootWhenNoInteractiveUser ^
  -AllowRebootAfterDsregLeave ^
  %IGNORE_RUN_GUARD_ARG% ^
  %RUN_ONCE_ARG% ^
  %PSEXEC_ARG% ^
  -ThrottleLimit %EHJIR_THROTTLE% ^
  -DelayBetweenCyclesMinutes %EHJIR_DELAY_BETWEEN_CYCLES_MINUTES% ^
  -IntuneRetrySleepMinutes %EHJIR_INTUNE_RETRY_SLEEP_MINUTES% ^
  -IntuneRetryMaxRetries %EHJIR_INTUNE_RETRY_MAX_RETRIES% ^
  -StaleCleanupDelaySeconds %EHJIR_STALE_CLEANUP_DELAY_SECONDS% ^
  -RebootDelaySeconds %EHJIR_REBOOT_DELAY_SECONDS% ^
  -PsExecTimeoutMinutes %EHJIR_PSEXEC_TIMEOUT_MINUTES% ^
  %*

set "EXITCODE=%ERRORLEVEL%"

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
popd
exit /b %EXITCODE%
