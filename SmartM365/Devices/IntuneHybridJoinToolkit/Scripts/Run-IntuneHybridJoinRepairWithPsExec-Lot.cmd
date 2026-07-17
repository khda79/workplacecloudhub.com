@echo off
setlocal EnableExtensions EnableDelayedExpansion



rem Shared LOT launcher. LOT definitions live under %ROOT_DIR%\Lots and run data under %ROOT_DIR%\Runs.
rem The LOT wrapper sets EHJIR_LOT_DIR and optional execution switches.

set "LOT_DIR=%EHJIR_LOT_DIR%"
if "%LOT_DIR%"=="" set "LOT_DIR=%~dp0"

set "IGNORE_RUN_GUARD_ARG="
if /I "%EHJIR_IGNORE_RUN_GUARD%"=="1" (
    set "IGNORE_RUN_GUARD_ARG=-IgnoreRunGuard"
)

set "RUN_ONCE_ARG="
if /I "%EHJIR_RUN_ONCE%"=="1" (
    set "RUN_ONCE_ARG=-RunOnce"
)

for %%I in ("%LOT_DIR%") do set "LOT_DIR=%%~fI"
if not "%LOT_DIR:~-1%"=="\" set "LOT_DIR=%LOT_DIR%\"
for %%I in ("%LOT_DIR%..") do set "LOTS_DIR=%%~fI"
for %%I in ("%LOTS_DIR%") do set "LOTS_NAME=%%~nxI"
for %%I in ("%LOTS_DIR%\..") do set "ROOT_DIR=%%~fI"
for %%I in ("%LOT_DIR%.") do set "LOT_NAME=%%~nxI"

if /I not "%LOTS_NAME%"=="Lots" (
    echo ERROR: LOT wrappers must run from the strict layout:
    echo   %ROOT_DIR%\Lots\LOT-NAME
    set "EXITCODE=1"
    goto :END
)

set "POWERSHELL_EXE="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "POWERSHELL_EXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined POWERSHELL_EXE (
    for /f "delims=" %%P in ('where pwsh.exe 2^>nul') do if not defined POWERSHELL_EXE set "POWERSHELL_EXE=%%P"
)
if not defined POWERSHELL_EXE if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not defined POWERSHELL_EXE (
    for /f "delims=" %%P in ('where powershell.exe 2^>nul') do if not defined POWERSHELL_EXE set "POWERSHELL_EXE=%%P"
)
if not defined POWERSHELL_EXE (
    echo ERROR: Windows PowerShell or PowerShell 7 was not found.
    set "EXITCODE=9009"
    goto :END
)

for /f "delims=" %%T in ('""%POWERSHELL_EXE%" -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss""') do set "RUN_STAMP=%%T"
if "%EHJIR_RUN_DIR%"=="" set "EHJIR_RUN_DIR=%ROOT_DIR%\Runs\%LOT_NAME%\!RUN_STAMP!"
for %%I in ("%EHJIR_RUN_DIR%") do set "RUN_DIR=%%~fI"

set "SCRIPT=%ROOT_DIR%\Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1"
set "COMPUTERS=%LOT_DIR%Computers.txt"
set "PARENT_INTUNE_CSV=%ROOT_DIR%\DevicesIntune.csv"
set "LOT_INTUNE_CSV=%RUN_DIR%\DevicesIntune.csv"
set "PARENT_AD_CSV=%ROOT_DIR%\DevicesAD.csv"
set "LOT_AD_CSV=%RUN_DIR%\DevicesAD.csv"
set "AD_DOMAIN_FILE=%LOT_DIR%AdDomain.txt"
set "PSEXEC_LOGS=%RUN_DIR%\PsExecLogs"
set "REPORTS=%RUN_DIR%\Reports"
set "CENTRAL_LOGS=%RUN_DIR%\CentralLogs"
set "ARCHIVES=%RUN_DIR%\Archives"
set "PSEXEC_EXE=%ROOT_DIR%\Scripts\PsExec.exe"

for %%D in ("%RUN_DIR%" "%RUN_DIR%\Logs" "%PSEXEC_LOGS%" "%REPORTS%" "%CENTRAL_LOGS%" "%ARCHIVES%" "%RUN_DIR%\State") do if not exist "%%~D" mkdir "%%~D" >nul 2>&1

call :PrintStartupInfo
if /I "%EHJIR_RUN_ONCE%"=="1" (title SmartM365 HybridJoin - %LOT_NAME% - Once) else (title SmartM365 HybridJoin - %LOT_NAME% - Loop)

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



set "DRY_RUN_REQUESTED=0"
for %%A in (%*) do (
    if /I "%%~A"=="-DryRun" set "DRY_RUN_REQUESTED=1"
)

if not "%DRY_RUN_REQUESTED%"=="1" (
    net session >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Run this CMD from an elevated administrator command prompt.
        set "EXITCODE=1"
        goto :END
    )

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
if not defined EHJIR_GLOBAL_CONCURRENCY_LIMIT set "EHJIR_GLOBAL_CONCURRENCY_LIMIT=15"
if not defined EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES set "EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES=0"
if not defined EHJIR_DELAY_BETWEEN_CYCLES_MINUTES set "EHJIR_DELAY_BETWEEN_CYCLES_MINUTES=5"
if not defined EHJIR_INTUNE_RETRY_SLEEP_MINUTES set "EHJIR_INTUNE_RETRY_SLEEP_MINUTES=5"
if not defined EHJIR_INTUNE_RETRY_MAX_RETRIES set "EHJIR_INTUNE_RETRY_MAX_RETRIES=5"
if not defined EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS set "EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS=300"
if not defined EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS set "EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS=3"
if not defined EHJIR_STALE_CLEANUP_DELAY_SECONDS set "EHJIR_STALE_CLEANUP_DELAY_SECONDS=60"
if not defined EHJIR_REBOOT_DELAY_SECONDS set "EHJIR_REBOOT_DELAY_SECONDS=180"
if not defined EHJIR_PSEXEC_TIMEOUT_MINUTES set "EHJIR_PSEXEC_TIMEOUT_MINUTES=120"
if not defined EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES set "EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES=15"
if not defined EHJIR_DISABLE_LOT_RUN_MUTEX set "EHJIR_DISABLE_LOT_RUN_MUTEX=0"
if not defined EHJIR_DISABLE_NIGHT_PAUSE set "EHJIR_DISABLE_NIGHT_PAUSE=0"
if not defined EHJIR_NIGHT_PAUSE_START_HOUR set "EHJIR_NIGHT_PAUSE_START_HOUR=20"
if not defined EHJIR_NIGHT_PAUSE_END_HOUR set "EHJIR_NIGHT_PAUSE_END_HOUR=7"
if not defined EHJIR_ALLOW_DSREG_LEAVE set "EHJIR_ALLOW_DSREG_LEAVE=1"
if not defined EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT set "EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT=1"
if not defined EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER set "EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER=1"
if not defined EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE set "EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE=1"
if not defined EHJIR_SKIP_VIRTUAL_MACHINES set "EHJIR_SKIP_VIRTUAL_MACHINES=1"
if not defined EHJIR_SKIP_PRE_RUN_ARCHIVE set "EHJIR_SKIP_PRE_RUN_ARCHIVE=0"
if not defined EHJIR_CONTINUE_ON_DNS_PREFLIGHT_FAILURE set "EHJIR_CONTINUE_ON_DNS_PREFLIGHT_FAILURE=0"
if not defined EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY set "EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY=1"
if not defined EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY set "EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY=0"
if not defined EHJIR_TECHNICIAN_RUN_GUARD_HOURS set "EHJIR_TECHNICIAN_RUN_GUARD_HOURS=3"
if not defined EHJIR_CENTRAL_LOG_COLLECTION_MODE set "EHJIR_CENTRAL_LOG_COLLECTION_MODE=Standard"
if /I not "%EHJIR_CENTRAL_LOG_COLLECTION_MODE%"=="Full" set "EHJIR_CENTRAL_LOG_COLLECTION_MODE=Standard"

call :NormalizeInt EHJIR_THROTTLE 10 1 9999
call :NormalizeInt EHJIR_GLOBAL_CONCURRENCY_LIMIT 15 1 9999
call :NormalizeInt EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES 0 0 100000
call :NormalizeInt EHJIR_DELAY_BETWEEN_CYCLES_MINUTES 5 0 100000
call :NormalizeInt EHJIR_INTUNE_RETRY_SLEEP_MINUTES 5 0 100000
call :NormalizeInt EHJIR_INTUNE_RETRY_MAX_RETRIES 5 0 100000
call :NormalizeInt EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS 300 0 3600
call :NormalizeInt EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS 3 1 30
call :NormalizeInt EHJIR_STALE_CLEANUP_DELAY_SECONDS 60 0 100000
call :NormalizeInt EHJIR_REBOOT_DELAY_SECONDS 180 0 100000
call :NormalizeInt EHJIR_PSEXEC_TIMEOUT_MINUTES 120 1 100000
call :NormalizeInt EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES 15 0 1440
call :NormalizeInt EHJIR_NIGHT_PAUSE_START_HOUR 20 0 23
call :NormalizeInt EHJIR_NIGHT_PAUSE_END_HOUR 7 0 23
call :NormalizeInt EHJIR_TECHNICIAN_RUN_GUARD_HOURS 3 0 168

set "ALLOW_DSREG_LEAVE_ARG=-AllowDsregLeave:$false"
if /I "%EHJIR_ALLOW_DSREG_LEAVE%"=="1" set "ALLOW_DSREG_LEAVE_ARG=-AllowDsregLeave"

set "ALLOW_STALE_INTUNE_CLEANUP_ARG="
if /I "%EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT%"=="1" set "ALLOW_STALE_INTUNE_CLEANUP_ARG=-AllowRemoveStaleIntuneEnrollment"

set "ALLOW_REBOOT_NO_USER_ARG="
if /I "%EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER%"=="1" set "ALLOW_REBOOT_NO_USER_ARG=-AllowRebootWhenNoInteractiveUser"

set "ALLOW_REBOOT_AFTER_LEAVE_ARG="
if /I "%EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE%"=="1" set "ALLOW_REBOOT_AFTER_LEAVE_ARG=-AllowRebootAfterDsregLeave"

set "SKIP_VM_ARG="
if /I "%EHJIR_SKIP_VIRTUAL_MACHINES%"=="1" set "SKIP_VM_ARG=-SkipVirtualMachines"

set "SKIP_PRE_RUN_ARCHIVE_ARG="
if /I "%EHJIR_SKIP_PRE_RUN_ARCHIVE%"=="1" set "SKIP_PRE_RUN_ARCHIVE_ARG=-SkipPreRunArchive"

set "TECHNICIAN_RUN_GUARD_ARG="
if /I "%EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY%"=="1" set "TECHNICIAN_RUN_GUARD_ARG=-UseTechnicianRunGuardHistory"
if /I "%EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY%"=="1" set "TECHNICIAN_RUN_GUARD_ARG=-IgnoreTechnicianRunGuardHistory"

set "CONTINUE_DNS_PREFLIGHT_ARG="
if /I "%EHJIR_CONTINUE_ON_DNS_PREFLIGHT_FAILURE%"=="1" set "CONTINUE_DNS_PREFLIGHT_ARG=-ContinueOnDnsPreflightFailure"

set "NIGHT_PAUSE_ARG="
if /I "%EHJIR_DISABLE_NIGHT_PAUSE%"=="1" set "NIGHT_PAUSE_ARG=-DisableNightPause"

set "LOT_RUN_MUTEX_ARG="
if /I "%EHJIR_DISABLE_LOT_RUN_MUTEX%"=="1" set "LOT_RUN_MUTEX_ARG=-DisableLotRunMutex"

set "GLOBAL_CONCURRENCY_PROVIDED=0"
for %%A in (%*) do (
    if /I "%%~A"=="-GlobalConcurrencyLimit" set "GLOBAL_CONCURRENCY_PROVIDED=1"
)
set "GLOBAL_CONCURRENCY_ARG="
if not "%GLOBAL_CONCURRENCY_PROVIDED%"=="1" (
    set "GLOBAL_CONCURRENCY_ARG=-GlobalConcurrencyLimit %EHJIR_GLOBAL_CONCURRENCY_LIMIT% -GlobalConcurrencyLeaseTimeoutMinutes %EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES%"
)

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

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -ComputerListPath "%COMPUTERS%" %INTUNE_ARGS% %AD_ARGS% ^
  -LogRoot "%PSEXEC_LOGS%" ^
  -ReportRoot "%REPORTS%" ^
  -CentralLogRoot "%CENTRAL_LOGS%" ^
  -ArchiveRoot "%ARCHIVES%" -CentralLogCollectionMode %EHJIR_CENTRAL_LOG_COLLECTION_MODE% -TechnicianRunGuardHours %EHJIR_TECHNICIAN_RUN_GUARD_HOURS% %TECHNICIAN_RUN_GUARD_ARG% %ALLOW_DSREG_LEAVE_ARG% %ALLOW_STALE_INTUNE_CLEANUP_ARG% %ALLOW_REBOOT_NO_USER_ARG% %ALLOW_REBOOT_AFTER_LEAVE_ARG% %IGNORE_RUN_GUARD_ARG% %RUN_ONCE_ARG% %SKIP_VM_ARG% %SKIP_PRE_RUN_ARCHIVE_ARG% %CONTINUE_DNS_PREFLIGHT_ARG% %LOT_RUN_MUTEX_ARG% %PSEXEC_ARG% ^
  -ThrottleLimit %EHJIR_THROTTLE% %GLOBAL_CONCURRENCY_ARG% ^
  -DelayBetweenCyclesMinutes %EHJIR_DELAY_BETWEEN_CYCLES_MINUTES% %NIGHT_PAUSE_ARG% ^
  -NightPauseStartHour %EHJIR_NIGHT_PAUSE_START_HOUR% ^
  -NightPauseEndHour %EHJIR_NIGHT_PAUSE_END_HOUR% ^
  -IntuneRetrySleepMinutes %EHJIR_INTUNE_RETRY_SLEEP_MINUTES% ^
  -IntuneRetryMaxRetries %EHJIR_INTUNE_RETRY_MAX_RETRIES% ^
  -RetryAfterRebootDelaySeconds %EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS% ^
  -RetryAfterRebootMaxAttempts %EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS% ^
  -StaleCleanupDelaySeconds %EHJIR_STALE_CLEANUP_DELAY_SECONDS% ^
  -RebootDelaySeconds %EHJIR_REBOOT_DELAY_SECONDS% ^
  -PsExecTimeoutMinutes %EHJIR_PSEXEC_TIMEOUT_MINUTES% -CancellationDrainTimeoutMinutes %EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES% %*

set "EXITCODE=%ERRORLEVEL%"
if "%EXITCODE%"=="-1073741819" call :CapturePowerShellCrash
if "%EXITCODE%"=="3221225477" call :CapturePowerShellCrash

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
exit /b %EXITCODE%

:PrintStartupInfo
if not defined EHJIR_THROTTLE set "EHJIR_THROTTLE=10"
if not defined EHJIR_GLOBAL_CONCURRENCY_LIMIT set "EHJIR_GLOBAL_CONCURRENCY_LIMIT=15"
if not defined EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES set "EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES=0"
if not defined EHJIR_DELAY_BETWEEN_CYCLES_MINUTES set "EHJIR_DELAY_BETWEEN_CYCLES_MINUTES=5"
if not defined EHJIR_NIGHT_PAUSE_START_HOUR set "EHJIR_NIGHT_PAUSE_START_HOUR=20"
if not defined EHJIR_NIGHT_PAUSE_END_HOUR set "EHJIR_NIGHT_PAUSE_END_HOUR=7"
if not defined EHJIR_DISABLE_NIGHT_PAUSE set "EHJIR_DISABLE_NIGHT_PAUSE=0"
if not defined EHJIR_SKIP_PRE_RUN_ARCHIVE set "EHJIR_SKIP_PRE_RUN_ARCHIVE=0"
if not defined EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY set "EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY=1"
if not defined EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY set "EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY=0"
if not defined EHJIR_TECHNICIAN_RUN_GUARD_HOURS set "EHJIR_TECHNICIAN_RUN_GUARD_HOURS=3"
if not defined EHJIR_CENTRAL_LOG_COLLECTION_MODE set "EHJIR_CENTRAL_LOG_COLLECTION_MODE=Standard"
if not defined EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS set "EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS=300"
if not defined EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS set "EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS=3"
if not defined EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES set "EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES=15"
echo.
echo SmartM365 Intune Hybrid Join Toolkit - LOT PsExec launcher
echo Started       : %DATE% %TIME%
echo LOT           : %LOT_DIR%
echo Root          : %ROOT_DIR%
echo PowerShell    : %POWERSHELL_EXE%
echo Script        : %SCRIPT%
echo Computers     : %COMPUTERS%
echo PsExec logs   : %PSEXEC_LOGS%
echo Reports       : %REPORTS%
echo Central logs  : %CENTRAL_LOGS%
echo Central mode  : %EHJIR_CENTRAL_LOG_COLLECTION_MODE%
echo Tech guard    : Use=%EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY%; Ignore=%EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY%; Hours=%EHJIR_TECHNICIAN_RUN_GUARD_HOURS%
if /I "%EHJIR_RUN_ONCE%"=="1" (echo Mode          : Once) else (echo Mode          : Loop)
if /I "%EHJIR_IGNORE_RUN_GUARD%"=="1" (echo Run guard     : Ignored) else (echo Run guard     : Enabled)
echo Worker limit  : Local=%EHJIR_THROTTLE%; Global=%EHJIR_GLOBAL_CONCURRENCY_LIMIT%; LeaseTimeout=%EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES% minute(s)
echo Cycle delay   : %EHJIR_DELAY_BETWEEN_CYCLES_MINUTES% minute(s)
echo Reboot retry  : Delay=%EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS% second(s); MaxAttempts=%EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS%
echo Stop drain    : %EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES% minute(s); second Ctrl+C or second GUI stop forces local workers
if /I "%EHJIR_DISABLE_NIGHT_PAUSE%"=="1" (echo Night pause   : Disabled) else (echo Night pause   : Enabled %EHJIR_NIGHT_PAUSE_START_HOUR%:00-%EHJIR_NIGHT_PAUSE_END_HOUR%:00)
if /I "%EHJIR_SKIP_PRE_RUN_ARCHIVE%"=="1" (echo Pre-run archive: Disabled) else (echo Pre-run archive: Enabled)
if not "%EHJIR_AD_DOMAIN%"=="" echo AD domain     : %EHJIR_AD_DOMAIN%
echo.
exit /b 0

:CapturePowerShellCrash
set "CRASH_LOG=%RUN_DIR%\Logs\PowerShellCrash_%RUN_STAMP%.txt"
> "%CRASH_LOG%" echo SmartM365 launcher detected a native PowerShell crash.
>> "%CRASH_LOG%" echo Exit code      : %EXITCODE% ^(0xC0000005^)
>> "%CRASH_LOG%" echo PowerShell     : %POWERSHELL_EXE%
>> "%CRASH_LOG%" echo Captured       : %DATE% %TIME%
>> "%CRASH_LOG%" echo.
>> "%CRASH_LOG%" echo Recent Application events 1000, 1001, and 1023:
wevtutil qe Application /q:"*[System[(EventID=1000 or EventID=1001 or EventID=1023)]]" /rd:true /f:text /c:30 >> "%CRASH_LOG%" 2>&1
echo ERROR: PowerShell crashed with 0xC0000005. Diagnostic events: %CRASH_LOG%
exit /b 0

:NormalizeInt
set "_EHJIR_VAR_NAME=%~1"
set "_EHJIR_DEFAULT=%~2"
set "_EHJIR_MIN=%~3"
set "_EHJIR_MAX=%~4"
set "_EHJIR_VALUE=!%_EHJIR_VAR_NAME%!"
set "_EHJIR_VALUE=!_EHJIR_VALUE:"=!"
for /f "tokens=* delims= " %%V in ("!_EHJIR_VALUE!") do set "_EHJIR_VALUE=%%V"
if not defined _EHJIR_VALUE set "_EHJIR_VALUE=%_EHJIR_DEFAULT%"
set "_EHJIR_NON_NUMERIC="
for /f "delims=0123456789" %%N in ("!_EHJIR_VALUE!") do set "_EHJIR_NON_NUMERIC=1"
if defined _EHJIR_NON_NUMERIC set "_EHJIR_VALUE=%_EHJIR_DEFAULT%"
set /a "_EHJIR_NUMBER=!_EHJIR_VALUE!" >nul 2>&1
if errorlevel 1 set /a "_EHJIR_NUMBER=%_EHJIR_DEFAULT%"
if !_EHJIR_NUMBER! LSS %_EHJIR_MIN% set /a "_EHJIR_NUMBER=%_EHJIR_MIN%"
if !_EHJIR_NUMBER! GTR %_EHJIR_MAX% set /a "_EHJIR_NUMBER=%_EHJIR_MAX%"
set "%_EHJIR_VAR_NAME%=!_EHJIR_NUMBER!"
exit /b 0
