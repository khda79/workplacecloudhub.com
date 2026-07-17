@echo off
setlocal EnableExtensions EnableDelayedExpansion


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

rem Shared LOT launcher. LOT definitions live under %ROOT_DIR%\Lots and run data under %ROOT_DIR%\Runs.
rem The LOT wrapper sets W11UT_LOT_DIR and optional execution switches.

set "LOT_DIR=%W11UT_LOT_DIR%"
if "%LOT_DIR%"=="" set "LOT_DIR=%~dp0"

for %%I in ("%LOT_DIR%") do set "LOT_DIR=%%~fI"
if not "%LOT_DIR:~-1%"=="\" set "LOT_DIR=%LOT_DIR%\"
for %%I in ("%LOT_DIR%..") do set "LOTS_DIR=%%~fI"
for %%I in ("%LOTS_DIR%\..") do set "ROOT_DIR=%%~fI"
for %%I in ("%LOT_DIR%.") do set "LOT_NAME=%%~nxI"

if /I not "%LOTS_DIR:~-4%"=="Lots" (
    echo ERROR: LOT wrappers must run from the strict layout:
    echo   %ROOT_DIR%\Lots\LOT-NAME
    set "EXITCODE=1"
    goto :END
)

for /f "delims=" %%T in ('""%POWERSHELL_EXE%" -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss""') do set "RUN_STAMP=%%T"
if "%W11UT_RUN_DIR%"=="" set "W11UT_RUN_DIR=%ROOT_DIR%\Runs\%LOT_NAME%\!RUN_STAMP!"
for %%I in ("%W11UT_RUN_DIR%") do set "RUN_DIR=%%~fI"

set "SCRIPT=%ROOT_DIR%\Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1"
set "COMPUTERS=%LOT_DIR%Computers.txt"
set "PARENT_AD_CSV=%ROOT_DIR%\DevicesAD.csv"
set "LOT_AD_CSV=%RUN_DIR%\DevicesAD.csv"
set "PARENT_INTUNE_CSV=%ROOT_DIR%\DevicesIntune.csv"
set "LOT_INTUNE_CSV=%RUN_DIR%\DevicesIntune.csv"
set "AD_DOMAIN_FILE=%LOT_DIR%AdDomain.txt"
set "PSEXEC_LOGS=%RUN_DIR%\PsExecLogs"
set "REPORTS=%RUN_DIR%\Reports"
set "CENTRAL_LOGS=%RUN_DIR%\CentralLogs"
set "PSEXEC_EXE=%ROOT_DIR%\Scripts\PsExec.exe"
set "LOT_CONFIG_FILE=%LOT_DIR%Windows11UpgradeToolkit.config"
set "ROOT_CONFIG_FILE=%ROOT_DIR%\Windows11UpgradeToolkit.config"

for %%D in ("%RUN_DIR%" "%RUN_DIR%\Logs" "%PSEXEC_LOGS%" "%REPORTS%" "%CENTRAL_LOGS%" "%RUN_DIR%\State") do if not exist "%%~D" mkdir "%%~D" >nul 2>&1
call :LOAD_CONFIG "%LOT_CONFIG_FILE%"
call :LOAD_CONFIG "%ROOT_CONFIG_FILE%"

rem Default LOT behavior matches the operator GUI: guarded repair and upgrade are enabled.
rem Set any W11UT_* switch to 0 before launching a LOT to disable that action.
if not defined W11UT_ALLOW_POLICY_REPAIR set "W11UT_ALLOW_POLICY_REPAIR=1"
if not defined W11UT_ALLOW_WU_RESET set "W11UT_ALLOW_WU_RESET=1"
if not defined W11UT_ALLOW_FORCE_UPGRADE set "W11UT_ALLOW_FORCE_UPGRADE=1"
if not defined W11UT_ALLOW_SETUP_UPGRADE set "W11UT_ALLOW_SETUP_UPGRADE=1"
if not defined W11UT_DIRECT_SETUP_UPGRADE set "W11UT_DIRECT_SETUP_UPGRADE=0"
if not defined W11UT_ALLOW_REBOOT set "W11UT_ALLOW_REBOOT=1"
if not defined W11UT_SCHEDULE_RETRY_AFTER_REBOOT set "W11UT_SCHEDULE_RETRY_AFTER_REBOOT=1"
if not defined W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS set "W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS=3"
if not defined W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS set "W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS=300"
if not defined W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS set "W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS=7"
if not defined W11UT_SETUP_REBOOT_WHEN_NO_USER set "W11UT_SETUP_REBOOT_WHEN_NO_USER=1"
if not defined W11UT_ALLOW_SETUP_PROFILE_REPAIR set "W11UT_ALLOW_SETUP_PROFILE_REPAIR=1"
if not defined W11UT_SKIP_VIRTUAL_MACHINES set "W11UT_SKIP_VIRTUAL_MACHINES=1"
if not defined W11UT_ALLOW_DISK_CLEANUP set "W11UT_ALLOW_DISK_CLEANUP=1"
if not defined W11UT_ALLOW_ADVANCED_DISK_CLEANUP set "W11UT_ALLOW_ADVANCED_DISK_CLEANUP=0"
if /I "%W11UT_ALLOW_DISM_COMPONENT_CLEANUP%"=="1" set "W11UT_ALLOW_ADVANCED_DISK_CLEANUP=1"
if not defined W11UT_SETUP_EXECUTION_MODE set "W11UT_SETUP_EXECUTION_MODE=LocalCache"
if not defined W11UT_SETUP_MEDIA_ID set "W11UT_SETUP_MEDIA_ID=Win11"
if not defined W11UT_SETUP_LANGUAGE set "W11UT_SETUP_LANGUAGE=MatchSystem"
if not defined W11UT_SETUP_DYNAMIC_UPDATE set "W11UT_SETUP_DYNAMIC_UPDATE=Disable"
if not defined W11UT_SETUP_SOURCE_CANDIDATE_LIMIT set "W11UT_SETUP_SOURCE_CANDIDATE_LIMIT=5"
if not defined W11UT_SETUP_COPY_IPG_MS set "W11UT_SETUP_COPY_IPG_MS=20"
if not defined W11UT_SETUP_COPY_JITTER_SECONDS set "W11UT_SETUP_COPY_JITTER_SECONDS=300"
if not defined W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT set "W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT=0"
if not defined W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES set "W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES=240"
if not defined W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT set "W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT=1"
if not defined W11UT_SETUP_SUBNET_PREFIX_LENGTH set "W11UT_SETUP_SUBNET_PREFIX_LENGTH=Auto"
if not defined W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES set "W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES=90"
if /I "%W11UT_DIRECT_SETUP_UPGRADE%"=="1" set "W11UT_ALLOW_REBOOT=1"
if /I "%W11UT_DIRECT_SETUP_UPGRADE%"=="1" set "W11UT_SETUP_REBOOT_WHEN_NO_USER=1"

set "W11UT_ALLOW_SETUP_SOURCE_VALIDATION=0"
if /I "%W11UT_DIRECT_SETUP_UPGRADE%"=="1" set "W11UT_ALLOW_SETUP_SOURCE_VALIDATION=1"
if /I "%W11UT_ALLOW_SETUP_UPGRADE%"=="1" set "W11UT_ALLOW_SETUP_SOURCE_VALIDATION=1"
if /I "%W11UT_ALLOW_SETUP_SOURCE_VALIDATION%"=="1" if /I not "%W11UT_SKIP_SETUP_MEDIA_PRECOPY%"=="1" (
    if "%W11UT_SETUP_SOURCE%"=="" if "%W11UT_SETUP_SOURCE_MAP%"=="" (
        echo ERROR: W11UT_SETUP_SOURCE or W11UT_SETUP_SOURCE_MAP is required when setup upgrade is enabled.
        echo Use a UNC path reachable by target computers, for example:
        echo   set W11UT_SETUP_SOURCE=\\server\share\Windows11
        echo Or use a CSV map:
        echo   set W11UT_SETUP_SOURCE_MAP=\\server\share\SetupSourceMap.csv
        echo Or set W11UT_SKIP_SETUP_MEDIA_PRECOPY=1 when the target cache is already valid.
        set "EXITCODE=1"
        goto :END
    )
    if not "%W11UT_SETUP_SOURCE%"=="" if not "%W11UT_SETUP_SOURCE:~0,2%"=="\\" if /I not "%W11UT_CONFIRM_LOCAL_SETUP_SOURCE%"=="1" (
        echo ERROR: W11UT_SETUP_SOURCE is not a UNC path:
        echo   %W11UT_SETUP_SOURCE%
        echo In LOT/PsExec mode, target computers run as SYSTEM and must read the source themselves.
        echo Use a UNC path or set W11UT_CONFIRM_LOCAL_SETUP_SOURCE=1 for a local/direct test.
        set "EXITCODE=1"
        goto :END
    )
    if not "%W11UT_SETUP_SOURCE_MAP%"=="" if not "%W11UT_SETUP_SOURCE_MAP:~0,2%"=="\\" if /I not "%W11UT_CONFIRM_LOCAL_SETUP_SOURCE%"=="1" (
        echo ERROR: W11UT_SETUP_SOURCE_MAP is not a UNC path:
        echo   %W11UT_SETUP_SOURCE_MAP%
        echo In LOT/PsExec mode, target computers run as SYSTEM and must read the map themselves.
        echo Use a UNC path or set W11UT_CONFIRM_LOCAL_SETUP_SOURCE=1 for a local/direct test.
        set "EXITCODE=1"
        goto :END
    )
)

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

set "RUN_ONCE_ARG="
set "RUN_MODE_LABEL=Loop"
if /I "%W11UT_RUN_ONCE%"=="1" (
    set "RUN_ONCE_ARG=-RunOnce"
    set "RUN_MODE_LABEL=Once"
)

set "IGNORE_RUN_GUARD_ARG="
if /I "%W11UT_IGNORE_RUN_GUARD%"=="1" (
    set "IGNORE_RUN_GUARD_ARG=-IgnoreRunGuard"
    set "RUN_MODE_LABEL=%RUN_MODE_LABEL%IgnoreRunGuard"
)
title SmartM365 W11UT - %LOT_NAME% - %RUN_MODE_LABEL%
set "SETUP_ARGS="
if /I "%W11UT_ALLOW_SETUP_UPGRADE%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -AllowSetupUpgrade"
if /I "%W11UT_DIRECT_SETUP_UPGRADE%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -DirectSetupUpgrade"
if /I "%W11UT_SETUP_REBOOT_WHEN_NO_USER%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -AllowSetupCompletionRebootWhenNoUser"
if /I "%W11UT_ALLOW_SETUP_PROFILE_REPAIR%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -AllowSetupProfileRepair"
if not "%W11UT_SETUP_SOURCE%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSourcePath ""%W11UT_SETUP_SOURCE%"""
if not "%W11UT_SETUP_SOURCE_MAP%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSourceMapPath ""%W11UT_SETUP_SOURCE_MAP%"""
if not "%W11UT_SETUP_EXECUTION_MODE%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupExecutionMode ""%W11UT_SETUP_EXECUTION_MODE%"""
if not "%W11UT_SETUP_MEDIA_ID%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupMediaId ""%W11UT_SETUP_MEDIA_ID%"""
if not "%W11UT_SETUP_LANGUAGE%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupLanguage ""%W11UT_SETUP_LANGUAGE%"""
if not "%W11UT_SETUP_DYNAMIC_UPDATE%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupDynamicUpdate ""%W11UT_SETUP_DYNAMIC_UPDATE%"""
if not "%W11UT_SETUP_SOURCE_CANDIDATE_LIMIT%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSourceCandidateLimit %W11UT_SETUP_SOURCE_CANDIDATE_LIMIT%"
if not "%W11UT_SETUP_COPY_IPG_MS%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupMediaCopyIpGapMilliseconds %W11UT_SETUP_COPY_IPG_MS%"
if not "%W11UT_SETUP_COPY_JITTER_SECONDS%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupMediaCopyJitterSeconds %W11UT_SETUP_COPY_JITTER_SECONDS%"
if not "%W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSourceConcurrencyLimit %W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT%"
if not "%W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSourceConcurrencyLeaseMinutes %W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES%"
if not "%W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSourceConcurrencyGateRoot ""%W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT%"""
if not "%W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSubnetConcurrencyLimit %W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT%"
if not "%W11UT_SETUP_SUBNET_PREFIX_LENGTH%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSubnetPrefixLength %W11UT_SETUP_SUBNET_PREFIX_LENGTH%"
if not "%W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSubnetConcurrencyLeaseMinutes %W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES%"
if not "%W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSubnetConcurrencyGateRoot ""%W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT%"""
if /I "%W11UT_SKIP_SETUP_MEDIA_PRECOPY%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -SkipSetupMediaPreCopy"

set "ACTION_ARGS="
if /I "%W11UT_AUDIT_ONLY%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AuditOnly"
if /I "%W11UT_ALLOW_POLICY_REPAIR%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowPolicyRepair"
if /I "%W11UT_ALLOW_WU_RESET%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowWUReset"
if /I "%W11UT_ALLOW_FORCE_UPGRADE%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowForceUpgrade"
if /I "%W11UT_ALLOW_REBOOT%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowReboot"
if /I "%W11UT_SCHEDULE_RETRY_AFTER_REBOOT%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -ScheduleRetryAfterReboot"
if not "%W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS%"=="" set "ACTION_ARGS=%ACTION_ARGS% -RetryAfterRebootMaxAttempts %W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS%"
if not "%W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS%"=="" set "ACTION_ARGS=%ACTION_ARGS% -RetryAfterRebootDelaySeconds %W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS%"
if not "%W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS%"=="" set "ACTION_ARGS=%ACTION_ARGS% -ForceRequiredRebootWhenUptimeOverDays %W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS%"
if /I "%W11UT_SKIP_VIRTUAL_MACHINES%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -SkipVirtualMachines"
if /I "%W11UT_ALLOW_DISK_CLEANUP%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowDiskCleanup"
if /I "%W11UT_ALLOW_ADVANCED_DISK_CLEANUP%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowAdvancedDiskCleanup"

if not defined W11UT_GLOBAL_CONCURRENCY_LIMIT set "W11UT_GLOBAL_CONCURRENCY_LIMIT=15"
if not defined W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES set "W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES=0"
if not defined W11UT_THROTTLE set "W11UT_THROTTLE=10"
if not defined W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS set "W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS=0"
if not defined W11UT_DELAY_BETWEEN_CYCLES_MINUTES set "W11UT_DELAY_BETWEEN_CYCLES_MINUTES=10"
if not defined W11UT_PSEXEC_TIMEOUT_MINUTES set "W11UT_PSEXEC_TIMEOUT_MINUTES=360"
if not defined W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES set "W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES=15"
if not defined W11UT_CENTRAL_LOG_COLLECTION_MODE set "W11UT_CENTRAL_LOG_COLLECTION_MODE=Standard"
if not defined W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY set "W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY=1"
if not defined W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY set "W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY=0"
if not defined W11UT_RUN_GUARD_HOURS set "W11UT_RUN_GUARD_HOURS=3"

set "THROTTLE_PROVIDED=0"
for %%A in (%*) do (
    if /I "%%~A"=="-ThrottleLimit" set "THROTTLE_PROVIDED=1"
)
set "THROTTLE_ARG="
if not "%THROTTLE_PROVIDED%"=="1" (
    set "THROTTLE_ARG=-ThrottleLimit %W11UT_THROTTLE%"
)

set "GLOBAL_CONCURRENCY_PROVIDED=0"
for %%A in (%*) do (
    if /I "%%~A"=="-GlobalConcurrencyLimit" set "GLOBAL_CONCURRENCY_PROVIDED=1"
)
set "GLOBAL_CONCURRENCY_ARG="
if not "%GLOBAL_CONCURRENCY_PROVIDED%"=="1" (
    set "GLOBAL_CONCURRENCY_ARG=-GlobalConcurrencyLimit %W11UT_GLOBAL_CONCURRENCY_LIMIT% -GlobalConcurrencyLeaseTimeoutMinutes %W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES%"
)

if "%W11UT_AD_DOMAIN%"=="" (
    if exist "%AD_DOMAIN_FILE%" (
        for /f "usebackq tokens=* delims=" %%D in ("%AD_DOMAIN_FILE%") do if not defined W11UT_AD_DOMAIN set "W11UT_AD_DOMAIN=%%D"
    )
)

set "TECH_RUN_GUARD_USE_PROVIDED=0"
set "TECH_RUN_GUARD_IGNORE_PROVIDED=0"
set "TECH_RUN_GUARD_HOURS_PROVIDED=0"
for %%A in (%*) do (
    if /I "%%~A"=="-UseTechnicianRunGuardHistory" set "TECH_RUN_GUARD_USE_PROVIDED=1"
    if /I "%%~A"=="-IgnoreTechnicianRunGuardHistory" set "TECH_RUN_GUARD_IGNORE_PROVIDED=1"
    if /I "%%~A"=="-RunGuardHours" set "TECH_RUN_GUARD_HOURS_PROVIDED=1"
)

set "TECH_RUN_GUARD_ARGS="
if /I "%W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY%"=="1" if not "%TECH_RUN_GUARD_USE_PROVIDED%"=="1" set "TECH_RUN_GUARD_ARGS=%TECH_RUN_GUARD_ARGS% -UseTechnicianRunGuardHistory"
if /I "%W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY%"=="1" if not "%TECH_RUN_GUARD_IGNORE_PROVIDED%"=="1" set "TECH_RUN_GUARD_ARGS=%TECH_RUN_GUARD_ARGS% -IgnoreTechnicianRunGuardHistory"
if not "%W11UT_RUN_GUARD_HOURS%"=="" if not "%TECH_RUN_GUARD_HOURS_PROVIDED%"=="1" set "TECH_RUN_GUARD_ARGS=%TECH_RUN_GUARD_ARGS% -RunGuardHours %W11UT_RUN_GUARD_HOURS%"

set "AD_ARGS=-AdInventoryCsv ""%LOT_AD_CSV%"""
if exist "%PARENT_AD_CSV%" set "AD_ARGS=%AD_ARGS% -AdRootInventoryCsv ""%PARENT_AD_CSV%"""
if not "%W11UT_AD_DOMAIN%"=="" set "AD_ARGS=%AD_ARGS% -AdDomain ""%W11UT_AD_DOMAIN%"""

set "INTUNE_ARGS=-IntuneInventoryCsv ""%LOT_INTUNE_CSV%"""
if exist "%PARENT_INTUNE_CSV%" set "INTUNE_ARGS=%INTUNE_ARGS% -IntuneRootInventoryCsv ""%PARENT_INTUNE_CSV%"""
if not "%W11UT_INTUNE_TENANT_ID%"=="" set "INTUNE_ARGS=%INTUNE_ARGS% -IntuneTenantId ""%W11UT_INTUNE_TENANT_ID%"""
if not "%W11UT_INTUNE_INVENTORY_PAGE_SIZE%"=="" set "INTUNE_ARGS=%INTUNE_ARGS% -IntuneInventoryPageSize %W11UT_INTUNE_INVENTORY_PAGE_SIZE%"
if /I "%W11UT_SKIP_INTUNE_INVENTORY_REFRESH%"=="1" set "INTUNE_ARGS=%INTUNE_ARGS% -SkipIntuneInventoryRefresh"
echo PowerShell    : %POWERSHELL_EXE%
"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -ComputerListPath "%COMPUTERS%" -LogRoot "%PSEXEC_LOGS%" -ReportRoot "%REPORTS%" -CentralLogRoot "%CENTRAL_LOGS%" -LauncherLogRoot "%RUN_DIR%\Logs" -CentralLogCollectionMode "%W11UT_CENTRAL_LOG_COLLECTION_MODE%" %RUN_ONCE_ARG% %IGNORE_RUN_GUARD_ARG% %ACTION_ARGS% %SETUP_ARGS% %PSEXEC_ARG% %THROTTLE_ARG% %GLOBAL_CONCURRENCY_ARG% %AD_ARGS% %INTUNE_ARGS% %TECH_RUN_GUARD_ARGS% -DelayBetweenComputersSeconds %W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS% -DelayBetweenCyclesMinutes %W11UT_DELAY_BETWEEN_CYCLES_MINUTES% -PsExecTimeoutMinutes %W11UT_PSEXEC_TIMEOUT_MINUTES% -CancellationDrainTimeoutMinutes %W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES% %*

set "EXITCODE=%ERRORLEVEL%"
if "%EXITCODE%"=="-1073741819" call :CapturePowerShellCrash
if "%EXITCODE%"=="3221225477" call :CapturePowerShellCrash
if "%EXITCODE%"=="9009" (
    echo ERROR: Command not found while launching the Windows 11 Upgrade Toolkit.
    echo        Check PowerShell, PsExec.exe, the LOT wrapper arguments, and any custom command line suffix.
)

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
exit /b %EXITCODE%

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

:LOAD_CONFIG
set "CONFIG_FILE=%~1"
if not exist "%CONFIG_FILE%" exit /b 0

for /f "usebackq eol=# tokens=1* delims==" %%A in ("%CONFIG_FILE%") do (
    if /I "%%~A"=="W11UT_AUDIT_ONLY" if not defined W11UT_AUDIT_ONLY set "W11UT_AUDIT_ONLY=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_POLICY_REPAIR" if not defined W11UT_ALLOW_POLICY_REPAIR set "W11UT_ALLOW_POLICY_REPAIR=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_WU_RESET" if not defined W11UT_ALLOW_WU_RESET set "W11UT_ALLOW_WU_RESET=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_FORCE_UPGRADE" if not defined W11UT_ALLOW_FORCE_UPGRADE set "W11UT_ALLOW_FORCE_UPGRADE=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_SETUP_UPGRADE" if not defined W11UT_ALLOW_SETUP_UPGRADE set "W11UT_ALLOW_SETUP_UPGRADE=%%~B"
    if /I "%%~A"=="W11UT_DIRECT_SETUP_UPGRADE" if not defined W11UT_DIRECT_SETUP_UPGRADE set "W11UT_DIRECT_SETUP_UPGRADE=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_SETUP_PROFILE_REPAIR" if not defined W11UT_ALLOW_SETUP_PROFILE_REPAIR set "W11UT_ALLOW_SETUP_PROFILE_REPAIR=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_REBOOT" if not defined W11UT_ALLOW_REBOOT set "W11UT_ALLOW_REBOOT=%%~B"
    if /I "%%~A"=="W11UT_SCHEDULE_RETRY_AFTER_REBOOT" if not defined W11UT_SCHEDULE_RETRY_AFTER_REBOOT set "W11UT_SCHEDULE_RETRY_AFTER_REBOOT=%%~B"
    if /I "%%~A"=="W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS" if not defined W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS set "W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS=%%~B"
    if /I "%%~A"=="W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS" if not defined W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS set "W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS=%%~B"
    if /I "%%~A"=="W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS" if not defined W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS set "W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS=%%~B"
    if /I "%%~A"=="W11UT_SETUP_REBOOT_WHEN_NO_USER" if not defined W11UT_SETUP_REBOOT_WHEN_NO_USER set "W11UT_SETUP_REBOOT_WHEN_NO_USER=%%~B"
    if /I "%%~A"=="W11UT_SKIP_VIRTUAL_MACHINES" if not defined W11UT_SKIP_VIRTUAL_MACHINES set "W11UT_SKIP_VIRTUAL_MACHINES=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_DISK_CLEANUP" if not defined W11UT_ALLOW_DISK_CLEANUP set "W11UT_ALLOW_DISK_CLEANUP=%%~B"
    if /I "%%~A"=="W11UT_ALLOW_ADVANCED_DISK_CLEANUP" if not defined W11UT_ALLOW_ADVANCED_DISK_CLEANUP set "W11UT_ALLOW_ADVANCED_DISK_CLEANUP=%%~B"
    if /I "%%~A"=="W11UT_SKIP_SETUP_MEDIA_PRECOPY" if not defined W11UT_SKIP_SETUP_MEDIA_PRECOPY set "W11UT_SKIP_SETUP_MEDIA_PRECOPY=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SOURCE" if not defined W11UT_SETUP_SOURCE set "W11UT_SETUP_SOURCE=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SOURCE_MAP" if not defined W11UT_SETUP_SOURCE_MAP set "W11UT_SETUP_SOURCE_MAP=%%~B"
    if /I "%%~A"=="W11UT_SETUP_EXECUTION_MODE" if not defined W11UT_SETUP_EXECUTION_MODE set "W11UT_SETUP_EXECUTION_MODE=%%~B"
    if /I "%%~A"=="W11UT_SETUP_MEDIA_ID" if not defined W11UT_SETUP_MEDIA_ID set "W11UT_SETUP_MEDIA_ID=%%~B"
    if /I "%%~A"=="W11UT_SETUP_LANGUAGE" if not defined W11UT_SETUP_LANGUAGE set "W11UT_SETUP_LANGUAGE=%%~B"
    if /I "%%~A"=="W11UT_SETUP_DYNAMIC_UPDATE" if not defined W11UT_SETUP_DYNAMIC_UPDATE set "W11UT_SETUP_DYNAMIC_UPDATE=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SOURCE_CANDIDATE_LIMIT" if not defined W11UT_SETUP_SOURCE_CANDIDATE_LIMIT set "W11UT_SETUP_SOURCE_CANDIDATE_LIMIT=%%~B"
    if /I "%%~A"=="W11UT_SETUP_COPY_IPG_MS" if not defined W11UT_SETUP_COPY_IPG_MS set "W11UT_SETUP_COPY_IPG_MS=%%~B"
    if /I "%%~A"=="W11UT_SETUP_COPY_JITTER_SECONDS" if not defined W11UT_SETUP_COPY_JITTER_SECONDS set "W11UT_SETUP_COPY_JITTER_SECONDS=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT" if not defined W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT set "W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES" if not defined W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES set "W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT" if not defined W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT set "W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT" if not defined W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT set "W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SUBNET_PREFIX_LENGTH" if not defined W11UT_SETUP_SUBNET_PREFIX_LENGTH set "W11UT_SETUP_SUBNET_PREFIX_LENGTH=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES" if not defined W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES set "W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT" if not defined W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT set "W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT=%%~B"
    if /I "%%~A"=="W11UT_CONFIRM_LOCAL_SETUP_SOURCE" if not defined W11UT_CONFIRM_LOCAL_SETUP_SOURCE set "W11UT_CONFIRM_LOCAL_SETUP_SOURCE=%%~B"
    if /I "%%~A"=="W11UT_AD_DOMAIN" if not defined W11UT_AD_DOMAIN set "W11UT_AD_DOMAIN=%%~B"
    if /I "%%~A"=="W11UT_INTUNE_TENANT_ID" if not defined W11UT_INTUNE_TENANT_ID set "W11UT_INTUNE_TENANT_ID=%%~B"
    if /I "%%~A"=="W11UT_INTUNE_INVENTORY_PAGE_SIZE" if not defined W11UT_INTUNE_INVENTORY_PAGE_SIZE set "W11UT_INTUNE_INVENTORY_PAGE_SIZE=%%~B"
    if /I "%%~A"=="W11UT_SKIP_INTUNE_INVENTORY_REFRESH" if not defined W11UT_SKIP_INTUNE_INVENTORY_REFRESH set "W11UT_SKIP_INTUNE_INVENTORY_REFRESH=%%~B"
    if /I "%%~A"=="W11UT_THROTTLE" if not defined W11UT_THROTTLE set "W11UT_THROTTLE=%%~B"
    if /I "%%~A"=="W11UT_GLOBAL_CONCURRENCY_LIMIT" if not defined W11UT_GLOBAL_CONCURRENCY_LIMIT set "W11UT_GLOBAL_CONCURRENCY_LIMIT=%%~B"
    if /I "%%~A"=="W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES" if not defined W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES set "W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS" if not defined W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS set "W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS=%%~B"
    if /I "%%~A"=="W11UT_DELAY_BETWEEN_CYCLES_MINUTES" if not defined W11UT_DELAY_BETWEEN_CYCLES_MINUTES set "W11UT_DELAY_BETWEEN_CYCLES_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_PSEXEC_TIMEOUT_MINUTES" if not defined W11UT_PSEXEC_TIMEOUT_MINUTES set "W11UT_PSEXEC_TIMEOUT_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES" if not defined W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES set "W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_CENTRAL_LOG_COLLECTION_MODE" if not defined W11UT_CENTRAL_LOG_COLLECTION_MODE set "W11UT_CENTRAL_LOG_COLLECTION_MODE=%%~B"
    if /I "%%~A"=="W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY" if not defined W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY set "W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY=%%~B"
    if /I "%%~A"=="W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY" if not defined W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY set "W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY=%%~B"
    if /I "%%~A"=="W11UT_RUN_GUARD_HOURS" if not defined W11UT_RUN_GUARD_HOURS set "W11UT_RUN_GUARD_HOURS=%%~B"
)
exit /b 0
