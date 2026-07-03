@echo off
setlocal EnableExtensions EnableDelayedExpansion

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

rem Shared LOT launcher. Keep LOT folders with tiny wrappers only.
rem The LOT wrapper sets W11UT_LOT_DIR and optional execution switches.

set "LOT_DIR=%W11UT_LOT_DIR%"
if "%LOT_DIR%"=="" set "LOT_DIR=%CD%\.."

for %%I in ("%LOT_DIR%") do set "LOT_DIR=%%~fI\"
set "ROOT_DIR=%LOT_DIR%.."
for %%I in ("%ROOT_DIR%") do set "ROOT_DIR=%%~fI"

set "SCRIPT=%ROOT_DIR%\Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1"
set "COMPUTERS=%LOT_DIR%Computers.txt"
set "PARENT_AD_CSV=%ROOT_DIR%\DevicesAD.csv"
set "LOT_AD_CSV=%LOT_DIR%DevicesAD.csv"
set "AD_DOMAIN_FILE=%LOT_DIR%AdDomain.txt"
set "PSEXEC_LOGS=%LOT_DIR%PsExecLogs"
set "REPORTS=%LOT_DIR%Reports"
set "CENTRAL_LOGS=%LOT_DIR%CentralLogs"
set "PSEXEC_EXE=%ROOT_DIR%\Scripts\PsExec.exe"
set "LOT_CONFIG_FILE=%LOT_DIR%Windows11UpgradeToolkit.config"
set "ROOT_CONFIG_FILE=%ROOT_DIR%\Windows11UpgradeToolkit.config"

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
if not defined W11UT_SETUP_REBOOT_WHEN_NO_USER set "W11UT_SETUP_REBOOT_WHEN_NO_USER=1"
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
if not defined W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES set "W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES=60"

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

set "RUN_ONCE_ARG="
if /I "%W11UT_RUN_ONCE%"=="1" set "RUN_ONCE_ARG=-RunOnce"

set "IGNORE_RUN_GUARD_ARG="
if /I "%W11UT_IGNORE_RUN_GUARD%"=="1" set "IGNORE_RUN_GUARD_ARG=-IgnoreRunGuard"

set "SETUP_ARGS="
if /I "%W11UT_ALLOW_SETUP_UPGRADE%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -AllowSetupUpgrade"
if /I "%W11UT_DIRECT_SETUP_UPGRADE%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -DirectSetupUpgrade"
if /I "%W11UT_SETUP_REBOOT_WHEN_NO_USER%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -AllowSetupCompletionRebootWhenNoUser"
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
if /I "%W11UT_SKIP_VIRTUAL_MACHINES%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -SkipVirtualMachines"
if /I "%W11UT_ALLOW_DISK_CLEANUP%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowDiskCleanup"
if /I "%W11UT_ALLOW_ADVANCED_DISK_CLEANUP%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowAdvancedDiskCleanup"

if not defined W11UT_GLOBAL_CONCURRENCY_LIMIT set "W11UT_GLOBAL_CONCURRENCY_LIMIT=15"
if not defined W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES set "W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES=0"
if not defined W11UT_THROTTLE set "W11UT_THROTTLE=10"
if not defined W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS set "W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS=0"
if not defined W11UT_DELAY_BETWEEN_CYCLES_MINUTES set "W11UT_DELAY_BETWEEN_CYCLES_MINUTES=5"
if not defined W11UT_PSEXEC_TIMEOUT_MINUTES set "W11UT_PSEXEC_TIMEOUT_MINUTES=180"
if not defined W11UT_CENTRAL_LOG_COLLECTION_MODE set "W11UT_CENTRAL_LOG_COLLECTION_MODE=Standard"

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

set "AD_ARGS=-AdInventoryCsv ""%LOT_AD_CSV%"""
if exist "%PARENT_AD_CSV%" set "AD_ARGS=%AD_ARGS% -AdRootInventoryCsv ""%PARENT_AD_CSV%"""
if not "%W11UT_AD_DOMAIN%"=="" set "AD_ARGS=%AD_ARGS% -AdDomain ""%W11UT_AD_DOMAIN%"""
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -ComputerListPath "%COMPUTERS%" ^
  -LogRoot "%PSEXEC_LOGS%" ^
  -ReportRoot "%REPORTS%" ^
  -CentralLogRoot "%CENTRAL_LOGS%" ^
  -CentralLogCollectionMode "%W11UT_CENTRAL_LOG_COLLECTION_MODE%" ^
  %RUN_ONCE_ARG% ^
  %IGNORE_RUN_GUARD_ARG% ^
  %ACTION_ARGS% ^
  %SETUP_ARGS% ^
  %PSEXEC_ARG% ^
  %THROTTLE_ARG% ^
  %GLOBAL_CONCURRENCY_ARG% ^
  %AD_ARGS% ^
  -DelayBetweenComputersSeconds %W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS% ^
  -DelayBetweenCyclesMinutes %W11UT_DELAY_BETWEEN_CYCLES_MINUTES% ^
  -PsExecTimeoutMinutes %W11UT_PSEXEC_TIMEOUT_MINUTES% ^
  %*

set "EXITCODE=%ERRORLEVEL%"

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
popd
exit /b %EXITCODE%

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
    if /I "%%~A"=="W11UT_ALLOW_REBOOT" if not defined W11UT_ALLOW_REBOOT set "W11UT_ALLOW_REBOOT=%%~B"
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
    if /I "%%~A"=="W11UT_THROTTLE" if not defined W11UT_THROTTLE set "W11UT_THROTTLE=%%~B"
    if /I "%%~A"=="W11UT_GLOBAL_CONCURRENCY_LIMIT" if not defined W11UT_GLOBAL_CONCURRENCY_LIMIT set "W11UT_GLOBAL_CONCURRENCY_LIMIT=%%~B"
    if /I "%%~A"=="W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES" if not defined W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES set "W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS" if not defined W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS set "W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS=%%~B"
    if /I "%%~A"=="W11UT_DELAY_BETWEEN_CYCLES_MINUTES" if not defined W11UT_DELAY_BETWEEN_CYCLES_MINUTES set "W11UT_DELAY_BETWEEN_CYCLES_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_PSEXEC_TIMEOUT_MINUTES" if not defined W11UT_PSEXEC_TIMEOUT_MINUTES set "W11UT_PSEXEC_TIMEOUT_MINUTES=%%~B"
    if /I "%%~A"=="W11UT_CENTRAL_LOG_COLLECTION_MODE" if not defined W11UT_CENTRAL_LOG_COLLECTION_MODE set "W11UT_CENTRAL_LOG_COLLECTION_MODE=%%~B"
)
exit /b 0
