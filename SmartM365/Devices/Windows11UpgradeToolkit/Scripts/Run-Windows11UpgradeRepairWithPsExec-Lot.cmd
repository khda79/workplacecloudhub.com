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
set "PSEXEC_LOGS=%LOT_DIR%PsExecLogs"
set "REPORTS=%LOT_DIR%Reports"
set "CENTRAL_LOGS=%LOT_DIR%CentralLogs"
set "PSEXEC_EXE=%ROOT_DIR%\Scripts\PsExec.exe"

rem Default LOT behavior matches the operator GUI: guarded repair and upgrade are enabled.
rem Set any W11UT_* switch to 0 before launching a LOT to disable that action.
if not defined W11UT_ALLOW_POLICY_REPAIR set "W11UT_ALLOW_POLICY_REPAIR=1"
if not defined W11UT_ALLOW_WU_RESET set "W11UT_ALLOW_WU_RESET=1"
if not defined W11UT_ALLOW_FORCE_UPGRADE set "W11UT_ALLOW_FORCE_UPGRADE=1"
if not defined W11UT_ALLOW_SETUP_UPGRADE set "W11UT_ALLOW_SETUP_UPGRADE=1"
if not defined W11UT_ALLOW_REBOOT set "W11UT_ALLOW_REBOOT=1"
if not defined W11UT_SKIP_VIRTUAL_MACHINES set "W11UT_SKIP_VIRTUAL_MACHINES=1"
if not defined W11UT_SETUP_SOURCE set "W11UT_SETUP_SOURCE=%ROOT_DIR%\SetupSource"
if not defined W11UT_SETUP_EXECUTION_MODE set "W11UT_SETUP_EXECUTION_MODE=LocalCache"
if not defined W11UT_SETUP_MEDIA_ID set "W11UT_SETUP_MEDIA_ID=Win11"
if not defined W11UT_SETUP_LANGUAGE set "W11UT_SETUP_LANGUAGE=MatchSystem"

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
if not "%W11UT_SETUP_SOURCE%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupSourcePath ""%W11UT_SETUP_SOURCE%"""
if not "%W11UT_SETUP_EXECUTION_MODE%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupExecutionMode ""%W11UT_SETUP_EXECUTION_MODE%"""
if not "%W11UT_SETUP_MEDIA_ID%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupMediaId ""%W11UT_SETUP_MEDIA_ID%"""
if not "%W11UT_SETUP_LANGUAGE%"=="" set "SETUP_ARGS=%SETUP_ARGS% -SetupLanguage ""%W11UT_SETUP_LANGUAGE%"""
if /I "%W11UT_SKIP_SETUP_MEDIA_PRECOPY%"=="1" set "SETUP_ARGS=%SETUP_ARGS% -SkipSetupMediaPreCopy"

set "ACTION_ARGS="
if /I "%W11UT_AUDIT_ONLY%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AuditOnly"
if /I "%W11UT_ALLOW_POLICY_REPAIR%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowPolicyRepair"
if /I "%W11UT_ALLOW_WU_RESET%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowWUReset"
if /I "%W11UT_ALLOW_FORCE_UPGRADE%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowForceUpgrade"
if /I "%W11UT_ALLOW_REBOOT%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowReboot"
if /I "%W11UT_SKIP_VIRTUAL_MACHINES%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -SkipVirtualMachines"

if not defined W11UT_GLOBAL_CONCURRENCY_LIMIT set "W11UT_GLOBAL_CONCURRENCY_LIMIT=15"
if not defined W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES set "W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES=0"
if not defined W11UT_DELAY_BETWEEN_CYCLES_MINUTES set "W11UT_DELAY_BETWEEN_CYCLES_MINUTES=5"
if not defined W11UT_PSEXEC_TIMEOUT_MINUTES set "W11UT_PSEXEC_TIMEOUT_MINUTES=180"

set "GLOBAL_CONCURRENCY_PROVIDED=0"
for %%A in (%*) do (
    if /I "%%~A"=="-GlobalConcurrencyLimit" set "GLOBAL_CONCURRENCY_PROVIDED=1"
)
set "GLOBAL_CONCURRENCY_ARG="
if not "%GLOBAL_CONCURRENCY_PROVIDED%"=="1" (
    set "GLOBAL_CONCURRENCY_ARG=-GlobalConcurrencyLimit %W11UT_GLOBAL_CONCURRENCY_LIMIT% -GlobalConcurrencyLeaseTimeoutMinutes %W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES%"
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -ComputerListPath "%COMPUTERS%" ^
  -LogRoot "%PSEXEC_LOGS%" ^
  -ReportRoot "%REPORTS%" ^
  -CentralLogRoot "%CENTRAL_LOGS%" ^
  %RUN_ONCE_ARG% ^
  %IGNORE_RUN_GUARD_ARG% ^
  %ACTION_ARGS% ^
  %SETUP_ARGS% ^
  %PSEXEC_ARG% ^
  %GLOBAL_CONCURRENCY_ARG% ^
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
