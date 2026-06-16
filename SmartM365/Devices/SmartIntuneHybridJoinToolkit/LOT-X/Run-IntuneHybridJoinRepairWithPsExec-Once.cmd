@echo off
setlocal EnableExtensions EnableDelayedExpansion

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)


set "EHJIR_LOT_DIR=%CD%\"
set "EHJIR_IGNORE_RUN_GUARD=0"
set "EHJIR_RUN_ONCE=1"
call "%CD%\..\Scripts\Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
