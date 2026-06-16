@echo off
setlocal
set "W11UT_LOT_DIR=%~dp0"
set "W11UT_RUN_ONCE=0"
call "%~dp0..\Scripts\Run-Windows11UpgradeRepairWithPsExec-Lot.cmd" %*
