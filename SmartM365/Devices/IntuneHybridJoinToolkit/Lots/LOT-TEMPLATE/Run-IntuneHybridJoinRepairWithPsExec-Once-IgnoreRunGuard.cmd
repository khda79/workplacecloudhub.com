@echo off
setlocal
set "EHJIR_LOT_DIR=%~dp0"
set "EHJIR_IGNORE_RUN_GUARD=1"
set "EHJIR_RUN_ONCE=1"
call "%~dp0..\..\Scripts\Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd" %*
