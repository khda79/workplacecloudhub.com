@echo off
setlocal EnableExtensions
rem SmartInventory single entry point (catalog: Orchestrator\Orchestrator-Jobs.json.template).
rem   SmartInventory.cmd                      lists the jobs and manual launches
rem   SmartInventory.cmd <Name> [options]     runs the entry now on this computer
rem   SmartInventory.cmd <Job> -Request       runs the job through the resident orchestrators
rem   SmartInventory.cmd <Name> -WhatIf       shows the resolved command
rem Windows PowerShell 5.1 is enough here (Exchange servers); the engine starts PowerShell 7 when the entry needs it.
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "POWERSHELL5=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Orchestrator\Invoke-SmartM365InventoryJob.ps1" %*
exit /b %ERRORLEVEL%
