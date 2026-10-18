@echo off
setlocal
start "SmartM365 Orchestrator Demo" "C:\Program Files\PowerShell\7\pwsh.exe" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0SmartM365-Inventory-Orchestrator-GUI.ps1" -Demo %*
exit /b 0
