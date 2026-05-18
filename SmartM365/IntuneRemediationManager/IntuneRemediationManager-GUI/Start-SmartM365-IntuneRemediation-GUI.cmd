@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
start "" /min "C:\Program Files\PowerShell\7\pwsh.exe" -STA -NoProfile -WindowStyle Hidden -File "%SCRIPT_DIR%SmartM365-IntuneRemediation-GUI.ps1"
