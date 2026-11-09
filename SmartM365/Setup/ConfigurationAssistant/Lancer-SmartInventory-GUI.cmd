@echo off
rem SmartInventory GUI launcher version 0.2.0
where pwsh.exe >nul 2>&1
if errorlevel 1 (
  echo PowerShell 7 was not found. Install it before starting the assistant.
  pause
  exit /b 1
)
start "" pwsh.exe -NoProfile -STA -WindowStyle Hidden -File "%~dp0SmartM365-Inventory-Configuration-GUI.ps1" %*
exit /b %errorlevel%
