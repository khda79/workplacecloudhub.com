@echo off
setlocal
start "Smart Device Reboot Manager Demo" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0SmartM365-DeviceRebootManager-GUI.ps1" -Demo %*
exit /b 0
