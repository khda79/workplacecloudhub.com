@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%POWERSHELL5%" set "POWERSHELL5=powershell.exe"

rem A headless console host (Windows 10 2004 and later) keeps the PowerShell console from flashing.
set "HOST_PREFIX="
set "OS_BUILD=0"
for /f "tokens=3" %%B in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v CurrentBuildNumber 2^>nul ^| find "CurrentBuildNumber"') do set "OS_BUILD=%%B"
if exist "%SystemRoot%\System32\conhost.exe" if %OS_BUILD% GEQ 19041 set "HOST_PREFIX="%SystemRoot%\System32\conhost.exe" --headless"
start "Smart Device Reboot Manager Demo" %HOST_PREFIX% "%POWERSHELL5%" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-DeviceRebootManager-GUI.ps1" -Demo %*
exit /b 0
