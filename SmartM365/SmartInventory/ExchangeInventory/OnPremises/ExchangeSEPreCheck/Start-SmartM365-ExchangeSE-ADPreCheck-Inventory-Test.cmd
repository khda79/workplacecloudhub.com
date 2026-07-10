@echo off
setlocal EnableExtensions
set "SCRIPT_DIR=%~dp0"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "POWERSHELL5=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%POWERSHELL5%" (
    echo Windows PowerShell 5.1 ^(powershell.exe^) was not found.
    echo This Exchange on-premises launcher requires Windows PowerShell 5.1 and Active Directory tools.
    echo Checked:
    echo   %SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
    echo   %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe
    exit /b 1
)
"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-ExchangeSE-ADPreCheck-Inventory.ps1" -Tenant test %*
exit /b %ERRORLEVEL%
