@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%CD%\"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "POWERSHELL5=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%POWERSHELL5%" (
    echo Windows PowerShell 5.1 ^(powershell.exe^) was not found.
    echo This Exchange on-premises launcher requires Windows PowerShell 5.1 and the Exchange Management Tools.
    echo Checked:
    echo   %SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
    echo   %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe
    pause
    popd
    exit /b 1
)
echo WARNING: this launcher can modify Exchange proxy addresses.
echo It runs SmartM365-Check-ProxyAddresses-Exchange.ps1 with -AddMissingAddress -SkipAllowListCsv.
set /p CONFIRM_WRITE=Type WRITE to continue: 
if /i not "%CONFIRM_WRITE%"=="WRITE" (
    echo Write mode cancelled.
    pause
    popd
    exit /b 1
)
"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-Check-ProxyAddresses-Exchange.ps1" -Tenant test -AddMissingAddress -SkipAllowListCsv
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
