@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    echo Launcher path:
    echo   %~dp0
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

set "SMARTM365_ROOT=%SCRIPT_DIR%..\..\..\..\"
for %%I in ("%SMARTM365_ROOT%") do set "SMARTM365_ROOT=%%~fI\"

set "CACHE_BASE=%ProgramData%\SmartM365\LauncherCache\ExchangeLocalMailboxes"
mkdir "%CACHE_BASE%" >nul 2>&1
if errorlevel 1 set "CACHE_BASE=%TEMP%\SmartM365LauncherCache\ExchangeLocalMailboxes"
mkdir "%CACHE_BASE%" >nul 2>&1
if errorlevel 1 (
    echo Failed to create local cache base folder:
    echo   %CACHE_BASE%
    pause
    popd
    exit /b 1
)

set "LOCAL_SMARTM365_ROOT=%CACHE_BASE%\SmartM365\"
set "LOCAL_SCRIPT_DIR=%LOCAL_SMARTM365_ROOT%SmartInventory\ExchangeInventory\OnPremises\Mailboxes\"
set "CACHE_LOG=%CACHE_BASE%\PrepareCache.log"

if exist "%LOCAL_SMARTM365_ROOT%" rmdir /s /q "%LOCAL_SMARTM365_ROOT%" >nul 2>&1
mkdir "%LOCAL_SCRIPT_DIR%" >nul 2>&1
mkdir "%LOCAL_SMARTM365_ROOT%Config" >nul 2>&1
mkdir "%LOCAL_SMARTM365_ROOT%Modules\SmartM365.Core" >nul 2>&1

echo Preparing local SmartM365 PowerShell runtime cache...
echo Source script folder:
echo   %SCRIPT_DIR%
echo Local cache folder:
echo   %LOCAL_SMARTM365_ROOT%

robocopy "%SCRIPT_DIR%" "%LOCAL_SCRIPT_DIR%" /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /LOG:"%CACHE_LOG%"
set "ROBOCOPY_EXIT=%ERRORLEVEL%"
if %ROBOCOPY_EXIT% GEQ 8 (
    echo Failed to prepare local script cache from:
    echo   %SCRIPT_DIR%
    echo Robocopy exit code: %ROBOCOPY_EXIT%
    echo Robocopy log:
    echo   %CACHE_LOG%
    pause
    popd
    exit /b 1
)

robocopy "%SMARTM365_ROOT%Config" "%LOCAL_SMARTM365_ROOT%Config" /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /LOG+:"%CACHE_LOG%"
set "ROBOCOPY_EXIT=%ERRORLEVEL%"
if %ROBOCOPY_EXIT% GEQ 8 (
    echo Failed to prepare local Config cache from:
    echo   %SMARTM365_ROOT%Config
    echo Robocopy exit code: %ROBOCOPY_EXIT%
    echo Robocopy log:
    echo   %CACHE_LOG%
    pause
    popd
    exit /b 1
)

robocopy "%SMARTM365_ROOT%Modules\SmartM365.Core" "%LOCAL_SMARTM365_ROOT%Modules\SmartM365.Core" /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /LOG+:"%CACHE_LOG%"
set "ROBOCOPY_EXIT=%ERRORLEVEL%"
if %ROBOCOPY_EXIT% GEQ 8 (
    echo Failed to prepare local SmartM365.Core cache from:
    echo   %SMARTM365_ROOT%Modules\SmartM365.Core
    echo Robocopy exit code: %ROBOCOPY_EXIT%
    echo Robocopy log:
    echo   %CACHE_LOG%
    pause
    popd
    exit /b 1
)

"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%LOCAL_SCRIPT_DIR%','%LOCAL_SMARTM365_ROOT%Config','%LOCAL_SMARTM365_ROOT%Modules\SmartM365.Core' -Include *.ps1,*.psm1,*.psd1 -File -Recurse -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1
"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -File "%LOCAL_SCRIPT_DIR%SmartM365-Exchange-Local-Mailboxes-Inventory.ps1" -Tenant test -OnlyADPermission
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
