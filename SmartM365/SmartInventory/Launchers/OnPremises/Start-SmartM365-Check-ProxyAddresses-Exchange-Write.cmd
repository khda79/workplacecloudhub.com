@echo off
setlocal EnableExtensions

set "UNC_WORK_DIR=%~dp0..\..\ExchangeInventory\OnPremises\ProxyAddresses\."

set "SCRIPT_DIR=%UNC_WORK_DIR%"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "POWERSHELL5=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%POWERSHELL5%" (
    echo Windows PowerShell 5.1 ^(powershell.exe^) was not found.
    echo This Exchange on-premises launcher requires Windows PowerShell 5.1 and the Exchange Management Tools.
    echo Checked:
    echo   %SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
    echo   %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe
    pause
    exit /b 1
)

echo WARNING: this launcher can modify Exchange proxy addresses.
echo It runs SmartM365-Check-ProxyAddresses-Exchange.ps1 with -AllOrganizationalUnit -AddMissingAddress.
set /p CONFIRM_WRITE=Type WRITE to continue: 
if /i not "%CONFIRM_WRITE%"=="WRITE" (
    echo Write mode cancelled.
    pause
    exit /b 1
)
set "SMARTM365_ROOT=%SCRIPT_DIR%\..\..\..\.."
for %%I in ("%SMARTM365_ROOT%") do set "SMARTM365_ROOT=%%~fI"

set "CACHE_BASE="
call :TryCacheBase "%ProgramData%\SmartM365\LauncherCache\ExchangeProxyAddresses"
if not defined CACHE_BASE call :TryCacheBase "%SystemRoot%\Temp\SmartM365LauncherCache\ExchangeProxyAddresses"
if not defined CACHE_BASE call :TryCacheBase "C:\Temp\SmartM365LauncherCache\ExchangeProxyAddresses"
if not defined CACHE_BASE call :TryCacheBase "%TEMP%\SmartM365LauncherCache\ExchangeProxyAddresses"
if not defined CACHE_BASE (
    echo Failed to create a writable local cache base folder.
    echo Tried:
    echo   %ProgramData%\SmartM365\LauncherCache\ExchangeProxyAddresses
    echo   %SystemRoot%\Temp\SmartM365LauncherCache\ExchangeProxyAddresses
    echo   C:\Temp\SmartM365LauncherCache\ExchangeProxyAddresses
    echo   %TEMP%\SmartM365LauncherCache\ExchangeProxyAddresses
    pause
    exit /b 1
)
goto :CacheBaseReady

:TryCacheBase
if defined CACHE_BASE exit /b 0
mkdir "%~1" >nul 2>&1
if not exist "%~1\" exit /b 0
type nul > "%~1\.smartm365-cache-write-test" 2>nul
if errorlevel 1 exit /b 0
del "%~1\.smartm365-cache-write-test" >nul 2>&1
set "CACHE_BASE=%~1"
exit /b 0

:CacheBaseReady

set "LOCAL_SMARTM365_ROOT=%CACHE_BASE%\SmartM365"
set "LOCAL_SCRIPT_DIR=%LOCAL_SMARTM365_ROOT%\SmartInventory\ExchangeInventory\OnPremises\ProxyAddresses"
set "CACHE_LOG=%CACHE_BASE%\PrepareCache.log"

if exist "%LOCAL_SMARTM365_ROOT%" rmdir /s /q "%LOCAL_SMARTM365_ROOT%" >nul 2>&1
mkdir "%LOCAL_SCRIPT_DIR%" >nul 2>&1
mkdir "%LOCAL_SMARTM365_ROOT%\Config" >nul 2>&1
mkdir "%LOCAL_SMARTM365_ROOT%\Modules\SmartM365.Core" >nul 2>&1

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
    exit /b 1
)

robocopy "%SMARTM365_ROOT%\Config" "%LOCAL_SMARTM365_ROOT%\Config" /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /LOG+:"%CACHE_LOG%"
set "ROBOCOPY_EXIT=%ERRORLEVEL%"
if %ROBOCOPY_EXIT% GEQ 8 (
    echo Failed to prepare local Config cache from:
    echo   %SMARTM365_ROOT%\Config
    echo Robocopy exit code: %ROBOCOPY_EXIT%
    echo Robocopy log:
    echo   %CACHE_LOG%
    pause
    exit /b 1
)

robocopy "%SMARTM365_ROOT%\Modules\SmartM365.Core" "%LOCAL_SMARTM365_ROOT%\Modules\SmartM365.Core" /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /LOG+:"%CACHE_LOG%"
set "ROBOCOPY_EXIT=%ERRORLEVEL%"
if %ROBOCOPY_EXIT% GEQ 8 (
    echo Failed to prepare local SmartM365.Core cache from:
    echo   %SMARTM365_ROOT%\Modules\SmartM365.Core
    echo Robocopy exit code: %ROBOCOPY_EXIT%
    echo Robocopy log:
    echo   %CACHE_LOG%
    pause
    exit /b 1
)

"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%LOCAL_SCRIPT_DIR%','%LOCAL_SMARTM365_ROOT%\Config','%LOCAL_SMARTM365_ROOT%\Modules\SmartM365.Core' -Include *.ps1,*.psm1,*.psd1 -File -Recurse -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1
"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -File "%LOCAL_SCRIPT_DIR%\SmartM365-Check-ProxyAddresses-Exchange.ps1" -Tenant prod -AllOrganizationalUnit -AddMissingAddress
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
