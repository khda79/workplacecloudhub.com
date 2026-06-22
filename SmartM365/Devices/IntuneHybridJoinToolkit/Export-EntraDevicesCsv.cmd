@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)


rem Root launcher.
rem Uses Scripts\SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1 and stores DevicesEntra.csv in the root folder.
rem Always exports the full Microsoft Entra devices inventory.
set "ROOT_DIR=%CD%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1"
set "OUTPUT=%ROOT_DIR%DevicesEntra.csv"
call :PrintStartupInfo

if not exist "%SCRIPT%" (
    echo ERROR: Script not found:
    echo "%SCRIPT%"
    set "EXITCODE=1"
    goto :END
)

net session >nul 2>&1
if errorlevel 1 (
    echo ERROR: Run this CMD from an elevated administrator command prompt.
    echo The output file is written in this folder:
    echo "%OUTPUT%"
    set "EXITCODE=1"
    goto :END
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -OutputPath "%OUTPUT%" ^
  -PageSize 999 ^
  %*

set "EXITCODE=%ERRORLEVEL%"

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
popd
exit /b %EXITCODE%
:PrintStartupInfo
echo.
echo SmartM365 Intune Hybrid Join Toolkit - Entra export launcher
echo Started : %DATE% %TIME%
echo Root    : %ROOT_DIR%
echo Script  : %SCRIPT%
echo Output  : %OUTPUT%
echo Page    : 999
echo.
exit /b 0
