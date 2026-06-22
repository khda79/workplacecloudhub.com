@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)


rem Root launcher.
rem Refreshes the tiny CMD wrappers in all LOT-* folders.
set "ROOT_DIR=%CD%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1"
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
    echo It updates CMD files inside LOT-* folders under:
    echo "%ROOT_DIR%"
    set "EXITCODE=1"
    goto :END
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -RootPath "%ROOT_DIR%." ^
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
echo SmartM365 Intune Hybrid Join Toolkit - LOT wrapper refresh launcher
echo Started : %DATE% %TIME%
echo Root    : %ROOT_DIR%
echo Script  : %SCRIPT%
echo Target  : LOT-* folders under %ROOT_DIR%
echo.
exit /b 0
