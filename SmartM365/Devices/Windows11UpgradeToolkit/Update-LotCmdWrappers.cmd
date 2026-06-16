@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the toolkit directory.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\SmartM365-Windows11Upgrade-Update-LotCmdWrappers.ps1" %*
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Finished with exit code %EXITCODE%.
pause
popd
exit /b %EXITCODE%
