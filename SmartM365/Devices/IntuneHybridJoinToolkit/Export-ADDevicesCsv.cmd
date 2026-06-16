@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

rem Root launcher.
rem Uses Scripts\SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1 and stores DevicesAD.csv in the root folder.
rem Without -Domain, exports all domains in the current AD forest.
rem Pass -Domain contoso.local, or set EHJIR_AD_DOMAIN before launching, to limit the export to one domain.
set "ROOT_DIR=%CD%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1"
set "OUTPUT=%ROOT_DIR%DevicesAD.csv"

if not exist "%SCRIPT%" (
    echo ERROR: Script not found:
    echo "%SCRIPT%"
    set "EXITCODE=1"
    goto :END
)

set "AD_DOMAIN_ARG="
if not "%EHJIR_AD_DOMAIN%"=="" (
    set "AD_DOMAIN_ARG=-Domain ""%EHJIR_AD_DOMAIN%"""
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -OutputPath "%OUTPUT%" ^
  %AD_DOMAIN_ARG% ^
  %*

set "EXITCODE=%ERRORLEVEL%"

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
popd
exit /b %EXITCODE%
