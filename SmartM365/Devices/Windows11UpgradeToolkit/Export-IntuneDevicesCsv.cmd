@echo off
setlocal EnableExtensions

set "UNC_WORK_DIR=%~dp0."

rem Root launcher.
rem Uses Scripts\SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1 and stores DevicesIntune.csv in the root folder.
rem Exports Intune managed devices through Microsoft Graph. Missing devices do not block AD fallback in LOT runs.
rem Uses delegated interactive Microsoft Graph authentication. Set W11UT_INTUNE_TENANT_ID only to constrain the sign-in tenant.
set "ROOT_DIR=%UNC_WORK_DIR%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1"
set "OUTPUT=%ROOT_DIR%DevicesIntune.csv"
call :PrintStartupInfo

if not exist "%SCRIPT%" (
    echo ERROR: Script not found:
    echo "%SCRIPT%"
    set "EXITCODE=1"
    goto :END
)

set "INTUNE_TENANT_ID_ARG="
if not "%W11UT_INTUNE_TENANT_ID%"=="" (
    set "INTUNE_TENANT_ID_ARG=-TenantId ""%W11UT_INTUNE_TENANT_ID%"""
)

set "INTUNE_PAGE_SIZE_ARG="
if not "%W11UT_INTUNE_INVENTORY_PAGE_SIZE%"=="" (
    set "INTUNE_PAGE_SIZE_ARG=-PageSize %W11UT_INTUNE_INVENTORY_PAGE_SIZE%"
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -OutputPath "%OUTPUT%" ^
  %INTUNE_TENANT_ID_ARG% ^
  %INTUNE_PAGE_SIZE_ARG% ^
  %*

set "EXITCODE=%ERRORLEVEL%"

:END
echo.
echo Finished with exit code %EXITCODE%.
pause
exit /b %EXITCODE%

:PrintStartupInfo
echo.
echo Auth    : Delegated interactive Microsoft Graph
echo SmartM365 Windows 11 Upgrade Toolkit - Intune export launcher
echo Started : %DATE% %TIME%
echo Root    : %ROOT_DIR%
echo Script  : %SCRIPT%
echo Output  : %OUTPUT%
if not "%W11UT_INTUNE_TENANT_ID%"=="" echo Tenant  : %W11UT_INTUNE_TENANT_ID%
if not "%W11UT_INTUNE_INVENTORY_PAGE_SIZE%"=="" echo PageSize: %W11UT_INTUNE_INVENTORY_PAGE_SIZE%
echo.
exit /b 0
