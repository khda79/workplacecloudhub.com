@echo off
setlocal EnableExtensions

set "SCRIPT=%~dp0..\..\M365Inventory\IntuneInventory\Devices\SmartM365-Create-DeviceSystemRemediation.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
    for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH_FOUND set "PWSH_FOUND=%%I"
    if not defined PWSH_FOUND exit /b 1
    set "PWSH=%PWSH_FOUND%"
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Tenant prod %*
exit /b %ERRORLEVEL%
