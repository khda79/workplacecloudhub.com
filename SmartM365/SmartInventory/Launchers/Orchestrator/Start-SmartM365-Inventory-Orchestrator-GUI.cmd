@echo off
setlocal
for /f "tokens=1-2 delims=." %%a in ("%time%") do set "SMART_TIME=%%a:%%b"
echo [%date% %SMART_TIME%] Starting SmartM365 Orchestrator GUI...
set "SCRIPT=%~dp0..\..\Orchestrator\SmartM365-Inventory-Orchestrator-GUI.ps1"
if not exist "%SCRIPT%" (
    for /f "tokens=1-2 delims=." %%a in ("%time%") do set "SMART_TIME=%%a:%%b"
    echo [%date% %SMART_TIME%] ERROR: GUI script not found: %SCRIPT%
    exit /b 2
)
"C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT%" -Tenant prod
set "EXIT_CODE=%ERRORLEVEL%"
for /f "tokens=1-2 delims=." %%a in ("%time%") do set "SMART_TIME=%%a:%%b"
echo [%date% %SMART_TIME%] SmartM365 Orchestrator GUI closed with exit code %EXIT_CODE%.
exit /b %EXIT_CODE%
