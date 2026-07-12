@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "INSTALL_SCRIPT=%SCRIPT_DIR%Install-WorkplaceCloudHub-CodeSigningCertificate.ps1"

if not exist "%INSTALL_SCRIPT%" (
    echo [ERROR] Certificate installer not found:
    echo   %INSTALL_SCRIPT%
    exit /b 1
)

echo Installing workplacecloudhub.com public code-signing certificate for the current user...
echo Script:
echo   %INSTALL_SCRIPT%
echo.

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%INSTALL_SCRIPT%" -StoreLocation CurrentUser
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo [ERROR] Certificate installation failed with exit code %EXIT_CODE%.
    pause
    exit /b %EXIT_CODE%
)

echo.
echo [OK] Certificate installation completed.
exit /b 0
