"""Create a PowerShell 7 launcher beside a generated cleanup script."""

import argparse
import re
from pathlib import Path


LAUNCHER_CONTENT = r'''@echo off
setlocal
set "TARGET_SCRIPT=%~dpn0.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=pwsh.exe"

if not exist "%TARGET_SCRIPT%" (
    powershell.exe -NoProfile -Command "Write-Host ('[{0}] Generated cleanup script not found: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:TARGET_SCRIPT)"
    exit /b 1
)

"%PWSH%" -NoProfile -Command "if ($PSVersionTable.PSVersion -ge [version]'7.4') { exit 0 } else { exit 1 }" >nul 2>&1
if errorlevel 1 (
    powershell.exe -NoProfile -Command "Write-Host ('[{0}] PowerShell 7.4 or later is required for generated cleanup scripts.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
    exit /b 1
)

"%PWSH%" -NoProfile -Command "$content = [IO.File]::ReadAllText($env:TARGET_SCRIPT); $required = @('SmartM365-CleanupSafety: RecycleDefault'); if ($env:TARGET_SCRIPT -match 'ExtraFilesRemove') { $required += 'SmartM365-CleanupSafety: HomeAspxProtected' }; if (@($required | Where-Object { -not $content.Contains($_) }).Count -gt 0) { exit 1 }" >nul 2>&1
if errorlevel 1 (
    powershell.exe -NoProfile -Command "Write-Host ('[{0}] Cleanup safety checks are missing. Regenerate the PowerShell script with the current toolkit.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
    exit /b 1
)

"%PWSH%" -NoProfile -Command "Write-Host ('[{0}] Script: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:TARGET_SCRIPT); Write-Host ('[{0}] Mode: EXECUTE - RECYCLE. The script will request explicit confirmation.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%TARGET_SCRIPT%" -Execute -Recycle
exit /b %ERRORLEVEL%
'''


def launcher_path(script_path: Path) -> Path:
    return Path(script_path).with_suffix(".cmd")


def write_ps7_launcher(script_path: Path) -> Path:
    result = launcher_path(script_path)
    result.write_bytes(LAUNCHER_CONTENT.replace("\n", "\r\n").encode("ascii"))
    return result


def remove_ps7_launcher(script_path: Path) -> None:
    launcher_path(script_path).unlink(missing_ok=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Add a PS7 launcher to an existing generated cleanup script.")
    parser.add_argument("script", type=Path)
    args = parser.parse_args()
    if not args.script.is_file() or not re.fullmatch(
        r"SmartM365-(?:.+-)?SharePointTarget-Extra(?:Files|Folders|Libraries)Remove-.+\.ps1",
        args.script.name,
        flags=re.IGNORECASE,
    ):
        parser.error("Expected an existing generated SharePoint target cleanup .ps1 script.")
    print(f"PowerShell 7 launcher created: {write_ps7_launcher(args.script)}")
