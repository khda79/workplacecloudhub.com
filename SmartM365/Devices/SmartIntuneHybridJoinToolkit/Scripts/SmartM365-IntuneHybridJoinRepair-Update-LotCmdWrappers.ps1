<#
.SYNOPSIS
Refreshes tiny CMD wrappers in all LOT-* folders.

.DESCRIPTION
The real launcher logic lives in Scripts\*.cmd. LOT folders only need tiny repair
wrappers that delegate to the shared launcher. Run this after creating a new LOT folder
or to convert older LOT folders to the shared-wrapper model.

The Intune inventory export is global and must be launched from the repository root,
so this script removes obsolete Export-IntuneDevicesCsv.cmd wrappers from LOT folders.
It also creates a blank AdDomain.txt file when missing so each LOT has a visible
place for optional per-domain AD inventory configuration.
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$RootPath,
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"

$ScriptDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

if ([string]::IsNullOrWhiteSpace($RootPath)) {
    $RootPath = Split-Path -Parent $ScriptDir
}
else {
    $RootPath = $RootPath.Trim().Trim([char]34)
}

$rootItem = Get-Item -LiteralPath $RootPath -ErrorAction Stop
$lotFolders = @(Get-ChildItem -LiteralPath $rootItem.FullName -Directory -Filter "LOT-*" -ErrorAction Stop)

if ($lotFolders.Count -eq 0) {
    Write-Host ("No LOT-* folders found under: {0}" -f $rootItem.FullName) -ForegroundColor Yellow
    return
}

$wrappers = @{
    "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd" = @'
@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "EHJIR_LOT_DIR=%CD%\"
set "EHJIR_IGNORE_RUN_GUARD=0"
set "EHJIR_RUN_ONCE=0"
call "%CD%\..\Scripts\Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
'@
    "Run-IntuneHybridJoinRepairWithPsExec-Once.cmd" = @'
@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "EHJIR_LOT_DIR=%CD%\"
set "EHJIR_IGNORE_RUN_GUARD=0"
set "EHJIR_RUN_ONCE=1"
call "%CD%\..\Scripts\Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
'@
    "Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd" = @'
@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "EHJIR_LOT_DIR=%CD%\"
set "EHJIR_IGNORE_RUN_GUARD=1"
set "EHJIR_RUN_ONCE=1"
call "%CD%\..\Scripts\Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
'@
    "Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd" = @'
@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "EHJIR_LOT_DIR=%CD%\"
set "EHJIR_IGNORE_RUN_GUARD=1"
set "EHJIR_RUN_ONCE=0"
call "%CD%\..\Scripts\Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
'@
}

$obsoleteWrappers = @(
    "Export-IntuneDevicesCsv.cmd",
    "Run-IntuneHybridJoinRepairWithPsExec.cmd",
    "Run-IntuneHybridJoinRepairWithPsExec-IgnoreRunGuard.cmd"
)

foreach ($lotFolder in $lotFolders) {
    $adDomainPath = Join-Path $lotFolder.FullName "AdDomain.txt"
    if (-not (Test-Path -LiteralPath $adDomainPath)) {
        if ($WhatIf) {
            Write-Host ("Would create AD domain config: {0}" -f $adDomainPath)
        }
        else {
            New-Item -ItemType File -Path $adDomainPath -Force -ErrorAction Stop | Out-Null
            Write-Host ("Created AD domain config: {0}" -f $adDomainPath) -ForegroundColor Green
        }
    }

    foreach ($obsoleteWrapper in $obsoleteWrappers) {
        $obsoletePath = Join-Path $lotFolder.FullName $obsoleteWrapper
        if (-not (Test-Path -LiteralPath $obsoletePath)) {
            continue
        }

        if ($WhatIf) {
            Write-Host ("Would remove obsolete wrapper: {0}" -f $obsoletePath)
            continue
        }

        Remove-Item -LiteralPath $obsoletePath -Force -ErrorAction Stop
        Write-Host ("Removed obsolete wrapper: {0}" -f $obsoletePath) -ForegroundColor Yellow
    }

    foreach ($wrapperName in $wrappers.Keys) {
        $targetPath = Join-Path $lotFolder.FullName $wrapperName
        if ($WhatIf) {
            Write-Host ("Would update: {0}" -f $targetPath)
            continue
        }

        Set-Content -LiteralPath $targetPath -Value $wrappers[$wrapperName] -Encoding ASCII -Force
        Write-Host ("Updated: {0}" -f $targetPath) -ForegroundColor Green
    }
}
