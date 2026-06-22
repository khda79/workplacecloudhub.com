<#
.SYNOPSIS
Refreshes tiny CMD wrappers in operational LOT-* folders.

.DESCRIPTION
The real launcher logic lives in Scripts\*.cmd. Operational LOT folders only need tiny
repair wrappers that delegate to the shared launcher. Run this after creating a new LOT
folder or to convert older LOT folders to the shared-wrapper model. The versioned LOT-X
template is intentionally skipped so Git does not see template wrapper churn.

The Intune inventory export is global and must be launched from the repository root,
so this script removes obsolete Export-IntuneDevicesCsv.cmd wrappers from LOT folders.
It also creates a blank AdDomain.txt file when missing so each LOT has a visible
place for optional per-domain AD inventory configuration.

.VERSION
1.0
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$RootPath,
    [switch]$WhatIf,
    [switch]$CheckOnly
)

$ErrorActionPreference = "Stop"
$script:LotWrapperChangesRequired = $false
$script:LotWrapperChangesApplied = $false

function ConvertTo-CrLfText {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return $null
    }

    $normalized = $Text -replace "`r`n|`r|`n", "`n"
    $normalized = $normalized.TrimEnd("`n") + "`n"
    return $normalized -replace "`n", "`r`n"
}

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
$lotFolders = @(
    Get-ChildItem -LiteralPath $rootItem.FullName -Directory -Filter "LOT-*" -ErrorAction Stop |
        Where-Object { $_.Name -ine "LOT-X" }
)

if ($lotFolders.Count -eq 0) {
    Write-Host ("No operational LOT-* folders found under: {0}" -f $rootItem.FullName) -ForegroundColor Yellow
    return
}

$wrappers = @{
    "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd" = @'
@echo off
setlocal EnableExtensions EnableDelayedExpansion

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
setlocal EnableExtensions EnableDelayedExpansion

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
setlocal EnableExtensions EnableDelayedExpansion

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
setlocal EnableExtensions EnableDelayedExpansion

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
        $script:LotWrapperChangesRequired = $true
        if ($WhatIf -or $CheckOnly) {
            Write-Host ("Would create AD domain config: {0}" -f $adDomainPath)
        }
        else {
            New-Item -ItemType File -Path $adDomainPath -Force -ErrorAction Stop | Out-Null
            $script:LotWrapperChangesApplied = $true
            Write-Host ("Created AD domain config: {0}" -f $adDomainPath) -ForegroundColor Green
        }
    }

    foreach ($obsoleteWrapper in $obsoleteWrappers) {
        $obsoletePath = Join-Path $lotFolder.FullName $obsoleteWrapper
        if (-not (Test-Path -LiteralPath $obsoletePath)) {
            continue
        }

        $script:LotWrapperChangesRequired = $true
        if ($WhatIf -or $CheckOnly) {
            Write-Host ("Would remove obsolete wrapper: {0}" -f $obsoletePath)
            continue
        }

        Remove-Item -LiteralPath $obsoletePath -Force -ErrorAction Stop
        $script:LotWrapperChangesApplied = $true
        Write-Host ("Removed obsolete wrapper: {0}" -f $obsoletePath) -ForegroundColor Yellow
    }

    foreach ($wrapperName in $wrappers.Keys) {
        $targetPath = Join-Path $lotFolder.FullName $wrapperName
        $expectedContent = ConvertTo-CrLfText -Text $wrappers[$wrapperName]
        $currentContent = $null
        if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
            $currentContent = ConvertTo-CrLfText -Text ([System.IO.File]::ReadAllText($targetPath))
        }

        if ($null -ne $currentContent -and $currentContent -ceq $expectedContent) {
            Write-Host ("Up to date: {0}" -f $targetPath) -ForegroundColor DarkGray
            continue
        }

        $script:LotWrapperChangesRequired = $true
        if ($WhatIf -or $CheckOnly) {
            Write-Host ("Would update: {0}" -f $targetPath)
            continue
        }

        [System.IO.File]::WriteAllText($targetPath, $expectedContent, [System.Text.Encoding]::ASCII)
        $script:LotWrapperChangesApplied = $true
        Write-Host ("Updated: {0}" -f $targetPath) -ForegroundColor Green
    }
}

if ($CheckOnly -and $script:LotWrapperChangesRequired) {
    exit 2
}
