<#
.SYNOPSIS
    Refreshes Smart Intune Windows 11 Upgrade Toolkit LOT CMD wrappers.

.DESCRIPTION
    Copies the standard tiny CMD wrappers from LOT-X to operational LOT-* folders.
    LOT-X is the versioned template and is not modified by this script.

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version: 0.1.0
#>

#requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ToolkitRoot
)

$ErrorActionPreference = 'Stop'

function Get-ToolkitRoot {
    if (-not [string]::IsNullOrWhiteSpace($ToolkitRoot)) {
        return (Get-Item -LiteralPath $ToolkitRoot -ErrorAction Stop).FullName
    }

    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    return (Get-Item -LiteralPath (Split-Path -Parent $scriptDir) -ErrorAction Stop).FullName
}

$root = Get-ToolkitRoot
$template = Join-Path $root 'LOT-X'
if (-not (Test-Path -LiteralPath $template -PathType Container)) {
    throw "LOT-X template not found: $template"
}

$wrapperNames = @(
    'Run-Windows11UpgradeRepairWithPsExec-Loop.cmd',
    'Run-Windows11UpgradeRepairWithPsExec-Once.cmd',
    'Run-Windows11UpgradeRepairWithPsExec-Loop-IgnoreRunGuard.cmd',
    'Run-Windows11UpgradeRepairWithPsExec-Once-IgnoreRunGuard.cmd'
)

$configTemplate = Join-Path $root 'Windows11UpgradeToolkit.config.template'

$lots = @(
    Get-ChildItem -LiteralPath $root -Directory -Filter 'LOT-*' -ErrorAction Stop |
        Where-Object { $_.Name -ine 'LOT-X' } |
        Sort-Object Name
)

if ($lots.Count -eq 0) {
    Write-Host "No operational LOT-* folder found."
    exit 0
}

foreach ($lot in $lots) {
    Write-Host ("Updating {0}" -f $lot.Name)
    foreach ($wrapperName in $wrapperNames) {
        $source = Join-Path $template $wrapperName
        $destination = Join-Path $lot.FullName $wrapperName
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
            throw "Template wrapper missing: $source"
        }

        if ($PSCmdlet.ShouldProcess($destination, "Copy $wrapperName")) {
            Copy-Item -LiteralPath $source -Destination $destination -Force -ErrorAction Stop
        }
    }

    $computersPath = Join-Path $lot.FullName 'Computers.txt'
    if (-not (Test-Path -LiteralPath $computersPath -PathType Leaf)) {
        if ($PSCmdlet.ShouldProcess($computersPath, 'Create Computers.txt')) {
            New-Item -ItemType File -Path $computersPath -Force -ErrorAction Stop | Out-Null
        }
    }

    $lotConfigPath = Join-Path $lot.FullName 'Windows11UpgradeToolkit.config'
    if (-not (Test-Path -LiteralPath $lotConfigPath -PathType Leaf)) {
        if ($PSCmdlet.ShouldProcess($lotConfigPath, 'Create LOT config')) {
            if (Test-Path -LiteralPath $configTemplate -PathType Leaf) {
                Copy-Item -LiteralPath $configTemplate -Destination $lotConfigPath -Force -ErrorAction Stop
            }
            else {
                Set-Content -LiteralPath $lotConfigPath -Value @(
                    '# SmartM365 Windows 11 Upgrade Toolkit LOT defaults.'
                    '# W11UT_SETUP_SOURCE should be a UNC path reachable from target computers.'
                    'W11UT_SETUP_SOURCE='
                ) -Encoding ASCII -Force
            }
        }
    }
}

Write-Host ("Updated {0} LOT folder(s)." -f $lots.Count)
