<#
.SYNOPSIS
Refreshes SmartM365 Intune Hybrid Join Toolkit LOT CMD wrappers.

.DESCRIPTION
Copies the standard tiny CMD wrappers from Lots\LOT-TEMPLATE to operational Lots\LOT-* folders.
Lots\LOT-TEMPLATE is the versioned template and is not modified by this script.

The Intune inventory export is global and must be launched from the repository root,
so this script removes obsolete Export-IntuneDevicesCsv.cmd wrappers from LOT folders.
It also creates Computers.txt and AdDomain.txt when missing so each operational LOT
keeps only its configuration and computer list beside the template wrappers.

.VERSION
1.2
#>

#requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$RootPath
)

$ErrorActionPreference = 'Stop'

function Get-ToolkitRoot {
    if (-not [string]::IsNullOrWhiteSpace($RootPath)) {
        return (Get-Item -LiteralPath ($RootPath.Trim().Trim([char]34)) -ErrorAction Stop).FullName
    }

    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    return (Get-Item -LiteralPath (Split-Path -Parent $scriptDir) -ErrorAction Stop).FullName
}

$root = Get-ToolkitRoot
$lotsRoot = Join-Path $root 'Lots'
$template = Join-Path $lotsRoot 'LOT-TEMPLATE'
if (-not (Test-Path -LiteralPath $lotsRoot -PathType Container)) {
    throw "Lots folder not found: $lotsRoot"
}
if (-not (Test-Path -LiteralPath $template -PathType Container)) {
    throw "Lots\LOT-TEMPLATE template not found: $template"
}

$wrapperNames = @(
    'Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd',
    'Run-IntuneHybridJoinRepairWithPsExec-Once.cmd',
    'Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd',
    'Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd'
)

$obsoleteWrappers = @(
    'Export-IntuneDevicesCsv.cmd',
    'Run-IntuneHybridJoinRepairWithPsExec.cmd',
    'Run-IntuneHybridJoinRepairWithPsExec-IgnoreRunGuard.cmd'
)

$lotFolders = @(
    Get-ChildItem -LiteralPath $lotsRoot -Directory -Filter 'LOT-*' -ErrorAction Stop |
        Where-Object { $_.Name -ine 'LOT-TEMPLATE' } |
        Sort-Object Name
)

if ($lotFolders.Count -eq 0) {
    Write-Host ("No operational LOT-* folders found under: {0}" -f $lotsRoot) -ForegroundColor Yellow
    return
}

foreach ($lot in $lotFolders) {
    foreach ($wrapperName in $wrapperNames) {
        $source = Join-Path $template $wrapperName
        $destination = Join-Path $lot.FullName $wrapperName
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
            throw "Template wrapper missing: $source"
        }

        if ($PSCmdlet.ShouldProcess($destination, "Copy $wrapperName")) {
            Copy-Item -LiteralPath $source -Destination $destination -Force -ErrorAction Stop
            Write-Host ("Updated: {0}" -f $destination) -ForegroundColor Green
        }
    }

    foreach ($configName in @('Computers.txt', 'AdDomain.txt')) {
        $destination = Join-Path $lot.FullName $configName
        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            continue
        }

        $source = Join-Path $template $configName
        if ($PSCmdlet.ShouldProcess($destination, "Create $configName")) {
            if (Test-Path -LiteralPath $source -PathType Leaf) {
                Copy-Item -LiteralPath $source -Destination $destination -Force -ErrorAction Stop
            }
            else {
                New-Item -ItemType File -Path $destination -Force -ErrorAction Stop | Out-Null
            }
            Write-Host ("Created: {0}" -f $destination) -ForegroundColor Green
        }
    }

    foreach ($obsoleteWrapper in $obsoleteWrappers) {
        $obsoletePath = Join-Path $lot.FullName $obsoleteWrapper
        if (-not (Test-Path -LiteralPath $obsoletePath -PathType Leaf)) {
            continue
        }

        if ($PSCmdlet.ShouldProcess($obsoletePath, 'Remove obsolete wrapper')) {
            Remove-Item -LiteralPath $obsoletePath -Force -ErrorAction Stop
            Write-Host ("Removed obsolete wrapper: {0}" -f $obsoletePath) -ForegroundColor Yellow
        }
    }
}