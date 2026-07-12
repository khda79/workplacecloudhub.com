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
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKWiIG4COpTRas
# lSeCUv7EZ4fK4IaUwN4jXiTO6JzaQqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDXQ+G7o6RmUNXuULBUVZ+4x/Q1UAz+J/o4uqStiXI59zANBgkqhkiG9w0B
# AQEFAASCAYC4qTS4TLncLeQwwx1UTYnIV7CZgQ5R+1sC39Fjy5/RoyzdANXq9bP8
# ecSMfSGaTmbHB1cWL08/tB+bizeB/V3Ib/Vfs7TvuI9EmeDGuBW1YeeWfzzvV297
# dg/IW3Vy3D7rBmmOu3EH1U4Rou484rs/07MageK9kfmtGLtk/f7UZCeeLQVbaPlv
# O+NqUB+w6A6X7AtijjGwWjNxfGwLxGSS4Jc76/hLfgJgJ4/8HiJkIjRSEo2y/zVH
# gekCVmcvUeMYRzimbfKNBrHzB0WTTdagCJ4W4IN1tphnqHAmsty4hcOpzY7QNLOV
# +aX7YsGzDoAGFb5Uajt/loILgG1LZEPQ5VPhIehp8I4rPaYD5If4S7laxk5B9eyB
# w6wI87VPdw+StU6K3qhoARfzmMZ9EyP7K02qwRC286IWCugfD8DT8HcVNa46H8fG
# utZR4dfbndZEXkqcXw+Y56OBo05dfC/yJKfYFZzpaockF7pBhL0+kReQmW4vmqQH
# yfH+DbK2Obw=
# SIG # End signature block
