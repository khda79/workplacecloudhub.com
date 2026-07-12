<#
.SYNOPSIS
    Refreshes SmartM365 Windows 11 Upgrade Toolkit LOT CMD wrappers.

.DESCRIPTION
    Copies the standard tiny CMD wrappers from Lots\LOT-TEMPLATE to operational Lots\LOT-* folders.
    Lots\LOT-TEMPLATE is the versioned template and is not modified by this script.

.VERSION
    0.1.2

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
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
$lotsRoot = Join-Path $root 'Lots'
$template = Join-Path $lotsRoot 'LOT-TEMPLATE'
if (-not (Test-Path -LiteralPath $template -PathType Container)) {
    throw "Lots\LOT-TEMPLATE template not found: $template"
}

$wrapperNames = @(
    'Run-Windows11UpgradeRepairWithPsExec-Loop.cmd',
    'Run-Windows11UpgradeRepairWithPsExec-Once.cmd',
    'Run-Windows11UpgradeRepairWithPsExec-Loop-IgnoreRunGuard.cmd',
    'Run-Windows11UpgradeRepairWithPsExec-Once-IgnoreRunGuard.cmd'
)

$configTemplate = Join-Path $root 'Windows11UpgradeToolkit.config.template'

$lots = @(
    Get-ChildItem -LiteralPath $lotsRoot -Directory -Filter 'LOT-*' -ErrorAction Stop |
        Where-Object { $_.Name -ine 'LOT-TEMPLATE' } |
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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAseKUVFBA9I4vP
# Tc3P9O5Jg8USFf7OPm3NMkqFnqbp7aCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCC8AZYmJOizP6txXaOZ+8+jFlHgomFGwaGpLLV4BBqsHTANBgkqhkiG9w0B
# AQEFAASCAYAWhtvbrJpquZOCbx0EQvH/snOP2qdOijMCjVaeaOKLMMzYXM53Piv3
# NOTMT0Mt2Xo8z4Abn6JDG66INtD4aRxE5pgGxZ0ouPc1MSfO4HGv7wJM66fbHBFr
# Lh2PaTd9ijykAVBEOEfp+XFn1+/lNm/Bhk5OVJAj8Yeqsh1BKMXnhb7qgJTY53af
# 1XdkjLXcBaBwxRg2zA7c2GmQJJKvSI/xQgwDg6DKWVjtFFoWWRfxaMH4DUTevkpH
# IL3bBzxiLBWVoUDav290Oc71tjVU1kyDNYHTOZ5alxJJpV2ov9svTEHYVvuZf0qh
# VQ8Z9wAPdbkjmrC6XKu5bswYOhNX9hlTqIf6tfguwg1gafjqRgrfE2kIBfmLzRZ0
# G41O6xRMx5xyPhoWntx06zMtJuoPQ6YHuM0dEucRPgWtduRHs8Qaci+fnoYyvoT8
# ZUxqs7QrKjJamrF/WW0CMFacRSnCHLGKoZ5OnC3xtmFJfMKZQTxn3JUnb6RRldHp
# rj+UnDUMpmQ=
# SIG # End signature block
