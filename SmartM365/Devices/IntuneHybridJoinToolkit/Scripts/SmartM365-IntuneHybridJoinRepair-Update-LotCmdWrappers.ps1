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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKWiIG4COpTRas
# lSeCUv7EZ4fK4IaUwN4jXiTO6JzaQqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDXQ+G7o6RmUNXuULBU
# VZ+4x/Q1UAz+J/o4uqStiXI59zANBgkqhkiG9w0BAQEFAASCAYAbWE1iEwMWKX6t
# K8oDInXwCH9QYEIfAgY2UnU2HRoIt7hc04gR8uo9TCpoJo17oWmnbJapJrw//cea
# tuMuDrDv9vE2JEXUzEsv5X+SwxTueI+T1mI6gfdDVkQHDwAQA0853gMg8x79wcVC
# 4sbUAjlWbmJyV3iIWvg9gwLK59hplVS4kFqseZWprV155+SZux2TWR5dO2PVX0J9
# VKpp6SV5AYQ30PgJcVfYgCTD6mJjWSynN7vuEq2n5EGvgF94SCurTFJxh76uObda
# lOaqThQTPhD93EcKbHTsUPfNSGj05vJ0gP2TdGW7qMY5fkGiyYfB+msB0R7vutQ7
# 9qSYW50lpER40OL24nenKpe/0pNSnGVjePBZ63D17FqGByT/joNcq6CWej/xTgwg
# IuBJ6Rz5DMa5cYYJ7fqAAzIL3zqYx7rcezRYTFE8KzPjYQsTRFmVb+vDgKF1SEwj
# meiXOWf+80XvPj/6M/pF21wiNrkBs/S33YIzs9oJb+RxhjRUOsI=
# SIG # End signature block
