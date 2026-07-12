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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAseKUVFBA9I4vP
# Tc3P9O5Jg8USFf7OPm3NMkqFnqbp7aCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC8AZYmJOizP6txXaOZ
# +8+jFlHgomFGwaGpLLV4BBqsHTANBgkqhkiG9w0BAQEFAASCAYA4jTwgDEL1UEmT
# d8N9XZQlcpU9cg4zl8Kz30SC3xKqsZ3WKlokWxiv3Bx5yO5RDPtI9l8sJ+i7SNVn
# fmG1PMKtIwOua0+XgkgSp+plzhKzaiDb8AvZvzkAYMomrJTOGGmZmvMtoUVXVFPL
# CVfqWugfP158qXUEIgTaneCTZDjOSng4/BQcLVRzKAL5xP23oNJculzCmwzXvgJn
# d1Vfq1ro9bYNGR3jXzlfPhyGx+jjujEZsYvnYBjia5Yca+idp8SvUAK7CG2u8eMz
# sfU7V+Db8NeAwmEdS5bMGspy/FnTBFKYEkqDcBOhxFS7krqemf7aqTpRdLn8jINV
# 0kiMCeoZK32rKSh8hPdL23cHlEc5yJ/rmf9owZDuSEkGKM0frMogHtH/E0UkHIW+
# KbmwVCaBO2YZtz0CdxHGrhyuzrn7lj4tZF7hyeik3j/TCsF0HzGzCAISPrCC1a0g
# GhviHwbFT067A0zs2TMdof9u5QLwUk/5zrRQvDu4vgQSdTdpyo8=
# SIG # End signature block
