# Name: SmartM365-Upgrade-Staging-Health-Detection.ps1
# Version: 1.0
# Description: Detects stale Windows upgrade staging folders and missing upgrade image files.

$ErrorActionPreference = "Stop"

$ScriptName = "Detect-Upgrade-Staging-Health"
$Version = "1.0"
$RecentSetupActivityHours = 6

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$Issues,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Issues.Contains($Message)) {
        $Issues.Add($Message)
    }
}

function Test-RecentSetupActivity {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Hours
    )

    $setupIndicators = @(
        "C:\Windows\Panther\setupact.log",
        "C:\Windows\Panther\setuperr.log",
        'C:\$WINDOWS.~BT\Sources\Panther\setupact.log',
        'C:\$WINDOWS.~BT\Sources\Panther\setuperr.log'
    )

    foreach ($indicator in $setupIndicators) {
        if (Test-Path -LiteralPath $indicator) {
            $item = Get-Item -LiteralPath $indicator -ErrorAction SilentlyContinue

            if ($item -and ((Get-Date) - $item.LastWriteTime).TotalHours -le $Hours) {
                return $true
            }
        }
    }

    return $false
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $recentSetupActivity = Test-RecentSetupActivity -Hours $RecentSetupActivityHours
    $upgradeSourcesPath = 'C:\$WINDOWS.~BT\Sources'
    $upgradeResiduePaths = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS')

    if ($recentSetupActivity) {
        Write-Output ("{0} v{1}: recent setup activity detected; no remediation requested." -f $ScriptName, $Version)
        exit 0
    }

    if (Test-Path -LiteralPath $upgradeSourcesPath) {
        $upgradeImageFiles = Get-ChildItem -LiteralPath $upgradeSourcesPath -Recurse -Include "*.esd", "*.wim" -File -ErrorAction SilentlyContinue
        $validImageFiles = $upgradeImageFiles | Where-Object { $_.Length -gt 0 }

        if ($null -eq $validImageFiles -or $validImageFiles.Count -eq 0) {
            Add-Issue -Issues $issues -Message "Upgrade staging sources exist but no non-empty ESD or WIM file was found."
        }
    }

    foreach ($path in $upgradeResiduePaths) {
        if (Test-Path -LiteralPath $path) {
            Add-Issue -Issues $issues -Message "Potentially stale upgrade staging folder detected: $path"
        }
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: no stale Windows upgrade staging content detected." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBZUlChGQQW6KLv
# xELUhAUfcVMOShFa2Dn66VZcVNurHqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBJILJOwph1DCV2+rqz
# DgOJX/WzovwFIEFps3fGP6IHljANBgkqhkiG9w0BAQEFAASCAYCSZ1pLgR87N+9D
# e16wW5IFAGrstQxioaHw5WF/IJmdIu1Et3H53y6UwtXVPLQXyjC/EiFhUuMQB/co
# agKF/yFo4ue/rsf67SwpIDqH6CrIKUSdvqx33j3wRT/ZwANtjJgDoLYZ0CurJMDs
# 5ptayBtdn1ejmHCnMuzSVtCgMWysi0B1jshT73vtqywUCAxKKVfUkVPPpRVvWRVG
# FSBf1FJ797mGJWO1lB29PezTLgDeN+fCh5/TfFGeSotZSvCYCG/9OKDCsH5QnZNN
# dDE0gdOn2WxgCpkA4yZ7sMXGUQ86WeT63wqdFuDZWCvdbp/l8GqzNe9VIxcb2tBm
# SYM3CL0qPjvqCUgW1o+PlwjtLycS1dnwRx9dIOTBtToYHPj/H90ZU8CMM/fuAobu
# 1IgbDABVRDaPwOukgUI5m5UOXY+dcLEDKkKsWUZ+ZzmWsZwZ8GecLsqlCeRra+7L
# LRLqayJUEdEoQTBTKzBpzt/6trNiB2ZIS5VJKcRMQd1NFWCuMYs=
# SIG # End signature block
