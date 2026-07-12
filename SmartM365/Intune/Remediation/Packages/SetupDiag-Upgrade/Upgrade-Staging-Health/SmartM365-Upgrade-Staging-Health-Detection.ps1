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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBZUlChGQQW6KLv
# xELUhAUfcVMOShFa2Dn66VZcVNurHqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBJILJOwph1DCV2+rqzDgOJX/WzovwFIEFps3fGP6IHljANBgkqhkiG9w0B
# AQEFAASCAYBmftnFs/sA5PhvmDDdFa5CMG871J2gSFo9IgGRXqJ4ugstfPRKc6P9
# RUQE6tN/uR6iF3VFi3vsygVriLI7/1tYx0lIHgI7WaJtAkwj2eiyIbBdERCVox8b
# ouivHe61ixeex5Ffp24tQLno+2C6pGBX5akXfcXbeRWuq4QW7802BU3bdZyyJhjf
# i43AOR2La9jl4obH+JiB6cVh5CbqTuAWoLLoSN0RzRs20vMoT0U/mFaki8GWPhLH
# X92vjXTJMDC53V48cTHEV4+YO3UTCrqLs7Jn5tqZtgCHFghC5ltmcczWgTKbBxjX
# th8JXmnqKx3wT4NDbZlch3x/qmQ1KtWdT5VYN5Kxr6QUytAfnX5/OXe44WLg6xsX
# 6U4Yq81NnkpjUMlJEwZSIYA6wzfnfxLgCdvuRVNFn73iHwyWcu2d9wonme0ABt3s
# 3l/jFrYbyPgc3xitwwFVsE1f/rU/WdD7NAKurqPBX5oHifWY58yKTyt9DG8HCw4/
# 7TzfpjVwsjg=
# SIG # End signature block
