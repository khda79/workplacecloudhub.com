# Name: SmartM365-WindowsUpdate-Reset-Detection.ps1
# Version: 1.0
# Description: Detects Windows Update reset conditions such as legacy WSUS policies or recent Autopatch error 0x80244007.

$ErrorActionPreference = "Stop"

$ScriptName = "Detect-WindowsUpdate-Reset"
$Version = "1.0"
$ForceRemediation = $false

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

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $wuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $wuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

    if (Test-Path -Path $wuPolicyPath) {
        $wuPolicy = Get-ItemProperty -Path $wuPolicyPath -ErrorAction Stop

        foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate")) {
            if ($wuPolicy.PSObject.Properties.Name -contains $name) {
                $value = [string]$wuPolicy.$name

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    Add-Issue -Issues $issues -Message "$name is configured."
                }
            }
        }
    }

    if (Test-Path -Path $wuAuPolicyPath) {
        $wuAuPolicy = Get-ItemProperty -Path $wuAuPolicyPath -ErrorAction Stop

        if ($wuAuPolicy.PSObject.Properties.Name -contains "UseWUServer" -and [int]$wuAuPolicy.UseWUServer -eq 1) {
            Add-Issue -Issues $issues -Message "UseWUServer=1."
        }
    }

    $reportingEventsPath = Join-Path -Path $env:windir -ChildPath "SoftwareDistribution\ReportingEvents.log"

    if (Test-Path -LiteralPath $reportingEventsPath) {
        $recentError = Select-String -Path $reportingEventsPath -Pattern "0x80244007|Windows Update Client failed to detect" -ErrorAction SilentlyContinue |
            Select-Object -Last 1

        if ($recentError) {
            Add-Issue -Issues $issues -Message "Recent Windows Update detection failure 0x80244007 found."
        }
    }

    if ($ForceRemediation) {
        if ($issues.Count -gt 0) {
            Write-Output ("{0} v{1}: forced remediation required. Indicators: {2}" -f $ScriptName, $Version, ($issues -join " | "))
        }
        else {
            Write-Output ("{0} v{1}: forced remediation required. No blocking indicator found." -f $ScriptName, $Version)
        }

        exit 1
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: no remediation required." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgeW5MBs6rLTs5
# kuutYKd5mK8LK4ET86smDU8zjO6lOKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCx5Cfo78KBLRLKgzR9ABpiXmLR/RF0RxlIUKExEUYNlTANBgkqhkiG9w0B
# AQEFAASCAYDA7XLyH/H6qmKMwgiJD1eYFBmxEB4jUH/Yw1s3rCJdNuaR8+mzTlar
# 6r7rfpsstpBr3VrhOPqmIGsK/Z46ulMxXB98EHkrtqYwBVSeTGF6IxsJVwcMZmYY
# g2+prhb7meu7pskDtVjHNQJq84AuFfTs3e/LyOHprHAetTjFuwUPOmCc8nOu6uSk
# Rks49C6SUoEuaxC81tDEZYViPQpjAfr0Ctt+7z264YZNco0rCCaz4XZX5QCJ+t+X
# pjzV5B9kgqNDMWIahL5pu2BkxATgC21Jnap8VTg95cJj4TSKqUKLboQYU3zCkMml
# umNeqGJ3XFutnr9LWXsyEcXGHVcBN1lXWVGCYGuyDmdGacG/aHjCzLtP2Vvy/7r7
# HAsuaIIlP+/I+nGPWrmoDLBrHIsYlLvEuKNO+CNNhTKxijBkWqYzioFg2z3debyQ
# rBBWNp7NkOSWzMEhqeIWUJxOpRqTqtQ8te/4OrXJLJBo8Gob+hRdNwnuQIQQXqTh
# zjl+6dZOD9w=
# SIG # End signature block
