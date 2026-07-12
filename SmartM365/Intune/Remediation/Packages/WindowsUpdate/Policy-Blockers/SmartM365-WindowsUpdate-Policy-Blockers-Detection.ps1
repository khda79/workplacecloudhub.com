# Name: SmartM365-WindowsUpdate-Policy-Blockers-Detection.ps1
# Version: 1.0
# Description: Detects WSUS, Windows Update, and WUfB policy values that can block cloud-managed update flows.

$ErrorActionPreference = "Stop"

$RequireWUfBPolicyManager = $false
$ScriptName = "Detect-WindowsUpdate-Policy-Blockers"
$Version = "1.0"

function Add-Finding {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$Findings,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Findings.Contains($Message)) {
        $Findings.Add($Message)
    }
}

try {
    $policyManagerUpdatePath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"
    $wuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $wuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    $blockingFindings = New-Object System.Collections.Generic.List[string]
    $contextFindings = New-Object System.Collections.Generic.List[string]

    if ($RequireWUfBPolicyManager -and -not (Test-Path -Path $policyManagerUpdatePath)) {
        Add-Finding -Findings $blockingFindings -Message "WUfB PolicyManager configuration is missing."
    }

    if (Test-Path -Path $wuPolicyPath) {
        $wuPolicy = Get-ItemProperty -Path $wuPolicyPath -ErrorAction Stop

        foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate")) {
            if ($wuPolicy.PSObject.Properties.Name -contains $name) {
                $value = [string]$wuPolicy.$name

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    Add-Finding -Findings $blockingFindings -Message "$name=$value"
                }
            }
        }

        foreach ($name in @("DoNotConnectToWindowsUpdateInternetLocations", "DisableWindowsUpdateAccess", "SetDisableUXWUAccess")) {
            if ($wuPolicy.PSObject.Properties.Name -contains $name) {
                $value = $wuPolicy.$name
                Add-Finding -Findings $contextFindings -Message "$name=$value"

                if ($value -eq 1) {
                    Add-Finding -Findings $blockingFindings -Message "$name=1"
                }
            }
        }
    }

    if (Test-Path -Path $wuAuPolicyPath) {
        $wuAuPolicy = Get-ItemProperty -Path $wuAuPolicyPath -ErrorAction Stop

        foreach ($name in @("UseWUServer", "NoAutoUpdate")) {
            if ($wuAuPolicy.PSObject.Properties.Name -contains $name) {
                $value = $wuAuPolicy.$name
                Add-Finding -Findings $contextFindings -Message "$name=$value"

                if ($value -eq 1) {
                    Add-Finding -Findings $blockingFindings -Message "$name=1"
                }
            }
        }

        if ($wuAuPolicy.PSObject.Properties.Name -contains "AUOptions") {
            Add-Finding -Findings $contextFindings -Message "AUOptions=$($wuAuPolicy.AUOptions)"
        }
    }

    if ($blockingFindings.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. Blocking policies: {2}" -f $ScriptName, $Version, ($blockingFindings -join " | "))
        exit 1
    }

    if ($contextFindings.Count -gt 0) {
        Write-Output ("{0} v{1}: no blocking policy detected. Context: {2}" -f $ScriptName, $Version, ($contextFindings -join " | "))
        exit 0
    }

    Write-Output ("{0} v{1}: no WSUS or Windows Update policy blocker detected." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDule3atyPpYQwq
# f6Y6ovJ2MEI0VJem/3+gmfLP05u4XqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDTYKgpEZqusYWWpWX/dstuMwWYaDaXRvcRnyOmapx7DDANBgkqhkiG9w0B
# AQEFAASCAYBnqXOGK4p8F/WYBOqRQPuOBXUZjjEqxD85MwP9WxI3rzHjNT/yBYhz
# YmTy5tngEwBypzw/bxt1sm/5gxvI5c5i87wYaD61nNvXXWzv5R+MXO80LUylXRAy
# s+CwWncG5M5WfUbVbwhjh/I1z+xm2mIAjlxgx5OQpvRPZkTqxxpYaXcAJFezyVyo
# y1Y6RCK3CjnCFHktJwUV140EZLxJWzcjdHutGaTsQ+sJ37BIQ8Ie9LHMZmtUQeSw
# 41L9PYeZQkCOWD3mILcwDk535Uy0VD4it7hatGWq2fXP1VssT21wy3hcpvcoVMd0
# qdyyzMls//SW065+7nrIFEEiORMzv5mVDjdjgLorRMuOsN0oTw2d4fEcH1PoF2H4
# irEG7cWLdVWh4D81nMexF3lE750/JI6ZtsEVnUJ1yYafegaHCl9tWSxpO+snD3Uu
# 4T8s3HF6ngWmd6H8CR9gZCHvijp1W8AHpatf8zn5kx3aiGQJJ/qEVclP7JQJjQRf
# Up/nXFe93bc=
# SIG # End signature block
