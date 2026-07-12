# Name: SmartM365-WUfB-Identity-Drift-Detection.ps1
# Version: 1.1
# Description: Detects a mismatch between MDM Windows Update for Business policy presence and Windows Update PolicyState identity

$ErrorActionPreference = "Stop"

try {
    $policyStatePath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState"
    $mdmUpdatePolicyPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"

    function Get-RegistryPropertyValue {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$Name
        )

        if (-not (Test-Path -Path $Path)) {
            return $null
        }

        $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue

        if ($null -eq $item) {
            return $null
        }

        if ($item.PSObject.Properties.Name -contains $Name) {
            return $item.$Name
        }

        return $null
    }

    function Format-CompactText {
        param(
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$Text,

            [int]$MaxLength = 180
        )

        $compactText = ($Text -replace "\s+", " ").Trim()

        if ($compactText.Length -gt $MaxLength) {
            return ($compactText.Substring(0, $MaxLength) + "...")
        }

        return $compactText
    }

    function Write-IntuneResult {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Status,

            [hashtable]$Data = @{}
        )

        $parts = New-Object System.Collections.Generic.List[string]
        $parts.Add("Status=$Status")

        foreach ($key in ($Data.Keys | Sort-Object)) {
            $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 240
            $parts.Add(("{0}={1}" -f $key, $value))
        }

        Write-Output ($parts -join "; ")
    }

    $mdmWUfBPresent = Test-Path -Path $mdmUpdatePolicyPath
    $isWUfBConfigured = Get-RegistryPropertyValue -Path $policyStatePath -Name "IsWUfBConfigured"
    $policyStateValue = if ($null -eq $isWUfBConfigured) { "Unknown" } else { [string]$isWUfBConfigured }

    $commonData = @{
        MDMWUfBPolicyPresent = $mdmWUfBPresent
        PolicyStateIsWUfBConfigured = $policyStateValue
    }

    # Drift condition:
    # MDM WUfB policy exists, but Windows Update PolicyState says WUfB is not configured.
    if ($mdmWUfBPresent -and $isWUfBConfigured -eq 0) {
        $commonData["Reason"] = "MDM WUfB policy is present but PolicyState IsWUfBConfigured is 0"
        Write-IntuneResult -Status "WUfBIdentityDriftDetected" -Data $commonData
        exit 1
    }

    # Do not fail devices that do not have MDM WUfB policy.
    if (-not $mdmWUfBPresent) {
        $commonData["Reason"] = "No MDM WUfB policy detected"
        Write-IntuneResult -Status "NotApplicable" -Data $commonData
        exit 0
    }

    if ($null -eq $isWUfBConfigured) {
        $commonData["Reason"] = "MDM WUfB policy is present but PolicyState IsWUfBConfigured is unavailable"
        Write-IntuneResult -Status "HealthyWithUnknownPolicyState" -Data $commonData
        exit 0
    }

    $commonData["Reason"] = "MDM WUfB policy and PolicyState are aligned"
    Write-IntuneResult -Status "Healthy" -Data $commonData
    exit 0
}
catch {
    Write-Output ("Status=Error; Message={0}" -f (($_.Exception.Message -replace "\s+", " ").Trim()))
    exit 2
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB0i6ggxqMIqZEm
# Fifjv7DchAqJr95zo65dniogWSNy5qCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAcOYoEZEsYebfKoEJPIwWRofaYBBsRonLKaHUa5YFTgzANBgkqhkiG9w0B
# AQEFAASCAYCsLKc3eoZ70m0xlVk50sC08ubmZVPik2RmoFsTRD629oD4MFHy97aC
# uo22Y8cetE0btRq+7C/o+GHjndy+2eJBX2J6iOjNpsmh/OhLsMY8SENKbx2oq47X
# Jyvmtw7nlZwVSXsajsAelCikhmU2MOtggPxoIPjPHzwQNxASJq2uB41RbCQJTahX
# 8J+SsTSOIiWAujPfdxCG8Xq8Kh5IffQekBU9FF3IY5cjrACI51zioGHZGqHJ+bj9
# qNA4FdEejn7FjGn84rA/bcSJckfhlKpCfnVkEku4gfOgKOCneklfFG7xpAtYTq4K
# T650OGNRs6uIYJbWcZiUjl5KkL8C3z4SZ4Y7SEcE3Oyb1yyV7ch1lpcgAocQadbj
# gc1xRV8r0g4y746yDqKVPw25XlUnDKLISkD5wc0YkX/0pUcsKScvLqv7fyvkWu7M
# 2wQEse+3gT+AAgpT/+QcTYqeAK/Ea4h1piC+XndZ4iGTD5WhOzfa8zr895RAhqNE
# fQr4NxVFqDo=
# SIG # End signature block
