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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB0i6ggxqMIqZEm
# Fifjv7DchAqJr95zo65dniogWSNy5qCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAcOYoEZEsYebfKoEJP
# IwWRofaYBBsRonLKaHUa5YFTgzANBgkqhkiG9w0BAQEFAASCAYCS6IIPDMjGH3MM
# 4ORK7eSEg79VozLmnOub6Mpp+6XaYXRDjuvqbiqNKJks7Vcazeshwriv2avwdjoU
# mWE1n5OMN3Er3u/+vk5OqieZ1h+GHMagt7T4EI5d5aNS9jT+npJPdVYtHA6ZI6pL
# q+Pjwnn/MgV02qzwSqoUg9zeBPFhYgTcrmuEH6DU5eFTG+5fJ5qAZD5aVSEw+Zaw
# I/BoKwYYN9BW9w45qth1E01qpYAmw1TQlaiNGnTpqM/nxty8zvNaJftOMaDwtAvU
# lvLdXMNUiauXhdxADsSfzQzPTh6xmmU2lHuYRQwgMP2h/lYuL+jOvSKPPeiv4r0f
# 0DAELrSoV1AiHGQdKVXJgGXI9i5n/ppbidwPzunaoMPMqQ1oeoJ5sHUhxbe+fkzg
# /TfzC5FO8jNCNnZM2yF1addOsxqQljrWmxEvV+5rsydMhJIMRzj4iirxM879+AOy
# MnX9YabPhgoQTsh2r+97gTKZv6AOP+oGsuPtLHknV2cPRgy8fOA=
# SIG # End signature block
