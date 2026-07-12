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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDule3atyPpYQwq
# f6Y6ovJ2MEI0VJem/3+gmfLP05u4XqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDTYKgpEZqusYWWpWX/
# dstuMwWYaDaXRvcRnyOmapx7DDANBgkqhkiG9w0BAQEFAASCAYBHv43uvwUVRTuB
# Afupw29lpHv5LAxANXwIafIL9yQJtmJCS+ppdWL/kCiAQZy10F7txiiwGZFR/wXf
# Y5EWt7TU0hc1HPDuz1UuM07pDHVEm/sIj/NWyLBOe0D6DnwFFMHWI/1GEyN2Q35Q
# SD7DKeO1o7l7D3nc6aNbj/xuG/Sxh5EDfAb7lmv/mB4t/AtY1LtJaA1cYTLCpzf7
# wSf2Ggcp50xUNtr20bw5DOprcRzVnCQ9WiFJmNErKceC4inZL+3HHi3iOp34fxeh
# XZoNBAHTIMdYPjnUmJipxB9Mr8XaDvTCXVAMtFhVRtuajG71dEji87Bm0cFeNiII
# sG3b2QpHKLFfS3+pSEGMqK7j9abArV8M1cUYU6mxCuS1h0CTKFn7+4ahAQZl6bgm
# Tv8uZGjopBtBJzX/bOKpEymh4TDM9/0Qdex3RillpghygvQTFe2Bikh7OqaGbyLC
# 1eo9Gi0h+3vPBnwNTM/g8Rq3o3Vswua5mK5IAWMbZaX0KCVXmNg=
# SIG # End signature block
