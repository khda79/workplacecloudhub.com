# Name: SmartM365-MDM-Enrollment-Repair-Detection.ps1
# Version: 1.0
# Description: Detects stale local MDM enrollment, device registration, and EnterpriseMgmt scheduled task signals.

$ErrorActionPreference = "Stop"

$RequireHybridJoin = $false
$ScriptName = "Detect-MDM-Enrollment-Repair"
$Version = "1.0"

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

function Get-DsregValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($Text -match ("(?m)^\s*{0}\s*:\s*(.+?)\s*$" -f [regex]::Escape($Name))) {
        return $matches[1].Trim()
    }

    return $null
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $dsregText = (& "$env:SystemRoot\System32\dsregcmd.exe" /status 2>$null) -join "`n"

    $deviceId = Get-DsregValue -Text $dsregText -Name "DeviceId"
    $azureAdJoined = Get-DsregValue -Text $dsregText -Name "AzureAdJoined"
    $domainJoined = Get-DsregValue -Text $dsregText -Name "DomainJoined"

    if ([string]::IsNullOrWhiteSpace($deviceId)) {
        Add-Issue -Issues $issues -Message "DeviceId missing from dsregcmd status."
    }

    if ($azureAdJoined -ne "YES") {
        Add-Issue -Issues $issues -Message "AzureAdJoined is not YES."
    }

    if ($RequireHybridJoin -and $domainJoined -ne "YES") {
        Add-Issue -Issues $issues -Message "DomainJoined is not YES."
    }

    $enrollmentRoot = "HKLM:\SOFTWARE\Microsoft\Enrollments"

    if (-not (Test-Path -Path $enrollmentRoot)) {
        Add-Issue -Issues $issues -Message "MDM enrollment registry root is missing."
    }
    else {
        $enrollmentKeys = Get-ChildItem -Path $enrollmentRoot -ErrorAction SilentlyContinue

        if ($null -eq $enrollmentKeys -or $enrollmentKeys.Count -eq 0) {
            Add-Issue -Issues $issues -Message "No MDM enrollment registry records found."
        }
    }

    $enterpriseTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { $_.TaskPath -like "\Microsoft\Windows\EnterpriseMgmt\*" }

    if ($null -eq $enterpriseTasks -or $enterpriseTasks.Count -eq 0) {
        Add-Issue -Issues $issues -Message "EnterpriseMgmt scheduled tasks are missing."
    }
    else {
        $enabledTasks = $enterpriseTasks | Where-Object { $_.State -ne "Disabled" }

        if ($null -eq $enabledTasks -or $enabledTasks.Count -eq 0) {
            Add-Issue -Issues $issues -Message "EnterpriseMgmt scheduled tasks are present but disabled."
        }
    }

    $imeService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue

    if ($null -eq $imeService) {
        Add-Issue -Issues $issues -Message "Intune Management Extension service is missing."
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: MDM enrollment and device registration signals are healthy." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCByhc+7nHahA6uE
# a+1KOhsSe2XZvxv0JjPqmDsGTYnPeKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCABI3SwjMLOi32Pv+jc
# an/j+pJdnKtHc3mlydM3EVw0TDANBgkqhkiG9w0BAQEFAASCAYClA9ESud/pF6Hh
# SpOaffZhTXa0/P6Oc8B+/ySb2vcXHw/zobUagjlEBPUqqpKmZ6qB7VxnMURoSMVY
# J6w+wfH7h34pBeKCa/v0kHbc2ghk7FhIpNZzcRpn3MAVgIMwBQq8SZvVSRqiuXn/
# jDzBFZmdZOtrq/a6XUU9sit/u36pb2kgIbBz6d+h8i4zRR3v/Qe5WbFr7NUdeqO1
# K55NDxTs8sTQW0RXkbA+8IcXinwtCI7TSEZkE7HIIHbvLE+QLsnvb38p25XeB7lV
# BKnJmp6k3jMsCAzwJ1kSKx2QzxgHKHwoB1Y+cxUduOzrPwaKPfUWPSs0iPX3emCu
# /2qBLGcaLiEPugCgUmWSPUOOGc6Sdb3o424Vd0rU4gg2MoeEpIosNuHY3R6fve1j
# +a0yRkQieKGh2fxSfBWMzryNrCLGgqt8DysAYw+jhunbavf8L1+/wNJVhV9vNuvj
# AGh2gC8YsG7rVSepZHmqc4emXdOEn4F2BAg8UfUZkLveayn+MAg=
# SIG # End signature block
