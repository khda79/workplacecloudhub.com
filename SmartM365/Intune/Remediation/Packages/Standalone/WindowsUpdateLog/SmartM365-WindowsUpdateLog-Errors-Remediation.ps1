<#
.SYNOPSIS
    Version: 1.0
    Generates a fresh WindowsUpdate.log diagnostic artifact after actionable Windows Update log errors are detected.

.VERSION
    1.0

.DESCRIPTION
    This remediation is intentionally diagnostic-only.
    It does not reset Windows Update or change device state, because WindowsUpdate.log errors can represent many different causes.
    Repair actions are covered by the dedicated Windows Update remediation packages.
    The generated log is stored under C:\ProgramData\SmartM365\IntuneRemediation\Temp\WindowsUpdateLog.
#>

[CmdletBinding()]
param(
    [int]$MinimumLogSizeBytes = 1024,
    [int]$MaxReasonLength = 180
)

$ErrorActionPreference = "Stop"

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

try {
    $localFolder = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Temp\WindowsUpdateLog"
    $computerName = $env:COMPUTERNAME
    $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $outputFile = "WindowsUpdate-Remediation-$date-$computerName.log"
    $localPath = Join-Path -Path $localFolder -ChildPath $outputFile

    if (-not (Test-Path -Path $localFolder)) {
        New-Item -Path $localFolder -ItemType Directory -Force | Out-Null
    }

    try {
        Get-WindowsUpdateLog -LogPath $localPath -ErrorAction Stop | Out-Null
    }
    catch {
        Write-IntuneResult -Status "WindowsUpdateLogUnavailable" -Data @{
            Reason = (Format-CompactText -Text $_.Exception.Message -MaxLength $MaxReasonLength)
        }
        exit 0
    }

    if (-not (Test-Path -Path $localPath -PathType Leaf)) {
        Write-IntuneResult -Status "WindowsUpdateLogNotGenerated"
        exit 0
    }

    $logFile = Get-Item -Path $localPath -ErrorAction Stop

    if ($logFile.Length -lt $MinimumLogSizeBytes) {
        Write-IntuneResult -Status "WindowsUpdateLogTooSmall" -Data @{
            LogSizeBytes = $logFile.Length
        }
        exit 0
    }

    Write-IntuneResult -Status "WindowsUpdateLogGenerated" -Data @{
        LogPath = $localPath
        LogSizeBytes = $logFile.Length
    }
    exit 0
}
catch {
    Write-IntuneResult -Status "Error" -Data @{
        Message = (Format-CompactText -Text $_.Exception.Message -MaxLength $MaxReasonLength)
    }
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB8Otkdmt/hmH4T
# KaBoMi0zkgGYYgyNdjT5CrkTYPT1sKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA2yxtvSRiBQDSyUt6g
# nFO+MwzbuFmYdPUYCHyrlN0IKjANBgkqhkiG9w0BAQEFAASCAYB3P45e0MHKyaM0
# kjmRUAWsutZMuIEmN0CVpIVDDafisB7RWtAbwH71yaSeRZTV+wiiNDJi/lvDEt1e
# bMlJKXFTElWWfobWdCdp4LeKuqyL7CS6JUxZtJmngjLf/P6ZxxYQdRbO0GS1jQnh
# ZOw2CZWfRcPHB6tPReIVvAvGhyyDYeMsNXAQrA3mZ7TrI82p8zriWG77pokycHF+
# LcvqzPIBuxGIiKltIxR4Tn6ixMKVpMe5QGywc1tF1lOoAOetL4az9iDQ1LFF7X1a
# vc+gPYFdFweZbmfO8qbeQyFEkhkGKvFzrhK1KaGPDE/rRO0UP9THdlQBfQ0BzzdB
# z1LhgU3yEypJlGUtmDyOx2e4ldp2j8/c9voC8HnRjG/rQ/g2n32JGE/e2hnEWSfU
# cVy2qza7il0TBfxKejFU6v24gaPcVkvRWKp00IQudKa3M7w9NDJu67msNL2fUF86
# GcqJXU4VhaJ3YFTyB2vQPkk4pMSaEPopfCycw5K8h8BLrgYUEsk=
# SIG # End signature block
