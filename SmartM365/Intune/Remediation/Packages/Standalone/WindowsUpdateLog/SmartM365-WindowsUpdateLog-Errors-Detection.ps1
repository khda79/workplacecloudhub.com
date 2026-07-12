<#
.SYNOPSIS
    Version: 1.1
    Generates WindowsUpdate.log and detects recent Windows Update errors.

.VERSION
    1.1

.DESCRIPTION
    Generates WindowsUpdate.log in C:\ProgramData\SmartM365\IntuneRemediation\Temp.
    Extracts actionable Windows Update related error lines.
    Ignores benign trace lines such as "error 0", flags, product type values, and successful progress callbacks.
    Emits a compact single-line output for Intune readability.
#>

[CmdletBinding()]
param(
    [int]$MaxErrorsToDisplay = 3,
    [int]$MinimumLogSizeBytes = 1024,
    [int]$MaxLineLength = 140
)

$ErrorActionPreference = "Stop"

function Format-CompactText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [int]$MaxLength = 140
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

function Test-BenignWindowsUpdateLogLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Line
    )

    $benignPatterns = @(
        "\bNo error\b",
        "\berror\s*[:=]?\s*0\b",
        "\berror\s*[:=]?\s*0x0+\b",
        "\bSucceeded\b",
        "\bsuccessfully\b",
        "\bCompleted\b",
        "\bRefresh complete\b",
        "\* START \*",
        "\bFlags:\s*0X[0-9A-Fa-f]+\b",
        "\bOS Product Type\s*=\s*0x[0-9A-Fa-f]+\b",
        "\bauth token of type\s*0x[0-9A-Fa-f]+\b",
        "\bcode Call (progress|complete) and error 0\b"
    )

    foreach ($pattern in $benignPatterns) {
        if ($Line -match $pattern) {
            return $true
        }
    }

    return $false
}

function Test-ActionableWindowsUpdateLogLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Line
    )

    if (Test-BenignWindowsUpdateLogLine -Line $Line) {
        return $false
    }

    if ($Line -match "\*FAILED\*\s*\[[0-9A-Fa-f]{8}\]") {
        return $true
    }

    if ($Line -match "\b(fatal|failed|failure)\b") {
        return $true
    }

    if ($Line -match "\berror\b" -and $Line -match "\b(0x[0-9A-Fa-f]{8}|[78][0-9A-Fa-f]{7}|C[0-9A-Fa-f]{7}|80D[0-9A-Fa-f]{5})\b") {
        return $true
    }

    return $false
}

try {
    $localFolder = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Temp\WindowsUpdateLog"
    $computerName = $env:COMPUTERNAME
    $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $outputFile = "WindowsUpdate-$date-$computerName.log"
    $localPath = Join-Path -Path $localFolder -ChildPath $outputFile

    if (-not (Test-Path -Path $localFolder)) {
        New-Item -Path $localFolder -ItemType Directory -Force | Out-Null
    }

    try {
        Get-WindowsUpdateLog -LogPath $localPath -ErrorAction Stop | Out-Null
    }
    catch {
        Write-IntuneResult -Status "WindowsUpdateLogUnavailable" -Data @{
            Reason = $_.Exception.Message
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

    $candidatePatterns = @("\berror\b", "\bfailed\b", "\bfailure\b", "\bfatal\b", "\*FAILED\*")
    $rawMatches = @(Select-String -Path $localPath -Pattern $candidatePatterns -CaseSensitive:$false -ErrorAction Stop)
    $filteredErrors = New-Object System.Collections.Generic.List[string]

    foreach ($match in $rawMatches) {
        if (Test-ActionableWindowsUpdateLogLine -Line $match.Line) {
            $filteredErrors.Add((Format-CompactText -Text $match.Line -MaxLength $MaxLineLength))
        }
    }

    if ($filteredErrors.Count -gt 0) {
        $sampleErrors = @($filteredErrors | Select-Object -First $MaxErrorsToDisplay)

        Write-IntuneResult -Status "WindowsUpdateErrorsDetected" -Data @{
            ErrorCount = $filteredErrors.Count
            LogSizeBytes = $logFile.Length
            Samples = ($sampleErrors -join " | ")
        }

        exit 1
    }

    Write-IntuneResult -Status "NoActionableWindowsUpdateErrorsFound" -Data @{
        CandidateLineCount = $rawMatches.Count
        LogSizeBytes = $logFile.Length
    }
    exit 0
}
catch {
    Write-IntuneResult -Status "Error" -Data @{
        Message = $_.Exception.Message
    }
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD/wZ31uMwgPmOC
# qBqOgy65PHAp5GLkez00ECdlE29WqqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDKjTK1glZ+4WGWrACi
# B0esA21kUF1/2s0xKmbZbpEZOjANBgkqhkiG9w0BAQEFAASCAYBvtLprpAG8xK0J
# OYOC1VkCrdHR+zGzu2n3Y8IxxO/Z+LVuuQ+ManrWnxYGvbfh30lxMeTQEdj5Meqg
# S9NXNvtqhaQzdtLOB2A4b8hMcaT3O2iH/p6LtqIPEwAuLCIWYwtlOZyobQtZw0Ga
# nuxS7X4x6Fz20LJJyLA7BN6/QszE1STPOm6Fx7azOHcqGR+7Kjbmpr0+8IshKsqJ
# i0BIaoApU80xSlC6V8QhsPQ+wAElyeVz7sGInZfKp1uLGMMWXcL5WGloJVZN9cjt
# oNdDoj/NckCtQg1lWzjelYg0WfJWpoWIQxqVbF3k74OJ+lbec2EEK8iP9+j5HWO7
# yk0tBblibOQrki7gAsmpRImwagzxbhHPbVcJQ3QlJv33yjf+BvWUDMas16QLI6jR
# eJIVssDuaiv7l+2YcrXyIey/ZCggFToyFA+mPDVBMXoaBSDJIAC+dhcYadEeyRUZ
# 8CTFL6nFAhaDrlwzHp7VC0/3g589sXYMy981PwicMZAz6VYavGQ=
# SIG # End signature block
