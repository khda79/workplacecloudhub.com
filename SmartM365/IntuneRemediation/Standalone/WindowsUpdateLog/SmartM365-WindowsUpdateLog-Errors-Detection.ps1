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
