<#
.SYNOPSIS
    Version: 1.0
    Generates WindowsUpdate.log and detects recent Windows Update errors.

.VERSION
    3.0

.DESCRIPTION
    Generates WindowsUpdate.log in C:\ProgramData\SmartM365\IntuneRemediation\Temp.
    Extracts Windows Update related error lines.
    Limits output for Intune readability.
    Uses a configurable maximum error count.
#>

[CmdletBinding()]
param(
    [int]$MaxErrorsToDisplay = 20,
    [int]$MinimumLogSizeBytes = 1024
)

$ErrorActionPreference = "Stop"

try {
    # ================================
    # Variables
    # ================================
    $localFolder = Join-Path -Path $env:WINDIR -ChildPath "Temp"
    $computerName = $env:COMPUTERNAME
    $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $outputFile = "WindowsUpdate-$date-$computerName.log"
    $localPath = Join-Path -Path $localFolder -ChildPath $outputFile

    # ================================
    # Ensure local folder exists
    # ================================
    if (-not (Test-Path -Path $localFolder)) {
        New-Item -Path $localFolder -ItemType Directory -Force | Out-Null
    }

    # ================================
    # Generate WindowsUpdate.log
    # ================================
    Write-Output "Status=GeneratingWindowsUpdateLog"
    Get-WindowsUpdateLog -LogPath $localPath -ErrorAction Stop | Out-Null

    # ================================
    # Validate generated log
    # ================================
    if (-not (Test-Path -Path $localPath -PathType Leaf)) {
        Write-Output "Status=WindowsUpdateLogNotGenerated"
        exit 1
    }

    $logFile = Get-Item -Path $localPath -ErrorAction Stop

    if ($logFile.Length -lt $MinimumLogSizeBytes) {
        Write-Output "Status=WindowsUpdateLogTooSmall"
        Write-Output "LogPath=$localPath"
        Write-Output "LogSizeBytes=$($logFile.Length)"
        exit 1
    }

    Write-Output "Status=WindowsUpdateLogGenerated"
    Write-Output "LogPath=$localPath"
    Write-Output "LogSizeBytes=$($logFile.Length)"

    # ================================
    # Extract relevant errors
    # ================================
    $errorPatterns = @(
        "\berror\b",
        "\bfailed\b",
        "\bfailure\b",
        "\bfatal\b",
        "0x[0-9A-Fa-f]{8}"
    )

    $ignorePatterns = @(
        "No error",
        "error\s*=\s*0x00000000",
        "0x00000000",
        "Succeeded",
        "successfully"
    )

    $rawMatches = Select-String -Path $localPath -Pattern $errorPatterns -CaseSensitive:$false -ErrorAction Stop

    $filteredErrors = $rawMatches | Where-Object {
        $line = $_.Line

        $ignore = $false
        foreach ($ignorePattern in $ignorePatterns) {
            if ($line -match $ignorePattern) {
                $ignore = $true
                break
            }
        }

        -not $ignore
    } | Select-Object -First $MaxErrorsToDisplay

    if ($null -ne $filteredErrors -and $filteredErrors.Count -gt 0) {
        Write-Output "Status=WindowsUpdateErrorsDetected"
        Write-Output "DisplayedErrorCount=$($filteredErrors.Count)"
        Write-Output "----------------------------------------"

        foreach ($errorLine in $filteredErrors) {
            Write-Output ($errorLine.Line.Trim())
        }

        exit 1
    }

    Write-Output "Status=NoWindowsUpdateErrorsFound"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
