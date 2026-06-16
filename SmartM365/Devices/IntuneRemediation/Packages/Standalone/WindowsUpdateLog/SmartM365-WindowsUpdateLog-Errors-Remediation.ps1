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
