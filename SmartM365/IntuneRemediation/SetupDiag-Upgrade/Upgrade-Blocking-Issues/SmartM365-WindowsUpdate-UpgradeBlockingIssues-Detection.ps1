<#
.SYNOPSIS
    Version: 1.0
    Detects blocking Windows Update and Feature Update issues for Windows 11 readiness.

.VERSION
    5.0

.DESCRIPTION
    Enterprise-grade detection script for Microsoft Intune.

    This version avoids parsing WindowsUpdate.log because it generates
    excessive false positives in enterprise environments.

    Detection is based on:
    - Windows Update event logs
    - Setup event logs
    - SetupDiag results
    - Rollback traces
    - Known blocking HRESULTs

    Designed for:
    - Windows 10 / Windows 11
    - Windows Autopatch
    - WUfB readiness
    - Feature Update troubleshooting

.NOTES
    Exit codes:
    0 = Healthy / compliant
    1 = Blocking issue detected
    2 = Technical error
#>

[CmdletBinding()]
param(
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 15
)

$ErrorActionPreference = "Stop"

try {

    # =========================================================
    # Paths
    # =========================================================
    $ScriptName = "Detect-WindowsUpdate-UpgradeBlockingIssues"
    $LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$ScriptName"

    if (-not (Test-Path $LogRoot)) {
        New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    }

    $SetupDiagXml = Join-Path $env:WINDIR "Logs\SetupDiag\SetupDiagResults.xml"

    $RollbackPaths = @(
        (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
        (Join-Path $env:WINDIR 'Panther\NewOS\Rollback')
    )

    # =========================================================
    # Known blocking HRESULTs
    # =========================================================
    $BlockingCodes = @(
        "0xC1900208", # compatibility blocker
        "0xC1900101", # driver rollback
        "0x80070070", # insufficient disk space
        "0x80240020", # install pending
        "0x80242016", # operation in progress
        "0x800F0922", # servicing / reserved partition
        "0xC1900200", # unsupported hardware
        "0x8007042B", # migration failure
        "0x8007001F"  # driver / IO failure
    )

    $Issues = New-Object System.Collections.Generic.List[string]

    $CutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)

    Write-Output "Status=Scanning"
    Write-Output "LookbackDays=$LookbackDays"

    # =========================================================
    # SetupDiag correlation
    # =========================================================
    if (Test-Path $SetupDiagXml) {

        try {
            $xmlContent = Get-Content -Path $SetupDiagXml -Raw -ErrorAction Stop

            foreach ($code in $BlockingCodes) {
                if ($xmlContent -match [regex]::Escape($code)) {
                    $Issues.Add("SetupDiagDetected=$code")
                }
            }

            if ($xmlContent -match "Matching Profile found:\s*(.+)") {
                $Issues.Add("SetupDiagProfile=$($matches[1].Trim())")
            }

        }
        catch {
            $Issues.Add("SetupDiagParseFailure")
        }
    }

    # =========================================================
    # Rollback detection
    # =========================================================
    foreach ($rollbackPath in $RollbackPaths) {

        if (Test-Path $rollbackPath) {
            $Issues.Add("RollbackDetected=$rollbackPath")
        }
    }

    # =========================================================
    # Windows Update event logs
    # =========================================================
    $WUEvents = Get-WinEvent `
        -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" `
        -MaxEvents 500 `
        -ErrorAction SilentlyContinue | Where-Object {
            $_.TimeCreated -ge $CutoffDate
        }

    foreach ($event in $WUEvents) {

        foreach ($code in $BlockingCodes) {

            if ($event.Message -match [regex]::Escape($code)) {

                $msg = $event.Message -replace "`r|`n", " "

                if ($msg.Length -gt 250) {
                    $msg = $msg.Substring(0,250) + "..."
                }

                $Issues.Add("WUEvent Id=$($event.Id) Code=$code Message=$msg")
            }
        }
    }

    # =========================================================
    # Setup event logs
    # =========================================================
    $SetupEvents = Get-WinEvent `
        -LogName "Setup" `
        -MaxEvents 300 `
        -ErrorAction SilentlyContinue | Where-Object {
            $_.TimeCreated -ge $CutoffDate
        }

    foreach ($event in $SetupEvents) {

        foreach ($code in $BlockingCodes) {

            if ($event.Message -match [regex]::Escape($code)) {

                $msg = $event.Message -replace "`r|`n", " "

                if ($msg.Length -gt 250) {
                    $msg = $msg.Substring(0,250) + "..."
                }

                $Issues.Add("SetupEvent Id=$($event.Id) Code=$code Message=$msg")
            }
        }
    }

    # =========================================================
    # Deduplicate
    # =========================================================
    $Issues = $Issues | Select-Object -Unique

    # =========================================================
    # Final result
    # =========================================================
    if ($Issues.Count -gt 0) {

        Write-Output "Status=BlockingIssuesDetected"
        Write-Output "IssueCount=$($Issues.Count)"
        Write-Output "----------------------------------------"

        $Issues | Select-Object -First $MaxIssuesToDisplay | ForEach-Object {
            Write-Output $_
        }

        exit 1
    }

    Write-Output "Status=Healthy"
    exit 0
}
catch {

    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"

    exit 2
}

