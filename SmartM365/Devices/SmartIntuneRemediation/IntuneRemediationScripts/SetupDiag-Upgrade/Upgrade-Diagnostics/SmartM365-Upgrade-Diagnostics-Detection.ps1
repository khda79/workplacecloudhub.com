<#
    Name: SmartM365-Upgrade-Diagnostics-Detection.ps1
    Version: 1.0
    Description: Consolidated SetupDiag and Windows upgrade blocker detection.
#>

[CmdletBinding()]
param(
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 5,
    [int]$MaxIssueLength = 180,
    [int]$MaxAnalysisAgeDays = 7
)

$ErrorActionPreference = "Stop"
$SetupDiagResultPath = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Output\SetupDiag\SetupDiagResults.xml"
$SetupDiagToolPaths = @(
    (Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Tools\SetupDiag\SetupDiag.exe"),
    (Join-Path $env:ProgramData "SetupDiag\SetupDiag.exe")
)
$EnterpriseRegistryPath = "HKLM:\SOFTWARE\SmartM365\IntuneRemediation\SetupDiag"

$BlockingCodes = @(
    "0xC1900208",
    "0xC1900101",
    "0x80070070",
    "0x80240020",
    "0x80242016",
    "0x800F0922",
    "0xC1900200",
    "0x8007042B",
    "0x8007001F"
)

$RollbackPaths = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther\NewOS\Rollback"),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

$SetupLogRoots = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther"),
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Panther'),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Issues.Contains($Message)) {
        $Issues.Add($Message)
    }
}

function Add-EventIssue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Source,
        [AllowNull()]$Events,
        [Parameter(Mandatory = $true)][string[]]$Codes,
        [Parameter(Mandatory = $true)][int]$MaxIssueLength
    )

    foreach ($eventRecord in @($Events)) {
        if ($null -eq $eventRecord -or [string]::IsNullOrWhiteSpace($eventRecord.Message)) {
            continue
        }

        foreach ($code in $Codes) {
            if ($eventRecord.Message -match [regex]::Escape($code)) {
                $message = $eventRecord.Message -replace "`r|`n", " "
                if ($message.Length -gt $MaxIssueLength) {
                    $message = $message.Substring(0, $MaxIssueLength) + "..."
                }
                Add-Issue -Issues $Issues -Message "$Source Id=$($eventRecord.Id) Code=$code Message=$message"
            }
        }
    }
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $cutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)
    $hasSetupDiagResult = Test-Path -LiteralPath $SetupDiagResultPath -PathType Leaf
    $hasSetupDiagTool = $false
    foreach ($setupDiagToolPath in $SetupDiagToolPaths) {
        if (Test-Path -LiteralPath $setupDiagToolPath -PathType Leaf) {
            $hasSetupDiagTool = $true
            break
        }
    }
    $hasSetupLogs = $false
    $hasRollbackLogs = $false

    foreach ($setupLogRoot in $SetupLogRoots) {
        if (Test-Path -LiteralPath $setupLogRoot) {
            $hasSetupLogs = $true
            break
        }
    }

    foreach ($rollbackPath in $RollbackPaths) {
        if (Test-Path -LiteralPath $rollbackPath) {
            $hasRollbackLogs = $true
            Add-Issue -Issues $issues -Message "RollbackDetected=$rollbackPath"
        }
    }

    if ($hasSetupLogs -and -not $hasSetupDiagResult) {
        if ($hasSetupDiagTool) {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedWithoutSetupDiagAnalysis"
        }
        else {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedSetupDiagToolMissing"
        }
    }

    if ($hasRollbackLogs -and -not $hasSetupDiagResult) {
        Add-Issue -Issues $issues -Message "RollbackLogsDetectedWithoutSetupDiagAnalysis"
    }

    if ($hasSetupDiagResult) {
        $xmlAgeDays = (New-TimeSpan -Start (Get-Item -LiteralPath $SetupDiagResultPath).LastWriteTimeUtc -End (Get-Date).ToUniversalTime()).TotalDays
        if ($xmlAgeDays -gt $MaxAnalysisAgeDays) {
            Add-Issue -Issues $issues -Message ("SetupDiagResultOutdated AgeDays={0:N1}" -f $xmlAgeDays)
        }

        $xmlContent = Get-Content -LiteralPath $SetupDiagResultPath -Raw -ErrorAction SilentlyContinue
        foreach ($code in $BlockingCodes) {
            if ($xmlContent -match [regex]::Escape($code)) {
                Add-Issue -Issues $issues -Message "SetupDiagDetected=$code"
            }
        }

        if ($xmlContent -match "Matching Profile found:\s*(.+)") {
            Add-Issue -Issues $issues -Message "SetupDiagProfile=$($matches[1].Trim())"
        }
    }
    elseif (Test-Path -Path $EnterpriseRegistryPath) {
        $setupDiagState = Get-ItemProperty -Path $EnterpriseRegistryPath -Name LastRunUtc -ErrorAction SilentlyContinue
        if ($setupDiagState -and -not [string]::IsNullOrWhiteSpace($setupDiagState.LastRunUtc)) {
            $analysisAgeDays = (New-TimeSpan -Start ([datetime]::Parse($setupDiagState.LastRunUtc)) -End (Get-Date).ToUniversalTime()).TotalDays
            if ($analysisAgeDays -gt $MaxAnalysisAgeDays) {
                Add-Issue -Issues $issues -Message ("SetupDiagRegistryStateOutdated AgeDays={0:N1}" -f $analysisAgeDays)
            }
        }
    }

    $wuEvents = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 500 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "WUEvent" -Events $wuEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    $setupEvents = Get-WinEvent -LogName "Setup" -MaxEvents 300 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "SetupEvent" -Events $setupEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    if ($issues.Count -gt 0) {
        Write-Output "Status=UpgradeDiagnosticsRequired"
        Write-Output "IssueCount=$($issues.Count)"
        $issues | Select-Object -First $MaxIssuesToDisplay | ForEach-Object { Write-Output $_ }
        exit 1
    }

    Write-Output "Status=Healthy"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
