# Name: SmartM365-SetupDiag-Required-Detection.ps1
# Version: 1.0
# Description: Detects whether SetupDiag analysis is required after a Windows setup or upgrade failure and verifies whether SetupDiag.exe is available

$ErrorActionPreference = "Stop"

try {
    $setupDiagResultPath = "C:\ProgramData\SmartM365\IntuneRemediation\Output\SetupDiag\SetupDiagResults.xml"
    $enterpriseRegistryPath = "HKLM:\SOFTWARE\SmartM365\IntuneRemediation\SetupDiag"
    $maxAnalysisAgeDays = 7

    # Local paths where SetupDiag.exe may be staged by Intune, SCCM, or an internal remediation package
    $setupDiagExecutablePaths = @(
        "C:\ProgramData\SmartM365\IntuneRemediation\Output\SetupDiag\SetupDiag.exe",
        "C:\ProgramData\SmartM365\IntuneRemediation\Tools\SetupDiag\SetupDiag.exe"
    )

    # Use single-quoted strings for paths containing "$Windows.~BT" to prevent PowerShell variable expansion
    $rollbackPaths = @(
        'C:\$Windows.~BT\Sources\Rollback',
        "C:\Windows\Panther\NewOS\Rollback",
        'C:\Windows.old\$Windows.~BT\Sources\Rollback'
    )

    $setupLogRoots = @(
        'C:\$Windows.~BT\Sources\Rollback',
        "C:\Windows\Panther",
        'C:\$Windows.~BT\Sources\Panther',
        'C:\Windows.old\$Windows.~BT\Sources\Rollback'
    )

    $hasSetupDiagExecutable = $false
    foreach ($setupDiagExecutablePath in $setupDiagExecutablePaths) {
        if (Test-Path -Path $setupDiagExecutablePath) {
            $hasSetupDiagExecutable = $true
            break
        }
    }

    $hasSetupDiagResult = Test-Path -Path $setupDiagResultPath

    $hasRollbackLogs = $false
    foreach ($rollbackPath in $rollbackPaths) {
        if (Test-Path -Path $rollbackPath) {
            $hasRollbackLogs = $true
            break
        }
    }

    $hasSetupLogs = $false
    foreach ($setupLogRoot in $setupLogRoots) {
        if (Test-Path -Path $setupLogRoot) {
            $hasSetupLogs = $true
            break
        }
    }

    # If no setup logs and no SetupDiag result exist, there is nothing actionable for SetupDiag
    if (-not $hasSetupLogs -and -not $hasSetupDiagResult) {
        Write-Output "No Windows setup logs found and no SetupDiag analysis required"
        exit 0
    }

    # Setup logs exist but SetupDiag.exe is not staged locally
    if ($hasSetupLogs -and -not $hasSetupDiagResult -and -not $hasSetupDiagExecutable) {
        Write-Output "Windows setup logs detected, SetupDiag analysis is required, but SetupDiag.exe is not available locally"
        exit 1
    }

    # A rollback without a SetupDiag result should trigger remediation
    if ($hasRollbackLogs -and -not $hasSetupDiagResult) {
        Write-Output "Windows setup rollback logs detected without SetupDiag analysis"
        exit 1
    }

    # If setup logs exist but no SetupDiag result exists, analysis is required
    if ($hasSetupLogs -and -not $hasSetupDiagResult) {
        Write-Output "Windows setup logs detected without SetupDiag analysis"
        exit 1
    }

    # If a SetupDiag result exists, validate the enterprise tracking timestamp when available
    if (Test-Path -Path $enterpriseRegistryPath) {
        $setupDiagState = Get-ItemProperty -Path $enterpriseRegistryPath -Name LastRunUtc -ErrorAction SilentlyContinue

        if ($null -ne $setupDiagState -and -not [string]::IsNullOrWhiteSpace($setupDiagState.LastRunUtc)) {
            $lastRunUtc = [datetime]::Parse($setupDiagState.LastRunUtc)
            $analysisAgeDays = (New-TimeSpan -Start $lastRunUtc -End (Get-Date).ToUniversalTime()).TotalDays

            if ($analysisAgeDays -gt $maxAnalysisAgeDays) {
                if ($hasSetupDiagExecutable) {
                    Write-Output "SetupDiag analysis is outdated and SetupDiag.exe is available locally"
                }
                else {
                    Write-Output "SetupDiag analysis is outdated and SetupDiag.exe is not available locally"
                }
                exit 1
            }

            Write-Output "SetupDiag analysis is recent and setup logs are consistent"
            exit 0
        }
    }

    # Fallback: use the XML file timestamp if the enterprise registry timestamp is unavailable
    if ($hasSetupDiagResult) {
        $xmlLastWriteUtc = (Get-Item -Path $setupDiagResultPath -ErrorAction Stop).LastWriteTimeUtc
        $xmlAgeDays = (New-TimeSpan -Start $xmlLastWriteUtc -End (Get-Date).ToUniversalTime()).TotalDays

        if ($xmlAgeDays -gt $maxAnalysisAgeDays) {
            if ($hasSetupDiagExecutable) {
                Write-Output "SetupDiag result file is outdated and SetupDiag.exe is available locally"
            }
            else {
                Write-Output "SetupDiag result file is outdated and SetupDiag.exe is not available locally"
            }
            exit 1
        }

        Write-Output "SetupDiag result file is recent"
        exit 0
    }

    Write-Output "SetupDiag analysis state is inconsistent"
    exit 1
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}

