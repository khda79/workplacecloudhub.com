# Name: SmartM365-Upgrade-Staging-Health-Detection.ps1
# Version: 1.0
# Description: Detects stale Windows upgrade staging folders and missing upgrade image files.

$ErrorActionPreference = "Stop"

$ScriptName = "Detect-Upgrade-Staging-Health"
$Version = "1.0"
$RecentSetupActivityHours = 6

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

function Test-RecentSetupActivity {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Hours
    )

    $setupIndicators = @(
        "C:\Windows\Panther\setupact.log",
        "C:\Windows\Panther\setuperr.log",
        'C:\$WINDOWS.~BT\Sources\Panther\setupact.log',
        'C:\$WINDOWS.~BT\Sources\Panther\setuperr.log'
    )

    foreach ($indicator in $setupIndicators) {
        if (Test-Path -LiteralPath $indicator) {
            $item = Get-Item -LiteralPath $indicator -ErrorAction SilentlyContinue

            if ($item -and ((Get-Date) - $item.LastWriteTime).TotalHours -le $Hours) {
                return $true
            }
        }
    }

    return $false
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $recentSetupActivity = Test-RecentSetupActivity -Hours $RecentSetupActivityHours
    $upgradeSourcesPath = 'C:\$WINDOWS.~BT\Sources'
    $upgradeResiduePaths = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS')

    if ($recentSetupActivity) {
        Write-Output ("{0} v{1}: recent setup activity detected; no remediation requested." -f $ScriptName, $Version)
        exit 0
    }

    if (Test-Path -LiteralPath $upgradeSourcesPath) {
        $upgradeImageFiles = Get-ChildItem -LiteralPath $upgradeSourcesPath -Recurse -Include "*.esd", "*.wim" -File -ErrorAction SilentlyContinue
        $validImageFiles = $upgradeImageFiles | Where-Object { $_.Length -gt 0 }

        if ($null -eq $validImageFiles -or $validImageFiles.Count -eq 0) {
            Add-Issue -Issues $issues -Message "Upgrade staging sources exist but no non-empty ESD or WIM file was found."
        }
    }

    foreach ($path in $upgradeResiduePaths) {
        if (Test-Path -LiteralPath $path) {
            Add-Issue -Issues $issues -Message "Potentially stale upgrade staging folder detected: $path"
        }
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: no stale Windows upgrade staging content detected." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}
