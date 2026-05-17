# Name: SmartM365-WindowsUpdate-Reset-Detection.ps1
# Version: 1.0
# Description: Detects Windows Update reset conditions such as legacy WSUS policies or recent Autopatch error 0x80244007.

$ErrorActionPreference = "Stop"

$ScriptName = "Detect-WindowsUpdate-Reset"
$Version = "1.0"
$ForceRemediation = $false

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

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $wuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $wuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

    if (Test-Path -Path $wuPolicyPath) {
        $wuPolicy = Get-ItemProperty -Path $wuPolicyPath -ErrorAction Stop

        foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate")) {
            if ($wuPolicy.PSObject.Properties.Name -contains $name) {
                $value = [string]$wuPolicy.$name

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    Add-Issue -Issues $issues -Message "$name is configured."
                }
            }
        }
    }

    if (Test-Path -Path $wuAuPolicyPath) {
        $wuAuPolicy = Get-ItemProperty -Path $wuAuPolicyPath -ErrorAction Stop

        if ($wuAuPolicy.PSObject.Properties.Name -contains "UseWUServer" -and [int]$wuAuPolicy.UseWUServer -eq 1) {
            Add-Issue -Issues $issues -Message "UseWUServer=1."
        }
    }

    $reportingEventsPath = Join-Path -Path $env:windir -ChildPath "SoftwareDistribution\ReportingEvents.log"

    if (Test-Path -LiteralPath $reportingEventsPath) {
        $recentError = Select-String -Path $reportingEventsPath -Pattern "0x80244007|Windows Update Client failed to detect" -ErrorAction SilentlyContinue |
            Select-Object -Last 1

        if ($recentError) {
            Add-Issue -Issues $issues -Message "Recent Windows Update detection failure 0x80244007 found."
        }
    }

    if ($ForceRemediation) {
        if ($issues.Count -gt 0) {
            Write-Output ("{0} v{1}: forced remediation required. Indicators: {2}" -f $ScriptName, $Version, ($issues -join " | "))
        }
        else {
            Write-Output ("{0} v{1}: forced remediation required. No blocking indicator found." -f $ScriptName, $Version)
        }

        exit 1
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: no remediation required." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}
