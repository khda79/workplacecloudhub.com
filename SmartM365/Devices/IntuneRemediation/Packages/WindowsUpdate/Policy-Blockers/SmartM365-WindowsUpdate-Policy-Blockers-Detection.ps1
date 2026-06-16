# Name: SmartM365-WindowsUpdate-Policy-Blockers-Detection.ps1
# Version: 1.0
# Description: Detects WSUS, Windows Update, and WUfB policy values that can block cloud-managed update flows.

$ErrorActionPreference = "Stop"

$RequireWUfBPolicyManager = $false
$ScriptName = "Detect-WindowsUpdate-Policy-Blockers"
$Version = "1.0"

function Add-Finding {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$Findings,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Findings.Contains($Message)) {
        $Findings.Add($Message)
    }
}

try {
    $policyManagerUpdatePath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"
    $wuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $wuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    $blockingFindings = New-Object System.Collections.Generic.List[string]
    $contextFindings = New-Object System.Collections.Generic.List[string]

    if ($RequireWUfBPolicyManager -and -not (Test-Path -Path $policyManagerUpdatePath)) {
        Add-Finding -Findings $blockingFindings -Message "WUfB PolicyManager configuration is missing."
    }

    if (Test-Path -Path $wuPolicyPath) {
        $wuPolicy = Get-ItemProperty -Path $wuPolicyPath -ErrorAction Stop

        foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate")) {
            if ($wuPolicy.PSObject.Properties.Name -contains $name) {
                $value = [string]$wuPolicy.$name

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    Add-Finding -Findings $blockingFindings -Message "$name=$value"
                }
            }
        }

        foreach ($name in @("DoNotConnectToWindowsUpdateInternetLocations", "DisableWindowsUpdateAccess", "SetDisableUXWUAccess")) {
            if ($wuPolicy.PSObject.Properties.Name -contains $name) {
                $value = $wuPolicy.$name
                Add-Finding -Findings $contextFindings -Message "$name=$value"

                if ($value -eq 1) {
                    Add-Finding -Findings $blockingFindings -Message "$name=1"
                }
            }
        }
    }

    if (Test-Path -Path $wuAuPolicyPath) {
        $wuAuPolicy = Get-ItemProperty -Path $wuAuPolicyPath -ErrorAction Stop

        foreach ($name in @("UseWUServer", "NoAutoUpdate")) {
            if ($wuAuPolicy.PSObject.Properties.Name -contains $name) {
                $value = $wuAuPolicy.$name
                Add-Finding -Findings $contextFindings -Message "$name=$value"

                if ($value -eq 1) {
                    Add-Finding -Findings $blockingFindings -Message "$name=1"
                }
            }
        }

        if ($wuAuPolicy.PSObject.Properties.Name -contains "AUOptions") {
            Add-Finding -Findings $contextFindings -Message "AUOptions=$($wuAuPolicy.AUOptions)"
        }
    }

    if ($blockingFindings.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. Blocking policies: {2}" -f $ScriptName, $Version, ($blockingFindings -join " | "))
        exit 1
    }

    if ($contextFindings.Count -gt 0) {
        Write-Output ("{0} v{1}: no blocking policy detected. Context: {2}" -f $ScriptName, $Version, ($contextFindings -join " | "))
        exit 0
    }

    Write-Output ("{0} v{1}: no WSUS or Windows Update policy blocker detected." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}
