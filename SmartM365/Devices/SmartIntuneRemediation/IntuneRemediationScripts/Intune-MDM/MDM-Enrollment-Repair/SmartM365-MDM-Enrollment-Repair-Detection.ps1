# Name: SmartM365-MDM-Enrollment-Repair-Detection.ps1
# Version: 1.0
# Description: Detects stale local MDM enrollment, device registration, and EnterpriseMgmt scheduled task signals.

$ErrorActionPreference = "Stop"

$RequireHybridJoin = $false
$ScriptName = "Detect-MDM-Enrollment-Repair"
$Version = "1.0"

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

function Get-DsregValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($Text -match ("(?m)^\s*{0}\s*:\s*(.+?)\s*$" -f [regex]::Escape($Name))) {
        return $matches[1].Trim()
    }

    return $null
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $dsregText = (& "$env:SystemRoot\System32\dsregcmd.exe" /status 2>$null) -join "`n"

    $deviceId = Get-DsregValue -Text $dsregText -Name "DeviceId"
    $azureAdJoined = Get-DsregValue -Text $dsregText -Name "AzureAdJoined"
    $domainJoined = Get-DsregValue -Text $dsregText -Name "DomainJoined"

    if ([string]::IsNullOrWhiteSpace($deviceId)) {
        Add-Issue -Issues $issues -Message "DeviceId missing from dsregcmd status."
    }

    if ($azureAdJoined -ne "YES") {
        Add-Issue -Issues $issues -Message "AzureAdJoined is not YES."
    }

    if ($RequireHybridJoin -and $domainJoined -ne "YES") {
        Add-Issue -Issues $issues -Message "DomainJoined is not YES."
    }

    $enrollmentRoot = "HKLM:\SOFTWARE\Microsoft\Enrollments"

    if (-not (Test-Path -Path $enrollmentRoot)) {
        Add-Issue -Issues $issues -Message "MDM enrollment registry root is missing."
    }
    else {
        $enrollmentKeys = Get-ChildItem -Path $enrollmentRoot -ErrorAction SilentlyContinue

        if ($null -eq $enrollmentKeys -or $enrollmentKeys.Count -eq 0) {
            Add-Issue -Issues $issues -Message "No MDM enrollment registry records found."
        }
    }

    $enterpriseTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { $_.TaskPath -like "\Microsoft\Windows\EnterpriseMgmt\*" }

    if ($null -eq $enterpriseTasks -or $enterpriseTasks.Count -eq 0) {
        Add-Issue -Issues $issues -Message "EnterpriseMgmt scheduled tasks are missing."
    }
    else {
        $enabledTasks = $enterpriseTasks | Where-Object { $_.State -ne "Disabled" }

        if ($null -eq $enabledTasks -or $enabledTasks.Count -eq 0) {
            Add-Issue -Issues $issues -Message "EnterpriseMgmt scheduled tasks are present but disabled."
        }
    }

    $imeService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue

    if ($null -eq $imeService) {
        Add-Issue -Issues $issues -Message "Intune Management Extension service is missing."
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: MDM enrollment and device registration signals are healthy." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}
