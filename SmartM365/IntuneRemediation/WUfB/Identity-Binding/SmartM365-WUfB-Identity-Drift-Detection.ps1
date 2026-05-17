# Name: SmartM365-WUfB-Identity-Drift-Detection.ps1
# Version: 1.0
# Description: Detects a mismatch between MDM Windows Update for Business policy presence and Windows Update PolicyState identity

$ErrorActionPreference = "Stop"

try {
    $policyStatePath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState"
    $mdmUpdatePolicyPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"

    function Get-RegistryPropertyValue {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$Name
        )

        if (-not (Test-Path -Path $Path)) {
            return $null
        }

        $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue

        if ($null -eq $item) {
            return $null
        }

        if ($item.PSObject.Properties.Name -contains $Name) {
            return $item.$Name
        }

        return $null
    }

    $mdmWUfBPresent = Test-Path -Path $mdmUpdatePolicyPath
    $isWUfBConfigured = Get-RegistryPropertyValue -Path $policyStatePath -Name "IsWUfBConfigured"

    Write-Output "MDMWUfBPolicyPresent=$mdmWUfBPresent"

    if ($null -eq $isWUfBConfigured) {
        Write-Output "PolicyStateIsWUfBConfigured=Unknown"
    }
    else {
        Write-Output "PolicyStateIsWUfBConfigured=$isWUfBConfigured"
    }

    # Drift condition:
    # MDM WUfB policy exists, but Windows Update PolicyState says WUfB is not configured.
    if ($mdmWUfBPresent -and $isWUfBConfigured -eq 0) {
        Write-Output "Status=WUfBIdentityDriftDetected"
        Write-Output "Reason=MDM WUfB policy is present but PolicyState IsWUfBConfigured is 0"
        exit 1
    }

    # Do not fail devices that do not have MDM WUfB policy.
    if (-not $mdmWUfBPresent) {
        Write-Output "Status=NotApplicable"
        Write-Output "Reason=No MDM WUfB policy detected"
        exit 0
    }

    if ($null -eq $isWUfBConfigured) {
        Write-Output "Status=HealthyWithUnknownPolicyState"
        Write-Output "Reason=MDM WUfB policy is present but PolicyState IsWUfBConfigured is unavailable"
        exit 0
    }

    Write-Output "Status=Healthy"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 2
}
