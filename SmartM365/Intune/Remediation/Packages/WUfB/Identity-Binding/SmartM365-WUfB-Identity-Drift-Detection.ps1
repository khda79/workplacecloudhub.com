# Name: SmartM365-WUfB-Identity-Drift-Detection.ps1
# Version: 1.1
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

    $mdmWUfBPresent = Test-Path -Path $mdmUpdatePolicyPath
    $isWUfBConfigured = Get-RegistryPropertyValue -Path $policyStatePath -Name "IsWUfBConfigured"
    $policyStateValue = if ($null -eq $isWUfBConfigured) { "Unknown" } else { [string]$isWUfBConfigured }

    $commonData = @{
        MDMWUfBPolicyPresent = $mdmWUfBPresent
        PolicyStateIsWUfBConfigured = $policyStateValue
    }

    # Drift condition:
    # MDM WUfB policy exists, but Windows Update PolicyState says WUfB is not configured.
    if ($mdmWUfBPresent -and $isWUfBConfigured -eq 0) {
        $commonData["Reason"] = "MDM WUfB policy is present but PolicyState IsWUfBConfigured is 0"
        Write-IntuneResult -Status "WUfBIdentityDriftDetected" -Data $commonData
        exit 1
    }

    # Do not fail devices that do not have MDM WUfB policy.
    if (-not $mdmWUfBPresent) {
        $commonData["Reason"] = "No MDM WUfB policy detected"
        Write-IntuneResult -Status "NotApplicable" -Data $commonData
        exit 0
    }

    if ($null -eq $isWUfBConfigured) {
        $commonData["Reason"] = "MDM WUfB policy is present but PolicyState IsWUfBConfigured is unavailable"
        Write-IntuneResult -Status "HealthyWithUnknownPolicyState" -Data $commonData
        exit 0
    }

    $commonData["Reason"] = "MDM WUfB policy and PolicyState are aligned"
    Write-IntuneResult -Status "Healthy" -Data $commonData
    exit 0
}
catch {
    Write-Output ("Status=Error; Message={0}" -f (($_.Exception.Message -replace "\s+", " ").Trim()))
    exit 2
}
