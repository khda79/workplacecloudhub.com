#Requires -Version 7.0
<#
.SYNOPSIS
Collects Intune configuration definitions, assignments, and reported device states.
.VERSION
1.1.4
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$InteractiveAuth,
    [switch]$ValidateOnly,
    [switch]$EnableConfiguredExternalActions,
    [string]$OutputPath,
    [string]$LatestCsvFolderPath,
    [ValidateRange(0,10000000)][int]$MaxItems = 0,
    [ValidateRange(30,7200)][int]$ExportTimeoutSeconds = 1200
)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Runtime = $null
$failure = $null

function Get-IntuneConfigurationApplicationState {
    param([string]$PolicyStatus,[string]$ReportStatus)
    $state = if ([string]::IsNullOrWhiteSpace($ReportStatus)) { $PolicyStatus } else { $ReportStatus }
    switch ($state.Trim().ToLowerInvariant()) {
        { $_ -in @('success','succeeded','applied','compliant','2') } { return 'ReportedApplied' }
        { $_ -in @('error','failed','conflict','noncompliant','4','5','6') } { return 'ReportedIssue' }
        { $_ -in @('pending','inprogress') } { return 'Pending' }
        { $_ -in @('notapplicable','not applicable','1') } { return 'NotApplicable' }
        default { return 'Unknown' }
    }
}

try {
    $smartRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
    . (Join-Path $smartRoot 'Config/SmartM365-TenantContext.ps1')
    $effective = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.112' -Global -ErrorAction Stop
    Import-Module (Join-Path $smartRoot 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.17' -Force
    $script:Runtime = Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effective `
        -DefaultOutputRelativePath 'Intune/WorkplaceScope/Configurations' -OutputPath $OutputPath `
        -LatestCsvFolderPath $LatestCsvFolderPath -ValidateOnly:$ValidateOnly `
        -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath `
        -ReadOnly:($ValidateOnly -or $MaxItems -gt 0)
    # DATA-LAST changes only when this run completes; a failed or running run leaves the previous set intact.
    Enable-SmartM365StagedSourcePublication
    Enable-SmartM365DailyCaptureRetention -OutputRoot $script:Runtime.OutputPath
    Connect-SmartM365EvidenceGraph $script:Runtime @('DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.ReadWrite.All') -InteractiveAuth:$InteractiveAuth
    Invoke-SmartM365Preflight -ScriptName 'SmartM365-IntuneConfigurations-Inventory' `
        -OutputPaths @($script:Runtime.OutputPath,$script:Runtime.LatestCsvFolderPath) `
        -RequiredGraphApplicationPermissions @('DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.ReadWrite.All') `
        -GraphProbeUris @('https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$top=1') | Out-Null
    if ($ValidateOnly) {
        Write-SmartM365EvidenceLog 'Configuration prerequisites validated. Report availability requires a collection run.'
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }

    $families = @(
        @{ Name='SettingsCatalog'; Path='configurationPolicies' },
        @{ Name='DeviceConfiguration'; Path='deviceConfigurations' },
        @{ Name='DeviceCompliance'; Path='deviceCompliancePolicies' },
        @{ Name='WindowsFeatureUpdate'; Path='windowsFeatureUpdateProfiles' },
        @{ Name='WindowsQualityUpdate'; Path='windowsQualityUpdateProfiles' }
    )
    $policies = [Collections.Generic.List[object]]::new()
    $settings = [Collections.Generic.List[object]]::new()
    $assignments = [Collections.Generic.List[object]]::new()
    $seenPolicies = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($family in $families) {
        $base = 'https://graph.microsoft.com/beta/deviceManagement/' + $family.Path
        $items = @(Get-SmartM365EvidenceGraphCollection -Uri $base -MaxItems $MaxItems)
        foreach ($policy in $items) {
            $id = [string](Get-SmartM365EvidenceProperty $policy @('id'))
            if (-not $id -or -not $seenPolicies.Add($id)) { throw "Missing or duplicate Intune policy id in $($family.Name)." }
            $name = [string](Get-SmartM365EvidenceProperty $policy @('name','displayName'))
            $policies.Add([pscustomobject][ordered]@{
                PolicyId=$id; PolicyFamily=$family.Name; DisplayName=$name
                PolicyType=[string](Get-SmartM365EvidenceProperty $policy @('@odata.type'))
                Platforms=[string](Get-SmartM365EvidenceProperty $policy @('platforms'))
                Technologies=[string](Get-SmartM365EvidenceProperty $policy @('technologies'))
                CreatedDateTime=[string](Get-SmartM365EvidenceProperty $policy @('createdDateTime'))
                LastModifiedDateTime=[string](Get-SmartM365EvidenceProperty $policy @('lastModifiedDateTime'))
                NativePolicyJson=(ConvertTo-Json -InputObject $policy -Depth 50 -Compress)
                RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
            })
            $encoded = [uri]::EscapeDataString($id)
            $targets = @(Get-SmartM365EvidenceGraphCollection -Uri "$base/$encoded/assignments")
            $seenAssignments = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($assignment in $targets) {
                $assignmentId = [string](Get-SmartM365EvidenceProperty $assignment @('id'))
                $target = Get-SmartM365EvidenceProperty $assignment @('target') $null
                if (-not $assignmentId -or -not $seenAssignments.Add($assignmentId) -or $null -eq $target) { throw "Invalid assignment for policy $id." }
                $assignments.Add([pscustomobject][ordered]@{
                    PolicyId=$id; PolicyFamily=$family.Name; AssignmentId=$assignmentId
                    TargetType=[string](Get-SmartM365EvidenceProperty $target @('@odata.type'))
                    GroupId=[string](Get-SmartM365EvidenceProperty $target @('groupId'))
                    AssignmentFilterId=[string](Get-SmartM365EvidenceProperty $target @('deviceAndAppManagementAssignmentFilterId'))
                    AssignmentFilterType=[string](Get-SmartM365EvidenceProperty $target @('deviceAndAppManagementAssignmentFilterType'))
                    RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
                })
            }
            if ($family.Name -eq 'SettingsCatalog') {
                $nativeSettings = @(Get-SmartM365EvidenceGraphCollection -Uri "$base/$encoded/settings")
                $seenSettings = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($setting in $nativeSettings) {
                    $settingId = [string](Get-SmartM365EvidenceProperty $setting @('id'))
                    $instance = Get-SmartM365EvidenceProperty $setting @('settingInstance') $null
                    $definitionId = [string](Get-SmartM365EvidenceProperty $instance @('settingDefinitionId'))
                    if (-not $settingId -or -not $seenSettings.Add($settingId) -or -not $definitionId) { throw "Invalid setting for policy $id." }
                    $settings.Add([pscustomobject][ordered]@{
                        PolicyId=$id; PolicyFamily=$family.Name; SettingId=$settingId
                        SettingDefinitionId=$definitionId
                        SettingInstanceType=[string](Get-SmartM365EvidenceProperty $instance @('@odata.type'))
                        SettingInstanceJson=(ConvertTo-Json -InputObject $instance -Depth 50 -Compress)
                        RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
                    })
                }
            }
        }
    }
    $assignmentRaw = @(Invoke-SmartM365EvidenceExportReport -ReportName 'DeviceAssignmentStatusByConfigurationPolicyV3' -ApiVersion v1.0 -TimeoutSeconds $ExportTimeoutSeconds -MaxItems $MaxItems)
    $assignmentStates = @($assignmentRaw | ForEach-Object {
        $policyId = [string](Get-SmartM365EvidenceProperty $_ @('PolicyId'))
        $deviceId = [string](Get-SmartM365EvidenceProperty $_ @('IntuneDeviceId'))
        if (-not $policyId -or -not $deviceId) { throw 'Configuration assignment status lacks policy or device id.' }
        [pscustomobject][ordered]@{
            PolicyId=$policyId; IntuneDeviceId=$deviceId
            AssignmentStatus=[string](Get-SmartM365EvidenceProperty $_ @('AssignmentStatus'))
            ReportStatus=[string](Get-SmartM365EvidenceProperty $_ @('ReportStatus'))
            LastModifiedTimeUtc=[string](Get-SmartM365EvidenceProperty $_ @('PspdpuLastModifiedTimeUtc'))
            RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
        }
    })
    $assignmentRaw = $null
    $statusRaw = @(Invoke-SmartM365EvidenceExportReport -ReportName 'DeviceStatusesByConfigurationProfileV3' -ApiVersion v1.0 -TimeoutSeconds $ExportTimeoutSeconds -MaxItems $MaxItems)
    $deviceStates = @($statusRaw | ForEach-Object {
        $policyId = [string](Get-SmartM365EvidenceProperty $_ @('PolicyId'))
        $deviceId = [string](Get-SmartM365EvidenceProperty $_ @('IntuneDeviceId'))
        if (-not $policyId -or -not $deviceId) { throw 'Configuration device status lacks policy or device id.' }
        $policyStatus = [string](Get-SmartM365EvidenceProperty $_ @('PolicyStatus'))
        $reportStatus = [string](Get-SmartM365EvidenceProperty $_ @('ReportStatus'))
        [pscustomobject][ordered]@{
            PolicyId=$policyId; IntuneDeviceId=$deviceId
            UPN=[string](Get-SmartM365EvidenceProperty $_ @('UPN'))
            PolicyName=[string](Get-SmartM365EvidenceProperty $_ @('PolicyName'))
            DeviceName=[string](Get-SmartM365EvidenceProperty $_ @('DeviceName'))
            PolicyInCollectedCatalog=$seenPolicies.Contains($policyId)
            PolicyStatus=$policyStatus
            ReportedApplicationState=Get-IntuneConfigurationApplicationState $policyStatus $reportStatus
            ReportStatus=$reportStatus
            LastModifiedTimeUtc=[string](Get-SmartM365EvidenceProperty $_ @('PspdpuLastModifiedTimeUtc'))
            RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
        }
    })
    $statusRaw = $null
    $currentByKey = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    $undatedStateRows = 0
    foreach ($state in $deviceStates) {
        [datetimeoffset]$stateTime = [datetimeoffset]::MinValue
        if (-not [datetimeoffset]::TryParse([string]$state.LastModifiedTimeUtc,[ref]$stateTime)) {
            $undatedStateRows++
            continue
        }
        $key = '{0}|{1}' -f $state.PolicyId,$state.IntuneDeviceId
        if (-not $currentByKey.ContainsKey($key)) {
            $currentByKey[$key] = [pscustomobject]@{ Row=$state; Time=$stateTime }
            continue
        }
        $prior = $currentByKey[$key]
        if ($stateTime -gt $prior.Time) {
            $currentByKey[$key] = [pscustomobject]@{ Row=$state; Time=$stateTime }
        }
        elseif ($stateTime -eq $prior.Time -and $state.ReportedApplicationState -ne $prior.Row.ReportedApplicationState) {
            throw 'Conflicting configuration states share the latest timestamp for a policy/device key.'
        }
    }
    $currentStates = @($currentByKey.Values | ForEach-Object { $_.Row })
    $currentByKey = $null
    $unknownStates = @($deviceStates | Where-Object { $_.ReportedApplicationState -eq 'Unknown' }).Count
    if ($deviceStates.Count -gt 0 -and $unknownStates -eq $deviceStates.Count) {
        throw 'Every configuration device state is Unknown; source status values require qualification before publication.'
    }
    if ($deviceStates.Count -gt 0 -and $currentStates.Count -eq 0) {
        throw 'No dated configuration device state can establish the current state.'
    }
    $uncataloguedStates = @($currentStates | Where-Object { -not $_.PolicyInCollectedCatalog }).Count
    $assignmentReportState = if ($assignmentStates.Count -gt 0) { 'Reported' } elseif ($assignments.Count -gt 0) { 'NoRowsWithAssignments' } else { 'NoRowsNoAssignments' }
    if ($assignmentReportState -eq 'NoRowsWithAssignments') {
        Write-SmartM365EvidenceLog 'Configuration assignment status report returned no rows while policy assignments exist; per-device assignment status is not qualified.' -Level WARNING
    }
    if ($undatedStateRows -gt 0) {
        Write-SmartM365EvidenceLog "Configuration device status rows without a valid timestamp were excluded from current state: $undatedStateRows." -Level WARNING
    }
    $statusCoverage = @([pscustomobject][ordered]@{
        AssignmentStatusReportState=$assignmentReportState
        AssignmentStatusRows=$assignmentStates.Count
        DeviceStatusHistoryRows=$deviceStates.Count
        DeviceStatusCurrentRows=$currentStates.Count
        UnknownDeviceStatusRows=$unknownStates
        UndatedDeviceStatusRows=$undatedStateRows
        CurrentStatesOutsideCollectedCatalog=$uncataloguedStates
        RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
    })
    if ($MaxItems -gt 0) {
        Write-SmartM365EvidenceLog "Sample complete: policies=$($policies.Count); settings=$($settings.Count); assignments=$($assignments.Count); assignmentStates=$($assignmentStates.Count); deviceStateHistory=$($deviceStates.Count); deviceStateCurrent=$($currentStates.Count). No canonical CSV published."
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    $rules = $global:SmartM365CsvValidationRules
    if ($rules -isnot [hashtable]) { throw 'Core CSV validation rules are unavailable.' }
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ConfigurationPolicies' -CriticalFields @('PolicyId','PolicyFamily','RunId','CollectedAtUtc') -UniqueFields @('PolicyId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ConfigurationSettings' -CriticalFields @('PolicyId','SettingId','SettingDefinitionId','RunId','CollectedAtUtc') -UniqueFields @('PolicyId','SettingId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ConfigurationAssignments' -CriticalFields @('PolicyId','AssignmentId','RunId','CollectedAtUtc') -UniqueFields @('PolicyId','AssignmentId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ConfigurationAssignmentStates' -CriticalFields @('PolicyId','IntuneDeviceId','RunId','CollectedAtUtc') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ConfigurationDeviceStates' -CriticalFields @('PolicyId','IntuneDeviceId','RunId','CollectedAtUtc') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ConfigurationDeviceCurrentStates' -CriticalFields @('PolicyId','IntuneDeviceId','LastModifiedTimeUtc','RunId','CollectedAtUtc') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ConfigurationStatusCoverage' -CriticalFields @('AssignmentStatusReportState','RunId','CollectedAtUtc') -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    $deviceStateColumns = @('PolicyId','IntuneDeviceId','UPN','PolicyName','DeviceName','PolicyInCollectedCatalog','PolicyStatus','ReportedApplicationState','ReportStatus','LastModifiedTimeUtc','RunId','CollectedAtUtc')
    $exports = @(
        @{Name='Intune_ConfigurationPolicies';Rows=$policies.ToArray();Columns=@('PolicyId','PolicyFamily','DisplayName','PolicyType','Platforms','Technologies','CreatedDateTime','LastModifiedDateTime','NativePolicyJson','RunId','CollectedAtUtc')},
        @{Name='Intune_ConfigurationSettings';Rows=$settings.ToArray();Columns=@('PolicyId','PolicyFamily','SettingId','SettingDefinitionId','SettingInstanceType','SettingInstanceJson','RunId','CollectedAtUtc')},
        @{Name='Intune_ConfigurationAssignments';Rows=$assignments.ToArray();Columns=@('PolicyId','PolicyFamily','AssignmentId','TargetType','GroupId','AssignmentFilterId','AssignmentFilterType','RunId','CollectedAtUtc')},
        @{Name='Intune_ConfigurationAssignmentStates';Rows=$assignmentStates;Columns=@('PolicyId','IntuneDeviceId','AssignmentStatus','ReportStatus','LastModifiedTimeUtc','RunId','CollectedAtUtc')},
        @{Name='Intune_ConfigurationDeviceStates';Rows=$deviceStates;Columns=$deviceStateColumns},
        @{Name='Intune_ConfigurationDeviceCurrentStates';Rows=$currentStates;Columns=$deviceStateColumns},
        @{Name='Intune_ConfigurationStatusCoverage';Rows=$statusCoverage;Columns=@('AssignmentStatusReportState','AssignmentStatusRows','DeviceStatusHistoryRows','DeviceStatusCurrentRows','UnknownDeviceStatusRows','UndatedDeviceStatusRows','CurrentStatesOutsideCollectedCatalog','RunId','CollectedAtUtc')}
    )
    foreach ($export in $exports) { Assert-SmartM365CsvDataCompleteness -Data $export.Rows -BaseFileName $export.Name -Columns $export.Columns | Out-Null }
    $assignmentStateCount = $assignmentStates.Count
    $deviceStateHistoryCount = $deviceStates.Count
    $deviceStateCurrentCount = $currentStates.Count
    $assignmentStates = $null
    $deviceStates = $null
    $currentStates = $null
    [GC]::Collect()
    foreach ($export in $exports) {
        $rowCount = $export.Rows.Count
        Export-SmartM365EvidenceDataset $script:Runtime $export.Name $export.Rows $export.Columns -SingleSerialization | Out-Null
        $export.Rows = $null
        if ($rowCount -ge 100000) { [GC]::Collect() }
    }
    Write-SmartM365EvidenceLog "Intune configurations collected: policies=$($policies.Count); settings=$($settings.Count); assignments=$($assignments.Count); assignmentStates=$assignmentStateCount; deviceStateHistory=$deviceStateHistoryCount; deviceStateCurrent=$deviceStateCurrentCount; unknownStates=$unknownStates." -Level SUCCESS

    $smartM365RunGuard.ReachedTerminal=$true
} catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_
 $failure = $_; throw } finally {
    $smartM365RunGuard.InCleanup=$true

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false } catch {}
    if ($script:Runtime) { Complete-SmartM365EvidenceRuntime -Status $(if ($failure) {'Failed'} else {'Success'}) -ErrorRecord $failure -FailureStage 'IntuneConfigurationsInventory' }
}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhZje7B3LUoF0w
# 2tvdRNaIjaU6BmaqtJfkk8mOEnqDaqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIAn2PopOjpoXf5bFph1qDp+KZ66yZ5vZDdbxvBvAhTA/MA0GCSqG
# SIb3DQEBAQUABIIBgIAMGUWeZ3umyqYBCTNkurlkWz4XIoMRoKiOc299icy8Ynth
# bNZygaHkAV+j0vVcczD3c9uoCoB/sR3O6S85KGEkKaHsBq09N3O3UjhlgqrwHZv3
# SXVKVPzbuGQUTqohh9k21ol3MAim/uy/ntEPPF5IZa9flWgKQ8wNH9LAEHAnFbvQ
# IM4UheUkyWkrRGXZ99/8IY9zBI1s538MGsmAOz0LT0wBkXbCbhO8Jxb6fH/ua3+p
# PDqnnkfCww4/sU0vfA14MoSyrPUhSTBuJtyuZSABAE0c0puBevXulu6yXWsK3Gn5
# 7vfnLL7PefLGH4YWJgaCgCDLaJjfz1YZEChdax14+e6I7qNd9mgrZ6Y7IBDWz5vs
# TeBJu05p4x8dtPoebAKmYShGH0uIHkDjqOfNh+4WFQZfga64Cxjjhze83xTMj/Sn
# 7DxMEW9t/EaxRh5OgMuT9PfPReq89wLGevWJ9CjhsPmZ3vxod6V53n2Jz8wYtA2o
# 8+fm0EXN0zA84CpOLaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAwNjIx
# NDRaMC8GCSqGSIb3DQEJBDEiBCDEeuBuPCua7dWqgHppasTKzfAHB02JgZz6KvgK
# ncnbXjANBgkqhkiG9w0BAQEFAASCAgAoNtTzKuLkunbsO+FcmB5lxkFziCMAnBKw
# E/UZIdnGs+f8A5Fa0OFetCGsi9VFRx3MEyaKfrZgmVoipcIZ4Juh8J+Q4TMyrYu4
# Mhmhj17XH1PPmYrnHO0rpK6TlIhpnjEYeVSx1pR1072LxrqVcrN+xP93ROpPvQDR
# 66DURljJVkGIuqmM+R3thn9oacHInJNObvL0ViOQ0R3TRfOTqpnhuOHmq248SO7u
# FT/oaVLGNrp+In/8OJpsLVdWW9mz0LP1X18xf98Xd0OvHG7lax+/EYbQ2L8M+5x7
# BLTc1E7A4eVhr8oecT1TZri75YIp+WV/bGng/YzHxnLbYTHBmzkpiQTI5uawptop
# +mcKaJ2IglocTo6K1O3UuyGmSXBJB8Cf4fAmALpqYgNfhHzVT0IztbzSHf4Xp2XZ
# iPHV6LFpGYZHMa98ppPbIVjUtfsFItXVLJjvWfDeJ/QpKpvVOuEkMp0d6wCvLkf5
# sJQj7Ov3xOKjKwkpCLILJqUShYnoruZXbhzIg24s1zNhXI07FbwUlU47qmD0Xsjn
# xxSLa8796nfHU11IPuFOnx/0jYwazoGEjPl+kunobAoYoxaj++ty6ffy27jUoW0E
# xSshu06tpp4y0OWtpbO6SH4w4quk85cwMGp0rRr3wcCehq2WFMw1d4cTqx53nzBi
# LSgy6aVh4A==
# SIG # End signature block
