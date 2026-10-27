#Requires -Version 7.0
<#
.SYNOPSIS
Collects Intune managed apps, assignments, and reported installation states.
.VERSION
1.3.0
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

function Get-IntuneManagedAppDeploymentState {
    param([string]$InstallState)
    switch (($InstallState.Trim().ToLowerInvariant() -replace '[\s_-]+','')) {
        'installed' { return 'ReportedInstalled' }
        { $_ -in @('failed','error','uninstallfailed') } { return 'ReportedFailure' }
        'notinstalled' { return 'ReportedNotInstalled' }
        { $_ -in @('pendinginstall','installpending','installing','pending','uninstallpending') } { return 'Pending' }
        'notapplicable' { return 'NotApplicable' }
        default { return 'Unknown' }
    }
}

function Get-IntuneAssignedManagedApps {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Apps,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Assignments
    )
    $assignedIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($assignment in $Assignments) { [void]$assignedIds.Add([string]$assignment.ApplicationId) }
    return @($Apps | Where-Object { $assignedIds.Contains([string]$_.ApplicationId) })
}

function Get-IntuneManagedAppStatusSource {
    param([Parameter(Mandatory)][string]$ApplicationId,[int]$ObservedRows,[Parameter(Mandatory)]$AssignedIds)
    if (-not $AssignedIds.Contains($ApplicationId)) { return 'NotAssigned' }
    if ($ObservedRows -gt 0) { return 'DeviceInstallStatusByApp' }
    return 'NoReportRows'
}

function Get-IntuneManagedAppInstallStates {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Apps,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$CollectedAtUtc,
        # Tenant identity columns, written first so the large dataset skips generic enrichment.
        [Parameter(Mandatory)]$Identity,
        [int]$ExportTimeoutSeconds=1200,
        [int]$MaxItems=0,
        [ValidateRange(0,60)][int]$PauseSeconds=3,
        [scriptblock]$ExportReport
    )
    $states = [Collections.Generic.List[object]]::new()
    $unknownStateCounts = [Collections.Generic.Dictionary[string,int]]::new([StringComparer]::Ordinal)
    $unknownStateRows = 0
    $unknownStateOverflow = 0
    $reportColumns = @('ApplicationId','DeviceId','UserPrincipalName','AppVersion','Platform','InstallState','InstallStateDetail','ErrorCode','LastModifiedDateTime','AppInstallState','AppInstallStateDetails','HexErrorCode')
    $tenantKey = [string]$Identity.TenantKey; $organizationKey = [string]$Identity.OrganizationKey
    $environmentKey = [string]$Identity.EnvironmentKey; $tenantId = [string]$Identity.TenantId
    foreach ($requiredIdentity in @($tenantKey,$organizationKey,$environmentKey)) {
        if ([string]::IsNullOrWhiteSpace($requiredIdentity)) { throw 'TenantKey, OrganizationKey and EnvironmentKey are required for managed-app device states.' }
    }
    for ($index=0; $index -lt $Apps.Count; $index++) {
        $appId = [string]$Apps[$index].ApplicationId
        $parsedId = [guid]::Empty
        if (-not [guid]::TryParse($appId,[ref]$parsedId)) { throw "Invalid managed app id in catalog: $appId" }
        if ($index -gt 0 -and $PauseSeconds -gt 0) { Start-Sleep -Seconds $PauseSeconds }
        $filter = "(ApplicationId eq '$($parsedId.ToString('D'))')"
        try {
            $reportRows = if ($ExportReport) {
                @(& $ExportReport $appId $filter)
            } else {
                @(Invoke-SmartM365EvidenceExportReport -ReportName 'DeviceInstallStatusByApp' -Filter $filter -Select $reportColumns -ApiVersion beta -TimeoutSeconds $ExportTimeoutSeconds -MaxItems $MaxItems)
            }
        } catch {
            Write-SmartM365EvidenceLog "DeviceInstallStatusByApp export failed for app $appId ($($index+1)/$($Apps.Count))." -Level ERROR
            throw
        }
        foreach ($row in $reportRows) {
            $rowAppId = [string](Get-SmartM365EvidenceProperty $row @('ApplicationId'))
            if ($rowAppId -ine $appId) { throw "DeviceInstallStatusByApp returned an absent or mismatched ApplicationId for app $appId." }
            $installState = [string](Get-SmartM365EvidenceProperty $row @('InstallState'))
            $appInstallState = [string](Get-SmartM365EvidenceProperty $row @('AppInstallState'))
            $stateForClassification = if ([string]::IsNullOrWhiteSpace($appInstallState)) { $installState } else { $appInstallState }
            $reportedState = Get-IntuneManagedAppDeploymentState $stateForClassification
            if ($reportedState -eq 'Unknown') {
                $unknownStateRows++
                $appValue = if ([string]::IsNullOrWhiteSpace($appInstallState)) { '<empty>' } else { ($appInstallState -replace '[\r\n\t]',' ').Trim() }
                $installValue = if ([string]::IsNullOrWhiteSpace($installState)) { '<empty>' } else { ($installState -replace '[\r\n\t]',' ').Trim() }
                if ($appValue.Length -gt 60) { $appValue = $appValue.Substring(0,60) + '...' }
                if ($installValue.Length -gt 60) { $installValue = $installValue.Substring(0,60) + '...' }
                $stateKey = "AppInstallState='$appValue', InstallState='$installValue'"
                if ($unknownStateCounts.ContainsKey($stateKey)) { $unknownStateCounts[$stateKey]++ }
                elseif ($unknownStateCounts.Count -lt 20) { $unknownStateCounts[$stateKey] = 1 }
                else { $unknownStateOverflow++ }
            }
            $installStateDetail = [string](Get-SmartM365EvidenceProperty $row @('InstallStateDetail'))
            if ([string]::IsNullOrWhiteSpace($installStateDetail)) { $installStateDetail = [string](Get-SmartM365EvidenceProperty $row @('AppInstallStateDetails')) }
            $errorCode = [string](Get-SmartM365EvidenceProperty $row @('ErrorCode'))
            if ([string]::IsNullOrWhiteSpace($errorCode)) { $errorCode = [string](Get-SmartM365EvidenceProperty $row @('HexErrorCode')) }
            $states.Add([pscustomobject][ordered]@{
                TenantKey=$tenantKey; OrganizationKey=$organizationKey; EnvironmentKey=$environmentKey; TenantId=$tenantId
                ApplicationId=$rowAppId; DeviceId=[string](Get-SmartM365EvidenceProperty $row @('DeviceId'))
                UserPrincipalName=[string](Get-SmartM365EvidenceProperty $row @('UserPrincipalName'))
                AppVersion=[string](Get-SmartM365EvidenceProperty $row @('AppVersion'))
                Platform=[string](Get-SmartM365EvidenceProperty $row @('Platform'))
                InstallState=$installState
                ReportedDeploymentState=$reportedState
                InstallStateDetail=$installStateDetail
                ErrorCode=$errorCode
                LastModifiedDateTime=[string](Get-SmartM365EvidenceProperty $row @('LastModifiedDateTime'))
                RunId=$RunId; CollectedAtUtc=$CollectedAtUtc
            })
        }
        if ((($index+1) % 10 -eq 0) -or ($index+1 -eq $Apps.Count)) {
            Write-SmartM365EvidenceLog "Managed app installation reports completed: $($index+1)/$($Apps.Count); deviceStateRows=$($states.Count)."
        }
    }
    if ($unknownStateRows -gt 0) {
        $examples = @($unknownStateCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 | ForEach-Object { "$($_.Key): $($_.Value)" }) -join '; '
        Write-SmartM365EvidenceLog "Unmapped managed-app installation states: rows=$unknownStateRows; values=$examples; otherRows=$unknownStateOverflow." -Level WARNING
    }
    return $states.ToArray()
}

try {
    $smartRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
    . (Join-Path $smartRoot 'Config/SmartM365-TenantContext.ps1')
    $effective = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.108' -Global -ErrorAction Stop
    Import-Module (Join-Path $smartRoot 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.15' -Force
    $script:Runtime = Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effective `
        -DefaultOutputRelativePath 'Intune/WorkplaceScope/ManagedApps' -OutputPath $OutputPath `
        -LatestCsvFolderPath $LatestCsvFolderPath -ValidateOnly:$ValidateOnly `
        -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath `
        -ReadOnly:($ValidateOnly -or $MaxItems -gt 0)
    # DATA-LAST changes only when this run completes; a failed or running run leaves the previous set intact.
    Enable-SmartM365StagedSourcePublication
    Enable-SmartM365DailyCaptureRetention -OutputRoot $script:Runtime.OutputPath
    Connect-SmartM365EvidenceGraph $script:Runtime @('DeviceManagementApps.Read.All','DeviceManagementManagedDevices.ReadWrite.All') -InteractiveAuth:$InteractiveAuth
    Invoke-SmartM365Preflight -ScriptName 'SmartM365-IntuneManagedApps-Inventory' `
        -OutputPaths @($script:Runtime.OutputPath,$script:Runtime.LatestCsvFolderPath) `
        -RequiredGraphApplicationPermissions @('DeviceManagementApps.Read.All','DeviceManagementManagedDevices.ReadWrite.All') `
        -GraphProbeUris @('https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$top=1') | Out-Null
    if ($ValidateOnly) {
        Write-SmartM365EvidenceLog 'Managed app prerequisites validated. Report availability requires a collection run.'
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }

    $base = 'https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps'
    $nativeApps = @(Get-SmartM365EvidenceGraphCollection -Uri $base -MaxItems $MaxItems)
    $apps = [Collections.Generic.List[object]]::new()
    $assignments = [Collections.Generic.List[object]]::new()
    $seenApps = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($app in $nativeApps) {
        $id = [string](Get-SmartM365EvidenceProperty $app @('id'))
        if (-not $id -or -not $seenApps.Add($id)) { throw 'Managed app catalog contains a missing or duplicate app id.' }
        $apps.Add([pscustomobject][ordered]@{
            ApplicationId=$id
            DisplayName=[string](Get-SmartM365EvidenceProperty $app @('displayName'))
            Publisher=[string](Get-SmartM365EvidenceProperty $app @('publisher'))
            AppType=[string](Get-SmartM365EvidenceProperty $app @('@odata.type'))
            DisplayVersion=[string](Get-SmartM365EvidenceProperty $app @('displayVersion','version'))
            CreatedDateTime=[string](Get-SmartM365EvidenceProperty $app @('createdDateTime'))
            LastModifiedDateTime=[string](Get-SmartM365EvidenceProperty $app @('lastModifiedDateTime'))
            RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
        })
        $encoded = [uri]::EscapeDataString($id)
        $nativeAssignments = @(Get-SmartM365EvidenceGraphCollection -Uri "$base/$encoded/assignments")
        $seenAssignments = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($assignment in $nativeAssignments) {
            $assignmentId = [string](Get-SmartM365EvidenceProperty $assignment @('id'))
            $target = Get-SmartM365EvidenceProperty $assignment @('target') $null
            if (-not $assignmentId -or -not $seenAssignments.Add($assignmentId) -or $null -eq $target) { throw "Invalid assignment for app $id." }
            $assignments.Add([pscustomobject][ordered]@{
                ApplicationId=$id; AssignmentId=$assignmentId
                Intent=[string](Get-SmartM365EvidenceProperty $assignment @('intent'))
                TargetType=[string](Get-SmartM365EvidenceProperty $target @('@odata.type'))
                GroupId=[string](Get-SmartM365EvidenceProperty $target @('groupId'))
                AssignmentFilterId=[string](Get-SmartM365EvidenceProperty $target @('deviceAndAppManagementAssignmentFilterId'))
                AssignmentFilterType=[string](Get-SmartM365EvidenceProperty $target @('deviceAndAppManagementAssignmentFilterType'))
                RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
            })
        }
    }
    $assignedApps = @(Get-IntuneAssignedManagedApps -Apps $apps.ToArray() -Assignments $assignments.ToArray())
    $assignedIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($app in $assignedApps) { [void]$assignedIds.Add([string]$app.ApplicationId) }
    Write-SmartM365EvidenceLog "Managed app catalog collected: apps=$($apps.Count); assignments=$($assignments.Count); assignedApps=$($assignedApps.Count); unassignedApps=$($apps.Count-$assignedApps.Count). Starting installation reports for assigned apps."
    # Core resolves and checks the tenant identity once, with the same rules as its CSV enrichment.
    $tenantIdentity = Add-SmartM365TenantKey -InputObject ([pscustomobject]@{})
    $deviceStates = @(Get-IntuneManagedAppInstallStates -Apps $assignedApps -RunId $script:Runtime.RunId -CollectedAtUtc $script:Runtime.CollectedAtUtc -Identity $tenantIdentity -ExportTimeoutSeconds $ExportTimeoutSeconds -MaxItems $MaxItems)
    $statusByApp = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($app in $apps) {
        $statusByApp[$app.ApplicationId] = [ordered]@{ ReportedInstalledRows=0; ReportedFailureRows=0; ReportedNotInstalledRows=0; PendingRows=0; NotApplicableRows=0; UnknownRows=0; ObservedRows=0; InCollectedCatalog=$true }
    }
    foreach ($state in $deviceStates) {
        if (-not $statusByApp.ContainsKey($state.ApplicationId)) {
            $statusByApp[$state.ApplicationId] = [ordered]@{ ReportedInstalledRows=0; ReportedFailureRows=0; ReportedNotInstalledRows=0; PendingRows=0; NotApplicableRows=0; UnknownRows=0; ObservedRows=0; InCollectedCatalog=$false }
        }
        $counts = $statusByApp[$state.ApplicationId]
        $counts.ObservedRows++
        switch ($state.ReportedDeploymentState) {
            'ReportedInstalled' { $counts.ReportedInstalledRows++ }
            'ReportedFailure' { $counts.ReportedFailureRows++ }
            'ReportedNotInstalled' { $counts.ReportedNotInstalledRows++ }
            'Pending' { $counts.PendingRows++ }
            'NotApplicable' { $counts.NotApplicableRows++ }
            default { $counts.UnknownRows++ }
        }
    }
    $aggregate = @($statusByApp.Keys | ForEach-Object {
        $counts = $statusByApp[$_]
        [pscustomobject][ordered]@{
            ApplicationId=$_
            InCollectedCatalog=$counts.InCollectedCatalog
            StatusSource=Get-IntuneManagedAppStatusSource -ApplicationId $_ -ObservedRows $counts.ObservedRows -AssignedIds $assignedIds
            ObservedRows=$counts.ObservedRows
            ReportedInstalledRows=$counts.ReportedInstalledRows
            ReportedFailureRows=$counts.ReportedFailureRows
            ReportedNotInstalledRows=$counts.ReportedNotInstalledRows
            PendingRows=$counts.PendingRows
            NotApplicableRows=$counts.NotApplicableRows
            UnknownRows=$counts.UnknownRows
            RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
        }
    })
    if ($deviceStates.Count -gt 0 -and @($deviceStates | Where-Object { $_.ReportedDeploymentState -eq 'Unknown' }).Count -eq $deviceStates.Count) {
        throw 'Every managed-app device state is Unknown; source install-state values require qualification before publication.'
    }
    if ($MaxItems -gt 0) {
        Write-SmartM365EvidenceLog "Sample complete: apps=$($apps.Count); assignments=$($assignments.Count); assignedApps=$($assignedApps.Count); statusSummary=$($aggregate.Count); deviceStateRows=$($deviceStates.Count). No canonical CSV published."
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    $rules = $global:SmartM365CsvValidationRules
    if ($rules -isnot [hashtable]) { throw 'Core CSV validation rules are unavailable.' }
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedApps' -CriticalFields @('ApplicationId','RunId','CollectedAtUtc') -UniqueFields @('ApplicationId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedAppAssignments' -CriticalFields @('ApplicationId','AssignmentId','RunId','CollectedAtUtc') -UniqueFields @('ApplicationId','AssignmentId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedAppInstallSummary' -CriticalFields @('ApplicationId','StatusSource','RunId','CollectedAtUtc') -UniqueFields @('ApplicationId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedAppDeviceStates' -CriticalFields @('ApplicationId','RunId','CollectedAtUtc') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    $exports = @(
        @{Name='Intune_ManagedApps';Rows=$apps.ToArray();Columns=@('ApplicationId','DisplayName','Publisher','AppType','DisplayVersion','CreatedDateTime','LastModifiedDateTime','RunId','CollectedAtUtc')},
        @{Name='Intune_ManagedAppAssignments';Rows=$assignments.ToArray();Columns=@('ApplicationId','AssignmentId','Intent','TargetType','GroupId','AssignmentFilterId','AssignmentFilterType','RunId','CollectedAtUtc')},
        @{Name='Intune_ManagedAppInstallSummary';Rows=$aggregate;Columns=@('ApplicationId','InCollectedCatalog','StatusSource','ObservedRows','ReportedInstalledRows','ReportedFailureRows','ReportedNotInstalledRows','PendingRows','NotApplicableRows','UnknownRows','RunId','CollectedAtUtc')},
        @{Name='Intune_ManagedAppDeviceStates';Rows=$deviceStates;Options=@{NoTenantKey=$true;LocalSerialization=$true};Columns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','ApplicationId','DeviceId','UserPrincipalName','AppVersion','Platform','InstallState','ReportedDeploymentState','InstallStateDetail','ErrorCode','LastModifiedDateTime','RunId','CollectedAtUtc')}
    )
    foreach ($export in $exports) { Assert-SmartM365CsvDataCompleteness -Data $export.Rows -BaseFileName $export.Name -Columns $export.Columns | Out-Null }
    $deviceStateCount = $deviceStates.Count
    $deviceStates = $null
    foreach ($export in $exports) {
        $options = if ($export.ContainsKey('Options')) { $export.Options } else { @{} }
        Export-SmartM365EvidenceDataset $script:Runtime $export.Name $export.Rows $export.Columns -SingleSerialization @options | Out-Null
        $export.Rows = $null
    }
    Write-SmartM365EvidenceLog "Managed apps collected: apps=$($apps.Count); assignments=$($assignments.Count); assignedApps=$($assignedApps.Count); unassignedApps=$($apps.Count-$assignedApps.Count); statusSummary=$($aggregate.Count); deviceStateRows=$deviceStateCount. Summary counts are report rows, not unique devices." -Level SUCCESS

    $smartM365RunGuard.ReachedTerminal=$true
} catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_
 $failure = $_; throw } finally {
    $smartM365RunGuard.InCleanup=$true

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false } catch {}
    if ($script:Runtime) { Complete-SmartM365EvidenceRuntime -Status $(if ($failure) {'Failed'} else {'Success'}) -ErrorRecord $failure -FailureStage 'IntuneManagedAppsInventory' }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBBK5jJfl6JqrMG
# oeL9yddMvdYcf1imPahC/jgTkFR+RaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKLXsNAMi+9rfiWK5LTKcS4C8hsyUwwcxZKwFbPoAD4gMA0GCSqG
# SIb3DQEBAQUABIIBgEngi01h1XFD1L72XPo4VL2Cy+wL6C4HTUkSGSl0MAwURu/m
# LhmPBn4NVfng9tffXtWoEfGR1uqx1dtl82pwt9rdttX05rQdTo3BgzlvNAhyUA9G
# j1UOv5D5iwe09vk8vfkDtFcxZmh0EGbSsH4xXZdppcA5UGW6z5z5B6oUOizmMCsG
# oMzF9R3rrfgEzKz0BecV/DARKn+TQTrlzU5K7P+Ub4/Aeu94DiICPU8Mxyh515tZ
# sPK78MIyGas8Paz9OUs/Xd6kRrUBL+84AL//ginb1SyJVM4N3SU3juALXEVqUPcY
# Ou1eDkk77M1rX2DBSztZ06XyMQNHFuj51s0LhawWNvL4gTPi3Qts2WwXNtK/Lc+U
# pE6ewn37FCla/cvsM2jtX3ace17uoChbo84PYiuCEu9o+9YsS/1odzm3AyikS/FG
# h26KHZh3Hee7g1lXTxiXyD5R6XsntpunoYlJCl1u7whvrBJ2VHo46/wbKhJ9GjcE
# GJD5VdB5+sc3n0JQ4qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAwNjIx
# NDVaMC8GCSqGSIb3DQEJBDEiBCBBI7UHFY0OnW4U3/Xdnqf9Y+QxBgNiTyebuC9p
# OLR0JjANBgkqhkiG9w0BAQEFAASCAgBmzSDB/5beFE2DAbvVU6MZisc8KHmj9+Ae
# hYF+xqd7z6UfwR+AFKDxc5J2umixxo+T3YQ/dwXpoJ6+0710cntYBZvZK0QPrI2P
# WuPWxQ6BzFqUBas8/4rlIfMd/oAD4LLiGZuoUzkvizH6j4Lk/rOy35pyhQYgKhkK
# n2eX3vxnUEu9wtrqMYL20uWDVe8XGr53zByOsYo0819GIeZpkALaf9OFaAqfGJ1w
# 6cd0S1jEMDql02LQjf2tU4QNGve+7AN3X2rzHlyykaRzo74IDbTuMbRcs2042wiO
# W0R9A2q426iFe48sAG4oAN3rqbet8FUNMD0dxjFitFlmxoppmm1/hSzBcDhqG14E
# AMAAwAyjoIzQS2EsSumBGaHB10GadinzhxLCShgGQpiyTvl/5v/3uySqu7tUuZo1
# fuB+3ggd1feIXrTX9KBt9oYjPMlQnuP9sVXK6I7grCMIphZmD+4cSe9kDX2U1vVe
# USASlSujOvx7QmhUS/0K2ZhUuha+fEV5zCJyeBHq+DYeiNvhBRJCiPcBnJjj50Mp
# 9ok/NIi4iOCtzUNO6rwXRXL/UT54uK2PQ8qcjmXJSzuUI/hQcSl9DXK7gXMiaIYT
# B9vVhyTAXKrfYqObTBeVVXesZZSg4m6Ywh3/RYcB8J7sUM2WMcLkqO3BBkjGWq27
# H5m9SMHskQ==
# SIG # End signature block
