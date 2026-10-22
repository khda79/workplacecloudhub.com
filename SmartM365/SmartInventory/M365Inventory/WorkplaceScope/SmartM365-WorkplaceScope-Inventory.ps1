#Requires -Version 7.0
<#
.SYNOPSIS
Collect native Entra direct membership and Intune policy/assignment evidence.
.VERSION
1.0.12
.NOTES
Candidate collector. Not automatically scheduled until production qualification.
Graph beta membership avoids the documented v1.0 service-principal omission.
Hidden groups require Member.Read.Hidden; missing access is fatal, not empty.
#>
[CmdletBinding()]
param([string]$Tenant='test', [switch]$InteractiveAuth, [switch]$ValidateOnly,
    [switch]$EnableConfiguredExternalActions, [string]$OutputPath, [string]$LatestCsvFolderPath,
    [ValidateRange(0,10000000)][int]$MaxItems=0)
function ConvertTo-InventoryIsoDate {
    param([AllowNull()][object]$Value)
    # Round-trip preserves DateTime.Kind/DateTimeOffset.Offset. Unspecified stays offsetless.
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) {
        return $Value.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    # Do not guess a culture or timezone for string-valued/deserialized evidence.
    return ''
}

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
$script:Runtime = $null; $failure = $null; $transcriptStarted = $false
try {
    $smartRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
    . (Join-Path $smartRoot 'Config/SmartM365-TenantContext.ps1')
    $effective = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.92' -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot '../../Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.11' -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot '../../Common/SmartM365.WorkplaceSource.psd1') -MinimumVersion '1.0.0' -ErrorAction Stop
    $script:Runtime = Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effective `
        -DefaultOutputRelativePath 'M365/WorkplaceScope' -OutputPath $OutputPath -LatestCsvFolderPath $LatestCsvFolderPath `
        -ValidateOnly:$ValidateOnly -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365CmdbSourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath -ReadOnly:($ValidateOnly -or $MaxItems -gt 0)
    Enable-SmartM365DailyCaptureRetention -OutputRoot $script:Runtime.OutputPath
    Start-SmartM365RunTranscript -Path $global:logTranscriptFile -Append | Out-Null
    $transcriptStarted = $true
    Initialize-SmartM365ImportExcel | Out-Null
    foreach ($module in @('Microsoft.Graph.Beta.Groups','Microsoft.Graph.Beta.DeviceManagement')) {
        if (-not (Get-Module -ListAvailable $module)) { throw "Required SDK module missing. Install-Module $module -Scope CurrentUser" }
        Import-Module $module -ErrorAction Stop
    }
    Connect-SmartM365EvidenceGraph $script:Runtime @('Group.Read.All','DeviceManagementConfiguration.Read.All') -InteractiveAuth:$InteractiveAuth
    Invoke-SmartM365Preflight -ScriptName 'SmartM365-WorkplaceScope-Inventory' -OutputPaths @($script:Runtime.OutputPath) `
        -RequiredGraphApplicationPermissions @('Group.Read.All','DeviceManagementConfiguration.Read.All') `
        -GraphProbeUris @('https://graph.microsoft.com/beta/groups?$top=1','https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$top=1') | Out-Null
    if ($ValidateOnly) { Write-SmartM365EvidenceLog 'Workplace scope prerequisites validated; no full collection or CSV publication.'; $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return }
    if ($MaxItems -gt 0) {
        $global:SmartM365MaxItems=$MaxItems; $global:SmartM365TestMaxItems=$MaxItems; $global:SmartM365IsMaxItemsRun=$true
    }
    $memberships = [Collections.Generic.List[object]]::new()
    $groupScope = [Collections.Generic.List[object]]::new()
    $absentGroups = [Collections.Generic.List[string]]::new()
    function Test-WorkplaceGroupNotFound {
        param([Parameter(Mandatory)]$Record)
        $exception = $Record.Exception
        while ($null -ne $exception) {
            foreach ($name in @('ResponseStatusCode','StatusCode')) {
                $status = Get-WorkplaceSourceValue $exception @($name) $null
                if ($null -ne $status) { return ([int]$status -eq 404) }
            }
            $response = Get-WorkplaceSourceValue $exception @('Response') $null
            $status = Get-WorkplaceSourceValue $response @('StatusCode') $null
            if ($null -ne $status) { return ([int]$status -eq 404) }
            $exception = $exception.InnerException
        }
        # SDK error code is accepted, never arbitrary digits in a message or ID.
        return ([string]$Record.Exception.Message -match '\[Request_ResourceNotFound\]')
    }
    function Get-WorkplaceDirectMembers {
        param([Parameter(Mandatory)][string]$GroupId)
        $onlyNotFound = $true
        for ($attempt=1; $attempt -le 3; $attempt++) {
            try {
                # Buffer each complete traversal: discard pages from a failed attempt.
                $rows = @(Get-MgBetaGroupMember -GroupId $GroupId -All -ErrorAction Stop)
                return [pscustomobject]@{Absent=$false;Members=$rows}
            } catch {
                $notFound = Test-WorkplaceGroupNotFound $_
                $pageToken = Test-WorkplaceInvalidPageToken $_
                if (-not $notFound -and -not $pageToken) {
                    throw [InvalidOperationException]::new("Direct membership collection failed. GroupId=$GroupId; attempt=$attempt/3; $($_.Exception.Message)", $_.Exception)
                }
                if ($pageToken) { $onlyNotFound = $false }
                if ($attempt -lt 3) {
                    $delay = @(5,15)[$attempt-1]
                    $reason = if ($pageToken) { 'InvalidPageToken' } else { 'NotFound (404)' }
                    Write-SmartM365EvidenceLog "Group membership retry. GroupId=$GroupId; reason=$reason; attempt=$attempt/3; restarting complete traversal after ${delay}s."
                    Start-Sleep -Seconds $delay
                }
            }
        }
        # A pagination failure, even mixed with later 404s, never proves absence.
        if (-not $onlyNotFound) {
            throw "Direct membership incomplete after three traversals. GroupId=$GroupId; invalid page token encountered; no group exclusion permitted."
        }
        for ($probe=1; $probe -le 2; $probe++) {
            try {
                $parent = @(Get-MgBetaGroup -GroupId $GroupId -Property id -ErrorAction Stop)
            } catch {
                if (-not (Test-WorkplaceGroupNotFound $_)) { throw }
                if ($probe -eq 1) { Start-Sleep -Seconds 15 }
                continue
            }
            if ($parent.Count -ne 1 -or [string]$parent[0].Id -ne $GroupId) {
                throw "Ambiguous group existence response after membership 404: $GroupId"
            }
            throw "Group still exists but direct membership could not be collected: $GroupId"
        }
        return [pscustomobject]@{Absent=$true;Members=@()}
    }
    function Test-WorkplaceInvalidPageToken {
        param([Parameter(Mandatory)]$Record)
        $exception = $Record.Exception
        $badRequest = $false
        $messages = [Collections.Generic.List[string]]::new()
        while ($null -ne $exception) {
            $messages.Add([string]$exception.Message)
            foreach ($name in @('ResponseStatusCode','StatusCode')) {
                $status = Get-WorkplaceSourceValue $exception @($name) $null
                if ($null -ne $status) {
                    if ([int]$status -ne 400) { return $false }
                    $badRequest = $true
                }
            }
            $response = Get-WorkplaceSourceValue $exception @('Response') $null
            $status = Get-WorkplaceSourceValue $response @('StatusCode') $null
            if ($null -ne $status) {
                if ([int]$status -ne 400) { return $false }
                $badRequest = $true
            }
            $exception = $exception.InnerException
        }
        $text = $messages -join "`n"
        $directoryToken = $text -match '\[DirectoryPageTokenNotFound(?:Exception)?\]'
        $badRequest = $badRequest -or ($text -match '\[Request_BadRequest\]') -or $directoryToken
        return ($badRequest -and ($directoryToken -or $text -match '\bThe page token is not valid\b'))
    }
    # Catalog properties and membership scope share this single enumeration.
    $groups = @(Get-MgBetaGroup -All -Property id,visibility,displayName,mailEnabled,securityEnabled,groupTypes,onPremisesSecurityIdentifier -ErrorAction Stop)
    $groupCatalogCollectedAtUtc = [datetime]::UtcNow.ToString('o')
    if ($MaxItems -gt 0) { $groups = @($groups | Select-Object -First $MaxItems) }
    if (@($groups | Where-Object Visibility -eq 'HiddenMembership').Count) {
        Invoke-SmartM365Preflight -ScriptName 'SmartM365-WorkplaceScope-Inventory' -RequiredGraphApplicationPermissions @('Member.Read.Hidden') | Out-Null
    }
    Write-SmartM365EvidenceLog "Reading direct memberships for $($groups.Count) Entra groups. This can be a long-running inventory."
    # Membership rows are validated while they are collected, with the publication rule of
    # M365_EntraGroupMemberships_All: non-empty GroupId/MemberId and unique (GroupId, MemberId) after trim and
    # lowercase. Distinct groups and distinct members per group make every pair distinct without a global set.
    $membershipTimer = [Diagnostics.Stopwatch]::StartNew()
    $seenGroups = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($group in $groups) {
        if ([string]::IsNullOrWhiteSpace([string]$group.Id) -or -not $seenGroups.Add(([string]$group.Id).Trim().ToLowerInvariant())) { throw 'Missing or duplicate Entra group identity.' }
        $memberRead = Get-WorkplaceDirectMembers -GroupId $group.Id
        if ($memberRead.Absent) { $absentGroups.Add([string]$group.Id); continue }
        $members = @($memberRead.Members)
        $collected = [datetime]::UtcNow.ToString('o')
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($member in $members) {
            if ([string]::IsNullOrWhiteSpace([string]$member.Id) -or -not $seen.Add(([string]$member.Id).Trim().ToLowerInvariant())) { throw 'Missing or duplicate Entra member identity.' }
            $memberships.Add([pscustomobject][ordered]@{
                GroupId=$group.Id; MemberId=$member.Id
                MemberType=Get-WorkplaceSourceValue $member @('@odata.type','OdataType')
                MembershipKind='Direct'; CollectionStatus='Collected'
                RunId=$script:Runtime.RunId; CollectedAtUtc=$collected
            })
        }
        $groupScope.Add([pscustomobject][ordered]@{
            GroupId=$group.Id; Visibility=$group.Visibility; MemberCount=$seen.Count
            MemberCollectionStatus='Collected'; RunId=$script:Runtime.RunId; CollectedAtUtc=$collected
            DisplayName=$group.DisplayName; MailEnabled=$group.MailEnabled; SecurityEnabled=$group.SecurityEnabled
            GroupTypes=(@($group.GroupTypes) -join ';'); OnPremisesSecurityIdentifier=$group.OnPremisesSecurityIdentifier
            GroupCollectedAtUtc=$groupCatalogCollectedAtUtc
        })
    }
    if ($memberships.Count -gt 0 -and [string]::IsNullOrWhiteSpace([string]$script:Runtime.RunId)) { throw 'Missing RunId on Entra membership rows.' }
    Write-SmartM365EvidenceLog ("Direct membership collection completed in {0:N1} min: {1} membership row(s) validated while collected." -f $membershipTimer.Elapsed.TotalMinutes, $memberships.Count)
    if ($absentGroups.Count) {
        # Corroborate targeted 404s with a successful fresh full enumeration.
        $freshIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($currentGroup in @(Get-MgBetaGroup -All -Property id -ErrorAction Stop)) {
            if (-not $currentGroup.Id -or -not $freshIds.Add([string]$currentGroup.Id)) { throw 'Missing or duplicate identity in group absence verification.' }
        }
        foreach ($id in $absentGroups) {
            if ($freshIds.Contains($id)) { throw "Group reappeared during absence verification; membership remains unqualified: $id" }
            Write-SmartM365EvidenceLog "Group excluded from catalog and memberships after three membership 404s, two parent 404s and absence from a fresh full catalog. GroupId=$id; absence corroborated, deletion cause not proven." -Level WARNING
        }
    }
    $policies = [Collections.Generic.List[object]]::new()
    $assignments = [Collections.Generic.List[object]]::new()
    $families = @(
        @{Name='SettingsCatalog';Command='Get-MgBetaDeviceManagementConfigurationPolicy';Path='configurationPolicies'},
        @{Name='DeviceConfiguration';Command='Get-MgBetaDeviceManagementDeviceConfiguration';Path='deviceConfigurations'},
        @{Name='DeviceCompliance';Command='Get-MgBetaDeviceManagementDeviceCompliancePolicy';Path='deviceCompliancePolicies'},
        @{Name='WindowsFeatureUpdate';Command='Get-MgBetaDeviceManagementWindowsFeatureUpdateProfile';Path='windowsFeatureUpdateProfiles'},
        @{Name='WindowsQualityUpdate';Command='Get-MgBetaDeviceManagementWindowsQualityUpdateProfile';Path='windowsQualityUpdateProfiles'}
    )
    foreach ($family in $families) {
        $items = @(& $family.Command -All -ErrorAction Stop)
        if ($MaxItems -gt 0) { $items = @($items | Select-Object -First $MaxItems) }
        $seenPolicies = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($policy in $items) {
            $id = [string](Get-WorkplaceSourceValue $policy @('id'))
            if (-not $id -or -not $seenPolicies.Add($id)) { throw 'Missing or duplicate policy identity.' }
            $uri = 'https://graph.microsoft.com/beta/deviceManagement/'+$family.Path+'/'+[uri]::EscapeDataString($id)+'/assignments'
            $targets = @(Get-WorkplaceGraphCollection -Uri $uri -Invoker { param($address) Invoke-SmartM365EvidenceGraphRequest -Uri $address })
            $collected = [datetime]::UtcNow.ToString('o')
            $policyEvidence = ConvertTo-WorkplacePolicyEvidence $policy $family.Name $script:Runtime.RunId $collected
            foreach ($dateField in @('CreatedDateTime','LastModifiedDateTime')) {
                $policyEvidence | Add-Member -NotePropertyName ($dateField + 'Iso8601') -NotePropertyValue (ConvertTo-InventoryIsoDate $policyEvidence.$dateField)
            }
            $policies.Add($policyEvidence)
            $seenAssignments = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($assignment in $targets) {
                $target = Get-WorkplaceSourceValue $assignment @('target') $null
                $assignmentId = [string](Get-WorkplaceSourceValue $assignment @('id'))
                if (-not $assignmentId -or -not $seenAssignments.Add($assignmentId) -or $null -eq $target) { throw 'Malformed policy assignment.' }
                $assignments.Add([pscustomobject][ordered]@{
                    PolicyFamily=$family.Name; PolicyId=$id; AssignmentId=$assignmentId
                    TargetType=Get-WorkplaceSourceValue $target @('@odata.type')
                    GroupId=Get-WorkplaceSourceValue $target @('groupId')
                    AssignmentFilterId=Get-WorkplaceSourceValue $target @('deviceAndAppManagementAssignmentFilterId')
                    AssignmentFilterType=Get-WorkplaceSourceValue $target @('deviceAndAppManagementAssignmentFilterType')
                    NativeTargetJson=ConvertTo-Json -InputObject $target -Depth 30 -Compress
                    RunId=$script:Runtime.RunId; CollectedAtUtc=$collected
                })
            }
        }
    }
    # No canonical publication until every required family and child collection succeeded.
    $exports = @(
        @{Name='M365_EntraGroupMemberships_All';Rows=$memberships.ToArray();Columns=@('GroupId','MemberId','MemberType','MembershipKind','CollectionStatus','RunId','CollectedAtUtc')},
        @{Name='M365_EntraGroupMembershipScope';Rows=$groupScope.ToArray();Columns=@('GroupId','Visibility','MemberCount','MemberCollectionStatus','RunId','CollectedAtUtc','DisplayName','MailEnabled','SecurityEnabled','GroupTypes','OnPremisesSecurityIdentifier','GroupCollectedAtUtc')},
        @{Name='Intune_Policies_All';Rows=$policies.ToArray();Columns=@('PolicyId','PolicyFamily','DisplayName','Description','Platforms','Technologies','FeatureUpdateVersion','CreatedDateTime','LastModifiedDateTime','NativeEvidenceJson','AssignmentCollectionStatus','RunId','CollectedAtUtc','CreatedDateTimeIso8601','LastModifiedDateTimeIso8601')},
        @{Name='Intune_PolicyAssignments_All';Rows=$assignments.ToArray();Columns=@('PolicyFamily','PolicyId','AssignmentId','TargetType','GroupId','AssignmentFilterId','AssignmentFilterType','NativeTargetJson','RunId','CollectedAtUtc')}
    )
    # Memberships were validated while collected and are exported first: Publish-SmartM365Csv applies the full rule
    # again before any file is written, so a second in-memory pass here is not needed.
    foreach ($export in $exports) {
        if ($export.Name -eq 'M365_EntraGroupMemberships_All') { continue }
        Assert-SmartM365CsvDataCompleteness -Data $export.Rows -BaseFileName $export.Name -Columns $export.Columns | Out-Null
    }
    foreach ($export in $exports) {
        Export-SmartM365EvidenceDataset $script:Runtime $export.Name $export.Rows $export.Columns -NoWeeklyHistory -SingleSerialization | Out-Null
    }
    $qualifications = @('Direct membership only; all five required policy families collected.')
    if ($absentGroups.Count) { $qualifications += "Groups excluded after corroborated absence (not empty membership): $($absentGroups -join ','); membership attempts=3; parent 404s=2; fresh full catalog verified." }
    Set-SmartM365CmdbSourceScope -CompleteScope ($MaxItems -eq 0) -Scope 'CMDB:group_scope,group_members,policies,policy_assignments' -Qualifications $qualifications
    Write-SmartM365EvidenceLog "Workplace scope collected. Groups=$($groupScope.Count); excludedAbsentGroups=$($absentGroups.Count); memberships=$($memberships.Count); policies=$($policies.Count); assignments=$($assignments.Count)." -Level SUCCESS

    $smartM365RunGuard.ReachedTerminal=$true
} catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_
 $failure=$_; throw } finally {
    $smartM365RunGuard.InCleanup=$true

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false } catch {}
    # Completion owns transcript closure so SharePoint receives the terminal summary.
    if ($script:Runtime) { Complete-SmartM365EvidenceRuntime -Status $(if ($failure) {'Failed'} else {'Success'}) -ErrorRecord $failure -FailureStage 'WorkplaceScopeInventory' -CloseTranscriptBeforeUpload:$transcriptStarted }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB57Gzr7y/mySIx
# 8YQRBxzwIHsB1HqJa0fQHpFsDu1OJaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGzpJtdkGxwqJb/ehgTId7N96n61YUDXvdHaGuk4NTEmMA0GCSqG
# SIb3DQEBAQUABIIBgBu/+HSiAi/Unzo3YtHMNZZFOCbN4WDxIUk1evKCKLNp0O0f
# AQ+t5IwAEt6VSyWkNfP0IcakT6rU618HeCvmdACpPtjyN5jpqItl1dL3l9g6ljR/
# PSkOotn5URwCZIcLEXU7k0nsS68ObGpLFd6nFyScK80oPovEcvC5aQblzDr0LO1W
# 1F+csqR9f6mb/vXHOlw63lH4yYguVRzuWvqt5rrsqqrONPj5YwjHnkzphmkVO3On
# 67vy3Vn0AOR1hHvmMy5assLbunE9AheyNQF7mbVEohrmzyZQu3Zk+ZOeA1LdK5S1
# A2Vjy9N439Bl+RtrbeeICjYuT/ccox+7d6rJiKVWidzGccfbTi4u+z1e4YgqHc4K
# 1TC3C4UYqqBK0QGBeFcVemBPUxtBA7FLvrrPwKlko79Rk6JKsfzZFGom7fKYy/GV
# 1ClnsNo9iQ4QJZzcPnltBSbFLpC3JbsCLDZUubEIoG4vo+3M0S0aBkpK+o6e8x2v
# 743gp2V3hMU1G4VBxKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIyMTI2
# MjhaMC8GCSqGSIb3DQEJBDEiBCDYS3MbVVvb+vnvAur5UVh0fG8/LyEymWCueIes
# KF+JNDANBgkqhkiG9w0BAQEFAASCAgCrSgx6CoyxD8sr6YqTwu9gS1/I68iYQs/d
# 96ntxcYOPGpw94xHtf7e3O86Ufho3kEp24jAwZtQ9sR+z/ErzEicy9TBnElgllWU
# knP98UtbLkg8Tb7yyyvHii5VbJHWecOc8A4MJYzpdbxk2YAXurL2Cj8KsRXfFjm7
# 0OPt26S+c0gWXgdLUav9PO4yWERvqfAwAd390t9eufy+wxDs4kJ9SNaoQKy0iZ8G
# NLLIgWVE15CnGQ5PkiztU4WmB1MgalCtf8UvIiiU7+ucyE10R7jtNhJFouGzHTU9
# nCfT++RdRm2aYCuVArPsmYL9obEyQFbFtEofWKhd5t/VbTZJbkb+qcDK+pc2q4Jz
# MBF9338kvP6EBzWrVd63qNXNXr2W96ZUco3j/38n+VuprLOk2IxqJqQUMzNy8Q3U
# kvmXMbqSv9k4LnNTdlTqLsvbCRoIpYR+2JSMKBjvUWDxjnzXsAGadsgktlSKHioL
# 6kqw8wKxDwteAngSwFusm2d1Ts5G9JlUSObJE8VEQuMMJhz66hxt8yzW1QcQQ9ee
# AEdSZ9343uGFw2a5ONa0bk9o1HppIa6+zgIhmcH+B1kpOwVrzRiZu5CJRY1v2vVc
# cCUyRzTiBoFJPne+ui/xtzOgeHOUvMM2D5XhDGl7ZoQlunrNzgKZB/Gm7sioTChB
# q/FhfWJTAA==
# SIG # End signature block
