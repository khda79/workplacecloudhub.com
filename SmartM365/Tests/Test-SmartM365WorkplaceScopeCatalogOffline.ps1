#Requires -Version 7.0
<#
.SYNOPSIS
Offline coherent group catalog/membership acquisition tests. No tenant calls.
.VERSION
1.0.2
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$collectorPath = Join-Path $root 'SmartInventory/M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1'
$collector = Get-Content -LiteralPath $collectorPath -Raw
$start = $collector.IndexOf('    $memberships = ')
$end = $collector.IndexOf('    Write-SmartM365EvidenceLog "Workplace scope collected.', $start)
if ($start -lt 0 -or $end -le $start) { throw 'Collection block boundaries changed; update this offline test.' }
$collectionBlock = [scriptblock]::Create($collector.Substring($start, $end - $start))
$script:checks = 0
function Assert-Catalog([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}
function Invoke-CatalogFixture([object[]]$Groups, [switch]$FailMembers, [switch]$DuplicateMember, [switch]$FailPolicies,
    [int]$Member404Count=0, [string[]]$ParentResults=@('Absent','Absent'), [switch]$Reappears,
    [switch]$FailFinalCatalog, [switch]$DuplicateFinalCatalog, [switch]$CodeOnly404, [int]$MemberStatus=0,
    [ValidateSet('HttpStatus','ResponseStatus','ResponseObject','Inner')][string]$ErrorShape='HttpStatus',
    [string[]]$MemberErrors=@()) {
    $fixture = New-Module -ScriptBlock {
        param($rows, $block, $memberFailure, $duplicate, $policyFailure, $missingCount, $parents,
            $reappears, $finalFailure, $finalDuplicate, $codeOnly, $memberStatus, $errorShape, $memberErrors)
        Set-StrictMode -Version Latest
        $script:Runtime = [pscustomobject]@{RunId='synthetic-runtime-run'}
        $script:Groups = @($rows); $script:Block = $block
        $script:FailMembers = $memberFailure; $script:DuplicateMember = $duplicate
        $script:FailPolicies = $policyFailure
        $script:MissingCount=$missingCount; $script:Parents=@($parents); $script:Reappears=$reappears
        $script:FinalFailure=$finalFailure; $script:FinalDuplicate=$finalDuplicate
        $script:CodeOnly=$codeOnly; $script:MemberStatus=$memberStatus; $script:ErrorShape=$errorShape; $script:FirstMemberCalls=0
        $script:MemberErrors=@($memberErrors); $script:CallsByGroup=@{}
        $script:Exports = @{}; $script:GroupCalls = 0; $script:MemberCalls = 0
        $script:ParentCalls=0; $script:Sleeps=[Collections.Generic.List[int]]::new()
        $script:Logs=[Collections.Generic.List[object]]::new(); $script:Qualifications=@()
        $script:Properties = @(); $script:Scope = ''; $script:ErrorText = ''
        function Write-Synthetic404 {
            if ($script:CodeOnly) { throw '[Request_ResourceNotFound] : Synthetic unavailable resource.' }
            if ($script:ErrorShape -ne 'HttpStatus') {
                $error404=[InvalidOperationException]::new('Synthetic structured unavailable resource.')
                switch ($script:ErrorShape) {
                    'ResponseStatus' { $error404 | Add-Member -NotePropertyName ResponseStatusCode -NotePropertyValue 404 }
                    'ResponseObject' { $error404 | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{StatusCode=404}) }
                    'Inner' { $error404=[InvalidOperationException]::new('Synthetic wrapped unavailable resource.',[Net.Http.HttpRequestException]::new('Synthetic inner 404.',$null,[Net.HttpStatusCode]::NotFound)) }
                }
                throw $error404
            }
            throw [Net.Http.HttpRequestException]::new('Synthetic unavailable resource.', $null, [Net.HttpStatusCode]::NotFound)
        }
        function Get-MgBetaGroup {
            [CmdletBinding()] param([switch]$All, [string[]]$Property, [string]$GroupId)
            if ($GroupId) {
                $result=$script:Parents[$script:ParentCalls]; $script:ParentCalls++
                switch ($result) {
                    'Absent' { Write-Synthetic404 }
                    'Exists' { [pscustomobject]@{Id=$GroupId} }
                    'Forbidden' { throw [Net.Http.HttpRequestException]::new('Synthetic forbidden.', $null, [Net.HttpStatusCode]::Forbidden) }
                    'Unavailable' { throw 'Synthetic parent failure' }
                    'WrongId' { [pscustomobject]@{Id='other-group'} }
                    'Empty' { }
                    default { throw 'Unexpected parent fixture.' }
                }
                return
            }
            if (-not $All) { throw 'Full group enumeration required.' }
            $script:GroupCalls++
            if ($script:GroupCalls -eq 1) { $script:Properties = @($Property); $script:Groups; return }
            if ($script:FinalFailure) { throw 'Synthetic final catalog failure' }
            if ($script:FinalDuplicate) { $script:Groups; $script:Groups; return }
            $script:Groups | Where-Object { $_.Id -ne 'group-one' -or $script:Reappears }
        }
        function Get-MgBetaGroupMember {
            [CmdletBinding()] param([string]$GroupId, [switch]$All)
            if (-not $All) { throw 'Full membership enumeration required.' }
            $script:MemberCalls++
            if (-not $script:CallsByGroup.ContainsKey($GroupId)) { $script:CallsByGroup[$GroupId]=0 }
            $script:CallsByGroup[$GroupId]++
            if ($script:FailMembers) { throw 'Synthetic membership failure' }
            if ($GroupId -eq 'group-one') {
                $script:FirstMemberCalls++
                if ($script:FirstMemberCalls -le $script:MemberErrors.Count) {
                    $kind=$script:MemberErrors[$script:FirstMemberCalls-1]
                    # Simulate pages emitted before the SDK encounters a bad nextLink.
                    if ($kind -ne 'Success') { [pscustomobject]@{Id='discarded-pagination-member';OdataType='#microsoft.graph.user'} }
                    switch ($kind) {
                        'InvalidPageToken' { throw [Net.Http.HttpRequestException]::new('The page token is not valid.', $null, [Net.HttpStatusCode]::BadRequest) }
                        'CodeOnlyPageToken' { throw '[Request_BadRequest] : The page token is not valid.' }
                        'DirectoryPageToken' { throw '[DirectoryPageTokenNotFoundException] : Synthetic paging token failure.' }
                        'WrappedPageToken' { throw [InvalidOperationException]::new('Synthetic SDK wrapper.',[Net.Http.HttpRequestException]::new('The page token is not valid.',$null,[Net.HttpStatusCode]::BadRequest)) }
                        'GenericBadRequest' { throw [Net.Http.HttpRequestException]::new('Synthetic invalid query.', $null, [Net.HttpStatusCode]::BadRequest) }
                        'ForbiddenPageToken' { throw [Net.Http.HttpRequestException]::new('[Request_BadRequest] : The page token is not valid.', $null, [Net.HttpStatusCode]::Forbidden) }
                        'UnknownPageToken' { throw 'The page token is not valid.' }
                        'NotFound' { Write-Synthetic404 }
                        'Success' { }
                        default { throw "Unknown error fixture: $kind" }
                    }
                }
                if ($script:MemberStatus) { throw [Net.Http.HttpRequestException]::new('Synthetic member error with 404-looking ID.', $null, [Net.HttpStatusCode]$script:MemberStatus) }
                if ($script:FirstMemberCalls -le $script:MissingCount) {
                    [pscustomobject]@{Id='discarded-partial-member';OdataType='#microsoft.graph.user'}
                    Write-Synthetic404
                }
                [pscustomobject]@{Id='user-one';OdataType='#microsoft.graph.user'}
                [pscustomobject]@{Id=$(if ($script:DuplicateMember) {'user-one'} else {'principal-one'});OdataType='#microsoft.graph.servicePrincipal'}
            }
        }
        function Get-WorkplaceSourceValue {
            param($Row, [string[]]$Names, $Default='')
            if ($null -eq $Row) { return $Default }
            foreach ($name in $Names) {
                $property = $Row.PSObject.Properties[$name]
                if ($null -ne $property) { return $property.Value }
            }
            return $Default
        }
        function Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
        function Write-SmartM365EvidenceLog { param($Message, $Level) $script:Logs.Add([pscustomobject]@{Message=$Message;Level=$Level}) }
        function Invoke-SmartM365Preflight { throw 'Unexpected preflight in synthetic visible group collection.' }
        function Get-MgBetaDeviceManagementConfigurationPolicy {
            [CmdletBinding()] param([switch]$All)
            if ($script:FailPolicies) { throw 'Synthetic policy failure' }
        }
        function Get-MgBetaDeviceManagementDeviceConfiguration { [CmdletBinding()] param([switch]$All) }
        function Get-MgBetaDeviceManagementDeviceCompliancePolicy { [CmdletBinding()] param([switch]$All) }
        function Get-MgBetaDeviceManagementWindowsFeatureUpdateProfile { [CmdletBinding()] param([switch]$All) }
        function Get-MgBetaDeviceManagementWindowsQualityUpdateProfile { [CmdletBinding()] param([switch]$All) }
        function Assert-SmartM365CsvDataCompleteness { param($Data, $BaseFileName, $Columns) }
        function Export-SmartM365EvidenceDataset {
            param($Runtime, $Name, $Rows, $Columns, [switch]$NoWeeklyHistory, [switch]$SingleSerialization)
            if (-not $NoWeeklyHistory -or -not $SingleSerialization) { throw 'Export optimization/history contract changed.' }
            $script:Exports[$Name] = [pscustomobject]@{Rows=@($Rows);Columns=@($Columns)}
        }
        function Set-SmartM365CmdbSourceScope { param($CompleteScope, $Scope, $Qualifications) $script:Scope=$Scope; $script:Qualifications=@($Qualifications) }
        function Invoke-SyntheticCatalog {
            $MaxItems = 0
            try { & $script:Block } catch { $script:ErrorText = $_.Exception.Message }
            [pscustomobject]@{Exports=$script:Exports;GroupCalls=$script:GroupCalls;MemberCalls=$script:MemberCalls;
                             Properties=$script:Properties;Scope=$script:Scope;ErrorText=$script:ErrorText;
                             ParentCalls=$script:ParentCalls;Sleeps=$script:Sleeps.ToArray();Logs=$script:Logs.ToArray();Qualifications=$script:Qualifications;
                             CallsByGroup=$script:CallsByGroup}
        }
    } -ArgumentList @($Groups, $collectionBlock, [bool]$FailMembers, [bool]$DuplicateMember, [bool]$FailPolicies,
        $Member404Count, $ParentResults, [bool]$Reappears, [bool]$FailFinalCatalog, [bool]$DuplicateFinalCatalog, [bool]$CodeOnly404, $MemberStatus, $ErrorShape, $MemberErrors)
    try { & $fixture { Invoke-SyntheticCatalog } }
    finally { Remove-Module $fixture -ErrorAction SilentlyContinue }
}
$group = [pscustomobject]@{Id='group-one';Visibility='Private';DisplayName="Group, quoted`nsecond line";
    MailEnabled=$false;SecurityEnabled=$true;GroupTypes=@('DynamicMembership','Unified');OnPremisesSecurityIdentifier='S-1-5-21-1-2-3-513'}
$empty = [pscustomobject]@{Id='group-empty';Visibility='Public';DisplayName='Empty group';MailEnabled=$true;
    SecurityEnabled=$false;GroupTypes=@();OnPremisesSecurityIdentifier=$null}
$result = Invoke-CatalogFixture @($group, $empty)
Assert-Catalog (-not $result.ErrorText) 'Complete synthetic group collection failed.'
Assert-Catalog ($result.GroupCalls -eq 1 -and $result.MemberCalls -eq 2) 'Catalog caused an additional group or membership enumeration.'
Assert-Catalog ($result.Exports.Count -eq 4) 'Existing four CSV exports changed.'
Assert-Catalog ($result.Scope -ceq 'CMDB:group_scope,group_members,policies,policy_assignments') 'Consumer receipt scope changed.'
foreach ($property in @('id','visibility','displayName','mailEnabled','securityEnabled','groupTypes','onPremisesSecurityIdentifier')) {
    Assert-Catalog ($result.Properties -contains $property) "Native property not requested: $property"
}
$scope = $result.Exports['M365_EntraGroupMembershipScope']
$members = $result.Exports['M365_EntraGroupMemberships_All'].Rows
Assert-Catalog ($scope.Rows.Count -eq 2 -and $members.Count -eq 2) 'Group or member rows were lost.'
$first = $scope.Rows[0]
Assert-Catalog ($first.DisplayName -ceq $group.DisplayName) 'Native multiline display name changed.'
Assert-Catalog (-not $first.MailEnabled -and $first.SecurityEnabled) 'Native false/true flags changed.'
Assert-Catalog ($first.GroupTypes -ceq 'DynamicMembership;Unified') 'Native group types changed.'
Assert-Catalog ($first.OnPremisesSecurityIdentifier -ceq $group.OnPremisesSecurityIdentifier) 'Native synchronized SID changed.'
Assert-Catalog ($first.MemberCount -eq 2 -and $scope.Rows[1].MemberCount -eq 0) 'Exact or empty member counts changed.'
Assert-Catalog ($scope.Rows[1].GroupTypes -ceq '' -and $null -eq $scope.Rows[1].OnPremisesSecurityIdentifier) 'Missing evidence was invented.'
foreach ($member in $members) {
    Assert-Catalog ($member.RunId -ceq $first.RunId -and $member.CollectedAtUtc -ceq $first.CollectedAtUtc) 'Catalog and members do not share row lineage.'
}
Assert-Catalog ([datetimeoffset]$first.GroupCollectedAtUtc -le [datetimeoffset]$first.CollectedAtUtc) 'Catalog date is after membership acquisition.'
foreach ($column in @('GroupId','Visibility','MemberCount','MemberCollectionStatus','RunId','CollectedAtUtc',
                     'DisplayName','MailEnabled','SecurityEnabled','GroupTypes','OnPremisesSecurityIdentifier','GroupCollectedAtUtc')) {
    Assert-Catalog ($scope.Columns -contains $column) "Export column missing: $column"
}
$zero = Invoke-CatalogFixture @()
Assert-Catalog (-not $zero.ErrorText -and $zero.Exports['M365_EntraGroupMembershipScope'].Rows.Count -eq 0) 'Successful zero collection failed.'
Assert-Catalog ($zero.Exports['M365_EntraGroupMembershipScope'].Columns.Count -eq 12) 'Empty enriched schema lost headers.'
foreach ($case in @(
    @{Groups=@($group,$group);Pattern='duplicate Entra group'},
    @{Groups=@($group);FailMembers=$true;Pattern='Synthetic membership failure'},
    @{Groups=@($group);DuplicateMember=$true;Pattern='duplicate Entra member'},
    @{Groups=@($group);FailPolicies=$true;Pattern='Synthetic policy failure'}
)) {
    $pattern=$case.Pattern; $arguments=$case.Clone(); $arguments.Remove('Pattern')
    $failed=Invoke-CatalogFixture @arguments
    Assert-Catalog ($failed.ErrorText -match $pattern -and $failed.Exports.Count -eq 0) "Failed/duplicate acquisition contract: expected=$pattern; actual=$($failed.ErrorText); exports=$($failed.Exports.Count)."
}
foreach ($count in @(1,2)) {
    $recovered=Invoke-CatalogFixture @($group,$empty) -Member404Count $count
    Assert-Catalog (-not $recovered.ErrorText -and $recovered.Exports.Count -eq 4) 'Transient 404 did not recover.'
    Assert-Catalog ($recovered.GroupCalls -eq 1 -and $recovered.ParentCalls -eq 0 -and $recovered.MemberCalls -eq (2+$count)) 'Recovered traversal made unexpected existence calls.'
    $rows=$recovered.Exports['M365_EntraGroupMemberships_All'].Rows
    Assert-Catalog ($rows.Count -eq 2 -and @($rows | Where-Object MemberId -eq 'discarded-partial-member').Count -eq 0) 'Failed pagination leaked partial membership rows.'
    Assert-Catalog (($recovered.Sleeps -join ',') -ceq $(if($count -eq 1){'5'}else{'5,15'})) 'Membership retries are not bounded to approved delays.'
}
foreach ($codeOnly in @($false,$true)) {
    $gone=Invoke-CatalogFixture @($group,$empty) -Member404Count 9 -CodeOnly404:$codeOnly
    Assert-Catalog (-not $gone.ErrorText -and $gone.Exports.Count -eq 4) 'Corroborated absent group blocked successful collection.'
    Assert-Catalog ($gone.GroupCalls -eq 2 -and $gone.ParentCalls -eq 2 -and $gone.MemberCalls -eq 4) 'Absence verification did not use three traversals, two probes and one complete catalog.'
    Assert-Catalog (($gone.Sleeps -join ',') -ceq '5,15,15') 'Absence verification delay changed.'
    $catalog=$gone.Exports['M365_EntraGroupMembershipScope'].Rows
    Assert-Catalog ($catalog.Count -eq 1 -and $catalog[0].GroupId -ceq 'group-empty') 'Excluded group was retained as empty or another group was lost.'
    Assert-Catalog ($gone.Exports['M365_EntraGroupMemberships_All'].Rows.Count -eq 0) 'Absent group retained partial members.'
    Assert-Catalog (@($gone.Logs | Where-Object { $_.Level -eq 'WARNING' -and $_.Message -match 'group-one' }).Count -eq 1) 'Absent group warning lacks its native identity.'
    Assert-Catalog (($gone.Qualifications -join ' ') -match 'group-one.*parent 404s=2') 'Current receipt qualifications do not trace exclusion.'
}
foreach ($shape in @('ResponseStatus','ResponseObject','Inner')) {
    $gone=Invoke-CatalogFixture @($group,$empty) -Member404Count 9 -ErrorShape $shape
    Assert-Catalog (-not $gone.ErrorText -and $gone.ParentCalls -eq 2 -and $gone.Exports.Count -eq 4) "Structured SDK 404 not recognized: $shape; $($gone.ErrorText)"
}
$allGone=Invoke-CatalogFixture @($group) -Member404Count 9
Assert-Catalog (-not $allGone.ErrorText -and $allGone.Exports['M365_EntraGroupMembershipScope'].Rows.Count -eq 0 -and $allGone.Exports['M365_EntraGroupMembershipScope'].Columns.Count -eq 12) 'Corroborated zero remaining groups lost the explicit schema.'
$policyFailureAfterAbsence=Invoke-CatalogFixture @($group,$empty) -Member404Count 9 -FailPolicies
Assert-Catalog ($policyFailureAfterAbsence.ErrorText -match 'policy failure' -and $policyFailureAfterAbsence.Exports.Count -eq 0 -and -not $policyFailureAfterAbsence.Scope) 'Group exclusion bypassed a later required policy failure.'
foreach ($case in @(
    @{ParentResults=@('Exists');Pattern='still exists'},
    @{ParentResults=@('Absent','Exists');Pattern='still exists'},
    @{ParentResults=@('Empty');Pattern='Ambiguous'},
    @{ParentResults=@('WrongId');Pattern='Ambiguous'},
    @{ParentResults=@('Forbidden');Pattern='forbidden'},
    @{ParentResults=@('Unavailable');Pattern='parent failure'},
    @{Reappears=$true;Pattern='reappeared'},
    @{FailFinalCatalog=$true;Pattern='final catalog failure'},
    @{DuplicateFinalCatalog=$true;Pattern='duplicate identity'}
)) {
    $pattern=$case.Pattern; $arguments=$case.Clone(); $arguments.Remove('Pattern')
    $ambiguous=Invoke-CatalogFixture @($group,$empty) -Member404Count 9 @arguments
    Assert-Catalog ($ambiguous.ErrorText -match $pattern -and $ambiguous.Exports.Count -eq 0 -and -not $ambiguous.Scope) 'Ambiguous absence or failed verification published canonical evidence.'
}
foreach ($status in @(401,403,429,503)) {
    $other=Invoke-CatalogFixture @($group,$empty) -MemberStatus $status
    Assert-Catalog ($other.Exports.Count -eq 0 -and $other.MemberCalls -eq 1 -and $other.ParentCalls -eq 0 -and $other.Sleeps.Count -eq 0) 'Non-404 error was treated as group disappearance.'
}
foreach ($kind in @('InvalidPageToken','CodeOnlyPageToken','DirectoryPageToken','WrappedPageToken')) {
    foreach ($count in @(1,2)) {
        $sequence=@(1..$count | ForEach-Object { $kind }) + @('Success')
        $pageRecovered=Invoke-CatalogFixture @($empty,$group) -MemberErrors $sequence
        Assert-Catalog (-not $pageRecovered.ErrorText -and $pageRecovered.Exports.Count -eq 4) "Page-token traversal was not restarted: $kind; $($pageRecovered.ErrorText)"
        Assert-Catalog ($pageRecovered.CallsByGroup['group-empty'] -eq 1 -and $pageRecovered.CallsByGroup['group-one'] -eq (1+$count)) 'Pagination recovery restarted a previously completed group.'
        Assert-Catalog ($pageRecovered.ParentCalls -eq 0 -and $pageRecovered.GroupCalls -eq 1) 'Pagination recovery entered absence verification.'
        $pageRows=$pageRecovered.Exports['M365_EntraGroupMemberships_All'].Rows
        Assert-Catalog ($pageRows.Count -eq 2 -and @($pageRows | Where-Object MemberId -eq 'discarded-pagination-member').Count -eq 0) 'Rejected pages leaked into the recovered membership export.'
        $pageScope=$pageRecovered.Exports['M365_EntraGroupMembershipScope'].Rows
        Assert-Catalog ($pageScope.Count -eq 2 -and $pageScope[1].MemberCount -eq 2) 'Recovered pagination lost or fabricated group coverage.'
        Assert-Catalog (($pageRecovered.Sleeps -join ',') -ceq $(if($count -eq 1){'5'}else{'5,15'})) 'Pagination delay or attempt budget changed.'
        Assert-Catalog (@($pageRecovered.Logs | Where-Object { $_.Message -match 'GroupId=group-one; reason=InvalidPageToken; attempt=' }).Count -eq $count) 'Pagination retry diagnostics omit the affected group or attempt.'
        Assert-Catalog (($pageRecovered.Qualifications -join ' ') -notmatch 'excluded') 'Recovered pagination was described as an exclusion.'
    }
}
foreach ($sequence in @(
    ,@('InvalidPageToken','NotFound','Success')
    ,@('NotFound','InvalidPageToken','Success')
)) {
    $mixed=Invoke-CatalogFixture @($empty,$group) -MemberErrors $sequence
    Assert-Catalog (-not $mixed.ErrorText -and $mixed.Exports.Count -eq 4 -and $mixed.ParentCalls -eq 0) "Recoverable mixed errors: sequence=$($sequence -join ','); error=$($mixed.ErrorText); exports=$($mixed.Exports.Count)."
}
foreach ($sequence in @(
    ,@('InvalidPageToken','InvalidPageToken','InvalidPageToken')
    ,@('InvalidPageToken','NotFound','NotFound')
    ,@('NotFound','InvalidPageToken','NotFound')
    ,@('NotFound','NotFound','InvalidPageToken')
)) {
    $failedPage=Invoke-CatalogFixture @($empty,$group) -MemberErrors $sequence
    Assert-Catalog ($failedPage.ErrorText -match 'GroupId=group-one.*no group exclusion permitted' -and $failedPage.Exports.Count -eq 0 -and -not $failedPage.Scope) 'Terminal pagination did not fail closed with the affected group identity.'
    Assert-Catalog ($failedPage.CallsByGroup['group-one'] -eq 3 -and $failedPage.ParentCalls -eq 0 -and $failedPage.GroupCalls -eq 1) 'Mixed errors or retry exhaustion entered group exclusion.'
    Assert-Catalog (($failedPage.Sleeps -join ',') -ceq '5,15') 'Terminal pagination exceeded the restart budget.'
}
foreach ($kind in @('GenericBadRequest','ForbiddenPageToken','UnknownPageToken')) {
    $notPaging=Invoke-CatalogFixture @($group) -MemberErrors @($kind)
    Assert-Catalog ($notPaging.ErrorText -match 'GroupId=group-one; attempt=1/3' -and $notPaging.MemberCalls -eq 1 -and $notPaging.Sleeps.Count -eq 0 -and $notPaging.ParentCalls -eq 0 -and $notPaging.Exports.Count -eq 0) "Unqualified error was retried or excluded: $kind"
}
foreach ($option in @('DuplicateMember','FailPolicies')) {
    $switches=@{$option=$true}
    $lateFailure=Invoke-CatalogFixture @($empty,$group) -MemberErrors @('InvalidPageToken','Success') @switches
    Assert-Catalog ($lateFailure.ErrorText -and $lateFailure.Exports.Count -eq 0 -and -not $lateFailure.Scope) 'Recovered pagination weakened downstream duplicate or policy validation.'
}
Write-Output "PASS: $script:checks offline catalog/membership checks. No collectors, APIs, mail or live writes."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCC6kd7TcliQJ8Y
# Yavq4jczt/SDdX5jPSnSP7GFqC/u4qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMwXWAZG/aBNqz0shLTC+/kjb56IgCxN/5DI/rhcl8SdMA0GCSqG
# SIb3DQEBAQUABIIBgKHfZNqiLaBPvMLmgGq9aCM/pOt/CYJ74jw72imVZHosb9mR
# A3d3dLaZEsEC4JCLdCLk1fgd3ytA9GgBz+6ijmNLQjlvpbzAQcQ/OdRyNsWFvtDG
# rvZE1uaWfi1qeer7DzkjavAUQKSVWgOBSfJl6kVkPIb+qnSveopqvq6fuvh+z/G1
# FdBFUZvMLKPYpF4PWd9U9RHrxIZusQ6371kiQ9PpOqTovOPiwa0spLjhFpjS5B3H
# fxISUtBmSHXxPJC6/DtASjL8DMdeL7gyfg/jIvtMT9qLO2U9+Ko1onfmO7UIqoZy
# r/i8C0E18nCl5KuRguJnYpFmP5o75I0aqTw+CU6aqHS9uRj60Is9SUmLjUJcz74E
# JaFzI3VyUyIYSH6X16zu+g0PXFCMqGJBNIGfmrFs4wb2quKBjDxO3wwcKhzINpSW
# 2ln3dpYIK/y5e7sxD994GHOWTVgHI8y4/wx4xmF1cXbOSeLhF7/uzM+QhPBxcKp7
# OcAfWj++FE3B8eT5yqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDYxMDM4
# MThaMC8GCSqGSIb3DQEJBDEiBCDUvEVtJrrj6El/iVWsxoM2VM4G92dYxKzx4c33
# QjFDujANBgkqhkiG9w0BAQEFAASCAgBHobKzr4KrPAIHOuuUQQ0F7V1YoEQRPzih
# zlFTAZpI7NyvWS9C6QTOJlQOgk4pouu0/u8n5kV/71Xu1OF6cslpPNDZbYf5mgQo
# XB9WezHGuLiedpv2ZPL4KRwgXKlpGpcP2nD2O6PztoEMK88T575xctszHV2Z1nP7
# iTElmpir0CJqQUGefbenvqeo7Eoe7Q/SY1K0TG4qPLTl8wN8wDYfO4XwEkGKQger
# rNR+IhfUWpSTBYMbUhJbgBBqh7/aVk+h4gawL6CXV0OkpDWGutkUTfCC84b/DUX+
# Sgj05QT4JwhN5Be2ofGBL/KliEDC+BDjjaew9UhAhpR3THlf0gRoeq8jgMr7103a
# FwcccySN0aLcOI1w3uP6nHcH+HUVYkPqIPWzwMGWJrGOGrE8qHEJdtDe/q0LIGE3
# 1TcM693st4P5hWzwyKCofUfQ314jZwgQ0w1BsBBRlsN3NHLWfbidFPdoccUF4tY6
# v7JInf3C/VmasFnjkISnw0/u+NfTKqCYyHJvVOYuIr8b0cW40GlFvEeh7YdCFw4F
# xL3qc93XDyslCL4yiOX5IpBoziS0DQ/9S67ITPLwwOHf9Vv4AmmK/w8jkF9bhvYg
# 5Pjg7m6ZAtPE3L7fPGN2upxzouTgv1kHXVSZjVMZcUZR3ITaXye8Jzeb4ywkXXeM
# 0G6rUocGpw==
# SIG # End signature block
