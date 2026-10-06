#Requires -Version 5.1
<#
.SYNOPSIS
Offline Exchange mailbox scope, native identity and quality classification tests.
.DESCRIPTION
Extracts functions and statements through the AST. All Exchange queries are
mocked; no tenant context, collector, module, export or external action runs.
.VERSION
1.0.5
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$path = Join-Path $root 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
$tokens=$null; $parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$tests=New-Object 'Collections.Generic.List[object]'
function Assert-True([bool]$Condition,[string]$Message) { if (-not $Condition) { throw $Message } }
function Assert-Throws([scriptblock]$Body,[string]$Pattern) {
    $failure=$null
    try { & $Body | Out-Null } catch { $failure=$_ }
    Assert-True ($null -ne $failure -and $failure.Exception.Message -like $Pattern) "Expected failure: $Pattern"
}
function Case([string]$Name,[scriptblock]$Body) {
    try { & $Body; $tests.Add([pscustomobject]@{Name=$Name;Passed=$true;Error=''}) }
    catch { $tests.Add([pscustomobject]@{Name=$Name;Passed=$false;Error=($_.Exception.Message+' '+$_.ScriptStackTrace)}) }
}
foreach ($name in @('Get-SmartM365LocalMailboxIssueImpact','Add-SmartM365LocalMailboxIssue',
    'Assert-SmartM365MailboxNativePopulation','Invoke-SmartM365LocalMailboxPopulationQuery',
    'Get-SmartM365LocalMailboxQualification','Find-SmartM365MailboxSmtpConflict',
    'Resolve-SmartM365MailboxWarningNativeGuid','ConvertFrom-SmartM365ExchangeRemoteMailboxWarnings',
    'Get-SmartM365MailboxDomainPartition','Assert-SmartM365MailboxDomainCoverage',
    'New-SmartM365LocalMailboxIssueEmailSection','Initialize-SmartM365LocalMailboxForestPopulation',
    'Get-SmartM365LocalMailboxColumns','Invoke-SmartM365RemoteMailboxPopulationQuery')) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Assert-True ($null -ne $node) "Missing function: $name"
    Set-Item "Function:script:$name" ([scriptblock]::Create($node.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
}
function Reset-Fixture {
    $script:LocalMailboxIssues=New-Object 'Collections.Generic.List[object]'
    $script:LocalMailboxIssueSequence=0
    $script:LocalMailboxAcquisitions=New-Object 'Collections.Generic.List[object]'
    $script:mockRows=@(); $script:mockWarnings=@(); $script:mockFailure=$false
    $script:DetectAllDomains=$false
    $script:LocalMailboxExpectedScopes=@()
    $script:LocalMailboxForestPopulation=$null
    $script:RemoteMailboxForestPopulation=$null
    $script:mockRemoteRows=@(); $script:remoteQueryCount=0; $script:mockRemoteFailure=$false
    $script:LocalMailboxAcquisitionMode='Domain'
    $script:queryCount=0
    $script:LocalMailboxObservedNativeGuids=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $script:LocalMailboxOwnedNativeGuids=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
}
function ConvertTo-SmartM365EmailHtmlText {param($Value) [Net.WebUtility]::HtmlEncode([string]$Value)}
function WriteLog {param($Message,$Level)}
function Get-Mailbox {
    [CmdletBinding()]
    param($ResultSize,$OrganizationalUnit,[switch]$ReadFromDomainController)
    $script:queryCount++
    $script:queryArguments=@{ResultSize=$ResultSize;Scope=$OrganizationalUnit;ReadFromDomainController=[bool]$ReadFromDomainController}
    if ($script:mockFailure) { throw 'Synthetic acquisition failure.' }
    # Reproduce the live regression: this switch silently returns no objects.
    if ($ReadFromDomainController) { return }
    foreach ($warning in $script:mockWarnings) { Write-Warning $warning }
    $script:mockRows
}
function Get-RemoteMailbox {
    [CmdletBinding()]
    param($ResultSize,$OnPremisesOrganizationalUnit,[switch]$ReadFromDomainController)
    $script:remoteQueryCount++
    $script:remoteQueryArguments=@{ResultSize=$ResultSize;Scope=$OnPremisesOrganizationalUnit;ReadFromDomainController=[bool]$ReadFromDomainController}
    if ($script:mockRemoteFailure) { throw 'Synthetic remote acquisition failure.' }
    if ($ReadFromDomainController) { return }
    foreach ($warning in $script:mockWarnings) { Write-Warning $warning }
    $script:mockRemoteRows
}
$scope='DC=synthetic,DC=invalid'
$goodQuery=[pscustomobject]@{Scope=$scope;QueryCompleted=$true;ProjectionCompleted=$true;Rows=1}
function Get-TestQualification([object[]]$Issues=@(),[object[]]$Queries=@($goodQuery),[bool]$Remote=$true,[object[]]$Scopes=@($scope)) {
    Get-SmartM365LocalMailboxQualification -ExpectedScopes $Scopes -Acquisitions $Queries -RemoteAcquisitionComplete $Remote -Issues $Issues -PopulationCoverageComplete $true
}
foreach ($category in @('ExchangeObjectInconsistentState','MissingPrimarySmtpAddress','MissingExternalEmailAddress','InvalidExternalEmailAddress','InvalidDisplayName')) {
    foreach ($operation in @('Get-Mailbox','Get-RemoteMailbox')) {
        Case "Observed recipient quality / $operation / $category" {
            $impact=Get-SmartM365LocalMailboxIssueImpact $category $operation -NativeRecordRetained $true
            Assert-True (-not $impact.BlocksCmdbQualification -and $impact.CollectionImpact -eq 'RecipientDataQuality') 'Observed recipient was treated as absent.'
            $proof=Get-TestQualification -Issues @([pscustomobject]@{Category=$category;Operation=$operation;NativeRecordRetained=$true;ObjectGuid='00000000-0000-0000-0000-000000000001'})
            Assert-True $proof.CompleteScope 'Complete acquisition with recipient anomaly was rejected.'
        }
    }
}
foreach ($category in @('MailboxStatisticsFailure','ArchiveMailboxStatisticsFailure','MissingMailboxDatabase','ExchangeStoreUnavailable')) {
    Case "Statistics unavailable / $category" {
        $impact=Get-SmartM365LocalMailboxIssueImpact $category 'Get-MailboxStatistics'
        Assert-True (-not $impact.BlocksCmdbQualification -and $impact.CollectionImpact -eq 'FieldUnavailable') 'Supplementary statistics made the population incomplete.'
    }
}
foreach ($category in @('MissingMobileDeviceIdentity','RecipientAmbiguous','MobileDeviceLookupFailure')) {
    Case "Mobile association unavailable / $category" {
        Assert-True (-not (Get-SmartM365LocalMailboxIssueImpact $category 'Get-MobileDevice').BlocksCmdbQualification) 'Missing mobile field was treated as lost mailbox.'
    }
}
foreach ($issue in @([pscustomobject]@{Category='NewUnexpectedFailure';Operation='Get-RemoteMailbox'},
    [pscustomobject]@{Category='MailboxStatisticsFailure';Operation='Get-Mailbox'},
    [pscustomobject]@{Category='RecipientAmbiguous';Operation='Get-RemoteMailbox'},
    [pscustomobject]@{Category='ConflictingPrimarySmtpAddress';Operation='MailboxReconciliation'},
    [pscustomobject]@{Category='MissingPrimarySmtpAddress';Operation='Unknown'}, 'legacy unclassified issue')) {
    Case "Unknown or essential failure blocks / $issue" {
        Assert-True (-not (Get-TestQualification -Issues @($issue)).CompleteScope) 'Unknown failure was allowed through.'
    }
}
Case 'Issue classification is additive and preserves the evidence' {
    Reset-Fixture
    $issue=Add-SmartM365LocalMailboxIssue -Category MissingPrimarySmtpAddress -Operation Get-RemoteMailbox -MailboxIdentity 'native-identity' -ObjectGuid '00000000-0000-0000-0000-000000000001' -NativeRecordRetained $true -Message 'Observed missing SMTP.'
    Assert-True ($script:LocalMailboxIssues.Count -eq 1 -and $issue.MailboxIdentity -eq 'native-identity' -and $issue.Message -eq 'Observed missing SMTP.') 'Issue evidence changed.'
    Assert-True ($issue.Severity -eq 'Warning' -and -not $issue.BlocksCmdbQualification) 'Issue classification missing.'
}
Case 'Unrecognized Exchange warning is retained and blocks' {
    $warnings=@(ConvertFrom-SmartM365ExchangeRemoteMailboxWarnings @('Synthetic unfamiliar warning.'))
    Assert-True ($warnings.Count -eq 1 -and $warnings[0].Issue -eq 'UnclassifiedRecipientWarning') 'Unrecognized warning was discarded.'
}
Case 'Complete empty domain query is valid population evidence' {
    $zero=[pscustomobject]@{Scope=$scope;QueryCompleted=$true;ProjectionCompleted=$true;Rows=0}
    Assert-True (Get-TestQualification -Queries @($zero)).CompleteScope 'Successful empty domain was rejected.'
}
foreach ($scenario in @('NoExpectedScopes','MissingDomain','DuplicateDomain','FailedQuery','FailedProjection','MissingRemote')) {
    Case "Incomplete acquisition blocks / $scenario" {
        $queries=@([pscustomobject]@{Scope=$scope;QueryCompleted=$true;ProjectionCompleted=$true;Rows=1})
        $scopes=@($scope); $remote=$true
        switch ($scenario) {
            'NoExpectedScopes' {$scopes=@()}
            'MissingDomain' {$scopes+= 'DC=other,DC=invalid'}
            'DuplicateDomain' {$queries+=$queries[0]}
            'FailedQuery' {$queries[0].QueryCompleted=$false}
            'FailedProjection' {$queries[0].ProjectionCompleted=$false}
            'MissingRemote' {$remote=$false}
        }
        Assert-True (-not (Get-TestQualification -Queries $queries -Scopes $scopes -Remote $remote).CompleteScope) 'Incomplete population qualified.'
    }
}
$id1=[guid]'00000000-0000-0000-0000-000000000001'
$id2=[guid]'00000000-0000-0000-0000-000000000002'
$native=@([pscustomobject]@{Guid=$id1},[pscustomobject]@{Guid=$id2})
$export=@([pscustomobject]@{ObjectGuid=$id1;PrimarySmtpAddress='';IsValid=$false},[pscustomobject]@{ObjectGuid=$id2;PrimarySmtpAddress='';IsValid=$false})
Case 'Known recipient warning is blocking until its native object is proven retained' {
    $issue=[pscustomobject]@{Category='MissingPrimarySmtpAddress';Operation='Get-RemoteMailbox';NativeRecordRetained=$false}
    Assert-True (-not (Get-TestQualification -Issues @($issue)).CompleteScope) 'Known warning hid a missing recipient.'
    $issue.NativeRecordRetained='true'
    Assert-True (-not (Get-TestQualification -Issues @($issue)).CompleteScope) 'Text marker substituted for native observation.'
}
Case 'Native warning identity matches exactly, including quoted canonical identity' {
    $rows=@([pscustomobject]@{Guid=$id1;Identity='synthetic.invalid/Container/Recipient';DistinguishedName='CN=Recipient,DC=synthetic,DC=invalid'})
    $match=Resolve-SmartM365MailboxWarningNativeGuid -ObjectPath "'synthetic.invalid/Container/Recipient'" -NativeRows $rows
    Assert-True ($match -eq $id1.ToString('D')) 'Canonical Exchange identity was not matched.'
    Assert-True ((Resolve-SmartM365MailboxWarningNativeGuid -ObjectPath $rows[0].DistinguishedName -NativeRows $rows) -eq $match) 'Exact DN was not matched.'
}
Case 'Display-name-only or absent warning identity cannot prove retention' {
    $rows=@([pscustomobject]@{Guid=$id1;Identity='synthetic.invalid/Container/Recipient';DisplayName='Recipient'})
    Assert-True (-not (Resolve-SmartM365MailboxWarningNativeGuid -ObjectPath 'Recipient' -NativeRows $rows)) 'Display name guessed a native identity.'
    Assert-True (-not (Resolve-SmartM365MailboxWarningNativeGuid -ObjectPath '' -NativeRows $rows)) 'Absent warning identity guessed.'
}
Case 'Ambiguous warning identity cannot prove retention' {
    $rows=@([pscustomobject]@{Guid=$id1;Identity='same'},[pscustomobject]@{Guid=$id2;Identity='same'})
    Assert-True (-not (Resolve-SmartM365MailboxWarningNativeGuid -ObjectPath 'same' -NativeRows $rows)) 'Ambiguous recipient guessed.'
}
Case 'Retention marker without a usable native GUID cannot qualify a recipient warning' {
    $issue=[pscustomobject]@{Category='MissingPrimarySmtpAddress';Operation='Get-RemoteMailbox';NativeRecordRetained=$true;ObjectGuid=''}
    Assert-True (-not (Get-TestQualification -Issues @($issue)).CompleteScope) 'Retention marker replaced native identity.'
    Reset-Fixture
    $observed=Add-SmartM365LocalMailboxIssue -Category MissingPrimarySmtpAddress -Operation Get-RemoteMailbox -NativeRecordRetained $true -Message 'No native GUID.'
    Assert-True ($observed.BlocksCmdbQualification -and -not $observed.NativeRecordRetained) 'Invalid evidence flag persisted as retained.'
}
Case 'Invalid recipients and missing SMTP retain distinct native GUIDs' {
    Assert-SmartM365MailboxNativePopulation -NativeRows $native -ExportRows $export
    Assert-True ($export.Count -eq 2 -and -not $export[0].IsValid) 'Recipient values were fabricated.'
}
Case 'Empty native and exported populations are accepted' { Assert-SmartM365MailboxNativePopulation -NativeRows @() -ExportRows @() }
Case 'Dropped native object is rejected' { Assert-Throws { Assert-SmartM365MailboxNativePopulation -NativeRows $native -ExportRows @($export[0]) } '*lost native*' }
Case 'Repeated native GUID is rejected' { Assert-Throws { Assert-SmartM365MailboxNativePopulation -NativeRows @($native[0],$native[0]) -ExportRows $export } '*Repeated native*' }
Case 'Repeated exported GUID is rejected' { Assert-Throws { Assert-SmartM365MailboxNativePopulation -NativeRows $native -ExportRows @($export[0],$export[0]) } '*differs from native*' }
Case 'Unknown native GUID is rejected' { Assert-Throws { Assert-SmartM365MailboxNativePopulation -NativeRows @([pscustomobject]@{Guid=''}) -ExportRows @($export[0]) } '*no usable*' }
Case 'Empty GUID is rejected' { Assert-Throws { Assert-SmartM365MailboxNativePopulation -NativeRows @([pscustomobject]@{Guid=[guid]::Empty}) -ExportRows @($export[0]) } '*no usable*' }
Case 'Replacement exported GUID is rejected' { Assert-Throws { Assert-SmartM365MailboxNativePopulation -NativeRows @($native[0]) -ExportRows @($export[1]) } '*differs from native*' }
Case 'Primary SMTP conflicts never remove or alter records' {
    Reset-Fixture
    $rows=@([pscustomobject]@{ObjectGuid=$id1;PrimarySmtpAddress='same@synthetic.invalid'},[pscustomobject]@{ObjectGuid=$id2;PrimarySmtpAddress='SAME@synthetic.invalid'})
    $before=$rows | ConvertTo-Json -Compress
    Find-SmartM365MailboxSmtpConflict -Rows $rows
    Assert-True (($rows | ConvertTo-Json -Compress) -ceq $before) 'SMTP check rewrote native records.'
    Assert-True ($script:LocalMailboxIssues.Count -eq 1 -and $script:LocalMailboxIssues[0].BlocksCmdbQualification) 'Core address conflict was hidden.'
}
Case 'Blank SMTP is not collapsed or called a duplicate address' {
    Reset-Fixture
    Find-SmartM365MailboxSmtpConflict -Rows $export
    Assert-True ($script:LocalMailboxIssues.Count -eq 0) 'Blank SMTP became a false duplicate group.'
}
Case 'Actual population query is unlimited and leaves projection unproven' {
    Reset-Fixture; $script:mockRows=$native
    $observed=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $scope)
    Assert-True ($observed.Count -eq 2 -and $script:queryArguments.ResultSize -eq 'Unlimited') 'Population was limited.'
    Assert-True ($script:LocalMailboxAcquisitions[0].QueryCompleted -and -not $script:LocalMailboxAcquisitions[0].ProjectionCompleted) 'Query alone claimed complete projection.'
}
Case 'Actual query failure leaves explicit failed acquisition' {
    Reset-Fixture; $script:mockFailure=$true
    Assert-Throws { Invoke-SmartM365LocalMailboxPopulationQuery -Scope $scope } '*Synthetic acquisition failure*'
    Assert-True (-not $script:LocalMailboxAcquisitions[0].QueryCompleted) 'Failed query became a zero population.'
}
Case 'Actual query captures unfamiliar warnings without losing rows' {
    Reset-Fixture; $script:mockRows=$native; $script:mockWarnings=@('Synthetic unfamiliar warning.')
    $observed=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $scope -WarningAction SilentlyContinue)
    Assert-True ($observed.Count -eq 2 -and $script:LocalMailboxIssues.Count -eq 1) 'Warning or returned objects were lost.'
    Assert-True $script:LocalMailboxIssues[0].BlocksCmdbQualification 'Unclassified warning qualified.'
}
Case 'Actual query binds recipient warnings to returned native GUIDs' {
    Reset-Fixture
    $script:mockRows=@([pscustomobject]@{Guid=$id1;Identity='synthetic.invalid/Container/Recipient'})
    $script:mockWarnings=@("The object synthetic.invalid/Container/Recipient has been corrupted or isn't compatible with the server.",'There is no primary SMTP address.')
    $observed=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $scope -WarningAction SilentlyContinue)
    Assert-True ($observed.Count -eq 1 -and $script:LocalMailboxIssues.Count -eq 2) 'Returned object or warnings lost.'
    foreach($issue in $script:LocalMailboxIssues) {
        Assert-True ($issue.NativeRecordRetained -and $issue.ObjectGuid -eq $id1.ToString('D') -and -not $issue.BlocksCmdbQualification) 'Warning did not bind to retained native GUID.'
    }
}
Case 'Actual query warning about an absent object stays blocking' {
    Reset-Fixture; $script:mockRows=$native
    $script:mockWarnings=@("The object synthetic.invalid/Missing has been corrupted or isn't compatible with the server.")
    $observed=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $scope -WarningAction SilentlyContinue)
    Assert-True ($observed.Count -eq 2 -and $script:LocalMailboxIssues[0].BlocksCmdbQualification) 'Absent warned object was assumed retained.'
}
Case 'Missing required UserMailbox database stays linked to its retained object' {
    Reset-Fixture
    $script:mockRows=@([pscustomobject]@{Guid=$id1;Identity='synthetic.invalid/Container/Recipient';Database='';IsValid=$false})
    $script:mockWarnings=@("The object synthetic.invalid/Container/Recipient has been corrupted or isn't compatible with the server.",'Database is mandatory on UserMailbox.')
    $observed=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $scope -WarningAction SilentlyContinue)
    Assert-True ($observed.Count -eq 1 -and $script:LocalMailboxIssues.Count -eq 2) 'Object or warning disappeared.'
    $issue=$script:LocalMailboxIssues[1]
    Assert-True ($issue.Category -eq 'MissingRequiredMailboxDatabase' -and $issue.ObjectGuid -eq $id1.ToString('D') -and -not $issue.BlocksCmdbQualification) 'Database warning lost native object context.'
    Assert-True (-not $observed[0].IsValid -and $observed[0].Database -eq '') 'Invalid recipient was repaired or excluded.'
}
foreach ($warningCase in @('Standalone','UnknownIntervening','MissingNativeObject','AmbiguousNativeObject')) {
    Case "Missing required database is blocking without proven context / $warningCase" {
        Reset-Fixture
        $script:mockRows=@([pscustomobject]@{Guid=$id1;Identity='synthetic.invalid/Recipient'})
        $script:mockWarnings=@("The object synthetic.invalid/Recipient has been corrupted or isn't compatible with the server.")
        switch($warningCase) {
            'Standalone' {$script:mockWarnings=@()}
            'UnknownIntervening' {$script:mockWarnings+='Synthetic unfamiliar warning.'}
            'MissingNativeObject' {$script:mockRows=@()}
            'AmbiguousNativeObject' {$script:mockRows+= [pscustomobject]@{Guid=$id2;Identity='synthetic.invalid/Recipient'}}
        }
        $script:mockWarnings+='Database is mandatory on UserMailbox.'
        Invoke-SmartM365LocalMailboxPopulationQuery -Scope $scope -WarningAction SilentlyContinue | Out-Null
        Assert-True $script:LocalMailboxIssues[$script:LocalMailboxIssues.Count-1].BlocksCmdbQualification 'Unproven database warning qualified.'
    }
}
Case 'Database warning is not allowed for a different operation' {
    Assert-True (Get-SmartM365LocalMailboxIssueImpact -Category MissingRequiredMailboxDatabase -Operation Get-RemoteMailbox -NativeRecordRetained $true).BlocksCmdbQualification 'Database warning was broadly allowlisted.'
}
Case 'One GUID repeated is an export defect, not a distinct-object SMTP conflict' {
    Reset-Fixture
    $rows=@([pscustomobject]@{ObjectGuid=$id1;PrimarySmtpAddress='same@synthetic.invalid'},[pscustomobject]@{ObjectGuid=$id1;PrimarySmtpAddress='same@synthetic.invalid'})
    Find-SmartM365MailboxSmtpConflict -Rows $rows
    Assert-True ($script:LocalMailboxIssues.Count -eq 1 -and $script:LocalMailboxIssues[0].Category -eq 'RepeatedNativeMailboxObjectGuid' -and $script:LocalMailboxIssues[0].BlocksCmdbQualification) 'Repeated native GUID was hidden or called a distinct-object SMTP conflict.'
    Assert-True ($rows.Count -eq 2) 'Conflict check discarded a record.'
}
$rootScope='DC=synthetic,DC=invalid'; $childScope='DC=child,DC=synthetic,DC=invalid'
$rootNative=[pscustomobject]@{Guid=$id1;DistinguishedName='CN=Root recipient,CN=Users,DC=synthetic,DC=invalid';Identity='synthetic.invalid/Users/Root recipient'}
$childNative=[pscustomobject]@{Guid=$id2;DistinguishedName='CN=Child recipient,CN=Users,DC=child,DC=synthetic,DC=invalid';Identity='child.synthetic.invalid/Users/Child recipient'}
Case 'Full forest makes one live unlimited query without the regressing switch and partitions every domain' {
    Reset-Fixture; $script:DetectAllDomains=$true; $script:LocalMailboxExpectedScopes=@($rootScope,$childScope,'DC=empty,DC=invalid')
    $script:mockRows=@($rootNative,$childNative)
    Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes $script:LocalMailboxExpectedScopes
    $script:mockFailure=$true
    $parent=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $rootScope)
    $child=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $childScope)
    $empty=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope 'DC=empty,DC=invalid')
    Assert-True ($script:queryCount -eq 1 -and $script:queryArguments.ResultSize -eq 'Unlimited' -and -not $script:queryArguments.Scope -and -not $script:queryArguments.ReadFromDomainController) 'Forest query is scoped, limited, regressed or repeated.'
    Assert-True ($parent.Count -eq 1 -and $child.Count -eq 1 -and $empty.Count -eq 0 -and $script:LocalMailboxAcquisitions.Count -eq 1) 'In-memory partition lost objects or fabricated domain acquisitions.'
    Assert-True ([object]::ReferenceEquals($parent[0],$rootNative) -and [object]::ReferenceEquals($child[0],$childNative)) 'Acquired native objects were replaced.'
    Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($script:LocalMailboxObservedNativeGuids) -OwnedGuids @($script:LocalMailboxOwnedNativeGuids) -ExportRows $export
    Assert-True (-not $script:LocalMailboxAcquisitions[0].ProjectionCompleted) 'Acquisition alone claimed completed enrichment.'
    $script:LocalMailboxAcquisitions[0].ProjectionCompleted=$true
    $proof=Get-SmartM365LocalMailboxQualification -ExpectedScopes $script:LocalMailboxExpectedScopes -Acquisitions $script:LocalMailboxAcquisitions.ToArray() -RemoteAcquisitionComplete $true -Issues @() -PopulationCoverageComplete $true -AcquisitionMode Forest
    Assert-True $proof.CompleteScope 'Complete single-query population rejected.'
    Assert-True (($proof.Qualifications -join ' ') -like '*actual queries=1*discovered domains=3*') 'Receipt calls partitions independent queries.'
}
Case 'Forest acquisition is not repeated and does not accept a domain outside its evidence' {
    Assert-Throws {Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope,$childScope)} '*once before*'
    Assert-Throws {Invoke-SmartM365LocalMailboxPopulationQuery -Scope 'DC=other,DC=invalid'} '*outside*'
    Assert-True ($script:queryCount -eq 1) 'Rejected request triggered acquisition.'
}
Case 'Failed forest query cannot supply domain projections or qualify empty output' {
    Reset-Fixture; $script:mockFailure=$true
    Assert-Throws {Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope)} '*Synthetic acquisition failure*'
    Assert-True ($null -eq $script:LocalMailboxForestPopulation -and -not $script:LocalMailboxAcquisitions[0].QueryCompleted) 'Failed query produced a reusable population.'
    $proof=Get-SmartM365LocalMailboxQualification -ExpectedScopes @($rootScope) -Acquisitions $script:LocalMailboxAcquisitions.ToArray() -RemoteAcquisitionComplete $true -Issues @() -PopulationCoverageComplete $true -AcquisitionMode Forest
    Assert-True (-not $proof.CompleteScope) 'Failed forest query qualified.'
}
Case 'Forest acquisition rejects ambiguous native identity and unmapped DN before processing' {
    foreach ($rows in @(@($rootNative,$rootNative),@([pscustomobject]@{Guid=$id1;DistinguishedName='CN=Recipient,DC=other,DC=invalid'}))) {
        Reset-Fixture; $script:mockRows=$rows
        Assert-Throws {Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope)} '*native*'
        Assert-True ($null -eq $script:LocalMailboxForestPopulation) 'Invalid population supplied domain rows.'
    }
}
Case 'Known forest recipient warnings are bound once; unclassified warnings still block' {
    Reset-Fixture; $script:mockRows=@($rootNative,$childNative)
    $script:mockWarnings=@("The object $($rootNative.Identity) has been corrupted or isn't compatible with the server.",'Database is mandatory on UserMailbox.')
    Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope,$childScope) -WarningAction SilentlyContinue
    Invoke-SmartM365LocalMailboxPopulationQuery -Scope $rootScope | Out-Null
    Invoke-SmartM365LocalMailboxPopulationQuery -Scope $childScope | Out-Null
    Assert-True ($script:LocalMailboxIssues.Count -eq 2 -and @($script:LocalMailboxIssues | Where-Object BlocksCmdbQualification).Count -eq 0) 'Warnings were repeated or unbound during partitioning.'
    $script:LocalMailboxAcquisitions[0].ProjectionCompleted=$true
    Add-SmartM365LocalMailboxIssue -Category UnexpectedFailure -Operation Get-Mailbox -Message 'Synthetic unknown warning.' | Out-Null
    $proof=Get-SmartM365LocalMailboxQualification -ExpectedScopes @($rootScope,$childScope) -Acquisitions $script:LocalMailboxAcquisitions.ToArray() -RemoteAcquisitionComplete $true -Issues $script:LocalMailboxIssues.ToArray() -PopulationCoverageComplete $true -AcquisitionMode Forest
    Assert-True (-not $proof.CompleteScope) 'Unclassified forest warning qualified.'
}
Case 'Unconfirmed zero forest cannot supply partitions or qualify publication' {
    Reset-Fixture
    Assert-Throws { Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope,$childScope) } '*Unconfirmed empty local mailbox forest*'
    Assert-True ($null -eq $script:LocalMailboxForestPopulation -and $script:LocalMailboxAcquisitions[0].QueryCompleted -and $script:LocalMailboxAcquisitions[0].ObservedRows -eq 0) 'Empty forest produced usable evidence or misreported query completion.'
    $proof=Get-SmartM365LocalMailboxQualification -ExpectedScopes @($rootScope,$childScope) -Acquisitions $script:LocalMailboxAcquisitions.ToArray() -RemoteAcquisitionComplete $true -Issues @() -PopulationCoverageComplete $true -AcquisitionMode Forest
    Assert-True (-not $proof.CompleteScope -and $script:queryCount -eq 1) 'Empty forest was qualified or replaced with another population.'
}
foreach ($defect in @('Scoped','Limited','NotFresh','RegressingSwitch','ZeroPopulation','NoProjection','MissingDomain','ExtraDomain','WrongCount','NoCoverage','RepeatedQuery','UnknownQuery')) {
    Case "Forest qualification rejects invalid acquisition proof / $defect" {
        Reset-Fixture; $script:mockRows=@($rootNative)
        Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope,$childScope)
        $q=$script:LocalMailboxAcquisitions[0]; $q.ProjectionCompleted=$true
        $coverage=$true
        switch ($defect) {
            Scoped {$q.Scope=$rootScope}
            Limited {$q.ResultSize='1'}
            NotFresh {$q.AcquiredThisRun=$false}
            RegressingSwitch {$q.ReadFromDomainController=$true}
            ZeroPopulation {$q.Rows=0;$q.ObservedRows=0;$q.DomainCounts[$rootScope]=0}
            NoProjection {$q.ProjectionCompleted=$false}
            MissingDomain {$q.DomainCounts.Remove($childScope)}
            ExtraDomain {$q.DomainCounts['DC=extra,DC=invalid']=0}
            WrongCount {$q.DomainCounts[$rootScope]=2}
            NoCoverage {$coverage=$false}
            RepeatedQuery {$script:LocalMailboxAcquisitions.Add($q)}
            UnknownQuery {$script:LocalMailboxAcquisitions.Clear();$script:LocalMailboxAcquisitions.Add([pscustomobject]@{QueryCompleted=$true})}
        }
        $proof=Get-SmartM365LocalMailboxQualification -ExpectedScopes @($rootScope,$childScope) -Acquisitions $script:LocalMailboxAcquisitions.ToArray() -RemoteAcquisitionComplete $true -Issues @() -PopulationCoverageComplete $coverage -AcquisitionMode Forest
        Assert-True (-not $proof.CompleteScope) 'Incomplete, scoped or stale forest evidence qualified.'
    }
}
Case 'Remote forest preserves live objects and warnings without the regressing switch' {
    Reset-Fixture; $script:mockRemoteRows=@($rootNative,$childNative)
    $script:mockWarnings=@('Synthetic recipient warning.')
    $population=Invoke-SmartM365RemoteMailboxPopulationQuery -WarningAction SilentlyContinue
    Assert-True ($script:remoteQueryCount -eq 1 -and $script:remoteQueryArguments.ResultSize -eq 'Unlimited' -and -not $script:remoteQueryArguments.Scope -and -not $script:remoteQueryArguments.ReadFromDomainController) 'Remote query limited or regressed.'
    Assert-True ($population.Rows.Count -eq 2 -and [object]::ReferenceEquals($population.Rows[0],$rootNative) -and $population.Warnings.Count -eq 1) 'Remote native evidence changed.'
}
Case 'Remote zero and query failure cannot become an acquired forest population' {
    Reset-Fixture
    Assert-Throws { Invoke-SmartM365RemoteMailboxPopulationQuery } '*Unconfirmed empty remote mailbox population*'
    $script:mockRemoteFailure=$true
    Assert-Throws { Invoke-SmartM365RemoteMailboxPopulationQuery } '*Synthetic remote acquisition failure*'
    Assert-True ($null -eq $script:RemoteMailboxForestPopulation) 'Failure installed a forest population.'
}
Case 'Both forest populations are acquired before the first domain publication' {
    $body=$ast.Extent.Text
    $local=$body.IndexOf('Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes $script:LocalMailboxExpectedScopes')
    $remote=$body.IndexOf('$script:RemoteMailboxForestPopulation = Invoke-SmartM365RemoteMailboxPopulationQuery')
    $process=$body.IndexOf('Process-SpecificDomain -CurrentDomain $domain')
    Assert-True ($local -gt 0 -and $remote -gt $local -and $process -gt $remote) 'Population guard runs after domain writes.'
}
Case 'Actual full-mode guard stops domain processing when the remote population is empty' {
    Reset-Fixture; $script:mockRows=@($rootNative,$childNative)
    $script:LocalMailboxExpectedScopes=@($rootScope,$childScope)
    $OnlyADPermission=$false; $ForceOverwriteCSV=$true; $MaxItems=0; $IncludeRemoteMailboxes=$true
    $domainsToProcess=@([pscustomobject]@{Name='synthetic.invalid'})
    $script:domainCalls=0
    function Process-SpecificDomain {param($CurrentDomain) $script:domainCalls++}
    $guard=$ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '-not $OnlyADPermission -and $ForceOverwriteCSV -and $MaxItems -eq 0' -and $n.Extent.Text.Contains('Initialize-SmartM365LocalMailboxForestPopulation')},$true)
    $loop=$ast.Find({param($n) $n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Extent.Text.Contains('Process-SpecificDomain -CurrentDomain $domain')},$true)
    Assert-True ($null -ne $guard -and $null -ne $loop) 'Actual full-mode guard or processing loop is missing.'
    Assert-Throws { & ([scriptblock]::Create($guard.Extent.Text + "`n" + $loop.Extent.Text)) } '*Unconfirmed empty remote mailbox population*'
    Assert-True ($script:domainCalls -eq 0 -and $null -eq $script:RemoteMailboxForestPopulation -and $script:queryCount -eq 1 -and $script:remoteQueryCount -eq 1) 'Remote zero started domain processing or reused evidence.'
}
Case 'Remote projection consumes current-run forest objects without querying twice' {
    Reset-Fixture; $script:mockRemoteRows=@($rootNative,$childNative)
    $script:RemoteMailboxForestPopulation=Invoke-SmartM365RemoteMailboxPopulationQuery
    $script:mockRemoteFailure=$true
    $inventory=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SmartM365ExchangeRemoteMailboxInventory'},$true)
    $selection=$inventory.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$population'},$true)
    $IncludedLDAPPaths=@()
    & ([scriptblock]::Create($selection.Extent.Text + '; $script:observedRemotePopulation=$population'))
    Assert-True ($script:remoteQueryCount -eq 1 -and [object]::ReferenceEquals($script:observedRemotePopulation,$script:RemoteMailboxForestPopulation)) 'Remote acquisition was repeated or replaced.'
}
Case 'Native empty local schema matches every distinct field of the actual normal projector' {
    $process=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'MailboxesProcessing2'},$true)
    $columns=New-Object 'Collections.Generic.List[string]'
    $LargeItemThresholdMBValue=35
    foreach ($command in $process.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Add-Member'},$true)) {
        if ($command.Parent.Extent.Text -notmatch '^\$userObj') { continue }
        for ($index=0; $index -lt $command.CommandElements.Count-1; $index++) {
            if ($command.CommandElements[$index].Extent.Text -eq '-Name') {
                $field=& ([scriptblock]::Create($command.CommandElements[$index+1].Extent.Text))
                if (-not $columns.Contains($field)) { $columns.Add($field) }
            }
        }
    }
    Assert-True ((@(Get-SmartM365LocalMailboxColumns) -join '|') -ceq ($columns -join '|')) 'Empty schema changed native field names/order or omitted a projected field.'
}
Case 'Production acquires before processing, rejects global zero and preserves domain empty schemas' {
    $body=$ast.Extent.Text
    Assert-True ($body.IndexOf('Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes $script:LocalMailboxExpectedScopes') -lt $body.IndexOf('foreach ($domain in $domainsToProcess)')) 'Full mode still acquires inside each domain.'
    Assert-True ($body.Contains('-AcquisitionMode $script:LocalMailboxAcquisitionMode') -and -not $body.Contains('Export-SmartM365EmptyLocalMailboxPopulation -Path $globalCombinedCsvFile') -and $body.Contains('Export-SmartM365EmptyLocalMailboxPopulation -Path $perDomainCsvFullPath -DomainScope $distinguishedName')) 'Global zero guard or domain schema was not wired.'
    $remote=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SmartM365RemoteMailboxPopulationQuery'},$true)
    $queries=@($remote.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-RemoteMailbox'},$true))
    Assert-True ($queries.Count -eq 2) 'Remote scoped and forest queries were not checked.'
    foreach ($query in $queries) { Assert-True (-not $query.Extent.Text.Contains('-ReadFromDomainController')) 'Remote query reintroduced the regressing switch.' }
}
Case 'Root/child overlap partitions native objects before enrichment and retains CN containers' {
    $parent=Get-SmartM365MailboxDomainPartition -NativeRows @($rootNative,$childNative) -DomainScope $rootScope -ForestScopes @($rootScope,$childScope)
    $child=Get-SmartM365MailboxDomainPartition -NativeRows @($childNative) -DomainScope $childScope -ForestScopes @($rootScope,$childScope)
    Assert-True ($parent.Rows.Count -eq 1 -and $parent.Rows[0].Guid -eq $id1 -and $child.Rows.Count -eq 1 -and $child.Rows[0].Guid -eq $id2) 'Parent scope duplicated or lost a child/container recipient.'
    Assert-True ($parent.ObservedGuids.Count -eq 2) 'Raw returned population was discarded from coverage proof.'
    Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($parent.ObservedGuids+$child.ObservedGuids) -OwnedGuids @($parent.OwnedGuids+$child.OwnedGuids) -ExportRows $export
}
Case 'Domain suffix comparison is case-insensitive and boundary-aware' {
    $row=[pscustomobject]@{Guid=$id1;DistinguishedName='CN=Recipient,OU=Staff,DC=SYNTHETIC,DC=INVALID'}
    $partition=Get-SmartM365MailboxDomainPartition -NativeRows @($row) -DomainScope $rootScope -ForestScopes @($rootScope,$childScope)
    Assert-True ($partition.Rows.Count -eq 1) 'Native DN case lost a recipient.'
    $row.DistinguishedName='CN=Recipient,DC=notsynthetic,DC=invalid'
    Assert-Throws {Get-SmartM365MailboxDomainPartition -NativeRows @($row) -DomainScope $rootScope -ForestScopes @($rootScope,$childScope)} '*does not belong*'
}
Case 'Empty domain partition is valid and has complete empty metadata' {
    $partition=Get-SmartM365MailboxDomainPartition -NativeRows @() -DomainScope $rootScope -ForestScopes @($rootScope,$childScope)
    Assert-True ($partition.Rows.Count -eq 0 -and $partition.ObservedGuids.Count -eq 0) 'Empty partition fabricated data.'
    Assert-SmartM365MailboxDomainCoverage -ObservedGuids @() -OwnedGuids @() -ExportRows @()
}
Case 'Partition rejects repeated native GUIDs within a query' {
    Assert-Throws {Get-SmartM365MailboxDomainPartition -NativeRows @($rootNative,$rootNative) -DomainScope $rootScope -ForestScopes @($rootScope,$childScope)} '*Repeated native*'
}
Case 'Partition rejects absent native identity or DN and invalid domain scope' {
    Assert-Throws {Get-SmartM365MailboxDomainPartition -NativeRows @([pscustomobject]@{Guid='';DistinguishedName=$rootNative.DistinguishedName}) -DomainScope $rootScope -ForestScopes @($rootScope)} '*no usable*'
    Assert-Throws {Get-SmartM365MailboxDomainPartition -NativeRows @([pscustomobject]@{Guid=$id1;DistinguishedName=''}) -DomainScope $rootScope -ForestScopes @($rootScope)} '*does not belong*'
    Assert-Throws {Get-SmartM365MailboxDomainPartition -NativeRows @() -DomainScope $childScope -ForestScopes @($rootScope)} '*Invalid or repeated*'
    Assert-Throws {Get-SmartM365MailboxDomainPartition -NativeRows @() -DomainScope $rootScope -ForestScopes @($rootScope,$rootScope)} '*Invalid or repeated*'
}
Case 'Native domain coverage rejects loss, substitution and repeated export GUIDs' {
    Assert-Throws {Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($id1,$id2) -OwnedGuids @($id1) -ExportRows @($export[0])} '*lost or substituted*'
    Assert-Throws {Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($id1) -OwnedGuids @($id2) -ExportRows @($export[1])} '*lost or substituted*'
    Assert-Throws {Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($id1) -OwnedGuids @($id1,$id1) -ExportRows @($export[0])} '*Invalid or repeated owned*'
    Assert-Throws {Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($id1,$id2) -OwnedGuids @($id1,$id2) -ExportRows @($export[0],$export[0])} '*differs from native*'
}
Case 'Acquisition success alone cannot substitute for global coverage proof' {
    $proof=Get-SmartM365LocalMailboxQualification -ExpectedScopes @($scope) -Acquisitions @($goodQuery) -RemoteAcquisitionComplete $true -Issues @()
    Assert-True (-not $proof.CompleteScope) 'Missing global coverage guard qualified.'
}
Case 'Actual queries retain all observed GUIDs but return only their native domain' {
    Reset-Fixture; $script:DetectAllDomains=$true; $script:LocalMailboxExpectedScopes=@($rootScope,$childScope)
    $script:mockRows=@($rootNative,$childNative)
    $parent=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $rootScope)
    $script:mockRows=@($childNative)
    $child=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $childScope)
    Assert-True ($parent.Count -eq 1 -and $child.Count -eq 1 -and $script:LocalMailboxObservedNativeGuids.Count -eq 2 -and $script:LocalMailboxOwnedNativeGuids.Count -eq 2) 'Real acquisition function did not partition native scope.'
    Assert-True ($script:LocalMailboxAcquisitions[0].ObservedRows -eq 2 -and $script:LocalMailboxAcquisitions[0].Rows -eq 1) 'Raw versus owned counts were conflated.'
}
Case 'Actual acquisition rejects assigning one native GUID twice' {
    Reset-Fixture; $script:DetectAllDomains=$true; $script:LocalMailboxExpectedScopes=@($rootScope,$childScope); $script:mockRows=@($childNative)
    Invoke-SmartM365LocalMailboxPopulationQuery -Scope $childScope | Out-Null
    Assert-Throws {Invoke-SmartM365LocalMailboxPopulationQuery -Scope $childScope} '*more than one domain*'
    Assert-True (-not $script:LocalMailboxAcquisitions[1].QueryCompleted) 'Duplicate ownership became complete acquisition.'
}
Case 'Local finding email counts native objects, occurrences and unbound warnings separately' {
    $issues=@([pscustomobject]@{Category='MissingRequiredMailboxDatabase';Operation='Get-Mailbox';ObjectGuid=$id1;BlocksCmdbQualification=$false},[pscustomobject]@{Category='MissingRequiredMailboxDatabase';Operation='Get-Mailbox';ObjectGuid=$id1;BlocksCmdbQualification=$false},[pscustomobject]@{Category='Unexpected <finding>';Operation='Get-Mailbox';ObjectGuid='';BlocksCmdbQualification=$true})
    $section=New-SmartM365LocalMailboxIssueEmailSection -Issues $issues -IssuesCsvPath 'Synthetic & <path>.csv'
    Assert-True ($section.Html -like '*Affected objects*' -and $section.Html -like '*MissingRequiredMailboxDatabase*' -and $section.Html -like '*0 identified; 1 unbound occurrences*') 'Email conflated occurrences and objects.'
    Assert-True ($section.Html -like '*Unexpected &lt;finding&gt;*' -and $section.Html -like '*Synthetic &amp; &lt;path&gt;.csv*') 'Finding or path was not HTML-escaped.'
    Assert-True ($section.Html -notlike ('*'+$id1.ToString('D')+'*')) 'Aggregate section exposed raw native GUIDs.'
    Assert-True ($null -eq (New-SmartM365LocalMailboxIssueEmailSection -Issues @() -IssuesCsvPath 'Synthetic.csv')) 'Empty findings fabricated a section.'
}
Case 'Global coverage is enforced before combined publication and wired to completion' {
    $process=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'MailboxesProcessing2'},$true)
    Assert-True ($process.Extent.Text.IndexOf('Invoke-SmartM365LocalMailboxPopulationQuery') -lt $process.Extent.Text.IndexOf('Get-MailboxStatistics')) 'Partition happened only after enrichment.'
    $body=$ast.Extent.Text
    $guard=$body.IndexOf('Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($script:LocalMailboxObservedNativeGuids')
    $publish=$body.IndexOf('Export-CsvAtomic -InputObject $Global:ScriptOverallMailboxData -Path $globalCombinedCsvFile')
    Assert-True ($guard -gt 0 -and $guard -lt $publish -and $body.Contains('-PopulationCoverageComplete $script:LocalMailboxPopulationCoverageComplete')) 'Global guard was not wired before combined publication.'
    $mail=$ast.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-SmartM365ExchangeLocalMailboxReportEmailBody'},$true)
    Assert-True ($mail.Extent.Text.Contains('-LocalMailboxCollectionIssues') -and $mail.Extent.Text.Contains('-CollectionIssuesCsvPath')) 'Actual report omitted local findings.'
}
Case 'All-domain root query includes containers, not only first-level OUs' {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Process-SpecificDomain'},$true)
    Assert-True (-not $node.Extent.Text.Contains('-SearchScope OneLevel')) 'Forest root still excludes container recipients.'
    $statements=@($node.Body.ProcessBlock.Statements | Where-Object {
        ($_.Extent.Text -eq '$pathsForMailboxProcessing = @($distinguishedName)') -or
        ($_.Extent.Text -eq '$domainDataFromProcessing = @(MailboxesProcessing -IncludedLDAPPaths $pathsForMailboxProcessing)')
    })
    Assert-True ($statements.Count -eq 2) 'Domain-root acquisition topology changed.'
    $distinguishedName=$scope
    function MailboxesProcessing {param($IncludedLDAPPaths) $script:observedScopes=@($IncludedLDAPPaths)}
    & ([scriptblock]::Create(($statements.Extent.Text -join "`n")))
    Assert-True ($script:observedScopes.Count -eq 1 -and $script:observedScopes[0] -eq $scope) 'Domain DN was not used.'
}
Case 'Child-domain acquisition does not depend on a default-domain AD lookup' {
    Reset-Fixture
    $process=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'MailboxesProcessing2'},$true)
    $loop=$process.Find({param($n) $n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Extent.Text.Contains('Attempting to retrieve mailboxes from path:')},$true)
    Assert-True ($null -ne $loop) 'Scoped acquisition loop is missing.'
    $runScope=[scriptblock]::Create($loop.Extent.Text + "`n" + '$script:observedScopedMailboxes=@($AllMailbox)')
    function Write-LogMailboxesProcessing {param($Message)}
    function Get-ADDomain {throw 'Default-domain AD lookup must not run.'}
    function Get-ADOrganizationalUnit {throw 'Default-domain AD lookup must not run.'}
    $script:LocalMailboxForestPopulation=@{$rootScope=@($rootNative);$childScope=@($childNative)}
    $IncludedLDAPPaths=@($rootScope,$childScope);$totalOUs=2;$ouCounter=0;$AllMailbox=@();$LimitResultSize=$null
    & $runScope
    Assert-True ($script:observedScopedMailboxes.Count -eq 2 -and
        @($script:observedScopedMailboxes.Guid | Sort-Object -Unique).Count -eq 2) 'Child-domain mailbox projection was lost.'
    $IncludedLDAPPaths=@('');$totalOUs=1;$ouCounter=0;$AllMailbox=@()
    Assert-Throws {& $runScope} '*Mailbox acquisition scope is empty*'
    $IncludedLDAPPaths=@('DC=missing,DC=invalid');$ouCounter=0;$AllMailbox=@()
    Assert-Throws {& $runScope} '*Requested domain is outside the acquired forest population*'
}
Case 'Final scope uses computed acquisition proof and qualifications' {
    $command=$ast.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Set-SmartM365CmdbSourceScope'},$true)
    Assert-True ($command.Extent.Text.Contains('$script:CmdbLocalMailboxSourceComplete') -and $command.Extent.Text.Contains('-Qualifications $qualification.Qualifications')) 'Completeness evidence is not wired to receipt.'
    Assert-True (-not $ast.Extent.Text.Contains('_WithoutDuplicateSMTP.csv')) 'Destructive SMTP cleanup remains.'
}

# The real receipt writer, but only synthetic files and synthetic Core identity.
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('ExchangeLocal-Qualification-'+[guid]::NewGuid().ToString('N'))
$savedGlobals=@{}
foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')) {
    $existing=Get-Variable $name -Scope Global -ErrorAction SilentlyContinue
    $savedGlobals[$name]=@{Exists=($null -ne $existing);Value=$(if($existing){$existing.Value}else{$null})}
}
$receiptModule=$null
try {
    New-Item -ItemType Directory -Path $testRoot | Out-Null
    $global:SmartM365TenantKey='synthetic'; $global:SmartM365OrganizationKey='synthetic'
    $global:SmartM365EnvironmentKey='test'; $global:SmartM365TenantId='synthetic-tenant'
    $global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $receiptModule=New-Module -Name SyntheticExchangeQualification -ScriptBlock {
        param($HelperPath)
        . $HelperPath
        function Test-SmartM365MaxItemsMode { $false }
        function Get-SmartM365ScriptVersionFromFile { param($Path) 'synthetic-version' }
        function WriteLog {param($Message,$Level)}
        Export-ModuleMember -Function *
    } -ArgumentList (Join-Path $root 'Modules/SmartM365.Core/SmartM365-CmdbReceipt.ps1')
    $registry=Get-Content -LiteralPath (Join-Path $root 'Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt') -Raw | ConvertFrom-Json
    $producer=@($registry.Producers | Where-Object Script -eq 'SmartM365-Exchange-Local-Mailboxes-Inventory.ps1')[0]
    foreach ($file in $producer.Files) {
        $fixture=Join-Path $testRoot $file
        [IO.File]::WriteAllText($fixture, "TenantKey,ObjectGuid`r`nsynthetic,$id1`r`n", [Text.UTF8Encoding]::new($false))
        [void]$global:csvGeneratedPaths.Add($fixture)
    }
    $originalHashes=@{}
    foreach ($file in $producer.Files) { $originalHashes[$file]=(Get-FileHash (Join-Path $testRoot $file)).Hash }
    function Invoke-TestReceipt($Qualification) {
        & $receiptModule {param($Script,$Folder,$Scope,$Proof)
            Start-SmartM365CmdbSourceReceipt -ScriptPath $Script -SourceRootPath $Folder
            Set-SmartM365CmdbSourceScope -CompleteScope $Proof.CompleteScope -Scope $Scope -Qualifications $Proof.Qualifications
            Complete-SmartM365CmdbSourceReceipt -Status CompletedWithWarnings
        } $producer.Script $testRoot $producer.Scope $Qualification | Out-Null
        Get-Content -LiteralPath (Join-Path $testRoot $producer.Receipt) -Raw | ConvertFrom-Json
    }
    Case 'Real receipt preserves 45 known findings without claiming clean recipient health' {
        $issues=@(1..40 | ForEach-Object {[pscustomobject]@{Category='MissingPrimarySmtpAddress';Operation='Get-RemoteMailbox';NativeRecordRetained=$true;ObjectGuid=$id1.ToString('D')}})
        $issues+=@(1..3 | ForEach-Object {[pscustomobject]@{Category='MailboxStatisticsFailure';Operation='Get-MailboxStatistics'}})
        $issues+=@(1..2 | ForEach-Object {[pscustomobject]@{Category='RecipientAmbiguous';Operation='Get-MobileDevice'}})
        $receipt=Invoke-TestReceipt (Get-TestQualification -Issues $issues)
        Assert-True ($receipt.Status -eq 'Completed' -and -not $receipt.IsPartialInventory -and $receipt.Errors -eq 0) 'Complete population rejected.'
        Assert-True (($receipt.Qualifications -join ' ') -like '*occurrences: 45; blocking=0*') 'Findings disappeared from receipt.'
        Assert-True ($receipt.Files.Count -eq 2 -and $receipt.Files[0].Rows -eq 1) 'Native source records missing.'
    }
    Case 'Real receipt refuses unknown warning without changing native CSVs' {
        $receipt=Invoke-TestReceipt (Get-TestQualification -Issues @([pscustomobject]@{Category='UnknownFailure';Operation='Get-RemoteMailbox'}))
        Assert-True ($receipt.Status -eq 'Failed' -and $receipt.IsPartialInventory -and $receipt.Files.Count -eq 0) 'Unknown failure qualified.'
        foreach ($file in $producer.Files) { Assert-True ((Get-FileHash (Join-Path $testRoot $file)).Hash -eq $originalHashes[$file]) 'Failure mutated native files.' }
    }
    Case 'Real receipt refuses an omitted domain despite published CSVs' {
        $receipt=Invoke-TestReceipt (Get-TestQualification -Scopes @($scope,'DC=missing,DC=invalid'))
        Assert-True ($receipt.Status -eq 'Failed') 'Export presence replaced domain coverage proof.'
    }
    Case 'Real receipt refuses old export reuse despite complete query metadata' {
        $global:csvGeneratedPaths.Clear()
        $receipt=Invoke-TestReceipt (Get-TestQualification)
        Assert-True ($receipt.Status -eq 'Failed' -and $receipt.Error -like '*this run*') 'Old export qualified.'
    }
}
finally {
    if ($receiptModule) { Remove-Module $receiptModule -ErrorAction SilentlyContinue }
    foreach ($name in $savedGlobals.Keys) {
        if ($savedGlobals[$name].Exists) { Set-Variable $name -Scope Global -Value $savedGlobals[$name].Value }
        else { Remove-Variable $name -Scope Global -ErrorAction SilentlyContinue }
    }
    $resolved=[IO.Path]::GetFullPath($testRoot)
    if ((Split-Path $resolved -Parent) -eq [IO.Path]::GetTempPath().TrimEnd('\') -and (Split-Path $resolved -Leaf) -like 'ExchangeLocal-Qualification-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
$tests | Format-Table -AutoSize
$failedTests=@($tests | Where-Object {-not $_.Passed})
if ($failedTests.Count) { $failedTests | Format-List Name,Error; throw 'Offline Exchange local qualification failed.' }
[pscustomobject]@{Status='Passed';TestCount=$tests.Count;ProductionActions=0;PowerShell=[string]$PSVersionTable.PSVersion}

# SIG # Begin signature block
# MIIePgYJKoZIhvcNAQcCoIIeLzCCHisCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqhsDDUkgDP6uaLY+OWbHTrNZ
# FSWgghf7MIIEvTCCAyWgAwIBAgIQHm7vO8c44bNEOMjxAx/iaDANBgkqhkiG9w0B
# AQsFADBOMR4wHAYDVQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG
# 9w0BCQEWHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIuY29tMB4XDTI2MDcxMzA4
# MjIzNVoXDTI5MDcxMzA4MzIyOVowTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRo
# dWIuY29tMSwwKgYJKoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHVi
# LmNvbTCCAaIwDQYJKoZIhvcNAQEBBQADggGPADCCAYoCggGBALHul87REUsh5/Q1
# ao/EXb9Ko3OaMKr4ACBmmMQZn4kLqDOPBaOF7daL4vW6y1LX+eRuF4LYTTwqGp8X
# wkUYN9CFUcWg33T1yc2JoLDObV5bLjVOPZROxDyZN/oxDoEgrfDSs5lTyqTURtDZ
# itGvMJtdnHHmGTmZUmBNPmbU+sMuT0EIzFKiMV5xo7eB9J34GWWuw+BcgESjlYP1
# 6/bFKuYyJ987M74OI43m+G9AyibX2x7pIVnNmKFRYLMMVQVwliGgZ3xf3I6jHyvw
# mk6E9ra1W+IRuKnEN9bNZ0eJHqEWjsP78nemqaxLQrE9tjfycdHc+3yKNbVcryGl
# XB3TfS1t3IZ9Hp9DimeintTh/cH9qa5eCgOlsFc6wGazpV0wCLXSuw/ZCjvx3iXe
# D2bzud6MtS3ZjBi0s9ziwfXuau8qPK36ouSmvSszNgq0s89cqYZ/x/dMM+0NqZuB
# SJVnXI/OgP9JJoWIPuVqqm4NXa6z17uJxAcFjf3c6BrSmSh03QIDAQABo4GWMIGT
# MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzA/BgNVHREEODA2
# gR1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbYIVd29ya3BsYWNlY2xvdWRo
# dWIuY29tMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFFyDjgA0DO/F3zwJ3Iq4AhAn
# nYPYMA0GCSqGSIb3DQEBCwUAA4IBgQA4WVAdkeSfycNN9MaHYRFngkNL5yUWohLI
# zsbgK0ERCh95Qq0N1LFFZdvdnei50FEkx/4xe66LxGfiTdhkPywq10WgrsAhBmSs
# 5UhG7WFoX8A8o+hqSD6vGOEnl8o3auP66yNW5okyuEquxKAZRibl+pdp378dbU4q
# Eq9dRaA35ZSUexGGsEY/YQRSdIBtG9krmGTIFoGgynn3JyQ5jsHETQGVCCNLLPBH
# sgDbpFkpDTeSC2foig6UIGcYAyz04kYLgsseOTyOcK1U41f6UeyK7UiH6vFJWnZa
# 0mScsg59zVdgE7yGesDcW05jxXBNLjIPxGM8HU3pPgoFy908VCnpXSTOGA1j4pjQ
# UYTj18wQ4jkCI+ljbMLfyjNaAy5wNNAjbXIjzXZ/nV5Z3jL0xFGofcMrfKPLZC6a
# AcPijfgavytDtaj7Uu7pou9JeOaT4no/psG8Ks7XGqFQ+2vIIBz9VXaf652nzOBr
# cggBOm+1PWM4L6df2tZIe1335qlYUUYwggWNMIIEdaADAgECAhAOmxiO+dAt5+/b
# UOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMT
# G0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0z
# MTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/z
# G6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZ
# anMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7s
# Wxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL
# 2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfb
# BHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3
# JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3c
# AORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqx
# YxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0
# viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aL
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAP
# BgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYw
# eQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4
# X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc4
# 7/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0
# mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk
# 1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLB
# sln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGp
# n1eqXijiuZQwgga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3
# DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0
# ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGln
# aUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAy
# NSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphBcr48
# RsAcrHXbo0ZodLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6pvF4u
# GjwjqNjfEvUi6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNn
# QxqXmRinvuNgxVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEdgkFi
# DNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjUjsZv
# kgFkriK9tUKJm/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bRVFLe
# GkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeSLsJy
# goLPp66bkDX1ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz
# 7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36
# czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQb
# U2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0
# qVjPKOWug/G6X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgw
# BgEB/wIBADAdBgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0jBBgw
# FoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQM
# MAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6
# MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# Um9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJ
# KoZIhvcNAQELBQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/T8Ob
# XAZz8OjuhUxjaaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQE7jU
# /kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKC
# hHyfpzee5kH0F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA
# 0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gxdEkM
# x1NKU4uHQcKfZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qI
# ijanrUR3anzEwlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcytL5T
# TLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEBYTpt
# MSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+
# 7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiSuEtQ
# vLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZPubdc
# MIIG7TCCBNWgAwIBAgIQCE/cM09+RU7bww+P+ZIYNTANBgkqhkiG9w0BAQsFADBp
# MQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMT
# OERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2
# IDIwMjUgQ0ExMB4XDTI2MDgwNTAwMDAwMFoXDTM3MTEwNDIzNTk1OVowYzELMAkG
# A1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdp
# Q2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjYgMTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALZ7pvLJ/s1K+NSbTGWz/TjG
# MPh8CQ6RucZCLv5anHzWJjF/NWJrFIhy24fcpKXlgRiky4WAawDfU3YP0BMxt9l3
# Dm5oCG5Z69AqEN1kgHg2epx+l+lZBcmJCcN0ASURML5uFIS80sZsDwO3BSkUxDjL
# JhBI+qiZP3aixAC/qEGLjsBNlLol9VZ7pfGEXiMlneJIC5/YKuizVzNFKZZEeoy/
# 0B8Zm+nzKBgSWG52lCO1w+nCg6XpCtklTJXeIg283hw7TmmsZXR+SMbjbrEOvZ3f
# P2VxIgeR28Y90ZStd3F9VuA5RVynb/whITPAo9b75Zr4Ta6Mj3URm26QZYMn/Fnb
# uTegcoRcFEZ9FOqM5T6MTdtr/n74lIT/ug0eeOzmZ6QTFg33otX+bFRsIolvykE1
# jive4PuESaT8zzVeFWDAMDtozNgLctkGD1ZjkEyZtJrLl5ya0m5doH/ScpaZCZVl
# 6pNUOCybMc/kxC6EAmSJY24L0yYKD1Nkddsnb/ItVKi/2nXpQNMu1PT5prW83vV8
# d67WowuUs0HdY4H8AMLGvdL/WHEj3ZnqMqAQQP9u3Ai9t+5eQ02GDwy0ODjdzi0x
# lp70W+ow63/0++YDEX1M0iwgUHwbrJvfpklkZQvw3+kv3vUPItdwroczk9icflf5
# 5W1zOEKAcJVAIXpcMCU9AgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1Ud
# DgQWBBQUyWOKMC7USvtulPPm40B+9ezN4jAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYy
# MDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2
# MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0G
# CSqGSIb3DQEBCwUAA4ICAQCNxTphHp1SCt+ZrAmAfn0oQLFr0mLywSLaDXQIENoy
# KqxrFbJblzCVP/pkXmwXOdrOpWygLzlT12os5ipDCy35RBCg2UMeApEtrfGhz45F
# 4Wt4WGdNdIbRWt3YTYJmpR+b7lr4d7Uwn+H600u4D7RnOGf8Wj4UNgAdZkfHhHv1
# mx9EVh71SJelcEN/oORSjXzdjfw1iZH9d8Nh/thn6hH23d+VsPAr6GAYyzSA02nX
# D1nYLI7Ijmiv+xLCiYC41DSFYL3GhTiy0PxpawPtGRyaBVGzq+UiTfM8pD7KVyF5
# aQyWP4KhVGUUTnmm/RlYJoW3TiXA/+t0YcT2oRVBm3JETjajHug2AL+v5jhtKVnd
# 3D0rbHXEu27o+Q8p4sEWPMqKDB+qbceb6T/6WcwTwXmQ9lOCLLYcsQeSWmvKqzpA
# ec9etE14jOQAzLKWdE3w/TCaKtLRaRT7LCkRYVnhA2D73FLje1O5b3HR5eHs0NzU
# /+xX7NbEdcofy0W3Wdwd1XOqtlpg/JgwtKfZM5dqO94lbUveOiJBI+xZEbGRsMNb
# XmMREUTgu+Oca7Y73MPWcslIx2VhkSKSXjDbD6rgg39H5Mh7QfieAIjWagkJNt68
# Yfim6cjEzVSiLSeZfdkr5dtFPTW6jATlWJdYeeDRGCyatf8R1hSjzSvdN8yWQPT9
# gzGCBa0wggWpAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29t
# MSwwKgYJKoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQ
# Hm7vO8c44bNEOMjxAx/iaDAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAig
# AoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgEL
# MQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUki0kGnCmNwNsN/vjXeVz
# +2AN5Q0wDQYJKoZIhvcNAQEBBQAEggGAT/BDeZZeOvt0UUPyANvR/mvGwyMB3CQ7
# CoBpfS+rSWuv4hdb/1jxK4XLAStGx5SJbMMuYXtUv2shX+d+KPMAAeZKsOFyeSbO
# R4fjx0jOnH+yw7+689sCwrGdpEZH2+Gw9DkxNJ8AmsZLcQblZ7huHcqt3SKH2u5X
# eGPY88V5rJ/ikag/wl7kZfzwbMxU5lreS42NTw10uAeiIOGZENRnrfRkroRyU30L
# fn2QrUW4MiIje4igr7zV2VIvtAFuvvobwU8pbLH4XbaSbvX3XqZLjnqbggubYHhD
# W+fHaCBFZe7dS0Sk1EwkVQJN5VhuNc8rz1WwGqREL9cCadcTcdr9vSQqQpsr0NBu
# /ZhB6mUopTxfxapl0fT4pmV632o1t1ujBKUyqYCxa3UAnnOwvQ1sKlCKfOXvJLOQ
# vqOXn0z1CoQjvc5nn/VcMgxBj0x50jUz9gz8PcoPGunMVGh1+hJtd/+fw6xbJuuD
# Ea14a8NX4E2GK+lxFAr5ABUZnP1FXZ5HoYIDJjCCAyIGCSqGSIb3DQEJBjGCAxMw
# ggMPAgEBMH0waTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0
# MDk2IFNIQTI1NiAyMDI1IENBMQIQCE/cM09+RU7bww+P+ZIYNTANBglghkgBZQME
# AgEFAKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTI2MTAwNjE0NDgxM1owLwYJKoZIhvcNAQkEMSIEIOaZg28tnL4y7YP639BEKk5j
# MZuUZjVpWNy6NskdpJXpMA0GCSqGSIb3DQEBAQUABIICAJE0J8uxbAvVRkH2Ctj6
# TOqk863CkYQ0jSrvuN/9LAkU5seW3ElqkacNC7OA3VWqYD6b5LpPzqYOWI6/KkHS
# qQNFw5r0botYf5+SkJjYANramd2AZVV8Smspm7L5mkDYSxJ8MfOmXejUxyMBi6Sp
# 4hSx06bbN2kpRKqh505P/GRPnD0Mm2CmNrMcBtMhnIk+p7GuvUMGeYm7NgzxtmER
# cilknZPTrR874UDk8nT2BMLw88RyZTX/YTIWBOnDAg4gReIsSSDs6iBat9me9Spq
# e8FE3EGFNVetdxgX7Yrcc2v9wuJC8XBsJurj4HKIoSqKcmuqS3SVR2rpCHQNQpFd
# zJ4/Vw61D45dlLx2JS5ag77DTCzePiP7dMFhVC8pO5cVjMOKLPb1lHd1+62+G/g5
# rmKmPUAcOrmmxtST8NHY6/1ZrvsXcA/73KADzIlup97k6Ggy4AgV1R86tdKL1RIn
# DpkzLQ+U18wrOvUFUoOLROl59mDPaO+mqVaC5uHIQzRLC6RW7Q2Esg/Ncz1BHuF0
# GQAVaNmmag1IXvoy+yW9Q+yGbBY42WGhMmMU6GpvKRxcoPfzJz/Qjc4DSx8Bx78D
# Gy1dhz2fa8H7TM/9AdtVz+jATaCUQSpg/1wppyYqJ/1E9WZcAt9b9bndmwdA8rSD
# x6OPd7BDmVq/YZmTjqpBKZqX
# SIG # End signature block
