#Requires -Version 5.1
<#
.SYNOPSIS
Offline Exchange mailbox scope, native identity and quality classification tests.
.DESCRIPTION
Extracts functions and statements through the AST. All Exchange queries are
mocked; no tenant context, collector, module, export or external action runs.
.VERSION
1.0.2
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
    'Get-SmartM365LocalMailboxColumns')) {
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
    $script:LocalMailboxAcquisitionMode='Domain'
    $script:queryCount=0
    $script:LocalMailboxObservedNativeGuids=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $script:LocalMailboxOwnedNativeGuids=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
}
function ConvertTo-SmartM365EmailHtmlText {param($Value) [Net.WebUtility]::HtmlEncode([string]$Value)}
function Get-Mailbox {
    [CmdletBinding()]
    param($ResultSize,$OrganizationalUnit,[switch]$ReadFromDomainController)
    $script:queryCount++
    $script:queryArguments=@{ResultSize=$ResultSize;Scope=$OrganizationalUnit;ReadFromDomainController=[bool]$ReadFromDomainController}
    if ($script:mockFailure) { throw 'Synthetic acquisition failure.' }
    foreach ($warning in $script:mockWarnings) { Write-Warning $warning }
    $script:mockRows
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
Case 'Full forest makes one fresh unlimited query and partitions every domain without new queries' {
    Reset-Fixture; $script:DetectAllDomains=$true; $script:LocalMailboxExpectedScopes=@($rootScope,$childScope,'DC=empty,DC=invalid')
    $script:mockRows=@($rootNative,$childNative)
    Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes $script:LocalMailboxExpectedScopes
    $script:mockFailure=$true
    $parent=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $rootScope)
    $child=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope $childScope)
    $empty=@(Invoke-SmartM365LocalMailboxPopulationQuery -Scope 'DC=empty,DC=invalid')
    Assert-True ($script:queryCount -eq 1 -and $script:queryArguments.ResultSize -eq 'Unlimited' -and -not $script:queryArguments.Scope -and $script:queryArguments.ReadFromDomainController) 'Forest query is scoped, limited, stale or repeated.'
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
Case 'Complete empty forest has real zero-domain evidence and never starts extra queries' {
    Reset-Fixture
    Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope,$childScope)
    Assert-True ($script:LocalMailboxForestPopulation.Count -eq 2 -and $script:LocalMailboxAcquisitions[0].DomainCounts[$rootScope] -eq 0) 'Empty domain was omitted or fabricated.'
    Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($script:LocalMailboxObservedNativeGuids) -OwnedGuids @($script:LocalMailboxOwnedNativeGuids) -ExportRows @()
    $script:LocalMailboxAcquisitions[0].ProjectionCompleted=$true
    $proof=Get-SmartM365LocalMailboxQualification -ExpectedScopes @($rootScope,$childScope) -Acquisitions $script:LocalMailboxAcquisitions.ToArray() -RemoteAcquisitionComplete $true -Issues @() -PopulationCoverageComplete $true -AcquisitionMode Forest
    Assert-True ($proof.CompleteScope -and $script:queryCount -eq 1) 'Successful empty forest was rejected.'
}
foreach ($defect in @('Scoped','Limited','NotFresh','NoProjection','MissingDomain','ExtraDomain','WrongCount','NoCoverage','RepeatedQuery','UnknownQuery')) {
    Case "Forest qualification rejects invalid acquisition proof / $defect" {
        Reset-Fixture; $script:mockRows=@($rootNative)
        Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes @($rootScope,$childScope)
        $q=$script:LocalMailboxAcquisitions[0]; $q.ProjectionCompleted=$true
        $coverage=$true
        switch ($defect) {
            Scoped {$q.Scope=$rootScope}
            Limited {$q.ResultSize='1'}
            NotFresh {$q.ReadFromDomainController=$false}
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
Case 'Production full acquisition precedes domain processing and empty forest publication has explicit schema' {
    $body=$ast.Extent.Text
    Assert-True ($body.IndexOf('Initialize-SmartM365LocalMailboxForestPopulation -ForestScopes $script:LocalMailboxExpectedScopes') -lt $body.IndexOf('foreach ($domain in $domainsToProcess)')) 'Full mode still acquires inside each domain.'
    Assert-True ($body.Contains('-AcquisitionMode $script:LocalMailboxAcquisitionMode') -and $body.Contains('Export-SmartM365EmptyLocalMailboxPopulation -Path $globalCombinedCsvFile') -and $body.Contains('Export-SmartM365EmptyLocalMailboxPopulation -Path $perDomainCsvFullPath -DomainScope $distinguishedName')) 'Forest proof or empty schema was not wired.'
    $remote=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SmartM365ExchangeRemoteMailboxInventory'},$true)
    foreach ($query in $remote.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-RemoteMailbox'},$true)) { Assert-True ($query.Extent.Text.Contains('-ReadFromDomainController')) 'Remote query lost freshness flag.' }
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
        ($_.Extent.Text -eq '$domainDataFromProcessing = MailboxesProcessing -IncludedLDAPPaths $pathsForMailboxProcessing')
    })
    Assert-True ($statements.Count -eq 2) 'Domain-root acquisition topology changed.'
    $distinguishedName=$scope
    function MailboxesProcessing {param($IncludedLDAPPaths) $script:observedScopes=@($IncludedLDAPPaths)}
    & ([scriptblock]::Create(($statements.Extent.Text -join "`n")))
    Assert-True ($script:observedScopes.Count -eq 1 -and $script:observedScopes[0] -eq $scope) 'Domain DN was not used.'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD/5DSWP6VEZk2a
# Yfbk0dy6xCeWONDweYGj5SMYwlYy2qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGYA4ubXqdXPb40KIlkr5vbe5emh/y8wXwW7a1rJr53IMA0GCSqG
# SIb3DQEBAQUABIIBgKQSMaHDuTUTEKxMfcVNVYPKtbb8xF0PAhVtwDVF+IOgsCgH
# 6qYzliu+Rwh+OIP/QUVFzFSmiDgwNf89/4HuQhslJUXRxfWyBZkMfr10gLg0Xwhl
# IuMUHKTFnZLveQ/JmQu2A+C/ZQEH0vBdBvB+yBQo5FaHPtTwOxkC8LKZzZHsvFNu
# QQBcrEWmYd2H4dR9UVCWYX/92zhMAa2pB1GQk+7m/nFnP9TECn6DQ5rJuUGTXVrQ
# eZkWDnCqNI97f7E2B6SgpL+1sEMMaB3/BoFgFWkZ0jGNYwoPnrrcnLzKid90tJ6n
# hTdrzD+rX7ampTPsUm4Gb8NQTffYcupoRANlCuZwlEe2dGXQqwJyQoSnou9baq1c
# Ad3sQ7DUXfppkAX2dcN1ysOduhLxq8RTXP5tMdpWhTvEFtmw3dVtINQOoXHq/MmU
# Si9doMMcBos3CwDlvzqsDS16Jn+D9M7WtNPrz+B7PszQg1cA2OyZFguDztYgxZCd
# +vN2NjBgYxsM0NnRCqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDUwOTMy
# MDBaMC8GCSqGSIb3DQEJBDEiBCDRFM7GXhkFo4asCsG76u9eoqnAy/469VBRAFqt
# fR/mhjANBgkqhkiG9w0BAQEFAASCAgCzF8CzafmnB+SLd0PC9Q8yoU3cu8SVSICr
# khHBa0z1XG5ujWNrWkZzZwJGimqmsEyAyvI9n1p2Zu2zgXUpEPZTrpRJMZxuJB4O
# nFS4c6VmcPk3kES5pl1PH0mnmNJi2D/WsWhuCuI/PyDoYZTbD9Tx+DpLxNzCu1v/
# 0jBO5pjLFNKDCkK09M4PKJMfMU7sie6wPlHF5uAYkNepL3L3Gsv1izLdJAVmNNaZ
# ML7KZERGZijkFJs3pJKRYbI0AcbzYMwrUd1rK1q+kdGnoJ73WqYT7UtqCjQpeUhP
# rFeeQkHLE4NmVMuxKRr1DlirwRhOzU1N3nIYHeecpdQWftsLxHjmcpNbZiV+iacn
# JUaD/1JvxFY9d/7rS3H+2i951gFqXwGB1/FsxbGczXwyXkjvNwMxSJ6Cyy+fPNYq
# +AwnM2ytaIq3XlXOMbuT+bcaWfbGqVhwlTjnrn/iyAid0UuJlcP2JwBptDpZJPtB
# FiMQ8i33TCnPVzVDd9Bfv3Yn3/QR6kq/fHRpeJPUvp/WYt/5bKt8JgMestj/YLL5
# 6Fm6gID9Ml/5JPwvE2vxa6xEPt9vrjD03ycTp76VBJs1+vI4Ds/b3ZzGfY+oMTTQ
# XKJnj+WBZ/N+APZf7LGubETGeCBfxl84M7WrcLRbweD8iIKxc2dumV6T6hiXHVAk
# T8V6yvyJow==
# SIG # End signature block
