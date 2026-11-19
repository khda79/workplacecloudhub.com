#Requires -Version 5.1
<#
.SYNOPSIS
Offline Exchange mailbox scope, native identity and quality classification tests.
.DESCRIPTION
Extracts functions and statements through the AST. All Exchange queries are
mocked; no tenant context, collector, module, export or external action runs.
.VERSION
1.0.8
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
$reportAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'SmartInventory/Common/SmartM365.ReportRenderers.ps1'),[ref]$null,[ref]$null)
$tests=New-Object 'Collections.Generic.List[object]'
function Assert-True([bool]$Condition,[string]$Message) { if (-not $Condition) { throw $Message } }
function Assert-Throws([scriptblock]$Body,[string]$Pattern) {
    $failure=$null
    try { & $Body | Out-Null } catch { $failure=$_ }
    Assert-True ($null -ne $failure -and $failure.Exception.Message -like $Pattern) "Expected failure: $Pattern; observed: $(if ($null -ne $failure) { $failure.Exception.Message } else { 'none' })"
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
    'Get-SmartM365LocalMailboxColumns','Invoke-SmartM365RemoteMailboxPopulationQuery',
    'Get-SmartM365ExchangeMailboxDailySnapshotColumns','New-SmartM365ExchangeMailboxDailySnapshot',
    'Get-SmartM365ExchangeMailboxDailySnapshotHistory','New-SmartM365ExchangeMailboxDailyDiffSection',
    'Get-SmartM365MailboxReportValue','Get-SmartM365MailboxReportDomain','ConvertTo-SmartM365MailboxReportRecord',
    'Publish-SmartM365ExchangeMailboxDailySnapshot')) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if (-not $node) {
        $node=$reportAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    }
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
    $script:ExpectedNoOnPremRemoteMailboxes=$false
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
    Assert-True ($script:LocalMailboxIssues.Count -eq 2) 'Conflict lost a native identity.'
    Assert-True (@($script:LocalMailboxIssues | Where-Object BlocksCmdbQualification).Count -eq 0) 'Retained identities were treated as lost acquisition.'
    Assert-True (@($script:LocalMailboxIssues.ObjectGuid | Sort-Object -Unique).Count -eq 2) 'Conflict identities were not retained.'
    Assert-True (Get-TestQualification -Issues @($script:LocalMailboxIssues.ToArray())).CompleteScope 'Complete conflicting recipients were rejected.'
}
foreach ($badId in @('', 'not-a-guid', [string][guid]::Empty)) {
    Case "SMTP conflict with unproven GUID remains blocking / $badId" {
        Reset-Fixture
        Find-SmartM365MailboxSmtpConflict -Rows @([pscustomobject]@{ObjectGuid=$id1;PrimarySmtpAddress='same@synthetic.invalid'},[pscustomobject]@{ObjectGuid=$badId;PrimarySmtpAddress='same@synthetic.invalid'})
        Assert-True (-not (Get-TestQualification -Issues @($script:LocalMailboxIssues.ToArray())).CompleteScope) 'Unproven population was qualified.'
        Assert-True (@($script:LocalMailboxIssues | Where-Object NativeRecordRetained).Count -eq 0) 'Partially identified conflict was accepted.'
    }
}
Case 'Retained SMTP conflict does not hide failed acquisition' {
    Reset-Fixture
    Find-SmartM365MailboxSmtpConflict -Rows @([pscustomobject]@{ObjectGuid=$id1;PrimarySmtpAddress='same@synthetic.invalid'},[pscustomobject]@{ObjectGuid=$id2;PrimarySmtpAddress='same@synthetic.invalid'})
    $badQuery=[pscustomobject]@{Scope=$scope;QueryCompleted=$false;ProjectionCompleted=$false;Rows=2}
    Assert-True (-not (Get-TestQualification -Issues @($script:LocalMailboxIssues.ToArray()) -Queries @($badQuery)).CompleteScope) 'Failed acquisition became qualified.'
    Assert-True (-not (Get-TestQualification -Issues @($script:LocalMailboxIssues.ToArray()) -Remote $false).CompleteScope) 'Missing remote population became qualified.'
}
Case 'SMTP allowlist requires the exact reconciliation operation' {
    Assert-True (Get-SmartM365LocalMailboxIssueImpact -Category ConflictingPrimarySmtpAddress -Operation Get-Mailbox -NativeRecordRetained $true).BlocksCmdbQualification 'Unexpected operation became qualified.'
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
Case 'Mailbox email renders seven summary KPI cards without a metric table' {
    $reportNode=$reportAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-SmartM365ExchangeLocalMailboxReportEmailBody'},$true)
    Assert-True ($null -ne $reportNode) 'Mailbox report email function is missing.'
    Set-Item Function:script:New-SmartM365ExchangeLocalMailboxReportEmailBody ([scriptblock]::Create($reportNode.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
    $modulePath=Join-Path $root 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'
    $moduleTokens=$null;$moduleErrors=$null
    $moduleAst=[Management.Automation.Language.Parser]::ParseFile($modulePath,[ref]$moduleTokens,[ref]$moduleErrors)
    Assert-True (@($moduleErrors).Count -eq 0) 'Windows PowerShell 5.1 mail template does not parse.'
    $mailNode=$moduleAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-SmartM365EmailBody'},$true)
    Assert-True ($null -ne $mailNode) 'Shared mail template is missing.'
    Set-Item Function:script:New-SmartM365EmailBody ([scriptblock]::Create($mailNode.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
    foreach ($helperName in @('New-SmartM365EmailKpiGridHtml','New-SmartM365EmailTableHtml')) {
        $helperNode=$moduleAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $helperName},$true)
        Assert-True ($null -ne $helperNode) "Missing shared mail helper: $helperName"
        Set-Item "Function:script:$helperName" ([scriptblock]::Create($helperNode.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
    }
    function Get-SmartM365MailboxReportCsvRowCount {param($Path) 0}
    $script:Tenant='synthetic'
    $row=[pscustomobject]@{DomainName='synthetic & invalid';TotalMailboxCount=12;TotalLocalMailboxCount=5;TotalRemoteMailboxCount=7;EnabledAccounts=9;DisabledAccounts=3;TotalLocalMailboxSizeGB=17.5}
    $html=New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows @($row) -DailyStatsCsv 'synthetic-daily.csv' -SummaryCsv 'synthetic-summary.csv' -Title 'Synthetic report'
    Assert-True ([regex]::Matches($html,'font-size:28px;line-height:32px;font-weight:800').Count -eq 7) 'Expected seven KPI cards.'
    foreach ($label in @('Domains','Inventory objects','Local mailboxes','Remote mailboxes','Enabled accounts','Disabled accounts','Local mailbox size')) {
        Assert-True ($html.Contains(('>{0}</div>' -f $label))) "Missing KPI card: $label"
    }
    Assert-True ($html.Contains('colspan="2"') -and ($html.Contains('>17,50</div>') -or $html.Contains('>17.50</div>'))) 'Local size card or value is missing.'
    Assert-True ($html -notmatch '>Metric</th>' -and $html -match '(?s)>Summary</div>.*>Domain summary \(raw inventory objects\)</div>') 'Legacy summary table remained or section order changed.'
    Assert-True ($html.Contains('synthetic &amp; invalid')) 'Domain names were not HTML-escaped.'
    $snapshot=New-SmartM365ExchangeMailboxDailySnapshot -ReportRows @($row) -Records @([pscustomobject]@{RecipientTypeDetails='UserMailbox';TotalItemSizeKnown=$true},[pscustomobject]@{RecipientTypeDetails='SharedMailbox';TotalItemSizeKnown=$true})
    $diff=New-SmartM365ExchangeMailboxDailyDiffSection -Current $snapshot -History @()
    $withDiff=New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows @($row) -DailyStatsCsv 'synthetic-daily.csv' -SummaryCsv 'synthetic-summary.csv' -DailyDiffSection $diff -Title 'Synthetic report'
    Assert-True ($withDiff -match '(?s)>Summary</div>.*>Raw inventory diff against previous scans</div>.*>Domain summary \(raw inventory objects\)</div>') 'Daily difference section was not placed below KPI cards.'
}
Case 'Daily mailbox snapshot marks incomplete size statistics without losing population counts' {
    $zeroSize=ConvertTo-SmartM365MailboxReportRecord -Row ([pscustomobject]@{RecipientTypeDetails='UserMailbox';ADDomain='synthetic.invalid';TotalItemSizeToMB=0})
    Assert-True ($zeroSize.TotalItemSizeKnown -and $zeroSize.TotalItemSizeToMB -eq 0) 'A valid zero-size mailbox was marked as missing statistics.'
    $report=@([pscustomobject]@{TotalMailboxCount=3;TotalLocalMailboxCount=2;TotalRemoteMailboxCount=1;EnabledAccounts=2;DisabledAccounts=1;TotalLocalMailboxSizeGB=1.25})
    $records=@(
        [pscustomobject]@{RecipientTypeDetails='UserMailbox';TotalItemSizeKnown=$true}
        [pscustomobject]@{RecipientTypeDetails='SharedMailbox';TotalItemSizeKnown=$false}
        [pscustomobject]@{RecipientTypeDetails='RemoteUserMailbox';TotalItemSizeKnown=$false}
    )
    $snapshot=New-SmartM365ExchangeMailboxDailySnapshot -ReportRows $report -Records $records
    Assert-True ($snapshot.TotalMailboxCount -eq 3 -and $snapshot.TotalLocalMailboxCount -eq 2 -and $snapshot.LocalSizeKnownMailboxCount -eq 1 -and -not $snapshot.LocalSizeCoverageComplete) 'Missing local size was treated as complete or removed a mailbox.'
    Assert-True ($snapshot.TotalLocalMailboxSizeGB -eq '1.25' -and $snapshot.SchemaVersion -eq '1') 'Daily snapshot value or schema changed.'
}
Case 'Daily mailbox diff uses the prior completed day and exact J-7 and J-30 dates' {
    $current=[pscustomobject]@{SnapshotDate='2026-10-07';DomainCount=10;TotalMailboxCount=100;TotalLocalMailboxCount=20;TotalRemoteMailboxCount=80;EnabledAccounts=90;DisabledAccounts=10;TotalLocalMailboxSizeGB=12.5;LocalSizeCoverageComplete=$false}
    $history=@(
        [pscustomobject]@{SnapshotDate='2026-10-06';GeneratedAt=[datetimeoffset]'2026-10-06T03:00:00+02:00';DomainCount=10;TotalMailboxCount=99;TotalLocalMailboxCount=21;TotalRemoteMailboxCount=78;EnabledAccounts=90;DisabledAccounts=9;TotalLocalMailboxSizeGB=12.0;LocalSizeCoverageComplete=$true}
        [pscustomobject]@{SnapshotDate='2026-09-30';GeneratedAt=[datetimeoffset]'2026-09-30T03:00:00+02:00';DomainCount=10;TotalMailboxCount=95;TotalLocalMailboxCount=20;TotalRemoteMailboxCount=75;EnabledAccounts=85;DisabledAccounts=10;TotalLocalMailboxSizeGB=11.0;LocalSizeCoverageComplete=$true}
        [pscustomobject]@{SnapshotDate='2026-09-07';GeneratedAt=[datetimeoffset]'2026-09-07T03:00:00+02:00';DomainCount=9;TotalMailboxCount=90;TotalLocalMailboxCount=18;TotalRemoteMailboxCount=72;EnabledAccounts=82;DisabledAccounts=8;TotalLocalMailboxSizeGB=10.0;LocalSizeCoverageComplete=$true}
    )
    $section=New-SmartM365ExchangeMailboxDailyDiffSection -Current $current -History $history
    Assert-True ($section.Html.Contains('Previous (2026-10-06)') -and $section.Html.Contains('J-7 (2026-09-30)') -and $section.Html.Contains('J-30 (2026-09-07)')) 'Comparison dates are not exact.'
    Assert-True ($section.Html.Contains('>+1</strong>') -and $section.Html.Contains('>-1</strong>')) 'Signed count deltas were lost.'
    $sizeRow=[regex]::Match($section.Html,'(?s)<tr><td[^>]*>Local mailbox size \(GB\)</td>.*?</tr>').Value
    Assert-True ([regex]::Matches($sizeRow,'>n/a</td>').Count -eq 3) 'Incomplete size statistics produced a numeric delta.'
    $current.LocalSizeCoverageComplete=$true
    $complete=New-SmartM365ExchangeMailboxDailyDiffSection -Current $current -History $history
    $completeSizeRow=[regex]::Match($complete.Html,'(?s)<tr><td[^>]*>Local mailbox size \(GB\)</td>.*?</tr>').Value
    Assert-True ($completeSizeRow -match '\+0[,.]50' -and $completeSizeRow -notmatch '>n/a</td>') 'Complete size statistics did not produce a decimal delta.'
    $missing=New-SmartM365ExchangeMailboxDailyDiffSection -Current $current -History @($history[0])
    Assert-True ($missing.Html.Contains('J-7 (2026-09-30)') -and $missing.Html.Contains('J-30 (2026-09-07)') -and ([regex]::Matches($missing.Html,'>n/a</td>').Count -eq 14)) 'Missing exact dates were replaced with nearby scans.'
}
Case 'Daily mailbox history replaces the same date and rejects incompatible schema' {
    $folder=Join-Path ([IO.Path]::GetTempPath()) ('MailboxDailyHistory-'+[guid]::NewGuid().ToString('N'))
    New-Item -Path $folder -ItemType Directory -Force | Out-Null
    try {
        $path=Join-Path $folder 'Exchange_OnPrem_Mailboxes_DailySummary.csv'
        $makeRow={param($date,$generated,$count)
            [pscustomobject][ordered]@{TenantKey='synthetic';OrganizationKey='synthetic';EnvironmentKey='test';TenantId='synthetic';SnapshotDate=$date;GeneratedAt=$generated;SchemaVersion='1';CompleteScope='True';DomainCount=1;TotalMailboxCount=$count;TotalLocalMailboxCount=1;TotalRemoteMailboxCount=($count-1);EnabledAccounts=$count;DisabledAccounts=0;TotalLocalMailboxSizeGB='1.00';LocalSizeKnownMailboxCount=1;LocalSizeCoverageComplete='True'}
        }
        $old=& $makeRow '2026-10-06' '2026-10-06T03:00:00+02:00' 4
        $sameDay=& $makeRow '2026-10-07' '2026-10-07T03:00:00+02:00' 5
        @($old,$sameDay) | Export-Csv -LiteralPath $path -Delimiter ';' -Encoding UTF8 -NoTypeInformation
        $read=@(Get-SmartM365ExchangeMailboxDailySnapshotHistory -Path $path -TenantKey 'synthetic')
        Assert-True ($read.Count -eq 2 -and $read[0].TotalMailboxCount -eq 4) 'Qualified history was not read.'
        function Export-CsvAtomic {param($InputObject,$Path,$Columns,$Encoding,$Delimiter) $script:publishedDailyRows=@($InputObject)}
        $global:SmartM365TenantKey='synthetic'
        $replacement=& $makeRow '2026-10-07' '2026-10-07T09:00:00+02:00' 6
        Publish-SmartM365ExchangeMailboxDailySnapshot -Snapshot $replacement -Path $path
        Assert-True ($script:publishedDailyRows.Count -eq 2 -and @($script:publishedDailyRows | Where-Object { $_.SnapshotDate -eq '2026-10-07' }).Count -eq 1 -and @($script:publishedDailyRows | Where-Object { $_.TotalMailboxCount -eq 6 }).Count -eq 1) 'Same-day rerun was appended twice.'
        Set-Content -LiteralPath $path -Value 'Incompatible;Header' -Encoding UTF8
        Assert-Throws {Get-SmartM365ExchangeMailboxDailySnapshotHistory -Path $path -TenantKey 'synthetic'} '*schema is incompatible*'
    }
    finally {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $folder -Force -ErrorAction SilentlyContinue
        Remove-Variable -Name SmartM365TenantKey -Scope Global -ErrorAction SilentlyContinue
    }
}
Case 'Global coverage is enforced before combined publication and wired to completion' {
    $process=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'MailboxesProcessing2'},$true)
    Assert-True ($process.Extent.Text.IndexOf('Invoke-SmartM365LocalMailboxPopulationQuery') -lt $process.Extent.Text.IndexOf('Get-MailboxStatistics')) 'Partition happened only after enrichment.'
    $body=$ast.Extent.Text
    $guard=$body.IndexOf('Assert-SmartM365MailboxDomainCoverage -ObservedGuids @($script:LocalMailboxObservedNativeGuids')
    $publish=$body.IndexOf('Export-CsvAtomic -InputObject $Global:ScriptOverallMailboxData -Path $globalCombinedCsvFile')
    Assert-True ($guard -gt 0 -and $guard -lt $publish -and $body.Contains('-PopulationCoverageComplete $script:LocalMailboxPopulationCoverageComplete')) 'Global guard was not wired before combined publication.'
    $mail=$ast.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-SmartM365ExchangeLocalMailboxReportEmailBody'},$true)
    Assert-True ($mail.Extent.Text.Contains('@mailInputs') -and $body.Contains('LocalMailboxCollectionIssues=@($script:LocalMailboxIssues.ToArray())') -and $body.Contains('CollectionIssuesCsvPath=$issueCsvPath')) 'Actual report omitted local findings.'
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
    $assignment=$ast.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$script:LocalMailboxDailySummaryQualified' -and $n.Extent.Text.Contains('$script:CmdbLocalMailboxSourceComplete')},$true)
    Assert-True ($null -ne $assignment -and $assignment.Extent.Text.Contains('$script:CmdbLocalMailboxSourceComplete') -and $assignment.Extent.Text.Contains('$MaxItems -eq 0')) 'Qualified daily summary guard lost full-scope proof.'
    Assert-True ($command.Extent.Text.Contains('-CompleteScope $script:LocalMailboxDailySummaryQualified') -and $command.Extent.Text.Contains('-Qualifications $qualification.Qualifications')) 'Completeness evidence is not wired to receipt.'
    $reportCall=$ast.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Invoke-SmartM365ExchangeLocalMailboxReport' -and $n.Extent.Text.Contains('-UseCurrentInventoryData:$true')},$true)
    Assert-True ($null -ne $reportCall -and $assignment.Extent.StartOffset -lt $reportCall.Extent.StartOffset) 'Completeness proof is calculated only after mail generation.'
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
        $receiptPath=Join-Path $testRoot $producer.Receipt
        $previousHash=if(Test-Path -LiteralPath $receiptPath){(Get-FileHash -LiteralPath $receiptPath).Hash}else{''}
        & $receiptModule {param($Script,$Folder,$Scope,$Proof)
            Start-SmartM365CmdbSourceReceipt -ScriptPath $Script -SourceRootPath $Folder
            Set-SmartM365CmdbSourceScope -CompleteScope $Proof.CompleteScope -Scope $Scope -Qualifications $Proof.Qualifications
            Complete-SmartM365CmdbSourceReceipt -Status CompletedWithWarnings
        } $producer.Script $testRoot $producer.Scope $Qualification | Out-Null
        $run=Get-Content -LiteralPath ($receiptPath -replace '\.current\.json\.txt$','.run.json.txt') -Raw | ConvertFrom-Json
        if ($run.Status -eq 'Failed') {
            if($previousHash){Assert-True ((Get-FileHash -LiteralPath $receiptPath).Hash -eq $previousHash) 'Rejected run replaced the last qualified receipt.'}
            return $run
        }
        Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
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
    Case 'Real receipt qualifies retained SMTP conflicts without declaring addresses unique' {
        Reset-Fixture
        Find-SmartM365MailboxSmtpConflict -Rows @([pscustomobject]@{ObjectGuid=$id1;PrimarySmtpAddress='same@synthetic.invalid'},[pscustomobject]@{ObjectGuid=$id2;PrimarySmtpAddress='same@synthetic.invalid'})
        $receipt=Invoke-TestReceipt (Get-TestQualification -Issues @($script:LocalMailboxIssues.ToArray()))
        Assert-True ($receipt.Status -eq 'Completed' -and $receipt.ConsumerScopeQualified -and -not $receipt.FullInventoryQualified) 'Retained conflict did not preserve scoped qualification.'
        Assert-True (($receipt.Qualifications -join ' ') -like '*occurrences: 2; blocking=0*') 'SMTP findings disappeared from qualified receipt.'
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
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBnRk0PPJu/yrR1
# UxKUM5x+SdWVN5wTsFt9grKey+jWCqCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgrUuR1vojxGvb
# 8oLIRvMtcKxYvkqiiVPaEnmcxjT3I7owDQYJKoZIhvcNAQEBBQAEggGASr9Qn8ti
# rzQdkFY569++P8i+NUCaai2NaFEnw4l9fU8cHe1rqkSQ/Eat/BWiKtoqZewwQzxh
# Gd3F2AnY56lbC6dSG6miB+mza9pxbyIUsO08DoAiLhBfZzP1vFYaIYwKccp7posC
# pXt7G/WeU0S06x9j1ElE2eaqdGTofevEYqvrpiTiWMG5KRdkagJOLOTBv0IKxo+h
# MzwFowQe7Q4VkrZGaZNRG/EAflUj5u7USUwO+fvf+PAWny64MklVUbxwnYPIZV6z
# b+REQW4n87djqCtPrEq2gKQBcL3cugmL2hEMWmZXErjmCviPJu/IMoYa3RVQHZOK
# cZ+X1xId1vRP5/yEZ37Au4kVGUsj1ZRMsNBE06zYD2ctmpXHxe5wioZjX1kbJzjn
# x2cEDgzAmbNAF/YeXrV/ROKVp5d9wHtDZsG4aRl9C4KvsDUjel/s+MRqFS45jBn8
# FG3Wwk8qks6Fr1+FZJ9LCiw5tWIzZjTNmK5v+y+xy8gtd+BO4wLylMO8oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNTI4WjA/BgkqhkiG9w0BCQQxMgQwZPaS31v2/qVgBksJvbEmpSgY
# OlEfgpw+5h5ENcF9PpCyGzTAYlNNdPwH7UBzpJoJMA0GCSqGSIb3DQEBAQUABIIC
# AKfgCuL9HtbSc8K1tVr+arxboyWwY54GPP7XsHDP3mao1HTmnBhgevN82602aLpM
# QfY/sIItM9eVnw2/zOnvgBIIkfKftnsYfD33t63waPejbOrsh28Rw7YYBFWtIDix
# hS8YNIzM88Bnq23FYr3SQARgxcX29ZOA2B4fNMloHQ0agQFUXX7y8CdWA0tLpFBr
# 9M+1Y5FFJjp7hq3ANZCb1TsMWG97gAinREspbkYO6z2E9TcY+j06OdOHuSHO7Qem
# eVMbTrpV2TJuTQWZI9E1bqIfSBI8+tpLtTi/kTEvHj1xKoMk3hiu7660vc9MQyZT
# A6xWZ+HQnv0BqrLB8+mQyeTnKaYmFtlcxJfb1tbD9MHikoKy8bt2jaLe4soag8sy
# llVY3V1SLyGKYcL4XWiC4bDhtHRZuNpjF2orX8HuZqT15TDqhtK7YLPJo/rWfVk1
# n0MGY3GoJhw5JFUZ9Wb1+JJ+wu9ZIs+usqQrDP4FRiJP4lP6fVlapylOVB90bRMp
# vF62vzIDCuSjTNU/YvUaUzqrHXWvC8YuBLAWeXGZV7f885WLdlITATy7kBjfI2eV
# XHhc0MmW6PtV9IaGSyb0WmckU3HLpe4A4tbYZ/t+u7HlfH5hQv9TSWu0yViTAtX0
# 1iERvlkJkVlATQz98+BSURDqccoO6rGLxcaKGjNN8dvK
# SIG # End signature block
