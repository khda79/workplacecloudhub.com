#Requires -Version 5.1
<#
.SYNOPSIS
Offline Exchange mailbox scope, native identity and quality classification tests.
.DESCRIPTION
Extracts functions and statements through the AST. All Exchange queries are
mocked; no tenant context, collector, module, export or external action runs.
.VERSION
1.0.0
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
    'Resolve-SmartM365MailboxWarningNativeGuid','ConvertFrom-SmartM365ExchangeRemoteMailboxWarnings')) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Assert-True ($null -ne $node) "Missing function: $name"
    Set-Item "Function:script:$name" ([scriptblock]::Create($node.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
}
function Reset-Fixture {
    $script:LocalMailboxIssues=New-Object 'Collections.Generic.List[object]'
    $script:LocalMailboxIssueSequence=0
    $script:LocalMailboxAcquisitions=New-Object 'Collections.Generic.List[object]'
    $script:mockRows=@(); $script:mockWarnings=@(); $script:mockFailure=$false
}
function Get-Mailbox {
    [CmdletBinding()]
    param($ResultSize,$OrganizationalUnit)
    $script:queryArguments=@{ResultSize=$ResultSize;Scope=$OrganizationalUnit}
    if ($script:mockFailure) { throw 'Synthetic acquisition failure.' }
    foreach ($warning in $script:mockWarnings) { Write-Warning $warning }
    $script:mockRows
}
$scope='DC=synthetic,DC=invalid'
$goodQuery=[pscustomobject]@{Scope=$scope;QueryCompleted=$true;ProjectionCompleted=$true;Rows=1}
function Get-TestQualification([object[]]$Issues=@(),[object[]]$Queries=@($goodQuery),[bool]$Remote=$true,[object[]]$Scopes=@($scope)) {
    Get-SmartM365LocalMailboxQualification -ExpectedScopes $Scopes -Acquisitions $Queries -RemoteAcquisitionComplete $Remote -Issues $Issues
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
if (@($tests | Where-Object {-not $_.Passed}).Count) { throw 'Offline Exchange local qualification failed.' }
[pscustomobject]@{Status='Passed';TestCount=$tests.Count;ProductionActions=0;PowerShell=[string]$PSVersionTable.PSVersion}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVu99g83jwuSPm
# loghqGaBGw4dLT/i0DAx+wMyDWElV6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEX/r+/5i+hrp3D/7pNAmcHcCSlt8hA4+jWeyg/AOoFTMA0GCSqG
# SIb3DQEBAQUABIIBgBsfPMDgJBj7dxO9bv9coVUlWjOXW1BN8JLQ8lzRSB5LtOx2
# rJjVpZam8zM95Yp5EoiUrzfwtCo/0tE3Em8c22rNjgSnz8ZGdAuF7b5l1dMFkIFb
# Ke0aSRDhXn6lrb/+qs81Dnuck3myQr/mwMDu30WFPvg6UVLH0ZsU3b3UcWzzbzxx
# JQyeotRqhwImEW4YctunOmTu3NNxea03hK34ytBc2H8h7A8C9LOvDvaclYOaUR4M
# 5nZ5TNouvBIUb8jtMXEjheJdAQs/YJ6sR1feWDiPX/KxwYVQ6bqGh14pYKI0gHv9
# Z5/qYEbfG2PokjY9Tgu7QKAdBmSdaOYT1/Olozn0EUxDu4bavVRyaYUo7zrsqwLB
# v0BUIyb8ZP2yeCsRFqrQj+eqngDsJC1+nOIK9RdzECNlo0ZZVhgb5Iu/xGlfkbGm
# s626NG7kp+bB0MoHQ5hKPI1GS6wNwntOWcEhO4rfp7rj8i5d9k+xTwir7pdt/3W6
# MxIVtAwLWElJCLqVd6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMyMjUy
# MjFaMC8GCSqGSIb3DQEJBDEiBCAdHnCqVtuY2RrlM1GdqvTt/W0BMv2VskpEO8cc
# k4U20DANBgkqhkiG9w0BAQEFAASCAgBBVGNjnIiA/PLw50lJ8GIT+cQVSF6LFVUG
# DNK5W/dTtI5ZshVW8lUuAGQaDNIWFxe8ILqZVnA7upWvfv7EZMt0rFGOUdI6yjCt
# yg9OELb1A1bSTeub4pJMZeuMUhb/U5sTkmHpA2OtM7xKZc2dCcBA80boroAlw7zi
# CS6EuxNNg5eQ2CKy7WlgUNEe+iC1HSjy9Dhmn/Vjl2IUYRXEF3v1TDSeG3bPLBQC
# 5qRVzfQjWHdi+e7E7hMRGEn7rP/q2OcF12AUMbzKnJKTSNhhNCLPC9IL2bKxFbKm
# Hjo/lkAqKLd6irQiqL9KTfBbqHME1jLW3plu+/e+fQnuWxDrdRvHWkJsyvUC0ptb
# N9e59ddBOLrgS3gKkDx/+CYgcCR3iJ6oPIzSTIpq75nTctSyhc5FlRPEfx5uubq4
# 1cD7wxRlmW6QHZ5XIMXjfq04iPcTAEgv/3vjYLB2aV00R1x4fL0ckU1O709eGTLb
# TkWOeJwlK1Yp3UJIVL8sG5duRoogduYl6ZjU2QpIjPkhlx9ui6wf8HJedMPFDTZB
# RLirqVtSICYro0KVUpa2f5zj2ib7ZOPX1/py2DcH+z9FjVkQJPkao59EiBFhIGGS
# fPuTR3/f+22rIE4To/VHtrAnU146Sqy1Mf4cwhnSIL1BDmY9BtvuaiiV5R+Ar1Yv
# t/9l/A8HyQ==
# SIG # End signature block
