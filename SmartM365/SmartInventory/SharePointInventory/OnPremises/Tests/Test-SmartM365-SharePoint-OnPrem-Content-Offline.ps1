<#
.SYNOPSIS
    Offline SharePoint on-premises content inventory tests.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$global:EnableTeamsNotifications=$false
$global:SmartM365GlobalConfig=[pscustomobject]@{ProfileKey='test';EnableTeamsNotifications=$false}
function global:Invoke-OfflineSuppressedTeamsNotification { [CmdletBinding()]param([Parameter(ValueFromRemainingArguments=$true)]$Remaining) }
Set-Alias -Name Send-SmartM365TeamsNotification -Value Invoke-OfflineSuppressedTeamsNotification -Scope Global -Option AllScope
$folder = Split-Path $PSScriptRoot -Parent
. (Join-Path $folder '..\..\..\Config\SmartM365-TenantContext.ps1')
$scriptPath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Content-Inventory.ps1'
$templatePath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Content-Inventory.local.json.template'
$registryPath = Join-Path $folder '..\..\..\Modules\SmartM365.Core\SmartM365-SourceReceipts.json.txt'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count) { throw "Content parser errors: $($errors.Count)" }
$template = Get-Content $templatePath -Raw | ConvertFrom-Json
if ($template.EnableSharePointUpload -ne $true -or $template.EnableWeeklyHistory -ne $true -or $template.SendMailMode -ne 'Graph') { throw 'Content template policy is invalid.' }
$versionMatch = [regex]::Match($ast.Extent.Text, '(?s)\.VERSION\s+([0-9.]+)')
if (-not $versionMatch.Success -or [version]$versionMatch.Groups[1].Value -lt [version]'1.0.13') { throw 'Content version was not updated for the combined summary email.' }
if ($null -eq $template.PSObject.Properties['MySiteHostNames'] -or $null -eq $template.MySiteHostNames) { throw 'Content template lacks MySiteHostNames.' }
if ($template.ExcludePersonalSiteCollections -ne $true) { throw 'Personal collection exclusion must be enabled in the template.' }
if ($ast.Extent.Text -notmatch '\$rootWebTemplate\s*-ieq\s*''SPSPERS''' -or
    $ast.Extent.Text -notmatch '\$coverage\[''Status''\]\s*=\s*''ExcludedPersonal''' -or
    $ast.Extent.Text -notmatch 'excludedPersonalCount\+\+' -or
    $ast.Extent.Text -notmatch "'ExcludedPersonal'\)" -or
    $ast.Extent.Text -notmatch 'ExcludePersonalSiteCollections=\$excludePersonalCollections') { throw 'Auditable personal collection exclusion path is incomplete.' }
if ($ast.Extent.Text -match 'Complete-SmartM365SourceReceipt[^\r\n]*-PartialInventory') { throw 'Content must not extend the shared source receipt with SharePoint-specific partial state.' }
if ($ast.Extent.Text -notmatch '\$coverageLevel\s*=\s*if\s*\(\$globalTimedOut\).*?elseif\s*\(\$failureCount\s*-gt\s*0\).*?WARNING' -or $ast.Extent.Text -notmatch 'Collection coverage:[^\r\n]+-Level\s+\$coverageLevel') { throw 'Content coverage logging can misclassify a partial run.' }
$firstUploadConfigRead = $ast.Extent.Text.IndexOf('$global:SharePointSiteHostname =', [StringComparison]::Ordinal)
$resolverDefinition = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Resolve-InventoryConfigTokens' }, $true))
if ($firstUploadConfigRead -lt 0 -or $resolverDefinition.Count -ne 1 -or $resolverDefinition[0].Extent.EndOffset -ge $firstUploadConfigRead) { throw 'Content reads upload configuration before the token resolver is defined.' }
$registry = Get-Content $registryPath -Raw | ConvertFrom-Json
$producer = @($registry.Producers | Where-Object Script -eq 'SmartM365-SharePoint-OnPrem-Content-Inventory.ps1')
if ($producer.Count -ne 1 -or $producer[0].Files.Count -ne 4) { throw 'Content source receipt registration is invalid.' }
if ($ast.Extent.Text -match '(?im)^\s*(Set-SPSite|Set-SPWeb|Set-SPContentDatabase|Add-SPShellAdmin|Remove-SPSite)\b') { throw 'A SharePoint write command was found.' }
if ($ast.Extent.Text -match '(?im)^\s*Get-SPSiteAdministration\b') { throw 'Lock inspection must use SPSite and SPContentDatabase only.' }
if ($ast.Extent.Text -notmatch '(?m)^\s*\$site\s*=\s*\$_\s*$' -or $ast.Extent.Text -match '(?m)^\s*param\(\$site\)') { throw 'Streaming site pipeline input is not bound to the current site.' }
if ($ast.Extent.Text -notmatch '\$runBase\s*=\s*if\s*\(\$MaxItems\s*-gt\s*0\).*?TEST' -or $ast.Extent.Text -notmatch 'Select-Object\s+-First\s+\$remaining' -or $ast.Extent.Text -notmatch 'Flush-RunRows\s+-Kind\s+CollectionCoverage') { throw 'Limited or per-site coverage path is missing.' }
if ($ast.Extent.Text -notmatch '\$databaseLockLookup\s*=\s*Get-LockStateLookup\s+-ContentDatabase\s+\$database' -or $ast.Extent.Text -match 'Get-LockStateLookup\s+[^\r\n]*-SiteUrl') { throw 'Limited runs must use the database-scoped lock-state lookup.' }
foreach ($name in @('Get-InventoryConfigValue','Resolve-InventoryConfigTokens','Assert-InventoryPath','Get-MySiteHostNames','Get-ObservedProperty','Test-MissingObservation','Resolve-SiteCategory','Resolve-QuotaUsage','Resolve-SiteLockObservation','Get-LockStateLookup','Assert-Deadline','Get-CollectionFailureStatus','Get-DatabaseCoverageStatus','Get-ContentPublicationDecision','Save-ContentWeeklyHistory','Write-RunCsv','Flush-RunRows','Ensure-GraphAuthenticationModule','Invoke-DailySummaryMail','Send-InventorySummaryMail','Get-ContentMailFarmEdition','Get-VerifiedInfrastructureMailData','New-ContentMailHtml','Send-ContentRunMail','Publish-QualifiedCsvUploads')) {
    $functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true))
    if ($functionAst.Count -ne 1) { throw "Missing function: $name" }
    . ([scriptblock]::Create($functionAst[0].Extent.Text))
}
$script:EffectiveConfig = [pscustomobject]@{
    SmartM365RootPath='C:\SmartM365'; WorkspaceRootPath='{{SmartM365RootPath}}'; ProfileKey='prod'
    DataAllRootPath='{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-ALL'
    OutputRoot='{{DataAllRootPath}}\SharePoint\OnPrem\Content'
    WeeklyHistoryFolderPath='{{DataAllRootPath}}\SharePoint\OnPrem\Content\WeeklyHistory'
    MySiteHostNames=@('My.Example.Test','my.example.test')
}
if ((Get-InventoryConfigValue -Name 'ExcludePersonalSiteCollections' -DefaultValue $true) -ne $true) { throw 'Missing personal exclusion setting did not default to true.' }
$script:EffectiveConfig | Add-Member -NotePropertyName ExcludePersonalSiteCollections -NotePropertyValue $false
if ((Get-InventoryConfigValue -Name 'ExcludePersonalSiteCollections' -DefaultValue $true) -ne $false) { throw 'Explicit false personal exclusion setting was ignored.' }
$script:EffectiveConfig.PSObject.Properties.Remove('ExcludePersonalSiteCollections')
if (@(Get-MySiteHostNames).Count -ne 1 -or @(Get-MySiteHostNames)[0] -ne 'my.example.test') { throw 'My Site host normalization failed.' }
$script:EffectiveConfig.MySiteHostNames = @()
if (@(Get-MySiteHostNames).Count -ne 0) { throw 'Empty My Site hosts must remain empty.' }
$script:EffectiveConfig.MySiteHostNames = @('one.example.test','two.example.test')
if (@(Get-MySiteHostNames).Count -ne 2) { throw 'Multiple My Site hosts must remain distinct.' }
$script:EffectiveConfig.MySiteHostNames = @('https://my.example.test')
try { $null = @(Get-MySiteHostNames); throw 'A URL was accepted as a My Site host name.' }
catch { if ($_.Exception.Message -notlike 'MySiteHostNames must contain host names*') { throw } }
$script:EffectiveConfig.MySiteHostNames = @('my.example.test')
foreach ($case in @(
    @{Url='https://my.example.test/personal/a';Template='SPSPERS';Hosts=@();Category='Personal';Status='Observed'},
    @{Url='https://my.example.test/teams';Template='STS';Hosts=@('my.example.test');Category='OtherOnMyHost';Status='Observed'},
    @{Url='https://team.example.test/';Template='STS';Hosts=@('my.example.test');Category='Other';Status='Observed'},
    @{Url='https://team.example.test/';Template='STS';Hosts=@();Category='';Status='HostListUnconfigured'},
    @{Url='invalid';Template='STS';Hosts=@('my.example.test');Category='';Status='InvalidUrl'},
    @{Url='https://team.example.test/';Template='';Hosts=@('my.example.test');Category='';Status='TemplateUnavailable'}
)) {
    $result = Resolve-SiteCategory -Url $case.Url -RootWebTemplate $case.Template -MySiteHostNames $case.Hosts
    if ($result.Category -ne $case.Category -or $result.Status -ne $case.Status) { throw "Site category failed: $($case.Url), $($case.Template)" }
}
foreach ($case in @(
    @{Storage='50';Quota='200';Percent='25';Status='Observed'},
    @{Storage='250';Quota='200';Percent='125';Status='Observed'},
    @{Storage='200';Quota='0';Percent='';Status='NoQuota'},
    @{Storage='';Quota='200';Percent='';Status='StorageUnavailable'},
    @{Storage='200';Quota='';Percent='';Status='QuotaUnavailable'},
    @{Storage='-1';Quota='200';Percent='';Status='InvalidStorage'},
    @{Storage='200';Quota='-1';Percent='';Status='InvalidQuota'}
)) {
    $result = Resolve-QuotaUsage -StorageBytes $case.Storage -QuotaLimitBytes $case.Quota
    if ($result.Percent -ne $case.Percent -or $result.Status -ne $case.Status) { throw "Quota usage failed: $($case.Storage)/$($case.Quota)" }
}
$collectionSchema = [regex]::Match($ast.Extent.Text, 'SiteCollections\s*=\s*@\([^\r\n]+\)').Value
foreach ($column in @('QuotaUsagePercent','QuotaUsageStatus','SiteCategory','SiteCategoryStatus')) {
    if ($collectionSchema -notmatch [regex]::Escape("'$column'")) { throw "Content CSV schema lacks $column." }
}
if ((Get-InventoryConfigValue 'OutputRoot') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-ALL\SharePoint\OnPrem\Content') { throw 'Nested content OutputRoot tokens were not resolved.' }
if ((Get-InventoryConfigValue 'WeeklyHistoryFolderPath') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-ALL\SharePoint\OnPrem\Content\WeeklyHistory') { throw 'Nested content history tokens were not resolved.' }
$script:EffectiveConfig.OutputRoot = '{{MissingPath}}\SharePoint'
try { $null = Get-InventoryConfigValue 'OutputRoot'; throw 'Missing token was accepted.' }
catch { if ($_.Exception.Message -ne 'Unresolved configuration token in: OutputRoot') { throw } }
$root = 'C:\ProgramData\SmartM365\LauncherCache\SharePointContent\SmartM365'
try { $null = Assert-InventoryPath -Path (Join-Path $root 'Data\Tenants\prod\DATA-ALL') -Name 'OutputRoot'; throw 'Disposable content output path was accepted.' }
catch { if ($_.Exception.Message -notlike 'OutputRoot resolves inside the disposable launcher cache.*') { throw } }
if ((Assert-InventoryPath -Path 'C:\SmartM365\DATA' -Name 'OutputRoot') -ne 'C:\SmartM365\DATA') { throw 'Persistent content output path was rejected.' }
$site = [pscustomobject]@{ ReadOnly=$false; ReadLocked=$false; WriteLocked=$false; LockIssue='' }
$database = [pscustomobject]@{ IsReadOnly=$false }
$unknown = Resolve-SiteLockObservation -Site $site -ContentDatabase $database
if ($unknown.LockState -ne '' -or $unknown.LockStatus -ne 'Unverified' -or $unknown.IsReadOnly -ne 'False' -or $unknown.ContentDatabaseIsReadOnly -ne 'False') { throw 'Unknown lock state was converted to a false observation.' }
foreach ($state in @('Unlock','NoAdditions','ReadOnly','NoAccess')) {
    $result = Resolve-SiteLockObservation -Site ([pscustomobject]@{ LockState=$state; ReadOnly=$false; ReadLocked=$false; WriteLocked=$false; LockIssue='' }) -ContentDatabase $database
    if ($result.LockState -ne $state -or $result.LockStatus -ne 'Observed') { throw "Lock state mapping failed: $state" }
}
$writeLocked = Resolve-SiteLockObservation -Site ([pscustomobject]@{ReadLocked=$false;WriteLocked=$true;ReadOnly=$false}) -ContentDatabase $database
if ($writeLocked.LockState -ne '' -or $writeLocked.LockStatus -ne 'Unverified') { throw 'WriteLocked alone was misclassified as ReadOnly.' }
$noAdditions = Resolve-SiteLockObservation -Site ([pscustomobject]@{ReadLocked=$false;WriteLocked=$true;ReadOnly=$false}) -ContentDatabase $database -FilteredLockState 'NoAdditions'
if ($noAdditions.LockState -ne 'NoAdditions' -or $noAdditions.LockStatus -ne 'Observed') { throw 'Filtered NoAdditions mapping failed.' }
$conflicting = Resolve-SiteLockObservation -Site ([pscustomobject]@{LockState='ReadOnly';ReadLocked=$false;WriteLocked=$true;ReadOnly=$false}) -ContentDatabase $database -FilteredLockState 'NoAdditions'
if ($conflicting.LockState -ne '' -or $conflicting.LockStatus -ne 'Conflict') { throw 'Conflicting lock-state observations were accepted.' }
$script:GlobalDeadlineUtc = [datetime]::UtcNow.AddMinutes(1)
$script:MockLockSites = @(
    [pscustomobject]@{Id='site-a';Url='https://example.test/a';State='Unlock'},
    [pscustomobject]@{Id='site-b';Url='https://example.test/b';State='NoAdditions'}
)
$script:MockFailState = ''
$script:MockIgnoreFilter = $false
function Get-SPSite {
    [CmdletBinding()]
    param([string]$Identity, $ContentDatabase, [scriptblock]$Filter, [string]$Limit)
    if ($Identity) { throw 'Identity cannot be combined with the lock-state filter.' }
    if ($null -eq $ContentDatabase -or $Limit -ne 'All') { throw 'The lock-state query must cover the content database.' }
    if ($Filter.ToString() -notmatch "'(Unlock|NoAdditions|ReadOnly|NoAccess)'") { throw 'An unsupported lock-state filter was used.' }
    $requestedState = $Matches[1]
    if ($requestedState -eq $script:MockFailState) { throw 'MockFilterFailure' }
    foreach ($entry in $script:MockLockSites) {
        if ($Identity -and $entry.Url -ne $Identity) { continue }
        if (-not $script:MockIgnoreFilter -and $entry.State -ne $requestedState) { continue }
        $result = [pscustomobject]@{ Id=$entry.Id }
        $result | Add-Member -MemberType ScriptMethod -Name Dispose -Value { }
        $result
    }
}
$lookup = Get-LockStateLookup -ContentDatabase $database
if ($lookup.States.Count -ne 2 -or $lookup.States['site-a'] -ne 'Unlock' -or $lookup.States['site-b'] -ne 'NoAdditions' -or $lookup.Errors.Count -ne 0) { throw 'Database lock-state lookup failed.' }
$script:MockFailState = 'NoAccess'
$failedLookup = Get-LockStateLookup -ContentDatabase $database
if ($failedLookup.Errors.Count -ne 1 -or $failedLookup.States.Count -ne 2) { throw 'Lock-state query failure was hidden.' }
$script:MockFailState = ''
$script:MockIgnoreFilter = $true
$conflictLookup = Get-LockStateLookup -ContentDatabase $database
if ($conflictLookup.States.Count -ne 0 -or $conflictLookup.Conflicts.Count -ne 2) { throw 'Conflicting filter results were accepted.' }
Remove-Item Function:\Get-SPSite
if ((Get-CollectionFailureStatus 'Access is denied') -ne 'AccessDenied' -or (Get-CollectionFailureStatus 'CollectionTimeout') -ne 'TimedOut' -or (Get-CollectionFailureStatus 'unexpected') -ne 'Failed') { throw 'Per-collection failure classification failed.' }
if ((Get-DatabaseCoverageStatus -ExpectedCount 2 -ObservedCount 2) -ne 'Collected' -or (Get-DatabaseCoverageStatus -ExpectedCount 2 -ObservedCount 1) -ne 'DatabaseCoverageMismatch' -or (Get-DatabaseCoverageStatus -ExpectedCount '' -ObservedCount 0) -ne 'DatabaseCountUnavailable') { throw 'Database coverage qualification failed.' }
if ((Get-ContentPublicationDecision -FailureCount 0 -GlobalTimedOut $false -MaxItems 0) -ne 'Complete' -or
    (Get-ContentPublicationDecision -FailureCount 8 -GlobalTimedOut $false -MaxItems 0) -ne 'Partial' -or
    (Get-ContentPublicationDecision -FailureCount 8 -GlobalTimedOut $true -MaxItems 0) -ne 'Blocked' -or
    (Get-ContentPublicationDecision -FailureCount 0 -GlobalTimedOut $false -MaxItems 2) -ne 'TestOnly') { throw 'Content publication decision did not preserve full, partial, timeout and limited-run boundaries.' }
$script:GlobalDeadlineUtc = [datetime]::UtcNow.AddMinutes(-1)
$timedOut = $false
try { Assert-Deadline ([datetime]::UtcNow.AddMinutes(1)) } catch { $timedOut = $_.Exception.Message -eq 'GlobalTimeout' }
if (-not $timedOut) { throw 'Global timeout was not enforced.' }
$coreManifest = Join-Path $folder '..\..\..\Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module $coreManifest -Force -ErrorAction Stop
$global:SmartM365TenantKey = 'OFFLINE'
$global:SmartM365OrganizationKey = 'OFFLINE'
$global:SmartM365EnvironmentKey = 'TEST'
$global:SmartM365TenantId = '00000000-0000-0000-0000-000000000001'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SPContent-Test-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $path = Write-RunCsv -Name 'Empty.csv' -Rows @() -Columns @('TenantKey','FarmId','LockState') -Folder $tempRoot
    if (-not $global:csvGeneratedPaths.Contains($path)) { throw 'Content run CSV was not registered for execution summary.' }
    $text = [IO.File]::ReadAllText($path)
    if ($text -notmatch 'TenantKey.*;.*FarmId.*;.*LockState' -or @(Import-Csv $path -Delimiter ';').Count -ne 0) { throw 'Empty content CSV contract failed.' }
    $buffers = @{ SiteCollections = (New-Object 'System.Collections.Generic.List[object]') }
    $paths = @{ SiteCollections = $path }
    $buffers.SiteCollections.Add([pscustomobject][ordered]@{TenantKey='OFFLINE';FarmId='FARM';LockState='NoAdditions'})
    $buffers.SiteCollections.Add([pscustomobject][ordered]@{TenantKey='OFFLINE';FarmId='FARM';LockState='ReadOnly'})
    Flush-RunRows -Kind SiteCollections -Buffers $buffers -Paths $paths -MinimumCount 2
    $actual = @(Import-Csv $path -Delimiter ';')
    if ($actual.Count -ne 2 -or $actual[0].LockState -ne 'NoAdditions' -or $actual[1].LockState -ne 'ReadOnly' -or $buffers.SiteCollections.Count -ne 0) { throw 'Bounded content CSV batch failed.' }
    $latest = Join-Path $tempRoot 'DATA-LAST'
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest -ScopeParameters @{MaxItems=0}
    foreach ($name in $producer[0].Files) {
        if ($name -eq 'SharePoint_OnPrem_SiteCollections.csv') {
            Copy-SmartM365FileAtomically -SourcePath $path -DestinationPath (Join-Path $latest $name)
        } else {
            Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM'}) -Path (Join-Path $latest $name) -Columns @('TenantKey','FarmId') -Delimiter ';' -NoTenantKey
        }
    }
    $proofPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    $proof = Get-Content $proofPath -Raw | ConvertFrom-Json
    if ($proof.Status -ne 'Completed' -or $proof.Files.Count -ne 4) { throw 'Content receipt completion failed.' }
    $proofHash = (Get-FileHash $proofPath -Algorithm SHA256).Hash
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest -ReadOnly -ScopeParameters @{MaxItems=2}
    if ((Get-FileHash $proofPath -Algorithm SHA256).Hash -ne $proofHash) { throw 'MaxItems changed the content receipt.' }
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest
    $null = Complete-SmartM365SourceReceipt -Status Failed -ErrorCount 1
    if ((Get-FileHash $proofPath -Algorithm SHA256).Hash -ne $proofHash) { throw 'A failed run changed the preceding content receipt.' }
    $partialLatest = Join-Path $tempRoot 'PARTIAL-DATA-LAST'
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $partialLatest -ScopeParameters @{MaxItems=0}
    foreach ($name in $producer[0].Files) {
        if ($name -eq 'SharePoint_OnPrem_CollectionCoverage.csv') {
            Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM';RunId='PARTIAL-RUN';Status='DatabaseCoverageMismatch'}) -Path (Join-Path $partialLatest $name) -Columns @('TenantKey','FarmId','RunId','Status') -Delimiter ';' -NoTenantKey
        } else {
            Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM'}) -Path (Join-Path $partialLatest $name) -Columns @('TenantKey','FarmId') -Delimiter ';' -NoTenantKey
        }
    }
    $partialPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    $partialProof = Get-Content -LiteralPath $partialPath -Raw | ConvertFrom-Json
    $coverageRows = @(Import-Csv -LiteralPath (Join-Path $partialLatest 'SharePoint_OnPrem_CollectionCoverage.csv') -Delimiter ';')
    if ($partialProof.Status -ne 'Completed' -or $partialProof.Files.Count -ne 4 -or $null -ne $partialProof.IsPartialInventory -or $partialProof.FullInventoryQualified -ne $false -or $partialProof.ConsumerScopeQualified -ne $false -or $coverageRows.Count -ne 1 -or $coverageRows[0].Status -ne 'DatabaseCoverageMismatch') { throw 'Observed rows or their coverage issue were not published with file-integrity evidence.' }
    function WriteLog { param($Message, $Level) }
    $global:EnableSharePointUpload = $false
    $historyRoot = Join-Path $tempRoot 'WeeklyHistory'
    $partialCsvPaths = @($producer[0].Files | ForEach-Object { Join-Path $partialLatest $_ })
    $partialHistoryRoot = Save-ContentWeeklyHistory -RunPaths $partialCsvPaths -HistoryRoot $historyRoot -RetentionWeeks 2 -Partial $true
    if ($partialHistoryRoot -ne (Join-Path $historyRoot 'Partial')) { throw 'Partial history was not isolated from qualified history.' }
    $partialWeeks = @(Get-ChildItem -LiteralPath $partialHistoryRoot -Directory)
    if ($partialWeeks.Count -ne 1) { throw 'Partial weekly snapshot was not created.' }
    $partialFiles = @(Get-ChildItem -LiteralPath $partialWeeks[0].FullName -File -Filter '*.csv')
    $manifestPath = Join-Path $partialWeeks[0].FullName 'manifest.json'
    if ($partialFiles.Count -ne 4) { throw 'Partial weekly snapshot lacks its four CSVs.' }
    $manifestDocument = (Read-SmartM365JsonDocument -Path $manifestPath).Document
    $historyCoverage = @(Import-Csv -LiteralPath (Join-Path $partialWeeks[0].FullName 'SharePoint_OnPrem_CollectionCoverage.csv') -Delimiter ';')
    if ($manifestDocument.HistoryLabel -notmatch 'partial; incomplete coverage' -or $manifestDocument.Files.Count -ne 4 -or $historyCoverage[0].Status -ne 'DatabaseCoverageMismatch') { throw 'Partial weekly history concealed its incomplete coverage.' }
    $qualifiedCsvPaths = @($producer[0].Files | ForEach-Object { Join-Path $latest $_ })
    $qualifiedHistoryRoot = Save-ContentWeeklyHistory -RunPaths $qualifiedCsvPaths -HistoryRoot $historyRoot -RetentionWeeks 2 -Partial $false
    if ($qualifiedHistoryRoot -ne $historyRoot -or @(Get-ChildItem -LiteralPath $historyRoot -Directory | Where-Object Name -Match '^\d{4}-W\d{2}$').Count -ne 1) { throw 'A partial snapshot blocked standard weekly history.' }
    $incompleteRoot = Join-Path $tempRoot 'INCOMPLETE-PARTIAL'
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $incompleteRoot -ScopeParameters @{MaxItems=0}
    Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM'}) -Path (Join-Path $incompleteRoot $producer[0].Files[0]) -Columns @('TenantKey','FarmId') -Delimiter ';' -NoTenantKey
    $rejectedPartialPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    $rejectedPartial = Get-Content -LiteralPath $rejectedPartialPath -Raw | ConvertFrom-Json
    if ($rejectedPartial.Status -ne 'Failed' -or (Test-Path -LiteralPath (Join-Path $incompleteRoot $producer[0].Receipt))) { throw 'Partial inventory bypassed required current CSV validation.' }
    $marker = Join-Path $tempRoot 'Content-DailySummary.sent'
    $script:sendCount = 0
    $sendAction = { $script:sendCount++ }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction)) { throw 'First content email was skipped.' }
    if (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction) { throw 'Same-day content email was repeated.' }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction -Force) -or $script:sendCount -ne 2) { throw 'Forced content email failed.' }
    $failedMarker = Join-Path $tempRoot 'Content-Failed.sent'
    try { $null = Invoke-DailySummaryMail -MarkerPath $failedMarker -SendAction { throw 'Simulated SMTP failure' }; throw 'Mail failure was ignored.' }
    catch { if ($_.Exception.Message -ne 'Simulated SMTP failure') { throw } }
    if (Test-Path -LiteralPath $failedMarker) { throw 'Failed content email wrote a marker.' }
    $script:mailParameters = $null
    $script:mailCalls = 0
    $script:mailShouldFail = $false
    function SendEmailHtmlReport { [CmdletBinding()] param($From,$To,$Subject,$BodyHtml,$MailPurpose,$SmtpServer,$SmtpPort,$SendMailMode,$Cc) if ($script:mailShouldFail) { throw 'Simulated Graph mail failure' }; $script:mailParameters = $PSBoundParameters; $script:mailCalls++ }
    $script:EffectiveConfig = [pscustomobject]@{ From='sender@example.test'; To=''; ErrorMailTo='recipient@example.test'; SmtpServer='smtp.example.test'; SmtpPort=25; SendMailMode='SMTP'; Cc='' }
    Send-InventorySummaryMail -Subject 'Offline summary' -BodyHtml '<p>Summary</p>'
    if ($script:mailParameters.To -ne 'recipient@example.test' -or $script:mailParameters.From -ne 'sender@example.test' -or $script:mailParameters.SendMailMode -ne 'Graph' -or $script:mailParameters.ContainsKey('Attachments') -or $script:mailParameters.ContainsKey('SmtpServer')) { throw 'Content Graph mail routing or attachment policy failed.' }
    $script:BuildEditionMapPath = Join-Path $folder 'SharePoint-OnPrem-BuildEditions.psd1'
    $infraRoot = Join-Path $tempRoot 'CombinedMailInfrastructure'
    New-Item -ItemType Directory -Path $infraRoot -Force | Out-Null
    $now = [datetime]::UtcNow
    function New-InfraMailTestRow {
        param([hashtable]$Fields)
        $row = @{TenantKey='offline';FarmId='FARM';RunId='CSV-RUN';CollectedAtUtc=$now.ToString('o')}
        foreach ($key in $Fields.Keys) { $row[$key] = $Fields[$key] }
        return [pscustomobject]$row
    }
    $infraTestRows = @{
        Farms = @(New-InfraMailTestRow @{BuildVersion='16.0.10417.20198';ServerCount='2'})
        Servers = @((New-InfraMailTestRow @{ServerName='SP1';Role='WebFrontEnd';Status='Online'}),(New-InfraMailTestRow @{ServerName='SP2';Role='Application';Status='Offline'}),(New-InfraMailTestRow @{ServerName='SQL1';Role='Invalid';Status='Online'}))
        ServiceApplications = @(New-InfraMailTestRow @{Name='Search';Status='Stopped'})
        WebApplications = @(New-InfraMailTestRow @{WebApplicationId='WEB';DefaultUrl='https://example.test';ApplicationPool='Pool';ContentDatabaseCount='1'})
        WebApplicationZones = @(New-InfraMailTestRow @{WebApplicationId='WEB';Zone='Default';AuthenticationMode='Claims'})
        ContentDatabases = @(New-InfraMailTestRow @{Name='Content';DatabaseServer='SQL1';Status='Online';IsReadOnly='True';NeedsUpgradeIncludeChildren='True';NeedsUpgradeStatus='Collected';DatabaseSizeBytes='';DatabaseSizeStatus='Unavailable'})
    }
    $infraProofs = @()
    foreach ($name in @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')) {
        $fileName = "SharePoint_OnPrem_$name.csv"
        $path = Join-Path $infraRoot $fileName
        $infraTestRows[$name] | Export-Csv -LiteralPath $path -Delimiter ';' -NoTypeInformation -Encoding UTF8
        $infraProofs += [pscustomobject]@{File=$fileName;Status='Success';RunId='RECEIPT-RUN';Rows=@($infraTestRows[$name]).Count;SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
    }
    $infraReceiptPath = Join-Path $infraRoot 'SmartInventory_SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.current.json.txt'
    [pscustomobject]@{Owner='SmartInventory-SourceReceipt';Status='Completed';TenantKey='offline';RunId='RECEIPT-RUN';CompletedAtUtc=$now.AddMinutes(-20).ToString('o');Files=$infraProofs} |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $infraReceiptPath
    $verifiedInfra = Get-VerifiedInfrastructureMailData -LatestRoot $infraRoot -FarmId 'FARM' -TenantKey 'offline' -NowUtc $now
    if ($verifiedInfra.Status -ne 'Qualified' -or $verifiedInfra.Rows.Servers.Count -ne 3) { throw 'Combined mail rejected a valid infrastructure receipt.' }
    if ((Get-ContentMailFarmEdition -BuildVersion '16.0.0.0') -ne 'Unknown') { throw 'An unlisted farm build was assigned an edition.' }
    if ((Get-VerifiedInfrastructureMailData -LatestRoot $infraRoot -FarmId 'OTHER' -TenantKey 'offline' -NowUtc $now).Status -ne 'Unavailable') { throw 'Combined mail accepted another farm.' }
    if ((Get-VerifiedInfrastructureMailData -LatestRoot $infraRoot -FarmId 'FARM' -TenantKey 'offline' -ExpectedWebApplicationIds @('OTHER') -NowUtc $now).Status -ne 'Unavailable') { throw 'Combined mail accepted different web application scopes.' }
    if ((Get-VerifiedInfrastructureMailData -LatestRoot $infraRoot -FarmId 'FARM' -TenantKey 'offline' -NowUtc $now.AddHours(25)).Status -ne 'Stale') { throw 'Combined mail accepted stale infrastructure.' }
    $farmCsvPath = Join-Path $infraRoot 'SharePoint_OnPrem_Farms.csv'
    $originalFarmCsv = [IO.File]::ReadAllBytes($farmCsvPath)
    [IO.File]::AppendAllText($farmCsvPath, 'tampered')
    if ((Get-VerifiedInfrastructureMailData -LatestRoot $infraRoot -FarmId 'FARM' -TenantKey 'offline' -NowUtc $now).Status -ne 'Unavailable') { throw 'Combined mail accepted a changed CSV.' }
    [IO.File]::WriteAllBytes($farmCsvPath, $originalFarmCsv)
    $infraRunPath = Join-Path $infraRoot 'SmartInventory_SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.run.json.txt'
    [pscustomobject]@{Status='Failed';StartedAtUtc=$now.AddMinutes(-1).ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath $infraRunPath
    if ((Get-VerifiedInfrastructureMailData -LatestRoot $infraRoot -FarmId 'FARM' -TenantKey 'offline' -NowUtc $now).Status -ne 'Stale') { throw 'Combined mail ignored a newer failed infrastructure run.' }
    Remove-Item -LiteralPath $infraRunPath
    $sampleStats = @{
        WebApplications=2;ContentDatabases=3;SiteCollections=4;SiteAdministrators=5;Webs=6
        CollectionCoverageRows=5;CoverageIssues=1;ReadOnlyCollections=1;UnverifiedLocks=1
        CoverageSamples=@([pscustomobject]@{Status='AccessDenied';Url='https://example.test/failed';ErrorMessage='Access is denied'})
        ReadOnlySamples=@([pscustomobject]@{Url='https://example.test/readonly';LockState='ReadOnly';ContentDatabaseIsReadOnly='False'})
    }
    $html = New-ContentMailHtml -Status Partial -Stats $sampleStats -UploadFailures 1 -TenantName 'offline' -Duration ([timespan]::FromSeconds(35)) -HostName 'TEST-HOST'
    foreach ($expected in @('SMARTM365 SHAREPOINT ONPREM','PARTIAL','Duration: 00:00:35','Coverage issues','Read-only collections','Unverified locks','Upload failures','Collection coverage examples','Read-only collection examples','https://example.test/failed','https://example.test/readonly','SharePoint_OnPrem_CollectionCoverage.csv')) {
        if ($html -notmatch [regex]::Escape($expected)) { throw "Content summary email lacks $expected" }
    }
    $combinedHtml = New-ContentMailHtml -Status Partial -ContentStatus Partial -Infrastructure $verifiedInfra -Stats $sampleStats -TenantName 'offline' -HostName 'TEST-HOST'
    foreach ($expected in @('Source status','Infrastructure','SharePoint servers','Service applications','Farm edition','16.0.10417.20198','Server not online','Service application not started','Content database read-only','Content database needs upgrade','SharePoint_OnPrem_Farms.csv','SharePoint_OnPrem_CollectionCoverage.csv')) {
        if ($combinedHtml -notmatch [regex]::Escape($expected)) { throw "Combined summary email lacks $expected" }
    }
    if ($combinedHtml -match '>SQL1</td>' -or $combinedHtml -match 'Top 5 largest content databases') { throw 'Combined mail exposed external servers or an unmeasured database size.' }
    $staleHtml = New-ContentMailHtml -Status Partial -ContentStatus Qualified -Infrastructure ([pscustomobject]@{Status='Stale';Detail='Older than 24 hours';CollectedAtUtc='';Rows=$null}) -Stats $sampleStats -TenantName 'offline'
    if ($staleHtml -notmatch 'Infrastructure data unavailable' -or $staleHtml -match '>SharePoint servers<') { throw 'Unavailable infrastructure was presented as observed.' }
    $earlyStats = @{WebApplications=2;ContentDatabases=3;SiteCollections=$null;SiteAdministrators=$null;Webs=$null;CollectionCoverageRows=$null;CoverageIssues=$null;ReadOnlyCollections=$null;UnverifiedLocks=$null;CoverageSamples=@();ReadOnlySamples=@()}
    $failedHtml = New-ContentMailHtml -Status Failed -Stats $earlyStats -Failures @('Receipt lock is in use') -TenantName 'offline'
    if ($failedHtml -notmatch 'FAILED' -or $failedHtml -notmatch 'Receipt lock is in use' -or $failedHtml -notmatch '>n/a</div>') { throw 'Early content failure was not rendered without fabricated zeroes.' }
    $outputBase = $tempRoot
    $Tenant = 'offline'
    $script:StartedAt = Get-Date
    $script:MailStats = $sampleStats
    $script:farm = [pscustomobject]@{Id='FARM'}
    $script:webApplications = @([pscustomobject]@{Id='WEB'})
    $latestRoot = $infraRoot
    $global:SmartM365TenantKey = 'offline'
    $ForceSendDailySummary = $false
    Send-ContentRunMail -Status Qualified
    if (-not (Test-Path -LiteralPath (Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Content-DailySummary.sent')) -or $script:mailParameters.BodyHtml -notmatch 'SharePoint servers') { throw 'Qualified combined mail was not sent with infrastructure.' }
    Send-ContentRunMail -Status Failed -Failures @('Receipt lock is in use')
    if ($script:mailCalls -ne 3) { throw 'Content qualified and issue emails did not use independent daily markers.' }
    Send-ContentRunMail -Status Partial
    if ($script:mailCalls -ne 3) { throw 'Same-day content issue email was repeated.' }
    $ForceSendDailySummary = $true
    Send-ContentRunMail -Status Partial
    if ($script:mailCalls -ne 4) { throw 'Forced content issue email was skipped.' }
    $ForceSendDailySummary = $false
    $outputBase = Join-Path $tempRoot 'UploadIssue'
    Send-ContentRunMail -Status Qualified -UploadFailures 1
    if (-not (Test-Path -LiteralPath (Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Content-IssueDailySummary.sent')) -or (Test-Path -LiteralPath (Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Content-DailySummary.sent')) -or $script:mailParameters.BodyHtml -notmatch 'PARTIAL') { throw 'Content upload failure did not produce a separate partial combined mail.' }
    $outputBase = Join-Path $tempRoot 'MailFailure'
    $script:mailShouldFail = $true
    Send-ContentRunMail -Status Failed -Failures @('Mock failure')
    if (Test-Path -LiteralPath (Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Content-IssueDailySummary.sent')) { throw 'Failed content mail consumed the issue marker.' }
    $script:mailShouldFail = $false
    $script:graphAvailable = $false
    $script:graphInstallerAvailable = $true
    $script:graphInstallCount = 0
    $script:graphImportCount = 0
    function Get-Module { [CmdletBinding()] param([switch]$ListAvailable,[string]$Name) if ($script:graphAvailable -and $Name -eq 'Microsoft.Graph.Authentication') { [pscustomobject]@{Name=$Name;Version=[version]'2.0.0'} } }
    function Get-Command { [CmdletBinding()] param([string]$Name) if (($Name -eq 'Install-Module' -and $script:graphInstallerAvailable) -or ($Name -in @('Connect-MgGraph','Get-MgContext','Invoke-MgGraphRequest') -and $script:graphAvailable)) { [pscustomobject]@{Name=$Name} } }
    function Install-Module { [CmdletBinding()] param([string]$Name,[string]$Scope,[string]$Repository,[switch]$Force,[switch]$AllowClobber) if ($Name -ne 'Microsoft.Graph.Authentication' -or $Scope -ne 'CurrentUser' -or $Repository -ne 'PSGallery' -or -not $Force -or -not $AllowClobber) { throw 'Unexpected Graph installation parameters.' }; $script:graphInstallCount++; $script:graphAvailable = $true }
    function Import-Module { [CmdletBinding()] param([string]$Name) if ($Name -ne 'Microsoft.Graph.Authentication') { throw 'Unexpected Graph import.' }; $script:graphImportCount++ }
    $originalProtocol = [Net.ServicePointManager]::SecurityProtocol
    Ensure-GraphAuthenticationModule
    Ensure-GraphAuthenticationModule
    if ($script:graphInstallCount -ne 1 -or $script:graphImportCount -ne 2 -or [Net.ServicePointManager]::SecurityProtocol -ne $originalProtocol) { throw 'Content Graph bootstrap was not idempotent or changed the process TLS policy.' }
    $script:graphAvailable = $false
    $script:graphInstallerAvailable = $false
    try { Ensure-GraphAuthenticationModule; throw 'Missing Graph installer was accepted.' }
    catch { if ($_.Exception.Message -ne 'Install-Module is unavailable.') { throw } }
    Remove-Item Function:\Get-Module,Function:\Get-Command,Function:\Install-Module,Function:\Import-Module
    $script:uploadCalls = New-Object 'System.Collections.Generic.List[string]'
    function Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) $script:uploadCalls.Add($LocalFilePath); if ($LocalFilePath -eq 'failed.csv') { return }; return [pscustomobject]@{ LocalFilePath=$LocalFilePath } }
    $global:EnableSharePointUpload = $false
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv')) -ne 0 -or $script:uploadCalls.Count -ne 0) { throw 'Disabled content upload was attempted.' }
    $global:EnableSharePointUpload = $true
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv','run.csv','failed.csv')) -ne 1 -or $script:uploadCalls.Count -ne 3) { throw 'Content upload qualification failed.' }
    . (Join-Path $folder '..\..\..\Modules\SmartM365.Core\SmartM365-SharePointUploadPolicy.ps1')
    $outputBase=Join-Path $tempRoot 'policy-all';$latestRoot=Join-Path $tempRoot 'policy-last';$global:LogPath=Join-Path $tempRoot 'policy-log'
    $global:SmartM365CollectorUploadPolicies=@{DataAll=[pscustomobject]@{Enabled=$false};DataLast=[pscustomobject]@{Enabled=$true};LogAll=[pscustomobject]@{Enabled=$true}}
    $script:uploadCalls.Clear();$runPath=Join-Path $outputBase 'run.csv';$lastPath=Join-Path $latestRoot 'latest.csv'
    if ((Publish-QualifiedCsvUploads -Paths @($runPath,$lastPath)) -ne 0 -or $script:uploadCalls.Count -ne 1 -or $script:uploadCalls[0] -ne $lastPath) { throw 'Content upload category selection failed.' }
    Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue
} finally {
    if ([IO.Path]::GetFullPath($tempRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Output 'PASS: content parser, nested configuration, persistent paths, lock mapping, timeouts, per-site errors, empty CSV, weekly history, receipts, combined HTML, verified infrastructure, upload policy, and independent daily mail markers.'

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBWFW8kY44Lo6rV
# fN2RU7frD12YsO+zqnXX0udpcjKZ3aCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgE5FyYiAPMkBm
# hvRR2ZPk/mTZxsids0WgwbZ9KYrYbUswDQYJKoZIhvcNAQEBBQAEggGABy4OrbsF
# Dq1o4a/AGuLjxDljgZFN69UejdWtSE37R30fDE0QKyrNjRc5OWfNxseAr2IKKE3f
# p7qJZf4elwthz+4BJ6tXqOFNxCXtPF/rGh0MEQI9xBXYz/aDQblpf397ENG1QN/v
# pbP33lisj2ZwandG8OnybUiJ2m+zfFsBSkOxjDpKqLCmAw8wvzUSIQ+HmOiRsZo5
# /s0/8kn9LuTM7Qlesaksvf1HyzQu6irpBe9Y9AzL3chc0sWTsg6Ceo3VLVbC/nWN
# nnFJ/yEbvUU7f6u6ht8KXQ+AfN75tvpqtRSpbOXJBJF5D3m520qr287G58Gepdws
# 246i/jw+k5YFFUyKh5vIVL5auCI8xIgokN1tYcSNLmK5fEJIHUBD0nfxUJWnObbz
# 3Zg0NwflNH5pfM7vRpVKBuL6UmnQDpe28++xDsubi3FUkgbhoR1xdrwhynPkmPlW
# 1JiA/fCvtEzLLYusLbIL/cEitMX9MmhDci0ML8NzM94Kq8n7WZJ+NioUoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTAxMTY1OTI2WjA/BgkqhkiG9w0BCQQxMgQwa9P5AIypahxxVxcBca1oTx5m
# DgTjw3nEmAfgS5Qw3mEgHKH9gTTkebbZ+UJtePZTMA0GCSqGSIb3DQEBAQUABIIC
# AKtnCNkNVOtdMvsfvKBwar223SDIHUKVpGKshYTff+1Xvn71kBierZbvrEV+Q0UU
# punEViwzF+g5oz3At/t4w3KHSukIXXRPBW3kgOrvWUpnBkFz0+VftT4lCPVkVxuE
# x2wZre6BBqHpyC+8AOn2XO+6p5CD9736CKa74CgEfYQdZFX7xLDMzQNpyAL479eU
# 38nRveUEgZ4EeejuLirqTuBCXu6/P+3iMIWzMM19mm3G0Cue6Xsj4ldLD2r1Ho8J
# vxckSnCWr1m6j1n2q9Fpko/pkoFN+RD+TzQVRr+EyC+M7oAv92aY4hFdmb9H1lYg
# HGnzg0ZrDAPJdgckG1P0WXiesE3kjFfFa0pTVTWo66o9KpEouj62nf5uUX0S1vJi
# ZlP3x4TMNvQa2UgUmxoh5KupKBuewCFxHHgHfsUcgLQb5wZ6qKIWGEFYzHX1e/rN
# AxIObMFvY218mXM2Sh2QszsMjSb3k62W7A5bRTpy/U4/VapM7nEXXyYU1Bt4oKWF
# ECA4UnmIFp87mDfbxlgsRUbEaB2jqw1nPddu+LMWm6yo82yiFymWOb/8JD0mD/YH
# dihrPxqtBnBOssAbNFxN87GAPHKW/nmcHo9S4i/bpufMf0HP3XUhgPBpyrmNySeZ
# TS3SwiOBlQVNShQGCNP65piMDyZ3uunUnbQ0X3Oz4AwN
# SIG # End signature block
