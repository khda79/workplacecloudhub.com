[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$folder = Split-Path $PSScriptRoot -Parent
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
if (-not $versionMatch.Success -or [version]$versionMatch.Groups[1].Value -lt [version]'1.0.12') { throw 'Content version was not updated for category and quota usage.' }
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
foreach ($name in @('Get-InventoryConfigValue','Resolve-InventoryConfigTokens','Assert-InventoryPath','Get-MySiteHostNames','Get-ObservedProperty','Test-MissingObservation','Resolve-SiteCategory','Resolve-QuotaUsage','Resolve-SiteLockObservation','Get-LockStateLookup','Assert-Deadline','Get-CollectionFailureStatus','Get-DatabaseCoverageStatus','Get-ContentPublicationDecision','Save-ContentWeeklyHistory','Write-RunCsv','Flush-RunRows','Ensure-GraphAuthenticationModule','Invoke-DailySummaryMail','Send-InventorySummaryMail','New-ContentMailHtml','Send-ContentRunMail','Publish-QualifiedCsvUploads')) {
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
catch { if ($_.Exception.Message -ne 'Unresolved configuration value: OutputRoot') { throw } }
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
    $earlyStats = @{WebApplications=2;ContentDatabases=3;SiteCollections=$null;SiteAdministrators=$null;Webs=$null;CollectionCoverageRows=$null;CoverageIssues=$null;ReadOnlyCollections=$null;UnverifiedLocks=$null;CoverageSamples=@();ReadOnlySamples=@()}
    $failedHtml = New-ContentMailHtml -Status Failed -Stats $earlyStats -Failures @('Receipt lock is in use') -TenantName 'offline'
    if ($failedHtml -notmatch 'FAILED' -or $failedHtml -notmatch 'Receipt lock is in use' -or $failedHtml -notmatch '>n/a</div>') { throw 'Early content failure was not rendered without fabricated zeroes.' }
    $outputBase = $tempRoot
    $Tenant = 'offline'
    $script:StartedAt = Get-Date
    $script:MailStats = $sampleStats
    $ForceSendDailySummary = $false
    Send-ContentRunMail -Status Qualified
    Send-ContentRunMail -Status Failed -Failures @('Receipt lock is in use')
    if ($script:mailCalls -ne 3) { throw 'Content qualified and issue emails did not use independent daily markers.' }
    Send-ContentRunMail -Status Partial
    if ($script:mailCalls -ne 3) { throw 'Same-day content issue email was repeated.' }
    $ForceSendDailySummary = $true
    Send-ContentRunMail -Status Partial
    if ($script:mailCalls -ne 4) { throw 'Forced content issue email was skipped.' }
    $ForceSendDailySummary = $false
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
} finally {
    if ([IO.Path]::GetFullPath($tempRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Output 'PASS: content parser, nested configuration, persistent paths, lock mapping, timeouts, per-site errors, empty CSV, partial and qualified weekly history, receipts, upload policy, daily mail gate and routing.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDY5xu2RRNZmZkp
# Y37nQLP83YfxI35HPn94EtHWfUnmuKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIG0VatRrFaR0jpO0mdmTOwOje14oUmHqMYHYbQqMLEv9MA0GCSqG
# SIb3DQEBAQUABIIBgH2HkyfnvPnJoNHLdpL6l9AHKxBga2iNUD7PZqKqXpPUi99o
# 5otjGFPC4ARIY0r0EPpJ+Fu6xFS0vuYTCOi2umgvqHPcnWunqw1mD/5QJhQH8ffV
# cL5iTVuN8DRuZs9mgYVrs0autrDSPwbGAKJY1lij61uVu99FQkQvxi6YQveU3MLP
# vtV59lQWYjO+VOFi79+zoa5lFuxqytWNsI05Wuli8ig0XqeaeMHYt0LntsxaZQn2
# /pxLvSDEWCep43xwiU09NPQtdODF6O9UyV1WYEKCN8i7/69LdHQknOxkRz/g1l2i
# EyBm8eRC6ufOkUeQKdJ6iB7nlbcIMXenURGnKsxlt8Ezx9Aot1ODG5Hxqw4NAZ8r
# Ez2vC/ZroMwA7f1uQt+9fYYJgKOabxcJYTxkCFcIwJEAIT/j2NvTz68PVWzvKW7A
# nEv0WWozqexgh7rBlPl4QAH0dtOYT+9zNfpuowazdlBJZtvdc9LPXZARfQq9Gi7X
# zzE9xEQvT5/z0g9VIKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTAwMDQz
# MzdaMC8GCSqGSIb3DQEJBDEiBCCR7mno5Wy4OXA8ysl3Ma68hoLvYoHr6PWm9Tfi
# 45XU3zANBgkqhkiG9w0BAQEFAASCAgCkYkMbKc/PNZr073L1ejwqE83MQhRSdHOm
# rGlmC69/BaiKvGxLXP337S/CWlHmw7tfEs8pLzRzyyepXaGlBBlyBgU1KVfveqxi
# bRiWLUZdgNzrzwHd+3GiziS6/+shIQ1qdjVE48okAe8/EnuvIUfuRM7yVISORgNB
# WoeGSnGnQMEEBYM1W5wc5/GUmpFBqTssYQEWULWSaBmjM7/SRUFG9AIjLtIAdQj/
# k/yYnd/iHeHiR8ymiRzPCVJaldBzCtUV7QehqfWA65/yXq3ahSpfB3eKO50p+hiH
# jJGnxrBklKDNfq8cFOqG9tCYQUB6WQLFP17ETz7ATD7/jp2fXfxr0/EZECAXqx73
# /AJXT4OW417jgRJnX+WWUWhxlgQOVUseH4aGXHKcP1aoBflYAOQkI5HVeoXjX050
# 8/j9wjbOEoYLGb07BluHQObLEvZbIA617fRtpPYZ4zIR7JhIoAux2D6qWSv1PcgV
# xDh077tJ6cn/6chCcSwEzZBlJcucbPj1sDflewEfmtMJzcZZ75meYKiBViV7mP7g
# YlLW5+GIIaRjdFNdPf1OAYOKaNUK/t+nOmYD5gVqp9+UsVsF6q416QKQYSiLPwkD
# nhKOZRoAtmREWoiNQKEhcjCKPiT78099kaLERmPtz7CSULlIFUmpatiiQlsPmAD+
# sH3OUTdl9g==
# SIG # End signature block
