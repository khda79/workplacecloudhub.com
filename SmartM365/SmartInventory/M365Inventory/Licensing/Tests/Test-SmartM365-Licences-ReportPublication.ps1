#Requires -Version 7.0
<#
.SYNOPSIS
  Offline checks for the license report published by every full licensing run and declared as optional receipt outputs.
.VERSION
1.1.0
.NOTES
  Synthetic data only. Uses the real SmartM365.Core receipt commands and exports; no tenant, Graph, mail or SharePoint access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity and upload globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline mocks keep the production command signatures.')]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$rulesPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-LicenseDependencies.json.txt')).Path
$exclusionsPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-LicenseExclusions.json.txt')).Path
$coreManifest = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$names = @(
  'Get-LicensesSuiteProduct', 'Get-LicensesFocusedSummaryRows', 'Get-LicensesAdditionalOverviewRows', 'New-LicensesOverviewCardHtml', 'ConvertTo-LicensesActivityDate',
  'ConvertTo-LicensesMailboxSizeGb', 'Get-LicensesQualifiedPartialAdCoverage', 'Get-LicensesCsvSource', 'Import-LicensesSourceCsv',
  'Read-LicensesIndexedSource', 'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary', 'Get-LicensesFocusedUsageRows',
  'Format-LicensesMetric', 'Get-LicensesAssignmentIssues', 'Get-LicensesAlertSummary', 'New-LicensesRecoveryWorkbook', 'Publish-LicensesReportCsv',
  'Publish-LicensesReportSnapshot', 'Write-LicensesDailyMailState', 'Enter-LicensesDailyMailGate', 'Send-LicensesFocusedSummaryEmail',
  'Read-LicensesTenantSnapshot', 'Publish-LicensesRunReport', 'ConvertFrom-LicensesJsonElement', 'ConvertTo-LicensesCountTable',
  'Read-LicensesPublishedReport', 'Get-LicensesOptionalProperty', 'Get-LicensesDecisionSettings', 'Read-LicensesDependencyRules', 'Get-LicensesAccountState',
  'Get-LicensesMailboxGuard', 'Get-LicensesCopilotUsageState', 'Get-LicensesCopilotActivity', 'Get-LicensesSkuDecisions', 'Publish-LicensesSkuReportCsv'
)
$definitions = @($ast.FindAll({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
}, $true))
if ($definitions.Count -ne $names.Count) { throw 'License report functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }

# The production module: a missing export fails here exactly as it would in the collector.
$coreModule = Import-Module $coreManifest -MinimumVersion '1.0.90' -Force -PassThru

$script:checks = 0
function Check {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) { throw $Message }
  $script:checks++
}
function Get-ScriptLocalConfigValue {
  param($Config, [string]$Name, $DefaultValue)
  if ($Config -and $Config.ContainsKey($Name)) { return $Config[$Name] }
  return $DefaultValue
}
function Test-SmartM365MaxItemsMode { return $false }
function WriteLog { param([string]$Message, [string]$Level) $script:Log.Add("$Level $Message") }
function Initialize-SmartM365ImportExcel { Import-Module ImportExcel -ErrorAction Stop | Out-Null }
function Send-SmartM365Mail {
  param([string]$From, [string]$To, [string]$Subject, [string]$BodyHtml, [string]$MailPurpose, [string[]]$Attachments, [switch]$AllowAttachments, [switch]$SuppressAttachmentLinks)
  $script:SentMail.Add([pscustomobject]@{ Subject=$Subject; BodyHtml=$BodyHtml })
}
function New-SmartM365EmailBody {
  param([string]$Title, [string]$Category, [string]$HostName, [string]$GeneratedAt, [string]$BodyHtml, [string]$Footer)
  return "<main>$BodyHtml</main><footer>$Footer</footer>"
}
function Invoke-SmartM365SharePointCsvUpload {
  param([string]$LocalFilePath)
  $script:Uploads.Add([IO.Path]::GetFileName($LocalFilePath))
  return -not $script:FailUpload
}

$saved = @{}
foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths','EnableSharePointUpload')) {
  $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
  $saved[$name] = @{ Exists=($null -ne $variable); Value=$(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'prod'
$global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'
$global:SmartM365TenantId = 'synthetic-tenant'
Remove-Variable -Name EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue
$script:OrgDomain = 'example.invalid'
$script:ScriptLocalConfig = @{ To='reports@example.invalid'; From='sender@example.invalid'; EnableLicenseSummaryEmail=$true }
$script:Log = [System.Collections.Generic.List[string]]::new()
$script:SentMail = [System.Collections.Generic.List[object]]::new()
$script:Uploads = [System.Collections.Generic.List[string]]::new()
$script:FailUpload = $false
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-LicenseReportPublication-' + [guid]::NewGuid().ToString('N'))
$dataLast = Join-Path $root 'DATA-LAST'
New-Item -ItemType Directory -Path $dataLast -Force | Out-Null
$mailStatePath = Join-Path $root 'M365_Licenses_SummaryEmail_SendState.json.txt'
$reportNames = @('M365_Licenses_ReportSummary.csv','M365_Licenses_ReportCandidates.csv','M365_Licenses_ReportAssignmentIssues.csv','M365_Licenses_ReportGaps.csv','M365_Licenses_ReportSources.csv')
$decisionNames = @('M365_Licenses_ReportSkuCandidates.csv','M365_Licenses_ReportSkuCapacity.csv')

function Write-ForeignReceipt {
  param([string]$Producer, [string]$FileName)
  $path = Join-Path $dataLast $FileName
  @{ Status='Completed'; IsPartialInventory=$false; ConsumerScopeQualified=$true
     Files=@(@{ File=$FileName; Status='Success'; IsPartialInventory=$false; SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }) } |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $dataLast "SmartInventory_$Producer.current.json.txt")
}
function Write-FixtureCsv {
  param([string]$Name, [object[]]$Rows)
  $path = Join-Path $dataLast $Name
  $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8NoBOM
  return $path
}
function Invoke-LicensingRun {
  param([object[]]$TenantRowsOverride, [string]$DependencyRulesPath = $rulesPath, [string]$ExclusionsPath = $exclusionsPath)
  # Same order as the collector: publish canonical CSVs, publish the report, then complete the receipt.
  $global:csvGeneratedPaths = $null
  Start-SmartM365CmdbSourceReceipt -ScriptPath $inventoryPath -SourceRootPath $dataLast
  $collectedAt = [datetimeoffset]::UtcNow.ToString('o')
  $tenantRows = @(
    [pscustomobject]@{ TenantKey='prod'; Id='sku-e3'; TenantSkuPartNumber='SPE_E3'; TenantPrepaidEnabled=10; TenantConsumedUnits=8; CollectedAtUtc=$collectedAt }
    [pscustomobject]@{ TenantKey='prod'; Id='sku-f3'; TenantSkuPartNumber='SPE_F1'; TenantPrepaidEnabled=5; TenantConsumedUnits=5; CollectedAtUtc=$collectedAt }
  )
  $published = @(
    Write-FixtureCsv 'M365_Licenses_Tenant.csv' $tenantRows
    Write-FixtureCsv 'M365_Licenses_Users.csv' @(
      [pscustomobject]@{ TenantKey='prod'; UserId='u1'; SkuId='sku-e3'; SkuPartNumber='SPE_E3'; 'User principal name'='u1@example.invalid'; 'Display name'='User One'; Source='Direct'; GroupsAssigningSku='' }
      [pscustomobject]@{ TenantKey='prod'; UserId='u2'; SkuId='sku-e3'; SkuPartNumber='SPE_E3'; 'User principal name'='u2@example.invalid'; 'Display name'='User Two'; Source='Direct'; GroupsAssigningSku='' }
      [pscustomobject]@{ TenantKey='prod'; UserId='u3'; SkuId='sku-f3'; SkuPartNumber='SPE_F1'; 'User principal name'='u3@example.invalid'; 'Display name'='User Three'; Source='Direct'; GroupsAssigningSku='' }
    )
    Write-FixtureCsv 'M365_Licenses_AssignmentPaths.csv' @(
      [pscustomobject]@{ TenantKey='prod'; UserId='u1'; SkuId='sku-e3'; AssignedByGroupId=''; AssignmentRoute='Direct'; AssignmentState='Active'; AssignmentError='None'; LastUpdatedDateTime=$collectedAt }
      [pscustomobject]@{ TenantKey='prod'; UserId='u4'; SkuId='sku-f3'; AssignedByGroupId='g1'; AssignmentRoute='Group'; AssignmentState='Error'; AssignmentError='CountViolation'; LastUpdatedDateTime=$collectedAt }
    )
    Write-FixtureCsv 'M365_Licenses_ServicePlans_Catalog.csv' @([pscustomobject]@{ TenantKey='prod'; SkuId='sku-e3'; PlanId='plan-1'; PlanName='EXCHANGE_S_ENTERPRISE' })
    Write-FixtureCsv 'M365_Licenses_UserServicePlanStates.csv' @([pscustomobject]@{ TenantKey='prod'; UserId='u1'; PlanStates='E' })
    Write-FixtureCsv 'M365_EntraGroups_All.csv' @([pscustomobject]@{ TenantKey='prod'; GroupId='g1'; DisplayName='License group' })
  )
  foreach ($path in $published) { Register-SmartM365SourceCsv -Path $path }
  $rowsForReport = if ($TenantRowsOverride) { $TenantRowsOverride } else { $tenantRows }
  $report = Publish-LicensesRunReport -Folder $dataLast -TenantKey 'prod' -CollectedAtUtc $collectedAt -TenantRows $rowsForReport -DependencyRulesPath $DependencyRulesPath -ExclusionsPath $ExclusionsPath
  Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope 'CMDB:skus,license_paths,plans,user_plans,groups'
  $completed = Complete-SmartM365CmdbSourceReceipt -Status 'Completed'
  $runState = Get-Content -LiteralPath ($completed -replace '\.current\.json\.txt$','.run.json.txt') -Raw | ConvertFrom-Json
  return [pscustomobject]@{ Report=$report; Receipt=(Get-Content -LiteralPath $completed -Raw | ConvertFrom-Json); RunState=$runState }
}

try {
  $today = [datetime]::UtcNow.Date
  $old = $today.AddDays(-120).ToString('yyyy-MM-dd')
  Write-FixtureCsv 'M365_Users_Active.csv' @(
    [pscustomobject]@{ TenantKey='prod'; 'Object Id'='u1'; 'User principal name'='u1@example.invalid'; 'Display name'='User One'; UserType='Member'; AccountEnabled='False'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime='' }
    [pscustomobject]@{ TenantKey='prod'; 'Object Id'='u2'; 'User principal name'='u2@example.invalid'; 'Display name'='User Two'; UserType='Member'; AccountEnabled='True'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime=$old }
    [pscustomobject]@{ TenantKey='prod'; 'Object Id'='u3'; 'User principal name'='u3@example.invalid'; 'Display name'='User Three'; UserType='Member'; AccountEnabled='False'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime='' }
    [pscustomobject]@{ TenantKey='prod'; 'Object Id'='u4'; 'User principal name'='u4@example.invalid'; 'Display name'='User Four'; UserType='Member'; AccountEnabled='True'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime=$old }
  ) | Out-Null
  Write-ForeignReceipt 'SmartM365-ActiveUsers-Inventory' 'M365_Users_Active.csv'
  Write-FixtureCsv 'Exchange_EXO_Mailboxes_AllDomains.csv' @(
    [pscustomobject]@{ TenantKey='prod'; ExternalDirectoryObjectId='u1'; RecipientTypeDetails='UserMailbox'; NativeIdentityStatus='Observed'; TotalItemSizeGB='1'; ArchiveStatus='None'; LitigationHoldEnabled='False'; RetentionHoldEnabled='False' }
    [pscustomobject]@{ TenantKey='prod'; ExternalDirectoryObjectId='u2'; RecipientTypeDetails='UserMailbox'; NativeIdentityStatus='Observed'; TotalItemSizeGB='1'; ArchiveStatus='None'; LitigationHoldEnabled='False'; RetentionHoldEnabled='False' }
    [pscustomobject]@{ TenantKey='prod'; ExternalDirectoryObjectId='u3'; RecipientTypeDetails='UserMailbox'; NativeIdentityStatus='Observed'; TotalItemSizeGB='1'; ArchiveStatus='None'; LitigationHoldEnabled='False'; RetentionHoldEnabled='False' }
  ) | Out-Null
  Write-ForeignReceipt 'SmartM365-EXO-Mailboxes-Inventory' 'Exchange_EXO_Mailboxes_AllDomains.csv'

  # 1. Nominal full run: report published before completion and recorded as optional receipt outputs.
  $first = Invoke-LicensingRun
  Check $first.Report.Published 'Full run publishes the license report.'
  Check $first.Report.UploadSatisfied 'Disabled SharePoint upload satisfies the report publication.'
  Check ($first.Receipt.Status -eq 'Completed' -and $first.Receipt.ConsumerScopeQualified) 'Licensing receipt completes with the report.'
  Check (@($first.Receipt.RequiredFiles).Count -eq 6) 'Required licensing files are unchanged.'
  foreach ($name in $reportNames) {
    $record = @($first.Receipt.Files | Where-Object File -eq $name)
    Check ($record.Count -eq 1 -and -not $record[0].Required -and $record[0].SHA256 -eq (Get-FileHash -LiteralPath (Join-Path $dataLast $name) -Algorithm SHA256).Hash) "Report CSV is recorded as an optional receipt file: $name"
    Check (@($first.Receipt.OptionalFiles | Where-Object { $_.File -eq $name -and $_.Published }).Count -eq 1) "Report CSV is declared published: $name"
  }
  $snapshot = Get-Content -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
  Check ($snapshot.SchemaVersion -eq 2 -and $snapshot.LogicVersion -eq '1.44' -and $snapshot.SnapshotId -eq $first.Report.SnapshotId) 'Snapshot schema and logic versions are unchanged.'
  Check ($snapshot.ReceiptRun.RunId -eq $first.Receipt.RunId -and $first.Report.ReceiptRunId -eq $first.Receipt.RunId -and $snapshot.ReceiptRun.Producer -eq 'SmartM365-Licences-Inventory.ps1') 'Snapshot carries the real licensing receipt run.'
  Check ($null -eq $first.RunState.PSObject.Properties['Notes']) 'A published report adds no run note.'
  foreach ($name in $decisionNames) {
    $record = @($first.Receipt.Files | Where-Object File -eq $name)
    Check ($record.Count -eq 1 -and -not $record[0].Required -and $record[0].SHA256 -eq (Get-FileHash -LiteralPath (Join-Path $dataLast $name) -Algorithm SHA256).Hash) "SKU decision CSV is recorded as an optional receipt file: $name"
    Check (@($first.Receipt.OptionalFiles | Where-Object { $_.File -eq $name -and $_.Published }).Count -eq 1) "SKU decision CSV is declared published: $name"
    $decisionRows = @(Import-Csv -LiteralPath (Join-Path $dataLast $name))
    Check ($decisionRows.Count -gt 0 -and @($decisionRows | Where-Object { $_.SnapshotId -ne $first.Report.SnapshotId -or $_.ReceiptRunId -ne $first.Receipt.RunId }).Count -eq 0) "SKU decision rows carry the report snapshot and receipt run: $name"
  }
  $skuCandidates = @(Import-Csv -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportSkuCandidates.csv'))
  Check (@($skuCandidates | Where-Object { $_.Recommendation -eq 'Referenced' -and $_.CountedIn -eq 'ReportCandidates:Recovery' }).Count -eq 2) 'Recovery suites are referenced, not recounted, in the SKU decisions.'
  Check ($null -ne $snapshot.Rendering -and $snapshot.Rendering.UsageAvailable) 'Snapshot carries its rendering evidence.'
  $licenseUsersRendering = @($snapshot.Rendering.UsageSources | Where-Object Name -eq 'M365_Licenses_Users.csv')[0]
  Check ($licenseUsersRendering.Ready -and $licenseUsersRendering.Provenance -eq 'current run publication' -and -not $licenseUsersRendering.Forced) 'Licensing users are qualified by the current run.'
  Check $snapshot.AssignmentIssues.Available 'Assignment issues are qualified by the current run.'
  Check (@($snapshot.AssignmentIssues.Rows | Where-Object { $_.UserId -eq 'u4' -and $_.Category -eq 'Capacity wait' }).Count -eq 1) 'Capacity wait is reported.'
  $sources = @(Import-Csv -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportSources.csv'))
  $tenantSource = @($sources | Where-Object Name -eq 'M365_Licenses_Tenant.csv')[0]
  $tenantRecord = @($first.Receipt.Files | Where-Object File -eq 'M365_Licenses_Tenant.csv')[0]
  Check ($tenantSource.SHA256 -eq $tenantRecord.SHA256) 'Report source hash equals the licensing receipt hash.'
  Check (@($sources | Where-Object { $_.Name -eq 'M365_Licenses_Users.csv' -and $_.Ready -eq 'True' -and $_.Provisional -eq 'False' }).Count -eq 1) 'Licensing users source is ready and not provisional.'
  $candidates = @(Import-Csv -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportCandidates.csv'))
  Check (@($candidates | Where-Object CandidateType -eq 'Recovery').Count -eq 2) 'Disabled E3 and F3 users are recovery candidates.'
  Check (@($candidates | Where-Object SnapshotId -ne $first.Report.SnapshotId).Count -eq 0) 'Report rows carry the published snapshot identity.'
  [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $dataLast -ExpectedTenantKey 'prod' -MailStatePath $mailStatePath)
  Check ($script:SentMail.Count -eq 1 -and $script:SentMail[0].BodyHtml -like "*Snapshot: $($first.Report.SnapshotId)*") 'Email renders the published snapshot.'
  Check ($script:SentMail[0].BodyHtml -like '*M365_Licenses_Users.csv: Ready (*current run publication)*') 'Email states the current-run provenance.'

  # 2. Second run on the same day: report republished, email limited to one per day.
  $second = Invoke-LicensingRun
  Check ($second.Report.Published -and $second.Report.SnapshotId -ne $first.Report.SnapshotId) 'Every full run republishes the report.'
  [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $dataLast -ExpectedTenantKey 'prod' -MailStatePath $mailStatePath)
  Check ($script:SentMail.Count -eq 1) 'The daily mail gate still limits the email.'

  # 3. Report failure: receipt unaffected, previous report files left in place, email refused.
  $before = @{}
  foreach ($name in $reportNames) { $before[$name] = (Get-FileHash -LiteralPath (Join-Path $dataLast $name) -Algorithm SHA256).Hash }
  $mismatchedRows = @([pscustomobject]@{ TenantKey='prod'; Id='sku-e3'; TenantSkuPartNumber='SPE_E3'; TenantPrepaidEnabled=99; TenantConsumedUnits=8; CollectedAtUtc=[datetimeoffset]::UtcNow.ToString('o') })
  $failed = Invoke-LicensingRun -TenantRowsOverride $mismatchedRows
  Check (-not $failed.Report.Published -and $failed.Report.Reason -like '*differs from the published tenant license CSV*') 'An integrity failure prevents report publication.'
  Check ($failed.Receipt.Status -eq 'Completed') 'A report failure does not fail the licensing receipt.'
  Check (@($failed.Receipt.Files | Where-Object { $_.File -in $reportNames }).Count -eq 0) 'Unpublished report files are not recorded.'
  Check (@($failed.Receipt.OptionalFiles | Where-Object Published).Count -eq 0) 'Optional report files are declared unpublished.'
  Check (@($failed.Receipt.Files | Where-Object { $_.File -in $decisionNames }).Count -eq 0) 'SKU decisions are not published without a report.'
  foreach ($name in $reportNames) { Check ((Get-FileHash -LiteralPath (Join-Path $dataLast $name) -Algorithm SHA256).Hash -eq $before[$name]) "Previous report file is left untouched: $name" }
  $refused = ''
  try { [void](Read-LicensesPublishedReport -Folder $dataLast -TenantKey 'prod') } catch { $refused = $_.Exception.Message }
  Check ($refused -like '*belongs to another licensing receipt run*') 'A stale report is refused by the email reader.'
  Check (@($script:Log | Where-Object { $_ -like 'WARNING License report not republished in this run*belong to an earlier run and are not recorded by this receipt*' }).Count -eq 1) 'Report failure is logged as a warning.'
  Check (@($failed.RunState.Notes | Where-Object { $_ -like 'License report not republished in this run*not recorded by this receipt*' }).Count -eq 1) 'Report failure is recorded in the run state.'
  $staleSnapshot = Get-Content -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
  Check ($staleSnapshot.ReceiptRun.RunId -ne $failed.Receipt.RunId) 'The remaining snapshot names an earlier receipt run.'

  # 4. Missing usage source: report still published with N/D usage.
  Move-Item -LiteralPath (Join-Path $dataLast 'M365_Users_Active.csv') -Destination (Join-Path $root 'M365_Users_Active.csv')
  $missingUsage = Invoke-LicensingRun
  Check ($missingUsage.Report.Published -and $missingUsage.Receipt.Status -eq 'Completed') 'A missing usage source does not prevent publication.'
  $missingSnapshot = Get-Content -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
  Check (-not $missingSnapshot.AssignmentIssues.Available) 'Assignment issues become N/D without active users.'
  Check (@($missingSnapshot.Rendering.UsageSources | Where-Object { $_.Name -eq 'M365_Users_Active.csv' -and -not $_.Ready -and $_.Reason -eq 'missing' }).Count -eq 1) 'Missing usage source is disclosed.'
  Move-Item -LiteralPath (Join-Path $root 'M365_Users_Active.csv') -Destination (Join-Path $dataLast 'M365_Users_Active.csv')

  # 5. SharePoint upload failure: local publication recorded, email withheld by the collector.
  $global:EnableSharePointUpload = $true
  $script:FailUpload = $true
  $uploadFailure = Invoke-LicensingRun
  Check ($uploadFailure.Report.Published -and -not $uploadFailure.Report.UploadSatisfied) 'Upload failure is reported separately from local publication.'
  Check (@($uploadFailure.Receipt.Files | Where-Object { $_.File -in $reportNames }).Count -eq 5) 'Locally published report remains recorded by the receipt.'
  Check (@($uploadFailure.RunState.Notes | Where-Object { $_ -like 'License report *recorded by this receipt but was not uploaded to SharePoint*' }).Count -eq 1) 'Upload failure is recorded in the run state.'
  Check (@($uploadFailure.RunState.Notes | Where-Object { $_ -like 'License SKU decision file *not uploaded to SharePoint*' }).Count -eq 2) 'SKU decision upload failures are recorded in the run state.'
  $script:FailUpload = $false
  $script:Uploads.Clear()
  $uploaded = Invoke-LicensingRun
  Check ($uploaded.Report.UploadSatisfied -and $script:Uploads.Count -eq 8 -and $script:Uploads[0] -eq 'M365_Licenses_ReportSnapshot.json.txt' -and $script:Uploads[7] -eq 'M365_Licenses_ReportSkuCapacity.csv') 'Snapshot, five report CSVs and two SKU decision CSVs are uploaded after registration.'
  Remove-Variable -Name EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue

  # 5b. SKU decision failure: the report stays published, the decision files are not recorded.
  $decisionFailure = Invoke-LicensingRun -DependencyRulesPath (Join-Path $root 'missing-rules.json')
  Check ($decisionFailure.Report.Published -and $decisionFailure.Receipt.Status -eq 'Completed') 'A SKU decision failure does not affect the report.'
  Check (@($decisionFailure.Receipt.Files | Where-Object { $_.File -in $reportNames }).Count -eq 5 -and @($decisionFailure.Receipt.Files | Where-Object { $_.File -in $decisionNames }).Count -eq 0) 'Only the report CSVs are recorded after a SKU decision failure.'
  Check (@($decisionFailure.RunState.Notes | Where-Object { $_ -like 'License SKU decisions not published in this run*belong to an earlier run*' }).Count -eq 1) 'SKU decision failure is recorded in the run state.'
  $badExclusions = Join-Path $root 'invalid-exclusions.json'
  Set-Content -LiteralPath $badExclusions -Value '{'
  foreach ($case in @(@{ Name='missing'; Path=(Join-Path $root 'missing-exclusions.json') }, @{ Name='invalid'; Path=$badExclusions })) {
    $exclusionFailure = Invoke-LicensingRun -ExclusionsPath $case.Path
    Check ($exclusionFailure.Report.Published -and $exclusionFailure.Receipt.Status -eq 'Completed' -and @($exclusionFailure.Receipt.Files | Where-Object { $_.File -in $reportNames }).Count -eq 5) "Report is published with $($case.Name) versioned exclusions."
    Check (@($exclusionFailure.Receipt.Files | Where-Object { $_.File -in $decisionNames }).Count -eq 0 -and @($exclusionFailure.RunState.Notes | Where-Object { $_ -like 'License SKU decisions not published in this run*' }).Count -eq 1) "No SKU decision CSV and a run note with $($case.Name) versioned exclusions."
    Check (@($script:Log | Where-Object { $_ -like 'WARNING License SKU decisions not published*exclusions*' }).Count -ge 1) "Exclusion failure is logged as a warning ($($case.Name))."
  }

  # 6. Accepted partial AD coverage stays provisional in the report.
  $scope = 'CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members'
  $coverage = @{
    Kind='ADDomainCoverage'; Status='PartialAccepted'; Reason='NonBlockingDomainErrors'
    ExpectedDomains=@('observed.example.invalid','unavailable.example.invalid'); CollectedDomains=@('observed.example.invalid')
    UnavailableDomains=@('unavailable.example.invalid'); NonBlockingDomainErrors=@('unavailable.example.invalid')
  }
  $adPath = Write-FixtureCsv 'AD_Users_AllDomains.csv' @(
    [pscustomobject]@{ TenantKey='prod'; ImmutableId_AD='a2'; UserPrincipalName='u2@example.invalid'; Enabled='True'; LastLogonDate=$old; DomainName='observed.example.invalid' }
  )
  @{ Producer='SmartM365-ActiveDirectory-Inventory.ps1'; Status='Completed'; Errors=0; Scope=$scope; ScopeQualification='ConsumerScope'
     ConsumerScopeQualified=$true; FullInventoryQualified=$false; IsPartialInventory=$true; RequiredFiles=@('AD_Users_AllDomains.csv'); DomainCoverage=$coverage
     Files=@(@{ File='AD_Users_AllDomains.csv'; Status='Success'; Errors=0; Scope=$scope; IsPartialInventory=$true; SHA256=(Get-FileHash -LiteralPath $adPath).Hash; DomainCoverage=$coverage }) } |
    ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $dataLast 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt')
  $partial = Invoke-LicensingRun
  Check $partial.Report.Published 'Partial AD coverage does not prevent publication.'
  $adSource = @(Import-Csv -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportSources.csv') | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0]
  Check ($adSource.Ready -eq 'True' -and $adSource.Provisional -eq 'True') 'Accepted partial AD coverage is provisional.'
  $partialRead = Read-LicensesPublishedReport -Folder $dataLast -TenantKey 'prod'
  Check (@($partialRead.Usage.Sources | Where-Object { $_.Name -eq 'AD_Users_AllDomains.csv' -and $_.PartialScopeAccepted }).Count -eq 1) 'Email rendering keeps the partial AD qualification.'

  # 7. Without an active licensing receipt, no report can be qualified.
  $orphan = Publish-LicensesRunReport -Folder $dataLast -TenantKey 'prod' -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -TenantRows @()
  Check (-not $orphan.Published -and $orphan.Reason -like 'No active licensing source receipt*') 'A report needs the active licensing receipt.'

  # 8. Email-only mode rejects ForceAdCsvAnalysis and the removed bypass switch.
  $inventoryText = Get-Content -LiteralPath $inventoryPath -Raw
  Check ($inventoryText -notmatch 'BypassLicenseUsersReceipt') 'The licensing receipt bypass switch is removed.'
  Check ($inventoryText -match "if \(\`$ForceAdCsvAnalysis\) \{\s*throw '-ForceAdCsvAnalysis applies to a full collection run") 'Email-only mode rejects ForceAdCsvAnalysis.'
  Check ($inventoryText -match '\$MaxItems -eq 0 -and \$TopUsers -eq 0 -and -not \(Test-SmartM365MaxItemsMode\)\) \{\s*\$currentOperation = ''Publish license report''') 'Sampled runs do not publish a license report.'
}
finally {
  foreach ($name in $saved.Keys) {
    if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
    else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
"PASS: License report publication offline checks ($script:checks checks)."
