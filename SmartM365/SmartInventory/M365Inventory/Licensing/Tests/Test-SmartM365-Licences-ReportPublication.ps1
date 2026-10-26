#Requires -Version 7.0
<#
.SYNOPSIS
  Offline checks for the license report published by every full licensing run and declared as optional receipt outputs.
.VERSION
1.1.3
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
  'Get-LicensesSuiteProduct', 'Get-LicensesFocusedSummaryRows', 'Get-LicensesAdditionalOverviewRows', 'Get-LicensesOtherProduct', 'New-LicensesOverviewCardHtml', 'ConvertTo-LicensesActivityDate',
  'ConvertTo-LicensesMailboxSizeGb', 'Get-LicensesQualifiedPartialAdCoverage', 'Get-LicensesCsvSource', 'Get-LicensesPublicationPath', 'Import-LicensesSourceCsv',
  'Read-LicensesIndexedSource', 'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary', 'Get-LicensesFocusedUsageRows', 'Import-LicensesSourceCsvCached', 'Get-LicensesUserAdScopes', 'Get-LicensesUserScope', 'Get-LicensesScopeBreakdown', 'New-LicensesScopeTablesHtml', 'New-LicensesSummaryMailContent', 'Send-LicensesSummaryMailScope', 'Get-LicensesDomainMailSettings', 'Get-LicensesDomainMailStatePath', 'Get-LicensesDomainPublishedReport',
  'Format-LicensesMetric', 'Get-LicensesAssignmentIssues', 'Get-LicensesAlertSummary', 'New-LicensesRecoveryWorkbook', 'Publish-LicensesReportCsv',
  'Publish-LicensesReportSnapshot', 'Write-LicensesDailyMailState', 'Enter-LicensesDailyMailGate', 'Send-LicensesFocusedSummaryEmail',
  'Read-LicensesTenantSnapshot', 'Publish-LicensesRunReport', 'Invoke-LicensesReportUpload', 'ConvertFrom-LicensesJsonElement', 'ConvertTo-LicensesCountTable',
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
$coreModule = Import-Module $coreManifest -MinimumVersion '1.0.95' -Force -PassThru

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
  # A staged licensing run writes its own CSVs to the copy that its completion promotes.
  $path = Get-LicensesPublicationPath -Path (Join-Path $dataLast $Name)
  $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8NoBOM
  return $path
}
function Invoke-LicensingRun {
  param([object[]]$TenantRowsOverride, [string]$DependencyRulesPath = $rulesPath, [string]$ExclusionsPath = $exclusionsPath,
    [switch]$Staged, [string]$CompleteStatus = 'Completed', [scriptblock]$BeforeComplete, [int]$E3Enabled = 10)
  # Same order as the collector: publish canonical CSVs, publish the report, then complete the receipt.
  $global:csvGeneratedPaths = $null
  Start-SmartM365CmdbSourceReceipt -ScriptPath $inventoryPath -SourceRootPath $dataLast
  if ($Staged) { Enable-SmartM365StagedSourcePublication }
  $collectedAt = [datetimeoffset]::UtcNow.ToString('o')
  $tenantRows = @(
    [pscustomobject]@{ TenantKey='prod'; Id='sku-e3'; TenantSkuPartNumber='SPE_E3'; TenantPrepaidEnabled=$E3Enabled; TenantConsumedUnits=8; CollectedAtUtc=$collectedAt }
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
  if ($BeforeComplete) { & $BeforeComplete }
  $null = if ($CompleteStatus -eq 'Completed') { Complete-SmartM365CmdbSourceReceipt -Status 'Completed' } else { Complete-SmartM365CmdbSourceReceipt -Status $CompleteStatus -ErrorCount 1 }
  $completed = Join-Path $dataLast 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
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
  Check ($uploaded.Report.UploadSatisfied -and $script:Uploads.Count -eq 10 -and $script:Uploads[0] -eq 'M365_Licenses_ReportSnapshot.json.txt' -and $script:Uploads[6] -eq 'M365_Licenses_ReportByDomain.csv' -and $script:Uploads[7] -eq 'M365_Licenses_ReportByCountry.csv' -and $script:Uploads[9] -eq 'M365_Licenses_ReportSkuCapacity.csv') 'Snapshot, five report CSVs, two AD scope CSVs and two SKU decision CSVs are uploaded after registration.'
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

  # 9. Staged run: CSVs, report CSVs and snapshot reach DATA-LAST together with the receipt, never before.
  Check ($inventoryText -match "Start-SmartM365CmdbSourceReceipt [^\r\n]*\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") 'The collector enables staging right after its receipt.'
  $watched = @($reportNames + $decisionNames + @('M365_Licenses_ReportSnapshot.json.txt', 'M365_Licenses_Users.csv', 'M365_Licenses_Tenant.csv', 'M365_Licenses_UserServicePlanStates.csv'))
  $receiptFile = Join-Path $dataLast 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  function Get-WatchedHashTable { $hashes = @{}; foreach ($name in $watched) { $file = Join-Path $dataLast $name; $hashes[$name] = if (Test-Path -LiteralPath $file) { (Get-FileHash -LiteralPath $file).Hash } else { '' } }; return $hashes }
  $beforeStaged = Get-WatchedHashTable
  $receiptBefore = Get-Content -LiteralPath $receiptFile -Raw | ConvertFrom-Json
  $global:EnableSharePointUpload = $true
  $script:Uploads.Clear()
  $script:MidRun = $null
  $failedStaged = Invoke-LicensingRun -Staged -CompleteStatus 'Failed' -BeforeComplete { $script:MidRun = Get-WatchedHashTable }
  Check $failedStaged.Report.Published 'A staged run builds its report from its staged copies.'
  Check (@($watched | Where-Object { $script:MidRun[$_] -ne $beforeStaged[$_] }).Count -eq 0) 'A running staged licensing run changed DATA-LAST.'
  $afterFailed = Get-WatchedHashTable
  Check (@($watched | Where-Object { $afterFailed[$_] -ne $beforeStaged[$_] }).Count -eq 0) 'A failed staged licensing run changed DATA-LAST.'
  Check ((Get-Content -LiteralPath $receiptFile -Raw | ConvertFrom-Json).RunId -eq $receiptBefore.RunId) 'A failed staged licensing run replaced the receipt.'
  Check ($script:Uploads.Count -eq 0) 'A staged licensing run uploaded files before its promotion.'
  Check (@(Get-ChildItem -LiteralPath (Join-Path $dataLast '.staging/SmartM365-Licences-Inventory') -Force).Count -eq 0) 'A failed staged licensing run left its staging folder.'
  # Production case: this run's tenant capacity differs from the previous run's DATA-LAST tenant CSV.
  $staged = Invoke-LicensingRun -Staged -E3Enabled 12 -BeforeComplete { $script:MidRun = Get-WatchedHashTable }
  Check $staged.Report.Published 'A staged run with a changed tenant capacity did not publish its report.'
  Check (@($watched | Where-Object { $script:MidRun[$_] -ne $beforeStaged[$_] }).Count -eq 0) 'A staged licensing run changed DATA-LAST before its completion.'
  Check ($staged.Receipt.Status -eq 'Completed' -and $staged.Receipt.RunId -ne $receiptBefore.RunId) 'A complete staged licensing run did not publish its receipt.'
  foreach ($name in @($reportNames + $decisionNames + @('M365_Licenses_Users.csv', 'M365_Licenses_Tenant.csv'))) {
    $record = @($staged.Receipt.Files | Where-Object File -eq $name)
    Check ($record.Count -eq 1 -and $record[0].SHA256 -eq (Get-FileHash -LiteralPath (Join-Path $dataLast $name)).Hash) "Promoted licensing CSV and receipt disagree: $name"
  }
  Check (@(Import-Csv -LiteralPath (Join-Path $dataLast 'M365_Licenses_Tenant.csv') | Where-Object { $_.TenantSkuPartNumber -eq 'SPE_E3' -and $_.TenantPrepaidEnabled -eq '12' }).Count -eq 1) 'The changed tenant capacity was not promoted.'
  $promotedSnapshot = Get-Content -LiteralPath (Join-Path $dataLast 'M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
  Check ($promotedSnapshot.SnapshotId -eq $staged.Report.SnapshotId -and $promotedSnapshot.ReceiptRun.RunId -eq $staged.Receipt.RunId) 'The promoted report snapshot does not name the promoted receipt run.'
  Check ($staged.RunState.Status -eq 'Completed' -and -not $staged.RunState.UnqualifiedPublication -and $staged.RunState.StagedPublication) 'The staged licensing run state is not qualified.'
  Check ($staged.Report.UploadDeferred -and $script:Uploads.Count -eq 0) 'The staged report upload was not deferred to the promotion.'
  $uploadedStaged = Invoke-LicensesReportUpload -Result $staged.Report -ReportPaths $staged.Report.ReportPaths -SkuPaths $staged.Report.SkuPaths
  Check ($uploadedStaged.UploadSatisfied -and $script:Uploads.Count -eq 10 -and $script:Uploads[0] -eq 'M365_Licenses_ReportSnapshot.json.txt') 'The promoted report was not uploaded by canonical name.'
  Check (@($staged.Report.ReportPaths + $staged.Report.SkuPaths | Where-Object { $_ -like '*.staging*' }).Count -eq 0) 'A staging path reached the report upload.'
  Remove-Variable -Name EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue
  $read = Read-LicensesPublishedReport -Folder $dataLast -TenantKey 'prod'
  Check ($null -ne $read) 'The email reader accepts the promoted report.'
}
finally {
  # Releases the producer lock if a check failed while a receipt was active.
  try { Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
  foreach ($name in $saved.Keys) {
    if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
    else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
"PASS: License report publication offline checks ($script:checks checks)."

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDWE9ohpNCRTUPK
# Pz+DMdaIM7OaccMYCAaJMZ+V5dEXDqCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQghnPRjn0aLl+y
# 2gcslIoFUu4DClnt2kmG8DiVk/OBzRgwDQYJKoZIhvcNAQEBBQAEggGAneOD9ZXw
# QSsvgGdOXMJT94y4XC9s1LrnN2ihsW78kmJEAiyG7wa4JgkDK6v3X9odKzAUWUgk
# lMH2fEwPSkHdfgurAWvIdd46HbZB5DJhx37r2aaEmg817GtHIb/MqMkO2jTTeJsa
# yVUhjEjcfRgG41qQCZJdIySbEAbmtTibJdZy6J93s+Sl16UUH+jBYbmQe6LAKUY5
# QmQ4A2uXfJwcFcZEwpRuVrlu5S4UyyRSV8rGEnAXKXpfDu2B8oK0GKMtukkLLkM1
# 708ULIAvkfG+YzCxEplV0aRWZZKJtIlnydWwofMOBKeqImM2hZeHRarq8RWGcrSn
# 3O6PH5j7bDX2RBZKBt7h3bAY6yRMLkH+LPa4tRhEGKEY/UkMU0HiP2vjT84rN1Ki
# wNxmdhDSF9frVJAusaqfUnEweTrw1z5fCRlIyFUT7D7Pd+a6pTqZXKlzR3yWGIe7
# Mx+SXtTjXdhNtx67ylbiP6XQLF1iN+6GDVuJm21hwUT+u6kOfDu/WbnnoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMDIwMTkxMDUwWjA/BgkqhkiG9w0BCQQxMgQwK/9Iy0bHz29hM61TXVdqE8aQ
# 6ZTFxoLhPCenfg6rVQJPEb3+WDKgijyg8maeou6cMA0GCSqGSIb3DQEBAQUABIIC
# AGfrq/8kpFQkmESNpjsb8oaOfTNMpvzdDFceqzXeqaYx1pGMwhEh6i1Up42x5hCp
# +yQoolp6w3/lN9GUd+PtieNbaG6gzH4DLeMfP+N5aUrRUD0wl8othJh0p+xac/3q
# dUcefhbTbMCUij9o+1YLU6j+P9CLZz5LtBAZVkGR3lTh4/JX7vbHzhEa5pquFr2g
# UE6Z/Eo150kvDE6MzC0nsAsAQhV7HAoOtrKnEwGOr0AZxPwe5ybu9TTtFKXWOpuq
# /E4bzi60vjkunuW3mEs7JYNGIjlLZbSbZEHm8Id0kneg8YTYWu18ao5LhYydtxQe
# On8eT0xQM/pagzavuEU183E039IV1grW6Z8ZhwssfmPBYi9C9HQ0ms0OBGwEfOfH
# zLe2+YZhkmJoZ4ChlBmyUrjd14CHyYSKjC8f6FvVH2IWGs2qMqdpIffgDIsX9nlN
# X+QfQM/+iMUk0gxSbfp8Sz2bNWGaKsw4sTq+HDRZE5fu6RFdnmC69fbwgD37xgWR
# TvajemYeXEJxQjFjqpzUh5dEbIhuWxdHIbmupvWSX4Fmj5HYMIT3WvxK2y4KUurp
# llb7qb0Cx5r99AJigBlmLY1h9YWL41xSpTUkHDyywCSPqUFkEdkvXb9HU6+mLsW9
# hG2VNXQ5zSGjfZHmXHh3UXa0W6QSthTIrVgY7kj7dxcJ
# SIG # End signature block
