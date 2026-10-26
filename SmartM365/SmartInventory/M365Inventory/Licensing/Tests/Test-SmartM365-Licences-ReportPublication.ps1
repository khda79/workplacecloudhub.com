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
  'Read-LicensesIndexedSource', 'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary', 'Get-LicensesFocusedUsageRows', 'Import-LicensesSourceCsvCached', 'Get-LicensesUserAdScopes', 'Get-LicensesUserScope', 'Get-LicensesScopeBreakdown', 'New-LicensesScopeTablesHtml', 'New-LicensesSummaryMailContent', 'Send-LicensesSummaryMailScope', 'ConvertTo-LicensesMailSchedule', 'Get-LicensesMailSchedule', 'Get-LicensesDomainMailSettings', 'Get-LicensesDomainMailStatePath', 'Get-LicensesDomainPublishedReport',
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAzy+1HmMfvaYpu
# uEK2EThagDOWLiV6cvlS4EdVRnuDvqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFBYB4mucTKSRPbdvUwhTgdwUd3TLS7GPXljANrRIxMRMA0GCSqG
# SIb3DQEBAQUABIIBgHLSmvVySr3cIaq0N6kVbEpVb47TtffViQStTjYNmf/wlS7X
# QxfaNF5OdHMR1NwLVtlIgl/28X6BquL0C1RGMOqlATtH4xGfIV2wW1i0Wm3nQDvR
# MkCSIdYqUiDU/Cb534B43UXqCOiMe1y4YdcwaT0B0p7goTpTX8J2mPubWSSCuOSn
# Ah7RT3SJ8o7FOhAwr8lXXZLzwGMP975H8rS9IJNzPswsEyT4jqQ/YDnpY4C1QhKW
# 86RCaj62bY9vu5cGt5F1evyDLiCgk1EdnbQAsgw4ZW4o5BNgirQpZFnk8UgahER9
# Jfta+xiPqO/U9XncbtVDCdiqeHsrHoRBAmKbPagJ7LZmGYUVV9ee7kKsBPM5AdGN
# 5Rx2FzBQfygJAho0VNlTf/LzOk26KugMIpgEsCr3ARui9U1GzZEC7IK3DbNOwfhz
# eaibNeuTaNbydHNzlR/Yah5fHD4BYCF8wLafugTEQ405E9JzchhPolLnWYtpkEpB
# S8+mF1vRsZmGNkrIxKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYyMzA5
# MTRaMC8GCSqGSIb3DQEJBDEiBCDlUFTAuuQtWZM0SxcFg9p+ymGLxCI79UvcqXwT
# rvYwJTANBgkqhkiG9w0BAQEFAASCAgC1x8GqCQ+1i9nhc2Vk/+ffp6n7L9qZsATp
# vfMQocU6OnTzz/tJsnxKh+4IFJKhXgZahqY2VUbbvicunJ+CDHZJ+eyQOD00XR7B
# ddvrRw9HzMs1ZtK8LZwDVPSdC5nsfdxZTBc9Q94ns20SsxNOZdX3SnYlNEKQpoW1
# GiiPCPvAGmSlS9oN2vzguBtycfxQzOCIgyqWdObZrjH1/H55JshoU+3fB47BDHD7
# ewGEEERVXGMTNWeF5VXsGoddMGtq0Q0lflth1gjWsDYMbFIMeDzoTrAqhn+9R3F2
# RzbpusfM/rQJNMAaxhEnhciJRY4NdauE+sKzKsoq4hHQT7YI9PrTOUpy2ncvQACt
# GXIfoAy5X6JbMPpiQcelmHunUMasO+u0+JdmoDC7jhvuu6x8vIUYoJte7c0+GfDc
# spPTndPBj3hoIK4rkfheqg5oPyOAuUOhdnX/o9KO7a+7rfDn63Uyunx/UGnCd2Jt
# eOl0yVVO9XydW6kzivjLNGqCnDwg4XhOtIhaGfHi3Uiuz1A51AKoOnKbBiplSCDP
# jO8UxWJNP/FimnooPfXetv5rtbfEWZ83m5Lu2h3QVACAnluVecr82krSIdOS/RJ6
# MIdc4keurjd4pbrLYAm+cac77pSpLuQa2JMF+mfMaXS7d6f/IKk7xQLsr9XJrAfC
# e94I1dJp3w==
# SIG # End signature block
