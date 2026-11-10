<#
.SYNOPSIS
  Offline checks for section 08 of the license email: user mailboxes convertible to shared mailboxes.
  Synthetic prepared batch only; no tenant, Exchange or server access.
.VERSION
1.0.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$coreModulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
Import-Module -Name $coreModulePath -MinimumVersion '1.0.79' -Force -ErrorAction Stop
function Initialize-SmartM365ImportExcel { SmartM365.Core\Initialize-SmartM365ImportExcel -CheckOnly | Out-Null }
function WriteLog { param([string]$Message, [string]$Level) }
$names = @(
  'Get-LicensesOptionalProperty', 'Get-LicensesPublicationPath', 'Get-LicensesCsvSource', 'Import-LicensesSourceCsv', 'Import-LicensesSourceCsvCached',
  'Get-LicensesSharedConversionBatch', 'ConvertTo-LicensesSignInUtc', 'Get-LicensesSharedConversionDetails', 'Get-LicensesSharedConversionCounts',
  'Format-LicensesMetric', 'Get-LicensesAlertSummary', 'New-LicensesOverviewCardHtml', 'New-LicensesScopeTablesHtml', 'New-LicensesSummaryMailContent',
  'Get-LicensesDomainPublishedReport', 'ConvertTo-LicensesCountTable', 'ConvertFrom-LicensesJsonElement', 'New-LicensesRecoveryWorkbook'
)
$definitions = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name }, $true))
if ($definitions.Count -ne $names.Count) { throw 'Shared conversion functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }

$script:checks = 0
function Check {
  param([bool]$Condition, [string]$Label)
  $script:checks++
  if (-not $Condition) { throw "FAIL: $Label" }
}
function Get-Sha([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }

$tenant = 'contoso-test'
$batchId = '20261109T070000000Z-0a1b2c3d'
$asOf = [datetime]::new(2026, 11, 10, 6, 0, 0, [DateTimeKind]::Utc)
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SharedConversion-' + [guid]::NewGuid().ToString('N'))
$dataLast = Join-Path $root 'DATA-LAST'
$powerBi = Join-Path $root 'DATA-POWERBI'
$batchFolder = Join-Path (Join-Path $powerBi 'batches') $batchId
$evidencePath = Join-Path $batchFolder 'MailboxSharedConversionEvidence.csv'
$qualificationPath = Join-Path $batchFolder 'DirectoryCleanupSourceQualification.csv'
$manifestPath = Join-Path $batchFolder 'batch.json.txt'
$pointerPath = Join-Path $powerBi 'current.json.txt'

function New-ConversionRow {
  param([string]$Guid, [string]$State, [string]$Owner, [string]$Domain, [string]$Reason, [string]$OwnerSignIn = '', [string]$Licenses = '')
  [pscustomobject][ordered]@{
    'TenantKey'=$tenant; 'Shared Conversion Key'="key-$Guid"; 'Mailbox GUID'=$Guid; 'Primary SMTP Address'="$Guid@contoso.example"; 'Owner Entra Object ID'=$Owner
    'AD Domain'=$Domain; 'Country'='France'; 'Mailbox Size GB'='3.5'; 'Active Delegate Count'='2'; 'Delegate Count'='3'
    'Owner Last Sign-In'=$OwnerSignIn; 'Days Since Owner Sign-In'=$(if ($OwnerSignIn) { '250' } else { '' }); 'Owner Licenses'=$Licenses
    'Conversion State'=$State; 'Decision Reason'=$Reason; 'Secondary Reasons'='Organization retention policies and in-place holds not assessed'
  }
}
$baseRows = @(
  (New-ConversionRow 'g-s1' 'Strong candidate' 'u1' 'dom-a.example' 'Owner disabled, no sign-in for over 180 days, delegates active' '2026-03-01T08:00:00Z' 'SPE_F1')
  (New-ConversionRow 'g-s2' 'Strong candidate' 'u2' 'dom-b.example' 'Owner disabled, no sign-in recorded, delegates active')
  (New-ConversionRow 'g-s3' 'Strong candidate' 'u3' 'dom-a.example' 'Owner disabled, no sign-in recorded, delegates active')
  (New-ConversionRow 'g-s4' 'Strong candidate' 'u4' 'Unknown' 'Owner disabled, no sign-in recorded, delegates active')
  (New-ConversionRow 'g-s5' 'Strong candidate' 'u7' 'dom-a.example' 'Owner disabled, no sign-in recorded, delegates active')
  (New-ConversionRow 'g-r1' 'Review' 'u5' 'dom-a.example' 'Owner signed in between 90 and 180 days ago' '2026-07-01T08:00:00Z')
  (New-ConversionRow 'g-r2' 'Review' 'u6' 'dom-b.example' 'Owner signed in between 90 and 180 days ago' '2026-07-02T08:00:00Z')
  (New-ConversionRow 'g-r3' 'Review' '' 'Unknown' 'Owner not found')
  (New-ConversionRow 'g-n1' 'Not candidate' 'u8' 'dom-a.example' 'No delegate other than the owner')
)

function Write-Batch {
  param([object[]]$Rows = $baseRows, [string]$QualificationStatus = 'Observed', [string]$CreatedUtc = '2026-11-09T07:00:00Z', [string]$PointerTenant = $tenant,
        [switch]$OmitEvidence, [int]$RowDelta = 0, [switch]$TamperAfterManifest)
  if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
  New-Item -ItemType Directory -Path $dataLast, $batchFolder -Force | Out-Null
  $Rows | Export-Csv -LiteralPath $evidencePath -NoTypeInformation -Encoding utf8
  @(
    [pscustomobject][ordered]@{ TenantKey=$tenant; Source='Exchange Online mailboxes'; Status='Observed'; 'Reference UTC'='2026-11-09T06:55:00Z' }
    [pscustomobject][ordered]@{ TenantKey=$tenant; Source='Shared mailbox conversion (EXO user mailboxes)'; Status=$QualificationStatus; 'Reference UTC'='2026-11-09T06:55:00Z' }
  ) | Export-Csv -LiteralPath $qualificationPath -NoTypeInformation -Encoding utf8
  $files = @([ordered]@{ File='DirectoryCleanupSourceQualification.csv'; SHA256=(Get-Sha $qualificationPath); Rows=2 })
  if (-not $OmitEvidence) { $files += [ordered]@{ File='MailboxSharedConversionEvidence.csv'; SHA256=(Get-Sha $evidencePath); Rows=(@($Rows).Count + $RowDelta) } }
  $manifest = [ordered]@{ SchemaVersion=1; BatchId=$batchId; TenantKey=$tenant; CreatedUtc=$CreatedUtc; Files=$files; Artifacts=@() }
  [IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
  $pointer = [ordered]@{ TenantKey=$PointerTenant; BatchId=$batchId; ManifestSHA256=(Get-Sha $manifestPath); SchemaVersion=1 }
  [IO.File]::WriteAllText($pointerPath, ($pointer | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
  if ($TamperAfterManifest) { Add-Content -LiteralPath $evidencePath -Value '' }
}

try {
  # --- Qualified batch ---------------------------------------------------------------------
  Write-Batch
  $batch = Get-LicensesSharedConversionBatch -Folder $dataLast -TenantKey $tenant -AsOfUtc $asOf
  Check $batch.Ready "Qualified batch is ready ($($batch.Reason))"
  Check ($batch.BatchId -ceq $batchId) 'Batch identity is kept'
  Check ($batch.BatchCreatedUtc -ceq '2026-11-09T07:00:00Z') 'Batch creation time keeps UTC'
  Check ($batch.ReferenceUtc -ceq '2026-11-09T06:55:00Z') 'Reference time comes from the conversion qualification row'
  Check ($batch.NotCandidate -eq 1) 'Not candidate rows are counted, not kept'
  Check (@($batch.Rows).Count -eq 8) 'Strong candidate and Review rows are kept'
  Check (@($batch.Rows | Where-Object MailboxGuid -ceq 'g-s4')[0].Domain -ceq '') 'Unknown AD domain is not a domain'

  # --- Rejections: a missing or unqualified batch is not zero ---------------------------------
  $rejections = @(
    @{ Label='Stale batch'; Setup={ Write-Batch -CreatedUtc '2026-11-01T07:00:00Z' }; Reason='older than 7 days' }
    @{ Label='Future batch'; Setup={ Write-Batch -CreatedUtc '2026-11-11T07:00:00Z' }; Reason='older than 7 days or has an invalid creation time' }
    @{ Label='Other tenant pointer'; Setup={ Write-Batch -PointerTenant 'other' }; Reason='pointer identity' }
    @{ Label='File not in manifest'; Setup={ Write-Batch -OmitEvidence }; Reason='does not publish MailboxSharedConversionEvidence.csv' }
    @{ Label='Hash mismatch'; Setup={ Write-Batch -TamperAfterManifest }; Reason='differs from its batch manifest' }
    @{ Label='Row count mismatch'; Setup={ Write-Batch -RowDelta 1 }; Reason='row count mismatch' }
    @{ Label='Conversion not observed'; Setup={ Write-Batch -QualificationStatus 'Not ready' }; Reason='not observed' }
    @{ Label='Unknown state'; Setup={ $bad = @($baseRows[0..1]); $bad[1] = (New-ConversionRow 'g-x' 'Maybe' 'u9' 'dom-a.example' 'x'); Write-Batch -Rows $bad }; Reason='unknown state' }
    @{ Label='Duplicate GUID'; Setup={ Write-Batch -Rows @($baseRows[0], $baseRows[0]) }; Reason='duplicate mailbox GUID' }
    @{ Label='Missing pointer'; Setup={ Write-Batch; Remove-Item -LiteralPath $pointerPath }; Reason='no prepared batch pointer' }
  )
  foreach ($case in $rejections) {
    & $case.Setup
    $rejected = Get-LicensesSharedConversionBatch -Folder $dataLast -TenantKey $tenant -AsOfUtc $asOf
    Check (-not $rejected.Ready -and $rejected.Reason -like "*$($case.Reason)*") "$($case.Label) is rejected ($($rejected.Reason))"
    Check (@($rejected.Rows).Count -eq 0) "$($case.Label) publishes no row"
  }

  # --- Revalidation with this run's licenses and Entra export ----------------------------------
  Write-Batch
  $batch = Get-LicensesSharedConversionBatch -Folder $dataLast -TenantKey $tenant -AsOfUtc $asOf
  $activePath = Join-Path $dataLast 'M365_Users_Active.csv'
  @(
    [pscustomobject][ordered]@{ TenantKey=$tenant; 'Object Id'='U1'; 'User principal name'='u1@contoso.example'; AccountEnabled='False'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime=''; LastSignInDateTimeIso8601='2026-03-01T08:00:30Z'; LastNonInteractiveSignInDateTimeIso8601=''; LastSuccessfulSignInDateTimeIso8601='' }
    [pscustomobject][ordered]@{ TenantKey=$tenant; 'Object Id'='u2'; 'User principal name'='u2@contoso.example'; AccountEnabled='False'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime=''; LastSignInDateTimeIso8601=''; LastNonInteractiveSignInDateTimeIso8601=''; LastSuccessfulSignInDateTimeIso8601='' }
    [pscustomobject][ordered]@{ TenantKey=$tenant; 'Object Id'='u3'; 'User principal name'='u3@contoso.example'; AccountEnabled='True'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime=''; LastSignInDateTimeIso8601=''; LastNonInteractiveSignInDateTimeIso8601='2026-11-09T20:00:00Z'; LastSuccessfulSignInDateTimeIso8601='' }
    [pscustomobject][ordered]@{ TenantKey=$tenant; 'Object Id'='u5'; 'User principal name'='u5@contoso.example'; AccountEnabled='True'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime=''; LastSignInDateTimeIso8601=''; LastNonInteractiveSignInDateTimeIso8601=''; LastSuccessfulSignInDateTimeIso8601='' }
    [pscustomobject][ordered]@{ TenantKey=$tenant; 'Object Id'='u7'; 'User principal name'='u7@contoso.example'; AccountEnabled='False'; OnPremisesImmutableId=''; LastSuccessfulSignInDateTime=''; LastSignInDateTimeIso8601=''; LastNonInteractiveSignInDateTimeIso8601=''; LastSuccessfulSignInDateTimeIso8601='' }
  ) | Export-Csv -LiteralPath $activePath -NoTypeInformation -Encoding utf8
  $active = Get-LicensesCsvSource -Folder $dataLast -FileName 'M365_Users_Active.csv' -Columns @('TenantKey','Object Id','AccountEnabled') -AsOfUtc $asOf
  Check $active.Ready 'Synthetic Entra export is ready'
  $allSkus = @{ u1=@{ SPE_F1=$true }; u5=@{ SPE_E3=$true }; u7=@{ POWER_BI_PRO=$true } }
  $suites = @{ u1=@{ 'Microsoft 365 F3'=$true }; u5=@{ 'Microsoft 365 E3'=$true } }
  $recovery = @([pscustomobject]@{ UserId='u1'; License='Microsoft 365 F3' })
  $details = Get-LicensesSharedConversionDetails -Batch $batch -ActiveSource $active -ExpectedTenantKey $tenant -AllSkusByUser $allSkus -TargetSuitesByUser $suites -RecoveryDetails $recovery -ReadCache @{}
  Check $details.Available "Details are available ($($details.Reason))"
  $byGuid = @{}; foreach ($row in @($details.Rows)) { $byGuid[$row.MailboxGuid] = $row }
  Check ($byGuid['g-s1'].SameDayCheck -ceq 'Confirmed') 'Same sign-in within one minute stays confirmed'
  Check ($byGuid['g-s1'].Licenses -ceq 'SPE_F1' -and $byGuid['g-s1'].TargetSuites -ceq 'F3' -and $byGuid['g-s1'].RecoveryLicenses -ceq 'F3') 'Current licenses, suite and recovery overlap are joined'
  Check ($byGuid['g-s2'].SameDayCheck -ceq 'Confirmed' -and $byGuid['g-s2'].Licenses -ceq '') 'Unlicensed Strong candidate stays confirmed'
  Check ($byGuid['g-s3'].SameDayCheck -ceq 'Owner signed in after the prepared batch') 'Owner sign-in after the batch is a change'
  Check ($byGuid['g-s4'].SameDayCheck -ceq 'Owner not found in the current Entra export') 'Owner missing today is a change'
  Check ($byGuid['g-r3'].SameDayCheck -ceq 'Confirmed' -and $byGuid['g-r3'].UserId -ceq '') 'Review without owner is kept unchanged'
  Check ($details.NotCandidate -eq 1) 'Not candidate count is published for the tenant'

  $counts = Get-LicensesSharedConversionCounts -Rows @($details.Rows)
  Check ($counts.Strong -eq 5 -and $counts.StrongConfirmed -eq 3 -and $counts.StrongChanged -eq 2) 'Strong candidates split into confirmed and changed'
  Check ($counts.StrongLicensed -eq 2 -and $counts.StrongTargetSuite -eq 1 -and $counts.StrongOtherSkusOnly -eq 1) 'Licensed Strong candidates split by suite and other SKUs'
  Check ($counts.StrongSuites['F3'] -eq 1 -and $counts.StrongSuites['E3'] -eq 0) 'Strong suite counts'
  Check ($counts.StrongInRecovery -eq 1) 'Recovery overlap is counted'
  Check ($counts.Review -eq 3 -and $counts.ReviewLicensed -eq 1 -and $counts.ReviewSuites['E3'] -eq 1) 'Review is counted separately'
  Check (@($counts.ReviewReasons)[0].Reason -ceq 'Owner signed in between 90 and 180 days ago' -and @($counts.ReviewReasons)[0].Count -eq 2) 'Main review reason first'

  $notReady = Get-LicensesSharedConversionDetails -Batch ([pscustomobject]@{ Ready=$false; Reason='no prepared batch pointer'; BatchId=''; BatchCreatedUtc=''; ReferenceUtc=''; NotCandidate=0; Rows=@() }) -ActiveSource $active -ExpectedTenantKey $tenant -AllSkusByUser $allSkus -TargetSuitesByUser $suites
  Check (-not $notReady.Available -and $notReady.Reason -ceq 'prepared batch: no prepared batch pointer') 'Unqualified batch is N/D'
  $activeDown = [pscustomobject]@{ Ready=$false; Reason='older than 14 days' }
  $noEntra = Get-LicensesSharedConversionDetails -Batch $batch -ActiveSource $activeDown -ExpectedTenantKey $tenant -AllSkusByUser $allSkus -TargetSuitesByUser $suites
  Check (-not $noEntra.Available -and $noEntra.Reason -like 'Entra users source:*') 'Unqualified Entra export is N/D'

  # --- Snapshot round trip --------------------------------------------------------------------
  $json = [pscustomobject]@{ SharedConversion=$details } | ConvertTo-Json -Depth 16 -Compress
  $document = [Text.Json.JsonDocument]::Parse($json)
  try { $restored = (ConvertFrom-LicensesJsonElement -Element $document.RootElement).SharedConversion } finally { $document.Dispose() }
  $restoredCounts = Get-LicensesSharedConversionCounts -Rows @($restored.Rows)
  Check ($restoredCounts.StrongConfirmed -eq 3 -and $restoredCounts.Review -eq 3 -and $restored.NotCandidate -eq 1) 'Snapshot round trip keeps the evidence'

  # --- Email rendering ------------------------------------------------------------------------
  $OrgDomain = 'contoso.example'
  $summaryRows = @(foreach ($p in @('Microsoft 365 F1','Microsoft 365 F3','Microsoft 365 E3','Microsoft 365 E5')) { [pscustomobject]@{ Product=$p; Enabled=10; Consumed=5; Subscribed=$true } })
  $additionalRows = @(foreach ($p in @('Microsoft 365 Copilot','Dynamics 365','Power BI')) { [pscustomobject]@{ Product=$p; Enabled=0; Consumed=0; Subscribed=$false } })
  $newUsage = { param($Conversion) [pscustomobject]@{ Rows=@(); RecoveryDetails=@(); DowngradeReview=$null; DowngradeDetails=@(); Sources=@(); SharedSourceReady=$false; IntuneSourceReady=$false
    AdSourceForced=$false; MailboxGap=$null; AdGapActivity=$null; SharedConversion=$Conversion } }
  $newReport = { param($Usage, [string]$Domain = '') [pscustomobject]@{ SnapshotId='0123456789abcdef0123456789abcdef'; CollectedAtUtc='2026-11-10T05:00:00Z'; TenantFileCollectedAtUtc='2026-11-10T05:00:00Z'
    ExchangePlan2StandaloneStock=$null; ExchangePlan2EquivalentStock=$null; SharedMailboxMigrationEvidence=$null; SummaryRows=$summaryRows; AdditionalRows=$additionalRows
    Usage=$Usage; AssignmentIssues=$null; AdScope=$null; ScopeDomain=$Domain; UserScopes=$null } }
  $mail = New-LicensesSummaryMailContent -PublishedReport (& $newReport (& $newUsage $restored))
  Check ($mail.Body -like '*08 &nbsp; User mailboxes convertible to shared*') 'Section 08 is rendered'
  Check ($mail.Body -like '*Strong candidates confirmed today*' -and $mail.Body -like '*Changed since the prepared batch*') 'Section 08 KPIs are rendered'
  Check ($mail.Body -like '*Licenses releasable after conversion: <strong>1</strong>*') 'Releasable target suites are rendered'
  Check ($mail.Body -like '*1 of the licensed Strong candidates are already counted in section 02*') 'Recovery overlap is stated'
  Check ($mail.Body -like '*Not candidate: 1.*' -and $mail.Body -like "*$batchId*") 'Tenant mail shows Not candidate and the batch'
  Check ($mail.Body -like '*Organization retention policies and in-place holds are not assessed*' -and $mail.Body -like '*Human decision required*') 'Warnings are rendered'
  Check ($mail.Body -like '*are in section 08*') 'Recovery overview points to section 08'
  $missing = New-LicensesSummaryMailContent -PublishedReport (& $newReport (& $newUsage $null))
  Check ($missing.Body -like '*N/D - this report has no shared conversion evidence.*') 'Older snapshot renders N/D'
  $unqualified = [pscustomobject]@{ Available=$false; Reason='prepared batch: <older> than 7 days'; BatchId=''; BatchCreatedUtc=''; ReferenceUtc=''; NotCandidate=$null; Rows=@() }
  $nd = New-LicensesSummaryMailContent -PublishedReport (& $newReport (& $newUsage $unqualified))
  Check ($nd.Body -like '*N/D - prepared batch: &lt;older&gt; than 7 days*') 'Unqualified evidence renders an encoded N/D reason'

  # --- AD domain view -------------------------------------------------------------------------
  $tenantReport = & $newReport (& $newUsage $restored)
  $tenantReport.AdScope = [pscustomobject]@{ Available=$true; Provisional=$false; Reason=''; Users=@(); ByDomain=@(); ByCountry=@()
    Domains=@([pscustomobject]@{ Domain='dom-a.example'; Products=@(); E3ToF3Review=$null; MailboxGap=$null; AdGapActivity=$null; AssignmentCounts=$null }) }
  $domainView = Get-LicensesDomainPublishedReport -Report $tenantReport -Domain 'DOM-A.example'
  Check $domainView.Available "Domain view is available ($($domainView.Reason))"
  $domainConversion = $domainView.Report.Usage.SharedConversion
  Check (@($domainConversion.Rows).Count -eq 4 -and $null -eq $domainConversion.NotCandidate) 'Domain view keeps only its owners and no Not candidate count'
  $domainMail = New-LicensesSummaryMailContent -PublishedReport $domainView.Report
  Check ($domainMail.Body -like '*08 &nbsp; User mailboxes convertible to shared*' -and $domainMail.Body -notlike '*Not candidate:*') 'Domain mail renders section 08 without the tenant count'

  # --- Workbook -------------------------------------------------------------------------------
  $book = Join-Path $root 'recovery.xlsx'
  [void](New-LicensesRecoveryWorkbook -Path $book -SummaryRows $summaryRows -Usage (& $newUsage $restored) -AssignmentIssues $null -CollectedAtUtc '2026-11-10T05:00:00Z')
  $package = Open-ExcelPackage -Path $book
  try {
    $sheet = $package.Workbook.Worksheets['Shared conversion']
    Check ($null -ne $sheet) 'Workbook has the Shared conversion tab'
    Check ($sheet.Cells[1,1].Text -ceq 'Conversion state') 'Shared conversion header'
    Check ($sheet.Dimension.End.Row -eq 7) 'Five Strong candidates and one licensed Review are listed'
    Check ($sheet.Cells[2,1].Text -ceq 'Strong candidate' -and $sheet.Cells[7,1].Text -ceq 'Review') 'Strong candidates come first'
  }
  finally { Close-ExcelPackage -ExcelPackage $package -NoSave }
  $emptyBook = Join-Path $root 'empty.xlsx'
  [void](New-LicensesRecoveryWorkbook -Path $emptyBook -SummaryRows $summaryRows -Usage (& $newUsage $unqualified) -AssignmentIssues $null -CollectedAtUtc '2026-11-10T05:00:00Z')
  $package = Open-ExcelPackage -Path $emptyBook
  try { Check ($package.Workbook.Worksheets['Shared conversion'].Cells[1,1].Text -ceq 'Conversion state') 'Empty Shared conversion tab keeps its headers' }
  finally { Close-ExcelPackage -ExcelPackage $package -NoSave }
}
finally {
  if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
}
Write-Host "PASS: shared conversion section offline checks ($script:checks checks)."
