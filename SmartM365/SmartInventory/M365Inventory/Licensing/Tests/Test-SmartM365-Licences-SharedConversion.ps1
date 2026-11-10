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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBwkipGz+ONvteg
# 2frr9E0a4M/PmCmB/HPqmRessJ9zpKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEOHvteoF0WxcdL+2iKwxJPOMQsjkieSQ6BcVUZuSBKcMA0GCSqG
# SIb3DQEBAQUABIIBgAdRfyOaMSrSW12QgEkBwk26gFV4+g6+UuUGxrxn9Gkq3H0L
# jkncNzGmXdaTuK32LM5ECJPitboovirtCbPJ7wOVvR5KPKQg3zWY8L6hKbTF6vGg
# kS7Ij0nbp3M6sRK0R0P29Emv0cNUI06BeZ7XrIuFys1mG7pF40aDDBK2clykQw+F
# t+DSbbwwBmj+LOlIAmjwTq2Uc8QI8isnbqh+zTtO7UmKYxm7efD3+SOHwIBWM4hE
# QU6U+PYmsm/TiBhly6IT+JBzpftafBIYikBWf7mqSR8aC1wltulmo/fOZmXbqu80
# BpAk2Ok4YLOc5Hf5f5XPPyCHxa83ZMrCp3cmVpnrIm60Gpoz3WNkubdb/+C99QWq
# COlu08+hVjK+13zpgn+0GivzBrujWMC8sBNi6cvKAYrhvEbLn4rzSI3r7cZzaA5W
# Ka8WfnOHKVb3WZy9NgR78O1l2UltoV0xe9nIER3Y0cRAomZZnvh8/FBEUKRv/6GJ
# RZv+QoLT81DR7ue6XKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTAxNjM2
# MjJaMC8GCSqGSIb3DQEJBDEiBCBge2aaxrXkA7RhVGWJUpubk9BHRCTGEu33Z1X2
# ztznlDANBgkqhkiG9w0BAQEFAASCAgAFLE6GeeyOCDxfNxnvXCYxZIuo6JyG/wOC
# OazceO4oM68ohBsutwhz3D+1o658Sp08c7wOPoPddH+ACXguotfygxM8bmFbfbPO
# IWVxAU1Yf/w23bPq7Ml0i1WdKKbn7ONgR9VuXW96CxLl9ABAHWuc41XliOHJRT0y
# OEMShkUgE8aPjDGsJblLslGNsisdV42D67h6LoLXQq02DGXGtPntCia47TrmVHYK
# 5Vu9sP9OJxSTz7lYyeIg1d81wqW0aY5vRSmcbd4f2pGtr6NRoQEvI0poJdEzLjEQ
# deFEuPQ9EKt7fwtZoZCTy9hu8jvhM4QcQOAQ5TMPuUR4x+Y5WTohkeDuJekzL1nr
# GBz8gcd7l9n2CVjtRrgXoBZ12q6ylXBAoGcvccXMoergzG1cmolRE730T0UabVEm
# Ebx0lJzjS3SsLeO6aQ5IQI/F3KvSC6E3+e8ZmrhgRpz0GnfcrB8dM9XfpL2GlSFq
# v3dDJRKmCTUWO4BDc+0/sCnP+gFutSBW19deo/gh1sAaZ7aD3ahu5+y7hKrwnBkL
# h+kKpZbggbKbaeUtpQDy8MYuhHKLk026gNpr3M2c70EUmCc4l9KjjetmfcQffEeQ
# h8qhnJTZns+AiZOWG/iu36um0qzNhrmgNEiM3EqqYsG5C9dAuPaik2q02KtO3/iS
# VkuAoywVbQ==
# SIG # End signature block
