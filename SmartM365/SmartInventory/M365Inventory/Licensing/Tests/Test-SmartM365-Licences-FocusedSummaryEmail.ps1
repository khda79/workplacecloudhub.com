<#
.SYNOPSIS
  Offline checks for the license overview and recovery email.
.VERSION
1.0.1
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
$inventoryText = Get-Content -LiteralPath $inventoryPath -Raw
$coreModulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
$coreModule = Import-Module -Name $coreModulePath -MinimumVersion '1.0.79' -Force -PassThru -ErrorAction Stop
if (-not (Get-Command -Name Complete-SmartM365CmdbSourceReceipt -Module $coreModule.Name -ErrorAction SilentlyContinue)) {
  throw 'SmartM365.Core must export the source receipt completion command used before the email.'
}
if ($inventoryText.IndexOf("Complete-SmartM365CmdbSourceReceipt -Status 'Completed'") -lt 0 -or
    $inventoryText.IndexOf("Complete-SmartM365CmdbSourceReceipt -Status 'Completed'") -gt
    $inventoryText.IndexOf('Send-LicensesFocusedSummaryEmail -TenantRows $tenantRows.ToArray()') -or
    $inventoryText.IndexOf('Licensing source receipt could not be completed; summary email was not sent') -lt 0) {
  throw 'The full collector must qualify its source receipt before sending the email.'
}
$names = @(
  'Get-LicensesFocusedSummaryRows', 'Get-LicensesAdditionalOverviewRows', 'New-LicensesOverviewCardHtml',
  'ConvertTo-LicensesActivityDate', 'ConvertTo-LicensesMailboxSizeGb', 'Get-LicensesCsvSource',
  'Import-LicensesSourceCsv', 'Read-LicensesIndexedSource', 'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary', 'Get-LicensesFocusedUsageRows', 'Format-LicensesMetric',
  'New-LicensesRecoveryWorkbook', 'Publish-LicensesReportCsv', 'Publish-LicensesReportSnapshot', 'Write-LicensesDailyMailState', 'Enter-LicensesDailyMailGate',
  'Send-LicensesFocusedSummaryEmail', 'Read-LicensesTenantSnapshot'
)
$definitions = @($ast.FindAll({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
}, $true))
if ($definitions.Count -ne $names.Count) { throw 'Focused summary functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }

function Assert-Equal {
  param($Actual, $Expected, [string]$Label)
  if ($Actual -ne $Expected) { throw "$Label expected '$Expected', got '$Actual'." }
}
function Get-ScriptLocalConfigValue {
  param($Config, [string]$Name, $DefaultValue)
  if ($Config.ContainsKey($Name)) { return $Config[$Name] }
  return $DefaultValue
}
function Test-SmartM365MaxItemsMode { return $script:Sampled }
function WriteLog { param([string]$Message, [string]$Level) }
function Send-SmartM365Mail {
  param([string]$From, [string]$To, [string]$Subject, [string]$BodyHtml, [string]$MailPurpose, [string[]]$Attachments, [switch]$AllowAttachments, [switch]$SuppressAttachmentLinks)
  if ($script:FailMail) { throw 'Simulated mail transport failure.' }
  if (-not $AllowAttachments -or -not $SuppressAttachmentLinks -or $Attachments.Count -ne 1 -or -not (Test-Path -LiteralPath $Attachments[0])) { throw 'Recovery workbook was not attached privately.' }
  $bookSummary = @(Import-Excel -Path $Attachments[0] -WorksheetName Summary)
  $bookCandidates = @(Import-Excel -Path $Attachments[0] -WorksheetName 'Recovery candidates' -WarningAction SilentlyContinue)
  $bookDowngrade = @(Import-Excel -Path $Attachments[0] -WorksheetName 'E3 to F3 review' -WarningAction SilentlyContinue)
  $bookDates = @{}
  $package = Open-ExcelPackage -Path $Attachments[0] -ErrorAction Stop
  try {
    $sheet = $package.Workbook.Worksheets['Recovery candidates']
    for ($row=2; $row -le $sheet.Dimension.End.Row; $row++) {
      $bookDates[[string]$sheet.Cells[$row,2].Text] = [pscustomobject]@{
        AdValue = $sheet.Cells[$row,6].Value
        AdFormat = $sheet.Cells[$row,6].Style.Numberformat.Format
        M365Value = $sheet.Cells[$row,7].Value
        M365Format = $sheet.Cells[$row,7].Style.Numberformat.Format
      }
    }
  }
  finally { $package.Dispose() }
  $script:SentMail.Add([pscustomobject]@{From=$From;To=$To;Subject=$Subject;BodyHtml=$BodyHtml;MailPurpose=$MailPurpose;BookSummary=$bookSummary;BookCandidates=$bookCandidates;BookDowngrade=$bookDowngrade;BookDates=$bookDates;AttachmentPath=$Attachments[0]}) | Out-Null
}
function New-SmartM365EmailBody {
  param([string]$Title, [string]$Category, [string]$HostName, [string]$GeneratedAt, [string]$BodyHtml, [string]$Footer)
  $script:TemplateCalls.Add([pscustomobject]@{HostName=$HostName;GeneratedAt=$GeneratedAt;Footer=$Footer}) | Out-Null
  return "<!-- SmartM365EmailTemplate:v1 --><html><body><header>$Title</header><main>$BodyHtml</main><footer>$Footer</footer></body></html>"
}

$script:OrgDomain = 'example.invalid'
$script:ScriptLocalConfig = @{To='reports@example.invalid';From='sender@example.invalid';EnableLicenseSummaryEmail=$true}
$script:Sampled = $false
$script:FailMail = $false
$script:SentMail = [System.Collections.Generic.List[object]]::new()
$script:TemplateCalls = [System.Collections.Generic.List[object]]::new()
$mailGateRoot = Join-Path $env:TEMP ('SmartM365-LicenseMailGate-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $mailGateRoot -Force | Out-Null
$script:MailStatePath = Join-Path $mailGateRoot 'M365_Licenses_SummaryEmail_SendState.json.txt'
$rows = @(
  [pscustomobject]@{TenantSkuPartNumber='M365_F1';TenantPrepaidEnabled=10;TenantConsumedUnits=8}
  [pscustomobject]@{TenantSkuPartNumber='M365_F1_COMM';TenantPrepaidEnabled=5;TenantConsumedUnits=4}
  [pscustomobject]@{TenantSkuPartNumber='SPE_F1';TenantPrepaidEnabled=20;TenantConsumedUnits=17}
  [pscustomobject]@{TenantSkuPartNumber='SPE_E3';TenantPrepaidEnabled=30;TenantConsumedUnits=28}
  [pscustomobject]@{TenantSkuPartNumber='Microsoft_365_Copilot';TenantPrepaidEnabled=10;TenantConsumedUnits=8}
  [pscustomobject]@{TenantSkuPartNumber='VIRTUAL_AGENT_USL';TenantPrepaidEnabled=1;TenantConsumedUnits=1}
  [pscustomobject]@{TenantSkuPartNumber='DYN365_FINANCE';TenantPrepaidEnabled=20;TenantConsumedUnits=12}
  [pscustomobject]@{TenantSkuPartNumber='Dyn365_Operations_Activity';TenantPrepaidEnabled=5;TenantConsumedUnits=3}
  [pscustomobject]@{TenantSkuPartNumber='Dynamics_365_for_Operations_Sandbox_Tier2_SKU';TenantPrepaidEnabled=50;TenantConsumedUnits=0}
  [pscustomobject]@{TenantSkuPartNumber='PROJECT_MADEIRA_PREVIEW_IW_SKU';TenantPrepaidEnabled=10000;TenantConsumedUnits=1}
  [pscustomobject]@{TenantSkuPartNumber='POWER_BI_PRO';TenantPrepaidEnabled=10;TenantConsumedUnits=7}
  [pscustomobject]@{TenantSkuPartNumber='PBI_PREMIUM_PER_USER';TenantPrepaidEnabled=4;TenantConsumedUnits=4}
  [pscustomobject]@{TenantSkuPartNumber='POWER_BI_STANDARD';TenantPrepaidEnabled=1000000;TenantConsumedUnits=1000}
  [pscustomobject]@{TenantSkuPartNumber='OTHER_SKU';TenantPrepaidEnabled=99;TenantConsumedUnits=99}
)
$summary = @(Get-LicensesFocusedSummaryRows -TenantRows $rows)
Assert-Equal $summary.Count 4 'Product count'
Assert-Equal $summary[0].Enabled 15 'F1 enabled sum'
Assert-Equal $summary[0].Consumed 12 'F1 consumed sum'
Assert-Equal $summary[1].Product 'Microsoft 365 F3' 'F3 mapping'
Assert-Equal $summary[1].Enabled 20 'F3 enabled count'
Assert-Equal $summary[2].Enabled 30 'E3 enabled count'
Assert-Equal $summary[3].Subscribed $false 'Missing E5 status'
Assert-Equal $summary[3].Enabled 0 'Missing E5 enabled count'
$additional = @(Get-LicensesAdditionalOverviewRows -TenantRows $rows)
Assert-Equal $additional.Count 3 'Additional product count'
Assert-Equal $additional[0].Product 'Microsoft 365 Copilot' 'Copilot family'
Assert-Equal $additional[0].Enabled 10 'Copilot excludes Studio licenses'
Assert-Equal $additional[0].Consumed 8 'Copilot used units'
Assert-Equal $additional[1].Enabled 25 'Dynamics user units exclude sandbox and preview'
Assert-Equal $additional[1].Consumed 15 'Dynamics used units'
Assert-Equal $additional[2].Enabled 14 'Power BI Pro and Premium Per User units exclude free Standard'
Assert-Equal $additional[2].Consumed 11 'Power BI used units'
if ((New-LicensesOverviewCardHtml -Row $additional[1] -Width 33 -Accent '#7c3aed') -notmatch 'font-size:26px;line-height:30px;[^>]*>60%</div>.*<strong>15</strong> used of <strong>25</strong> enabled') {
  throw 'Dynamics overview card must emphasize the utilization percentage above the license counts.'
}
if ((New-LicensesOverviewCardHtml -Row $summary[3] -Width 25 -Accent '#475569') -notmatch 'font-size:26px;line-height:30px;[^>]*>N/A</div>.*Not subscribed') {
  throw 'An unsubscribed overview card must show N/A as its primary value.'
}

Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Sent mail count'
$sentState = Get-Content -LiteralPath $script:MailStatePath -Raw | ConvertFrom-Json
Assert-Equal $sentState.Status 'Sent' 'Daily send state'
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Second send on the same Paris day is skipped'
Assert-Equal $script:SentMail[0].BookSummary.Count 4 'Workbook summary has four target suites'
Assert-Equal $script:SentMail[0].BookCandidates.Count 0 'Workbook with unavailable sources has no invented candidates'
Assert-Equal (Test-Path -LiteralPath $script:SentMail[0].AttachmentPath) $false 'Temporary attachment removed after send'
Assert-Equal $script:SentMail[0].MailPurpose 'Report' 'Mail purpose'
Assert-Equal $script:SentMail[0].To 'reports@example.invalid' 'Report recipient'
foreach ($product in @('F1','F3','E3','E5')) {
  if ($script:SentMail[0].BodyHtml -notmatch ('>' + $product + '</td>')) { throw "Missing product '$product' in mail body." }
}
if ($script:SentMail[0].BodyHtml -like '*OTHER_SKU*' -or $script:SentMail[0].BodyHtml -like '*99*') {
  throw 'Unrelated SKU leaked into the focused summary email.'
}

$script:Sampled = $true
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Sampled run sends no mail'
$script:Sampled = $false
$script:ScriptLocalConfig.EnableLicenseSummaryEmail = $false
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Disabled summary sends no mail'
$script:Sampled = $true
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
Assert-Equal $script:SentMail.Count 2 'Explicit email-only mode sends mail'
if ($script:SentMail[1].Subject -ne 'Microsoft 365 license overview and recovery' -or $script:SentMail[1].BodyHtml -notlike '*No new inventory was run*') {
  throw 'Email-only mode did not identify the existing CSV snapshot.'
}
if ($script:SentMail[1].BodyHtml.IndexOf('Source: existing published CSV') -le $script:SentMail[1].BodyHtml.IndexOf('Source freshness') -or
    $script:SentMail[1].BodyHtml.IndexOf('Source: existing published CSV') -ge $script:SentMail[1].BodyHtml.IndexOf('<footer>Host:')) {
  throw 'Existing CSV provenance is not at the bottom of the email content.'
}
$script:Sampled = $false

$dayStatePath = Join-Path $mailGateRoot 'synthetic-day-state.json.txt'
$dayStart = [datetimeoffset]::Parse('2026-10-24T22:30:00Z')
$pendingGate = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc $dayStart
Assert-Equal $pendingGate.Day '2026-10-25' 'Paris calendar day'
$pendingGate.Stream.Dispose()
$pendingRetry = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc $dayStart
Assert-Equal $pendingRetry.Skip $true 'Pending delivery prevents an automatic duplicate'
$forcedGate = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc $dayStart -Force
Assert-Equal $forcedGate.Skip $false 'Force switch bypasses the daily limit'
Write-LicensesDailyMailState -Stream $forcedGate.Stream -Bytes ([Text.Encoding]::UTF8.GetBytes('{"TenantKey":"prod","LocalDate":"2026-10-25","Status":"Sent"}'))
$forcedGate.Stream.Dispose()
$nextDayGate = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc ([datetimeoffset]::Parse('2026-10-25T23:30:00Z'))
Assert-Equal $nextDayGate.Skip $false 'New Paris day allows another send'
$nextDayGate.Stream.Dispose()

$badRows = @([pscustomobject]@{TenantSkuPartNumber='SPE_E5';TenantPrepaidEnabled=$null;TenantConsumedUnits=1})
$threw = $false
try { Get-LicensesFocusedSummaryRows -TenantRows $badRows | Out-Null } catch { $threw = $true }
Assert-Equal $threw $true 'Missing count is rejected'
$failedPreparationPath = Join-Path $mailGateRoot 'failed-preparation.json.txt'
$threw = $false
try { Send-LicensesFocusedSummaryEmail -TenantRows $badRows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $failedPreparationPath -Manual | Out-Null } catch { $threw = $true }
Assert-Equal $threw $true 'Preparation failure is reported'
Assert-Equal (Get-Item -LiteralPath $failedPreparationPath).Length 0 'Preparation failure leaves no daily send marker'
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $failedPreparationPath -Manual | Out-Null
Assert-Equal $script:SentMail.Count 3 'Preparation failure can be retried without force'

$uncertainPath = Join-Path $mailGateRoot 'uncertain-send.json.txt'
$script:FailMail = $true
$threw = $false
try { Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $uncertainPath -Manual | Out-Null } catch { $threw = $true }
$script:FailMail = $false
Assert-Equal $threw $true 'Transport failure is reported'
Assert-Equal (Get-Content -LiteralPath $uncertainPath -Raw | ConvertFrom-Json).Status 'Pending' 'Uncertain delivery blocks automatic retry'
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $uncertainPath -Manual | Out-Null
Assert-Equal $script:SentMail.Count 3 'Uncertain delivery is not automatically repeated'
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -ExpectedTenantKey 'prod' -MailStatePath $uncertainPath -Manual -ForceLicenseSummaryEmail | Out-Null
Assert-Equal $script:SentMail.Count 4 'Explicit force retries uncertain delivery'

$adRecentRow = [pscustomobject]@{Enabled='True';LastLogonDate=([datetime]::UtcNow.AddDays(-2).ToString('yyyy-MM-dd'))}
$adOldRow = [pscustomobject]@{Enabled='True';LastLogonDate=([datetime]::UtcNow.AddDays(-100).ToString('yyyy-MM-dd'))}
$adDisabledRow = [pscustomobject]@{Enabled='False';LastLogonDate=''}
$adNoDateRow = [pscustomobject]@{Enabled='True';LastLogonDate=''}
$section06Members = @{
  'a'=[pscustomobject]@{OnPremisesImmutableId='imm-a';'User principal name'='a@example.invalid'}
  'b'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='b@example.invalid'}
  'c'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='c@example.invalid'}
  'd'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='d@example.invalid'}
  'e'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='e@example.invalid'}
  'f'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='f@example.invalid'}
}
$adGap = Get-LicensesAdAccountActivitySummary `
  -MailboxGap ([pscustomobject]@{NoUserMailbox=[pscustomobject]@{Available=$true;Members=$section06Members}}) `
  -AdSource ([pscustomobject]@{Ready=$true;Forced=$false;Reason=''}) `
  -AdByImmutable @{'imm-a'=$adRecentRow} `
  -AdByUpn @{'a@example.invalid'=$adRecentRow;'b@example.invalid'=$adOldRow;'c@example.invalid'=$adDisabledRow;'f@example.invalid'=$adNoDateRow} `
  -DuplicateAdUpns @{'e@example.invalid'=$true} -Cutoff ([datetime]::UtcNow.Date.AddDays(-90))
Assert-Equal $adGap.Members 6 'Section 07 AD member population'
Assert-Equal $adGap.AdObserved 4 'Unique AD matches'
Assert-Equal $adGap.AdEnabledRecent 1 'Recent AD logon'
Assert-Equal $adGap.AdEnabledInactive 1 'Inactive AD logon'
Assert-Equal $adGap.AdEnabledNoDate 1 'Missing AD logon date'
Assert-Equal $adGap.AdDisabled 1 'Disabled AD account'
Assert-Equal $adGap.NoObservedMatch 1 'Unmatched AD account kept separate'
Assert-Equal $adGap.Ambiguous 1 'Ambiguous AD account kept separate'

$testRoot = Join-Path $env:TEMP ('SmartM365-LicenseSummary-' + [guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
  $csvPath = Join-Path $testRoot 'M365_Licenses_Tenant.csv'
  @(
    [pscustomobject]@{TenantKey='prod';TenantSkuPartNumber='SPE_F1';TenantPrepaidEnabled=20;TenantConsumedUnits=17;CollectedAtUtc='2026-10-06T10:00:00Z'}
    [pscustomobject]@{TenantKey='prod';TenantSkuPartNumber='SPE_E3';TenantPrepaidEnabled=30;TenantConsumedUnits=28;CollectedAtUtc='2026-10-06T10:00:00Z'}
  ) | Export-Csv -LiteralPath $csvPath -NoTypeInformation
  $snapshot = Read-LicensesTenantSnapshot -Path $csvPath -ExpectedTenantKey 'prod'
  Assert-Equal $snapshot.Rows.Count 2 'Existing CSV row count'
  Assert-Equal $snapshot.CollectedAtUtc '2026-10-06T10:00:00.0000000+00:00' 'Existing CSV collection time'
  $threw = $false
  try { Read-LicensesTenantSnapshot -Path $csvPath -ExpectedTenantKey 'test' | Out-Null } catch { $threw = $true }
  Assert-Equal $threw $true 'Cross-tenant CSV is rejected'
  $csvRows = @(Import-Csv -LiteralPath $csvPath)
  $csvRows[1].CollectedAtUtc = '2026-10-05T10:00:00Z'
  $csvRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation
  $threw = $false
  try { Read-LicensesTenantSnapshot -Path $csvPath -ExpectedTenantKey 'prod' | Out-Null } catch { $threw = $true }
  Assert-Equal $threw $true 'Mixed snapshot timestamps are rejected'
  $tenantCollectedAt = [datetimeoffset]::UtcNow.ToString('o')
  $rows | ForEach-Object {
    [pscustomobject]@{ TenantKey='prod'; TenantSkuPartNumber=$_.TenantSkuPartNumber;
      TenantPrepaidEnabled=$_.TenantPrepaidEnabled; TenantConsumedUnits=$_.TenantConsumedUnits;
      CollectedAtUtc=$tenantCollectedAt }
  } | Export-Csv -LiteralPath $csvPath -NoTypeInformation

  $today = [datetime]::UtcNow.Date
  $recent = $today.AddDays(-2).ToString('yyyy-MM-dd')
  $old = $today.AddDays(-100).ToString('yyyy-MM-dd')
  $adRecent = $today.AddDays(-2).ToString('dd/MM/yyyy HH:mm:ss')
  $adOld = $today.AddDays(-100).ToString('dd/MM/yyyy HH:mm:ss')
  $refresh = $today.AddDays(-1).ToString('yyyy-MM-dd')
  @(
    [pscustomobject]@{TenantKey='prod';UserId='u1';SkuPartNumber='M365_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u1';SkuPartNumber='M365_F1_COMM'}
    [pscustomobject]@{TenantKey='prod';UserId='u2';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u2';SkuPartNumber='VISIOCLIENT'}
    [pscustomobject]@{TenantKey='prod';UserId='u3';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u3';SkuPartNumber='POWER_BI_PRO'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_E5'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u5';SkuPartNumber='M365_F1_COMM'}
    [pscustomobject]@{TenantKey='prod';UserId='u6';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u6';SkuPartNumber='POWER_BI_PRO'}
    [pscustomobject]@{TenantKey='prod';UserId='u7';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u8';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u9';SkuPartNumber='POWER_BI_PRO'}
    [pscustomobject]@{TenantKey='prod';UserId='u11';SkuPartNumber='VISIOCLIENT'}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_Users.csv') -NoTypeInformation
  $licenseUsersPath = Join-Path $testRoot 'M365_Licenses_Users.csv'
  $licenseManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  $licenseHash = (Get-FileHash -LiteralPath $licenseUsersPath -Algorithm SHA256).Hash
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='M365_Licenses_Users.csv';Status='Success';IsPartialInventory=$false;SHA256=$licenseHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $localGuid = [guid]::Parse('00112233-4455-6677-8899-aabbccddeeff')
  $localImmutable = [Convert]::ToBase64String($localGuid.ToByteArray())
  @(
    [pscustomobject]@{TenantKey='prod';'Object Id'='u1';'User principal name'='u1@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u2';'User principal name'='u2@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='a2';LastSuccessfulSignInDateTime=$old}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u3';'User principal name'='u3@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='a3';LastSuccessfulSignInDateTime=$recent}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u4';'User principal name'='u4@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='a4';LastSuccessfulSignInDateTime=$old}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u6';'User principal name'='u6@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=$recent}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u7';'User principal name'='u7@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u8';'User principal name'='u8@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u9';'User principal name'='u9@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u10';'User principal name'='u10@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u11';'User principal name'='u11@example.invalid';UserType='Guest';AccountEnabled='True';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u12';'User principal name'='u12@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId=$localImmutable;LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u13';'User principal name'='u13@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u15';'User principal name'='u15@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Users_Active.csv') -NoTypeInformation
  $onPremPath = Join-Path $testRoot 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
  @([pscustomobject]@{TenantKey='prod';ObjectGUID=$localGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}) |
    Export-Csv -LiteralPath $onPremPath -NoTypeInformation
  $onPremManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-Exchange-Local-Mailboxes-Inventory.current.json.txt'
  $onPremHash = (Get-FileHash -LiteralPath $onPremPath -Algorithm SHA256).Hash
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256=$onPremHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  @(
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a2';UserPrincipalName='u2@example.invalid';Enabled='True';LastLogonDate=$adOld}
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a3';UserPrincipalName='u3@example.invalid';Enabled='True';LastLogonDate=$adRecent}
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a4';UserPrincipalName='u4@example.invalid';Enabled='True';LastLogonDate=$adOld}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'AD_Users_AllDomains.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u2@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$old;IsDeleted='False'}
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u3@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$recent;IsDeleted='False'}
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u4@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$old;IsDeleted='False'}
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u6@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$old;IsDeleted='False'}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Users_Activity.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u2@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u3@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;'Is Deleted'='False'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u4@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False'}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Mailbox_Usage.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u2@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False';'Send Count'=0;'Read Count'=0}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u3@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;'Is Deleted'='False';'Send Count'=1;'Read Count'=0}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u4@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False';'Send Count'=0;'Read Count'=0}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Email_Activity.csv') -NoTypeInformation
  $appsPath = Join-Path $testRoot 'M365_Apps_Usage_180D.csv'
  @(
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u3@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;Windows='Yes';Mac='No'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u4@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$old;Windows='No';Mac='No'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u6@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;Windows='No';Mac='No'}
  ) | Export-Csv -LiteralPath $appsPath -NoTypeInformation
  $sharedPath = Join-Path $testRoot 'Exchange_EXO_Mailboxes_AllDomains.csv'
  @(
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u1';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='12,5';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u2';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u3';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='49,99';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u4';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='12,0';ArchiveStatus='Active';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u6';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='1,4';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u7';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u8';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='12,0';ArchiveStatus='Active';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u9';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u10';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u13';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='1';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u15';RecipientTypeDetails='RoomMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='1';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
  ) | Export-Csv -LiteralPath $sharedPath -NoTypeInformation

  $oneDrivePath = Join-Path $testRoot 'M365_OneDrive_Usage.csv'
  @(
    [pscustomobject]@{TenantKey='prod';'Owner Principal Name'='u6@example.invalid';'Is Deleted'='False';'Storage Used (Byte)'='1073741824';'Report Refresh Date'=$refresh;'Report Period'='180'}
  ) | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation

  $intunePath = Join-Path $testRoot 'Intune_Devices_Inventory.csv'
  @(
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u2-a';OS='Windows';UserId='u2';'Primary user UPN'='u2@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u2-b';OS='Windows';UserId='u2';'Primary user UPN'='u2@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u3';OS='Windows';UserId='u3';'Primary user UPN'='u3@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='phone-u1';OS='iOS';UserId='u1';'Primary user UPN'='u1@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u4';OS='Windows';UserId='u4';'Primary user UPN'='u4@example.invalid'}
  ) | Export-Csv -LiteralPath $intunePath -NoTypeInformation

  $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $usage.MailboxGap.UserMailboxes.Total 2 'User mailboxes without target suite'
  Assert-Equal $usage.MailboxGap.UserMailboxes.Universe 5 'Qualified EXO UserMailbox denominator'
  Assert-Equal $usage.MailboxGap.UserMailboxes.OtherSkus 1 'User mailbox with another SKU'
  Assert-Equal $usage.MailboxGap.UserMailboxes.NoSkus 1 'User mailbox with no SKU'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraStateAvailable $true 'User mailbox Entra state is qualified'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraEnabled 1 'Enabled Entra account with unlicensed UserMailbox'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraDisabled 1 'Disabled Entra account with unlicensed UserMailbox'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraStateUnknown 0 'Unlicensed UserMailbox Entra state reconciles'
  $activePath = Join-Path $testRoot 'M365_Users_Active.csv'
  $activeRows = @(Import-Csv -LiteralPath $activePath)
  $candidateWithoutState = @($activeRows | Where-Object { $_.'Object Id' -eq 'u10' })[0]
  $candidateWithoutState.AccountEnabled = ''
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  $unknownState = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $unknownState.MailboxGap.UserMailboxes.Total 2 'Unknown Entra state retains UserMailbox total'
  Assert-Equal $unknownState.MailboxGap.UserMailboxes.EntraStateUnknown 1 'Blank Entra state is N/D'
  $candidateWithoutState.AccountEnabled = 'True'
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  Move-Item -LiteralPath $activePath -Destination ($activePath + '.missing')
  try {
    $missingEntra = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
    Assert-Equal $missingEntra.MailboxGap.UserMailboxes.Total 2 'Missing Entra source retains UserMailbox total'
    Assert-Equal $missingEntra.MailboxGap.UserMailboxes.EntraStateAvailable $false 'Missing Entra source leaves state N/D'
  }
  finally { Move-Item -LiteralPath ($activePath + '.missing') -Destination $activePath }
  Assert-Equal $usage.MailboxGap.NoUserMailbox.Total 1 'On-premises UserMailbox is excluded from mailbox-free accounts'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.Universe 7 'Qualified nontechnical Entra account denominator'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.OtherSkus 1 'Mailbox-free account with another SKU'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.NoSkus 0 'On-premises mailbox account with no SKU is excluded'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.Guests 1 'Guest is identified within the total'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.MemberEnabled 0 'Enabled member segment'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.MemberDisabled 0 'On-premises mailbox disabled member is excluded'
  $localAccount = @($activeRows | Where-Object { $_.'Object Id' -eq 'u12' })[0]
  $localAccount.OnPremisesImmutableId = ''
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  $upnFallback = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $upnFallback.MailboxGap.NoUserMailbox.Total 1 'Unique UPN fallback excludes on-premises UserMailbox'
  $localAccount.OnPremisesImmutableId = $localImmutable
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  $byProduct = @{}
  foreach ($item in $usage.Rows) { $byProduct[$item.Product] = $item.Counts }
  Assert-Equal $byProduct['Microsoft 365 F1'].Assigned 2 'F1 distinct users'
  Assert-Equal $byProduct['Microsoft 365 F1'].Multiple 0 'F1 variants count as one target suite'
  Assert-Equal $byProduct['Microsoft 365 F1'].MultipleAll 0 'Shared F1 mailbox is excluded from user overlap'
  Assert-Equal $byProduct['Microsoft 365 F1'].Disabled 0 'Disabled shared mailbox is excluded from disabled users'
  Assert-Equal $byProduct['Microsoft 365 F1'].DisabledUnknown 1 'F1 unmatched account unknown'
  Assert-Equal $byProduct['Microsoft 365 F1'].SharedEligible 1 'F1 disabled shared mailbox under 50 GB'
  Assert-Equal $byProduct['Microsoft 365 F3'].Assigned 2 'F3 duplicate assignment path deduplicated'
  Assert-Equal $byProduct['Microsoft 365 F3'].Multiple 0 'Shared F3 mailbox is excluded from target-suite overlap'
  Assert-Equal $byProduct['Microsoft 365 F3'].MultipleAll 1 'F3 non-shared user with an add-on SKU'
  Assert-Equal $byProduct['Microsoft 365 F3'].AdEntraInactive 1 'Shared mailbox is excluded from AD and Entra inactivity'
  Assert-Equal $byProduct['Microsoft 365 F3'].MailboxInactive 1 'Shared mailbox is excluded from mailbox inactivity'
  Assert-Equal $byProduct['Microsoft 365 F3'].M365Inactive 1 'Shared mailbox is excluded from M365 inactivity'
  Assert-Equal $byProduct['Microsoft 365 F3'].SharedUnder50 1 'F3 licensed shared mailbox under 50 GB'
  Assert-Equal $byProduct['Microsoft 365 F3'].SharedEligible 0 'F3 archived shared mailbox excluded'
  Assert-Equal $byProduct['Microsoft 365 F3'].RecoveryCandidates 1 'Archived shared mailbox excluded from recovery union'
  Assert-Equal $byProduct['Microsoft 365 F3'].RecoveryPrimaryPc 1 'Two Windows PCs count once for a recovery candidate'
  Assert-Equal $byProduct['Microsoft 365 F1'].RecoveryPrimaryPc 0 'An iOS device is not an Intune PC'
  Assert-Equal $byProduct['Microsoft 365 E3'].LocalAppsInactive 1 'E3 local Apps inactive'
  Assert-Equal $byProduct['Microsoft 365 E3'].MailboxUnknown 1 'E3 missing mailbox row unknown'
  Assert-Equal $byProduct['Microsoft 365 E3'].M365Inactive 0 'Recent Apps activity prevents M365 false inactivity'
  Assert-Equal $byProduct['Microsoft 365 E3'].SharedEligible 1 'E3 active shared mailbox is a recovery candidate'
  Assert-Equal $byProduct['Microsoft 365 E3'].SharedLicensed 2 'Both E3 shared mailboxes remain in the dedicated shared table'
  Assert-Equal $byProduct['Microsoft 365 E3'].Disabled 1 'Only a disabled non-shared E3 account is counted'
  Assert-Equal $byProduct['Microsoft 365 E3'].RecoveryCandidates 2 'E3 shared and disabled user candidates are distinct'
  Assert-Equal $byProduct['Microsoft 365 E3'].RecoveryPrimaryPc 1 'Shared recovery candidate can be primary on an Intune PC'
  Assert-Equal $byProduct['Microsoft 365 E3'].Multiple 0 'E3 plus Power BI is one target suite'
  Assert-Equal $byProduct['Microsoft 365 E3'].MultipleAll 1 'Only non-shared E3 plus Power BI is counted'
  Assert-Equal $byProduct['Microsoft 365 E5'].Multiple 0 'Shared E5 mailbox is excluded from target-suite overlap'
  Assert-Equal $byProduct['Microsoft 365 E5'].LocalAppsInactive 0 'Shared E5 mailbox is excluded from local Apps inactivity'
  Assert-Equal $byProduct['Microsoft 365 E5'].RecoveryCandidates 0 'E5 archived shared mailbox is not recoverable'
  Assert-Equal $byProduct['Microsoft 365 E5'].RecoveryPrimaryPc 0 'A PC on a blocked shared mailbox is not a recovery intersection'
  Assert-Equal @($usage.RecoveryDetails).Count 4 'One detail row per qualified recovery candidate and target suite'
  Assert-Equal @($usage.RecoveryDetails | Where-Object { $_.License -eq 'Microsoft 365 F1' -and $_.UserId -eq 'u1' -and $_.RecoveryReason -eq 'Eligible shared mailbox under 50 GB' }).Count 1 'F1 shared mailbox recovery detail'
  Assert-Equal @($usage.RecoveryDetails | Where-Object { $_.License -eq 'Microsoft 365 F3' -and $_.UserId -eq 'u2' -and $_.RecoveryReason -eq 'No M365 activity in 90 days' -and $_.PrimaryOnIntuneWindowsPc -eq 'Yes' }).Count 1 'F3 inactive user and Intune detail'
  Assert-Equal @($usage.RecoveryDetails | Where-Object { $_.License -eq 'Microsoft 365 E3' -and $_.UserId -eq 'u7' -and $_.RecoveryReason -eq 'Disabled account' }).Count 1 'E3 disabled user recovery detail'
  $u2Detail = @($usage.RecoveryDetails | Where-Object { $_.UserId -eq 'u2' })[0]
  Assert-Equal $u2Detail.LastAdActivityDate $old 'F3 recovery detail has last AD logon date'
  Assert-Equal $u2Detail.LastM365ActivityDate $old 'F3 recovery detail has last M365 activity date'
  $u3Detail = @($usage.RecoveryDetails | Where-Object { $_.UserId -eq 'u3' })[0]
  Assert-Equal $u3Detail.LastAdActivityDate 'N/D' 'Shared mailbox has no AD activity date'
  Assert-Equal $u3Detail.LastM365ActivityDate 'N/D' 'Shared mailbox has no M365 user activity date'
  Assert-Equal $usage.DowngradeReview.Candidates 1 'One qualified E3 to F3 review candidate'
  Assert-Equal @($usage.DowngradeDetails).Count 1 'E3 to F3 detail count matches candidate KPI'
  Assert-Equal $usage.DowngradeDetails[0].UserId 'u6' 'E3 to F3 review lists the qualified user'
  Assert-Equal $usage.DowngradeDetails[0].MailboxSizeGB ([decimal]1.4) 'E3 to F3 detail keeps mailbox size'
  Assert-Equal $usage.DowngradeDetails[0].OneDriveUsedBytes 1073741824 'E3 to F3 detail keeps exact OneDrive storage'
  Assert-Equal $usage.DowngradeReview.RecoveryExcluded 2 'Recovery candidates excluded from downgrade review'
  Assert-Equal $usage.DowngradeReview.Unknown 0 'All E3 downgrade inputs are qualified'
  Assert-Equal $usage.DowngradeReview.Excluded 1 'Archived shared mailbox is excluded from downgrade review'
  $incompleteReview = $usage | Select-Object *
  $incompleteReview.DowngradeDetails = @()
  $reviewMismatchRejected = $false
  try {
    [void](New-LicensesRecoveryWorkbook -Path (Join-Path $testRoot 'IncompleteReview.xlsx') -SummaryRows $summary -Usage $incompleteReview -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')))
  }
  catch { $reviewMismatchRejected = $_.Exception.Message -like 'E3 to F3 review detail count differs from the email KPI*' }
  Assert-Equal $reviewMismatchRejected $true 'E3 to F3 workbook cannot omit email KPI candidates'
  Assert-Equal ((ConvertTo-LicensesActivityDate '01/09/2026 20:00:00').ToString('yyyy-MM-dd')) '2026-09-01' 'AD day/month parsing'
  Assert-Equal (ConvertTo-LicensesMailboxSizeGb '49,99') ([decimal]49.99) 'French mailbox size parsing'
  Assert-Equal (ConvertTo-LicensesMailboxSizeGb '50,00') ([decimal]50) '50 GB boundary parsing'

  $script:SentMail.Clear()
  $script:TemplateCalls.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  Assert-Equal $script:SentMail.Count 1 'Enriched email sent'
  $reportSnapshotPath = Join-Path $testRoot 'M365_Licenses_ReportSnapshot.json.txt'
  $reportSnapshot = Get-Content -LiteralPath $reportSnapshotPath -Raw | ConvertFrom-Json
  Assert-Equal $reportSnapshot.TenantKey 'prod' 'Snapshot tenant'
  Assert-Equal @($reportSnapshot.Products).Count 4 'Snapshot target products'
  Assert-Equal @($reportSnapshot.OtherProducts).Count 3 'Snapshot paid overview products'
  Assert-Equal @($reportSnapshot.RecoveryCandidates).Count 4 'Snapshot recovery detail reconciles with workbook'
  Assert-Equal @($reportSnapshot.DowngradeCandidates).Count 1 'Snapshot downgrade detail reconciles with workbook'
  Assert-Equal @($reportSnapshot.Products | Where-Object Product -eq 'Microsoft 365 E3')[0].Counts.RecoveryCandidates 2 'Snapshot E3 recovery KPI'
  Assert-Equal $reportSnapshot.E3ToF3Review.Candidates 1 'Snapshot E3 downgrade KPI'
  $reportSummaryCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSummary.csv'))
  $reportCandidatesCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportCandidates.csv'))
  $reportGapsCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportGaps.csv'))
  $reportSourcesCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSources.csv'))
  Assert-Equal $reportSummaryCsv.Count 7 'Power BI summary has all email products'
  Assert-Equal @($reportCandidatesCsv | Where-Object CandidateType -eq 'Recovery').Count 4 'Power BI recovery rows match the workbook'
  Assert-Equal @($reportCandidatesCsv | Where-Object CandidateType -eq 'E3 to F3 review').Count 1 'Power BI review rows match the workbook'
  Assert-Equal $reportGapsCsv.Count 25 'Power BI gaps have all mail indicators'
  Assert-Equal @($reportSourcesCsv | Where-Object Name -eq 'M365_Licenses_Tenant.csv').Count 1 'Power BI source freshness includes tenant source'
  foreach ($csvRow in @($reportSummaryCsv) + @($reportCandidatesCsv) + @($reportGapsCsv) + @($reportSourcesCsv)) {
    Assert-Equal $csvRow.SnapshotId $reportSnapshot.SnapshotId 'Power BI CSV snapshot identity'
    Assert-Equal $csvRow.TenantKey 'prod' 'Power BI CSV tenant identity'
  }
  if ($script:SentMail[0].BodyHtml -notmatch [regex]::Escape("Snapshot: $($reportSnapshot.SnapshotId)")) { throw 'Email does not identify its Power BI snapshot.' }
  Assert-Equal $script:SentMail[0].BookCandidates.Count 4 'Workbook row count matches recovery totals'
  Assert-Equal @($script:SentMail[0].BookCandidates | Where-Object License -eq 'Microsoft 365 E3').Count 2 'Workbook E3 rows match the email KPI'
  Assert-Equal @($script:SentMail[0].BookSummary | Where-Object License -eq 'Microsoft 365 E3')[0].RecoveryCandidates 2 'Workbook Summary E3 total'
  Assert-Equal @($script:SentMail[0].BookSummary | Where-Object License -eq 'Microsoft 365 E3')[0].E3toF3ReviewCandidates 1 'Workbook Summary E3 review total'
  Assert-Equal $script:SentMail[0].BookDowngrade.Count 1 'Workbook E3 to F3 rows match the email KPI'
  Assert-Equal $script:SentMail[0].BookDowngrade[0].UserId 'u6' 'Workbook E3 to F3 review identity'
  Assert-Equal $script:SentMail[0].BookDowngrade[0].MailboxSizeGB ([decimal]1.4) 'Workbook E3 to F3 review mailbox size'
  Assert-Equal $script:SentMail[0].BookDowngrade[0].OneDriveUsedBytes 1073741824 'Workbook E3 to F3 review exact OneDrive bytes'
  $bookU2 = $script:SentMail[0].BookDates['u2']
  Assert-Equal ([datetime]::FromOADate([double]$bookU2.AdValue)).ToString('yyyy-MM-dd') $old 'Workbook has AD activity date'
  Assert-Equal ([datetime]::FromOADate([double]$bookU2.M365Value)).ToString('yyyy-MM-dd') $old 'Workbook has M365 activity date'
  if ($script:SentMail[0].BookDates['u2'].AdValue -is [string] -or
      $script:SentMail[0].BookDates['u2'].M365Value -is [string] -or
      $null -eq $script:SentMail[0].BookDates['u2'].AdValue -or
      $null -eq $script:SentMail[0].BookDates['u2'].M365Value) {
    throw 'Workbook activity dates are stored as text or missing.'
  }
  Assert-Equal $script:SentMail[0].BookDates['u2'].AdFormat 'yyyy-mm-dd' 'AD date uses sortable Excel date format'
  Assert-Equal $script:SentMail[0].BookDates['u2'].M365Format 'yyyy-mm-dd' 'M365 date uses sortable Excel date format'
  Assert-Equal $script:SentMail[0].BookDates['u3'].AdValue $null 'Shared mailbox AD date cell is blank'
  Assert-Equal $script:SentMail[0].BookDates['u3'].M365Value $null 'Shared mailbox M365 date cell is blank'
  Assert-Equal (Test-Path -LiteralPath $script:SentMail[0].AttachmentPath) $false 'Temporary recovery workbook removed after send'
  if ($script:SentMail[0].BodyHtml -notlike '*Multiple assigned SKUs*' -or $script:SentMail[0].BodyHtml -notlike '*Multiple target suites*' -or $script:SentMail[0].BodyHtml -notlike '*Recovery candidates*' -or $script:SentMail[0].BodyHtml -notlike '*Removal candidates after archive and hold checks*') {
    throw 'Enriched KPI headers are missing from email.'
  }
  if ($script:SentMail[0].BodyHtml -notlike '*Candidates primary on Intune PC*' -or
      $script:SentMail[0].BodyHtml -notlike '*Intune_Devices_Inventory.csv*') {
    throw 'Intune PC recovery indicator or source freshness is missing.'
  }
  if ($script:SentMail[0].BodyHtml -notlike '*AD accounts and activity among section 07 Members*' -or
      $script:SentMail[0].BodyHtml -notlike '*No AD match means no unique match in the available AD export*') {
    throw 'AD account activity qualification is missing from section 07.'
  }
  if ($script:SentMail[0].BodyHtml -notlike '*03 &nbsp; E3 to F3 downgrade review*' -or
      $script:SentMail[0].BodyHtml -notlike '*E3 to F3 review tab listing every qualified review candidate*' -or
      $script:SentMail[0].BodyHtml -notlike '*OneDrive storage below 2 GB*' -or
      $script:SentMail[0].BodyHtml -notlike '*Activity dates are sortable Excel dates*') {
    throw 'Downgrade review or workbook activity dates are not described.'
  }
  $oneDriveRows = @(Import-Csv -LiteralPath $oneDrivePath)
  $oneDriveRows[0].'Storage Used (Byte)' = '2147483648'
  $oneDriveRows | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation
  $oneDriveBoundary = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $oneDriveBoundary.DowngradeReview.Candidates 0 'OneDrive at 2 GB is excluded'
  Assert-Equal @($oneDriveBoundary.DowngradeDetails).Count 0 'Boundary user is absent from E3 to F3 detail'
  Assert-Equal $oneDriveBoundary.DowngradeReview.Excluded 2 'OneDrive at 2 GB is counted as excluded'
  $emptyReviewPath = Join-Path $testRoot 'NoE3ToF3Review.xlsx'
  [void](New-LicensesRecoveryWorkbook -Path $emptyReviewPath -SummaryRows $summary -Usage $oneDriveBoundary -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')))
  $emptyReviewRows = @(Import-Excel -Path $emptyReviewPath -WorksheetName 'E3 to F3 review' -WarningAction SilentlyContinue)
  Assert-Equal $emptyReviewRows.Count 0 'Empty E3 to F3 review sheet has no candidate rows'
  $emptyPackage = Open-ExcelPackage -Path $emptyReviewPath
  try { Assert-Equal $emptyPackage.Workbook.Worksheets['E3 to F3 review'].Cells[1,1].Text 'UserId' 'Empty E3 to F3 review sheet retains headers' }
  finally { $emptyPackage.Dispose() }
  Remove-Item -LiteralPath $emptyReviewPath -Force
  $oneDriveRows[0].'Storage Used (Byte)' = ''
  $oneDriveRows | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation
  $oneDriveMissing = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $oneDriveMissing.DowngradeReview.Unknown 1 'Missing OneDrive size is N/D'
  $oneDriveRows[0].'Storage Used (Byte)' = '1073741824'
  $oneDriveRows | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation
  if ($script:SentMail[0].BodyHtml -notmatch '<td style="[^"]*">E3</td><td style="[^"]*">2</td><td style="[^"]*">1</td>' -or
      $script:SentMail[0].BodyHtml -notlike '*Disabled users and activity/overlap indicators exclude identified shared mailboxes*') {
    throw 'The email does not separate disabled user accounts from shared mailboxes.'
  }
  if ($script:SentMail[0].BodyHtml -notmatch 'License overview and recovery' -or $script:SentMail[0].BodyHtml -notmatch 'F3/F1 adds both suite counts') {
    throw 'KPI banner and assignment-grain note are missing.'
  }
  $overviewBody = $script:SentMail[0].BodyHtml
  if ($overviewBody.IndexOf('>License overview</h2>') -lt 0 -or
      $overviewBody -notmatch '(?s)>License overview</h2>\s*<h2[^>]*>License recovery overview</h2>' -or
      $overviewBody.IndexOf('>License recovery overview</h2>') -ge $overviewBody.IndexOf('MICROSOFT 365 SUITES') -or
      $overviewBody.IndexOf('Recovery candidates F3/F1') -ge $overviewBody.IndexOf('MICROSOFT 365 SUITES') -or
      $overviewBody.IndexOf('MICROSOFT 365 SUITES') -ge $overviewBody.IndexOf('COPILOT, DYNAMICS 365 AND POWER BI') -or
      $overviewBody.IndexOf('COPILOT, DYNAMICS 365 AND POWER BI') -ge $overviewBody.IndexOf('01 &nbsp; License capacity') -or
      $overviewBody -notlike '*MICROSOFT 365 SUITES*' -or
      $overviewBody -notlike '*COPILOT, DYNAMICS 365 AND POWER BI*') {
    throw 'Recovery KPIs are not immediately below License overview and above the license bands.'
  }
  foreach ($label in @('F1','F3','E3','E5','Copilot','Dynamics 365','Power BI')) {
    if ($overviewBody -notmatch ('>' + [regex]::Escape($label) + '</div>')) { throw "Missing overview card '$label'." }
  }
  if ($script:SentMail[0].BodyHtml -notmatch '>0</div><div[^>]*>Recovery candidates E5' -or
      $script:SentMail[0].BodyHtml -notmatch '>2</div><div[^>]*>Recovery candidates E3' -or
      $script:SentMail[0].BodyHtml -notmatch '>2 \(N/D: 1\)</div><div[^>]*>Recovery candidates F3/F1' -or
      $script:SentMail[0].BodyHtml -notmatch '>1</div><div[^>]*>E3 to F3 downgrade review') {
    throw 'Recovery and E3 downgrade KPI cards do not match the detailed sections.'
  }
  $body = $script:SentMail[0].BodyHtml
  if ($body.IndexOf('01 &nbsp; License capacity') -lt 0 -or
      $body.IndexOf('02 &nbsp; Recovery by license') -le $body.IndexOf('01 &nbsp; License capacity') -or
      $body.IndexOf('03 &nbsp; E3 to F3 downgrade review') -le $body.IndexOf('02 &nbsp; Recovery by license') -or
      $body.IndexOf('04 &nbsp; Activity and overlap') -le $body.IndexOf('03 &nbsp; E3 to F3 downgrade review') -or
      $body -notlike '*Consumed (used %)*' -or $body -notlike '*Available (free %)*') {
    throw 'License capacity order or percentages are missing.'
  }
  if ($body.IndexOf('06 &nbsp; User mailboxes without licence F1/F3/E3/E5') -le $body.IndexOf('05 &nbsp; Licensed shared mailboxes') -or
      $body.IndexOf('07 &nbsp; Without User mailboxes + without licence F1/F3/E3/E5') -le $body.IndexOf('06 &nbsp; User mailboxes without licence F1/F3/E3/E5') -or
      $body -notmatch '>2</strong> of 5 qualified EXO UserMailbox.*?<strong>40%</strong>' -or
      $body -notmatch 'Entra enabled</th>.*?Entra disabled</th>.*?Entra state N/D</th>.*?<td[^>]*>1</td><td[^>]*>1</td><td[^>]*>0</td>' -or
       $body -notmatch '>1</strong> of 7 qualified Entra accounts.*?<strong>14.3%</strong>' -or
       $body -notmatch 'Member enabled</th>.*?<td[^>]*>0</td><td[^>]*>0</td><td[^>]*>1</td><td[^>]*>1</td><td[^>]*>0</td><td[^>]*>0</td>') {
    throw 'Mailbox and account gap sections are missing or have incorrect counts.'
  }
  Assert-Equal $script:TemplateCalls.Count 1 'Focused email uses one shared template'
  Assert-Equal $script:TemplateCalls[0].HostName '' 'Host metadata removed from template header'
  Assert-Equal $script:TemplateCalls[0].GeneratedAt '' 'Generated metadata removed from template header'
  if ($script:TemplateCalls[0].Footer -notmatch '^Host: .+ \| Generated: \d{4}-\d{2}-\d{2}') { throw 'Host and generated time are missing from template footer.' }
  if ($body.IndexOf('<footer>Host:') -le $body.IndexOf('Source freshness')) { throw 'Host and generated time are not at the bottom of the email.' }
  if ($script:SentMail[0].BodyHtml -match '<table border="1"' -or $script:SentMail[0].BodyHtml -match '<th>License</th>') {
    throw 'Legacy unstyled table remains in email.'
  }
  if ($script:SentMail[0].BodyHtml -notmatch '<td style="[^"]*">F3</td><td style="[^"]*">20</td><td style="[^"]*">17 \(85%\)</td><td style="[^"]*">3 \(15%\)</td><td style="[^"]*">2</td><td style="[^"]*">Subscribed</td>') {
    throw 'F3 license capacity and percentages are missing from email.'
  }

  $mailboxRowsWithId = @(Import-Csv -LiteralPath $sharedPath)
  @($mailboxRowsWithId | Where-Object ExternalDirectoryObjectId -eq 'u10')[0].ExternalDirectoryObjectId = ''
  $mailboxRowsWithId | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $unjoined = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $unjoined.MailboxGap.UserMailboxes.Unknown 1 'User mailbox without Entra ID is unqualified'
  Assert-Equal $unjoined.MailboxGap.NoUserMailbox.Available $false 'Missing mailbox identity prevents proving account has no mailbox'
  @($mailboxRowsWithId | Where-Object RecipientTypeDetails -eq 'UserMailbox' | Where-Object { -not $_.ExternalDirectoryObjectId })[0].ExternalDirectoryObjectId = 'u10'
  $mailboxRowsWithId | Export-Csv -LiteralPath $sharedPath -NoTypeInformation

  $intuneManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-Devices-Inventory.current.json.txt'
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='Intune_Devices_Inventory.csv';Status='Failed';IsPartialInventory=$true})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $intuneManifestPath
  $partialIntune = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $partialIntune.IntuneSourceReady $false 'Partial Intune file receipt is rejected'
  Assert-Equal @($partialIntune.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0].Counts.RecoveryCandidates 1 'Intune source does not change recovery count'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  if ($script:SentMail[0].BodyHtml -notmatch '<td style="[^"]*">F3</td><td style="[^"]*">1</td><td style="[^"]*">0</td><td style="[^"]*">1</td><td style="[^"]*">N/D</td>') {
    throw 'Unqualified Intune PC source is not marked N/D in recovery table.'
  }
  Remove-Item -LiteralPath $intuneManifestPath -Force

  $misaligned = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -LicenseSnapshotUtc ([datetimeoffset]::UtcNow.AddDays(-3)) -AsOfUtc $today
  $f3Misaligned = @($misaligned.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $f3Misaligned.Available $false 'Misaligned license snapshots are not qualified'
  Assert-Equal $misaligned.Sources[0].Ready $false 'Misaligned license user source is not ready'
  Assert-Equal $misaligned.MailboxGap $null 'Misaligned license source makes both new sections N/D'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.AddDays(-3).ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  $misalignedSnapshot = Get-Content -LiteralPath $reportSnapshotPath -Raw | ConvertFrom-Json
  $misalignedF3 = @($misalignedSnapshot.Products | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $misalignedF3.UsageAvailable $false 'Snapshot rejects misaligned license assignments'
  Assert-Equal $misalignedF3.Counts $null 'Snapshot leaves unqualified metrics null, not zero'
  $misalignedSummaryCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSummary.csv'))
  $misalignedF3Csv = @($misalignedSummaryCsv | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $misalignedF3Csv.EvidenceStatus 'N/D' 'CSV marks unqualified usage N/D'
  Assert-Equal $misalignedF3Csv.RecoveryCandidates '' 'CSV leaves unqualified recovery count blank'
  Assert-Equal $script:SentMail.Count 1 'Unqualified usage still sends a stock summary'
  if ($script:SentMail[0].BodyHtml -notlike '*N/D*') { throw 'Unqualified usage is not marked N/D in email.' }

  $sharedRows = @(Import-Csv -LiteralPath $sharedPath)
  $u4Shared = @($sharedRows | Where-Object ExternalDirectoryObjectId -eq 'u4')[0]
  $u4Shared.ArchiveStatus = 'None'
  $u4Shared.LitigationHoldEnabled = 'True'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $held = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $heldE5 = @($held.Rows | Where-Object Product -eq 'Microsoft 365 E5')[0]
  Assert-Equal $heldE5.Counts.SharedEligible 0 'Litigation hold blocks shared recovery'
  Assert-Equal $heldE5.Counts.RecoveryCandidates 0 'Litigation hold blocks recovery union'

  $u4Shared.LitigationHoldEnabled = 'False'
  $u4Shared.TotalItemSizeGB = ''
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $unknownSize = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $unknownE5 = @($unknownSize.Rows | Where-Object Product -eq 'Microsoft 365 E5')[0]
  Assert-Equal $unknownE5.Counts.SharedUnknown 1 'Missing shared mailbox size is unknown'
  Assert-Equal $unknownE5.Counts.RecoveryUnknown 1 'Missing size cannot become a recovery candidate'

  $u3Shared = @($sharedRows | Where-Object ExternalDirectoryObjectId -eq 'u3')[0]
  $u3Shared.TotalItemSizeGB = '50,00'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $boundary = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $boundaryE3 = @($boundary.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Assert-Equal $boundaryE3.Counts.SharedEligible 0 '50 GB shared mailbox is excluded'
  Assert-Equal $boundaryE3.Counts.RecoveryCandidates 1 '50 GB shared mailbox is excluded while disabled user remains'

  $u3Shared.TotalItemSizeGB = '49,99'
  $u4Shared.TotalItemSizeGB = '12,0'
  $u4Shared.ArchiveStatus = 'Active'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $exoManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-EXO-Mailboxes-Inventory.current.json.txt'
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='Exchange_EXO_Mailboxes_AllDomains.csv';Status='Failed';IsPartialInventory=$true})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $exoManifestPath
  $partialExo = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $partialExo.SharedSourceReady $false 'Partial EXO file receipt is rejected'
  Assert-Equal $partialExo.MailboxGap.UserMailboxes.Available $false 'Partial EXO source makes user mailbox gap N/D'
  Assert-Equal $partialExo.MailboxGap.NoUserMailbox.Available $false 'Partial EXO source makes no-mailbox gap N/D'
  $partialE3 = @($partialExo.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Assert-Equal $partialE3.Counts.RecoveryCandidates 0 'Unavailable EXO source cannot yield recovery candidates'
  Assert-Equal $partialE3.Counts.RecoveryUnknown 4 'Unavailable EXO source makes all E3 mailbox types unknown'
  Assert-Equal $partialE3.Counts.MultipleAllUnknown 4 'Unqualified mailbox type is excluded from overlap counts'
  Remove-Item -LiteralPath $exoManifestPath -Force

  $adManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'
  @{Status='Failed';IsPartialInventory=$true} | ConvertTo-Json | Set-Content -LiteralPath $adManifestPath
  $partialAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $f3PartialAd = @($partialAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $f3PartialAd.Counts.AdEntraInactive 0 'Partial AD export cannot prove inactivity'
  Assert-Equal $f3PartialAd.Counts.AdEntraUnknown 1 'Only non-shared F3 account has unknown AD activity'
  Assert-Equal @($partialAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'Failed AD receipt is rejected'
  $forcedAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -ForceAdCsvAnalysis
  $forcedAdSource = @($forcedAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0]
  $forcedF3 = @($forcedAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $forcedAdSource.Ready $true 'Fresh AD CSV is accepted with explicit override'
  Assert-Equal $forcedAdSource.Forced $true 'AD override is recorded in source metadata'
  Assert-Equal $forcedAd.AdSourceForced $true 'AD override is reported to the email builder'
  Assert-Equal $forcedF3.Counts.AdEntraInactive 1 'Forced AD CSV contributes only non-shared user inactivity'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail -ForceAdCsvAnalysis | Out-Null
  if ($script:SentMail[0].BodyHtml -notlike '*Provisional AD/Entra indicator*' -or
      $script:SentMail[0].BodyHtml -notlike '*PROVISIONAL (*FORCED AD CSV*') {
    throw 'Forced AD usage is not visibly qualified in the email.'
  }
  $adPath = Join-Path $testRoot 'AD_Users_AllDomains.csv'
  $adLastWriteUtc = (Get-Item -LiteralPath $adPath).LastWriteTimeUtc
  [System.IO.File]::SetLastWriteTimeUtc($adPath, $today.AddDays(-20))
  $staleAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -ForceAdCsvAnalysis
  Assert-Equal @($staleAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'AD override does not bypass CSV freshness'
  Assert-Equal @($staleAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0].Counts.AdEntraInactive 0 'Stale AD CSV does not prove inactivity'
  [System.IO.File]::SetLastWriteTimeUtc($adPath, $adLastWriteUtc)
  $adRows = @(Import-Csv -LiteralPath $adPath)
  $adRows[0].TenantKey = 'another-tenant'
  $adRows | Export-Csv -LiteralPath $adPath -NoTypeInformation
  $crossTenantAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -ForceAdCsvAnalysis
  Assert-Equal @($crossTenantAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'AD override does not bypass tenant isolation'
  Assert-Equal $crossTenantAd.AdSourceForced $false 'Rejected AD CSV is not presented as forced evidence'
  $adRows[0].TenantKey = 'prod'
  $adRows | Export-Csv -LiteralPath $adPath -NoTypeInformation
  Remove-Item -LiteralPath $adManifestPath -Force

  $staleApps = @(Import-Csv -LiteralPath $appsPath)
  foreach ($row in $staleApps) { $row.'Report Refresh Date' = $today.AddDays(-20).ToString('yyyy-MM-dd') }
  $staleApps | Export-Csv -LiteralPath $appsPath -NoTypeInformation
  $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $e5 = @($usage.Rows | Where-Object Product -eq 'Microsoft 365 E5')[0]
  Assert-Equal $e5.Counts.LocalAppsInactive 0 'Stale Apps report is not treated as no use'
  Assert-Equal $e5.Counts.LocalAppsUnknown 0 'Shared E5 mailbox remains outside local Apps usage'

  $conflictGuid = [guid]::Parse('10213243-5465-7687-98a9-bacbdcedfe0f')
  $activeRows = @(Import-Csv -LiteralPath $activePath)
  @($activeRows | Where-Object { $_.'Object Id' -eq 'u11' })[0].OnPremisesImmutableId = [Convert]::ToBase64String($conflictGuid.ToByteArray())
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';ObjectGUID=$localGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}
    [pscustomobject]@{TenantKey='prod';ObjectGUID=$conflictGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='RoomMailbox';NativeIdentityStatus='Observed'}
    [pscustomobject]@{TenantKey='prod';ObjectGUID=([guid]::NewGuid().ToString());UserPrincipalName='u8@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}
  ) | Export-Csv -LiteralPath $onPremPath -NoTypeInformation
  $conflictHash = (Get-FileHash -LiteralPath $onPremPath -Algorithm SHA256).Hash
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256=$conflictHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $ambiguousOnPrem = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $ambiguousOnPrem.MailboxGap.NoUserMailbox.Available $true 'Unrelated accounts remain qualified when local identities conflict'
  Assert-Equal $ambiguousOnPrem.MailboxGap.NoUserMailbox.Unknown 2 'Identity conflicts qualify affected unlicensed accounts as unknown'
  Assert-Equal $ambiguousOnPrem.MailboxGap.NoUserMailbox.Total 0 'Conflicting local mailbox identities are not counted as mailbox-free'
  $activeRows = @(Import-Csv -LiteralPath $activePath)
  @($activeRows | Where-Object { $_.'Object Id' -eq 'u11' })[0].OnPremisesImmutableId = ''
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  @([pscustomobject]@{TenantKey='prod';ObjectGUID=$localGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}) |
    Export-Csv -LiteralPath $onPremPath -NoTypeInformation

  @{Status='Failed';IsPartialInventory=$true;ConsumerScopeQualified=$false;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $partialOnPrem = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $partialOnPrem.MailboxGap.UserMailboxes.Available $true 'EXO UserMailbox section remains available'
  Assert-Equal $partialOnPrem.MailboxGap.NoUserMailbox.Available $false 'Partial on-premises receipt makes section 07 N/D'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  if ($script:SentMail[0].BodyHtml -notlike '*Section 07 N/D*') { throw 'Unqualified on-premises source is not visible in the email.' }
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256='BADHASH'})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $changedOnPrem = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $changedOnPrem.MailboxGap.NoUserMailbox.Available $false 'Changed on-premises CSV hash makes section 07 N/D'
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256=$onPremHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath

  @{Status='Completed';IsPartialInventory=$false;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $oldReceipt = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $oldReceipt.Rows[0].Available $false 'Old licensing receipt needs an explicit bypass'
  $bypassedReceipt = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -BypassLicenseUsersReceipt
  Assert-Equal $bypassedReceipt.Rows[0].Available $true 'Bypass accepts fresh license-users CSV'
  Assert-Equal $bypassedReceipt.LicenseSourceForced $true 'Bypass is recorded as provisional'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail -BypassLicenseUsersReceipt | Out-Null
  if ($script:SentMail[0].BodyHtml -notlike '*Provisional license assignment indicators*' -or
      $script:SentMail[0].BodyHtml -notlike '*PROVISIONAL (*license-users CSV absent*') {
    throw 'The temporary licensing receipt bypass is not visible in the email.'
  }
  $licenseLastWriteUtc = (Get-Item -LiteralPath $licenseUsersPath).LastWriteTimeUtc
  @{Status='Running';IsPartialInventory=$true;StartedAtUtc=$licenseLastWriteUtc.AddMinutes(5).ToString('o');Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $runningWithoutBypass = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $runningWithoutBypass.Rows[0].Available $false 'Running receipt needs an explicit bypass'
  $runningWithBypass = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -BypassLicenseUsersReceipt
  Assert-Equal $runningWithBypass.Rows[0].Available $true "Bypass accepts the prior CSV while a new collection is running ($($runningWithBypass.Sources[0].Reason))"
  Assert-Equal $runningWithBypass.LicenseSourceForced $true 'Running receipt bypass is provisional'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail -BypassLicenseUsersReceipt | Out-Null
  if ($script:SentMail[0].BodyHtml -notlike '*Provisional license assignment indicators*' -or
      $script:SentMail[0].BodyHtml -notlike '*previous license-users CSV while a new collection is running*') {
    throw 'The running licensing receipt bypass is not visible in the email.'
  }
  @{Status='Running';IsPartialInventory=$true;StartedAtUtc=$licenseLastWriteUtc.AddMinutes(-5).ToString('o');Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $inProgressCsv = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -BypassLicenseUsersReceipt
  Assert-Equal $inProgressCsv.Rows[0].Available $false 'Bypass rejects a CSV updated after the new collection started'
  @{Status='Completed';IsPartialInventory=$true;StartedAtUtc=$licenseLastWriteUtc.AddMinutes(5).ToString('o');Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $partialReceipt = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -BypassLicenseUsersReceipt
  Assert-Equal $partialReceipt.Rows[0].Available $false 'Bypass does not accept a completed partial licensing receipt'
  @{Status='Failed';IsPartialInventory=$true;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $failedReceipt = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -BypassLicenseUsersReceipt
  Assert-Equal $failedReceipt.Rows[0].Available $false 'Bypass does not accept failed licensing receipt'
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='M365_Licenses_Users.csv';Status='Success';IsPartialInventory=$false;SHA256='BADHASH'})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $wrongHash = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -BypassLicenseUsersReceipt
  Assert-Equal $wrongHash.Rows[0].Available $false 'Bypass does not accept changed licensing CSV'
  @{Status='Completed';IsPartialInventory=$false;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  [System.IO.File]::SetLastWriteTimeUtc($licenseUsersPath, $today.AddDays(-20))
  $staleLicense = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -BypassLicenseUsersReceipt
  Assert-Equal $staleLicense.Rows[0].Available $false 'Bypass does not accept stale licensing CSV'
  [System.IO.File]::SetLastWriteTimeUtc($licenseUsersPath, $licenseLastWriteUtc)
}
finally {
  if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
  if (Test-Path -LiteralPath $mailGateRoot) { Remove-Item -LiteralPath $mailGateRoot -Recurse -Force }
}
'PASS: Focused license summary email offline checks.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCyeYJbJCAEFBVe
# FNVQpJlBJtHiPKX9JpyQD5eAj56rx6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJRgJLjur5LCuTGSO8eEKXbq05U/YqeKb29G053VTspbMA0GCSqG
# SIb3DQEBAQUABIIBgJACrGxEjxX+JsZFQXIe8qF7QFoIL7c+FrPKdchuGXBZROMP
# Njc4rDzK/SuRAa29Q672/Zc6rlbbMFGT+QKTt9C6HIs/R3CZHBS45QJrsIbLoYQK
# 5TLw1EafrEfeNhfXC9KkrNdeSyyWGMDexMXX0jnoOA4KDpKEnIZjrrpzcTjy0v4y
# J5EQB6BX74PsXYr1Nrw0Y6LH4P8dFA/uL58ZY27N0w+B/pgpcsUptJ7kko3pIWta
# jB2ZlLVUOZIQXxuY3fXwTg8mvhRqdjmdfuYx0RY7K+io/OKxySyDtNoPWV/YuS/X
# XFrji85eJ+vw3rq1pEKdLTvXnr3IJ7lMbSsKL9BLWM5f4+WHT5Icirh5fQZR+k0X
# oXvMPJ2nLR9fOrvFmT8FWFJr5DVcicNhq+snepsBFJVrR8wGKt/qJZdJTql1jehK
# /6CNOG+eoLaTEZgZpRPvR9j496igdCMxmlTKvy/tVwniBIvraW3JlEIEgE0YFukP
# NSDOfndjYL3CZj8db6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcyMDA4
# MjJaMC8GCSqGSIb3DQEJBDEiBCCkIk8xNDdiCZ+F54ZlyPJ98y71WqV4ufqfxay4
# M5XwCzANBgkqhkiG9w0BAQEFAASCAgBPBhyE1NHlUTJCffrzAuVJRLn4uVfrMhnA
# wGUi70pmP2hWgrzjfl1KHW3I2Opb5ZCx8l4KROU+xJiq1TK5R/llguX5q9hvjNaU
# 9RUEqE+YQvUHmZ+Bm5H5j2ZcpDCf3Ixqvl4iKQWgrLck3sNhnxtW2t2o0RqHdyGq
# n10bBBTpAYNpE7D9oe+dxlPvGVrA15QpkSeKEZhTE/Nbpls415EyziLc5OKXiKkf
# 8Az6A8s14NrvJcjGB6kNmRc1qAFd88i/qdooKij3dmkxqqZzIs7UgRMp+YtB41B/
# UALz2fZcT5kn/UxHs+QgJh044WukNJwNoeSoIgRpu+Vpwc7Irb7QmPzQNtDBwYRC
# AGndme2WvYf1D/ao/rDZpgqPt1pIZjhw3L8llWkkwWh6sGFfVm2GdEjIl5NQ3Bdj
# SS7jpew1SMQX0OA+9pLOX9tkplfc1SKLLxa8+GnhzXAjqUVQEclnY8B3UpqZmSsA
# fosFcF2h8c73x6hVY/tg1mAA7JAAWXHzE2ZoyeTl+0O17N3waAK+lkZCXeC15Eay
# ndaRfDxQF4BGww+5fhWT0Wiw0/Fh4pqucO6T3m4MtTMxn5mqNFGnpfEXayAMyoMh
# BXJIQzTKTMOSrv8LMOSwfL0smyXdlzL9P9ckSvYdaeZOPwgPSWF5GM5Tj7eNyewo
# J+zjD445Dw==
# SIG # End signature block
