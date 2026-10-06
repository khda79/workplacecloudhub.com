<#
.SYNOPSIS
  Offline checks for the F1/F3/E3/E5 license summary email.
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
$names = @(
  'Get-LicensesFocusedSummaryRows', 'ConvertTo-LicensesActivityDate', 'ConvertTo-LicensesMailboxSizeGb', 'Get-LicensesCsvSource',
  'Read-LicensesIndexedSource', 'Get-LicensesFocusedUsageRows', 'Format-LicensesMetric',
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
  param([string]$From, [string]$To, [string]$Subject, [string]$BodyHtml, [string]$MailPurpose)
  $script:SentMail.Add([pscustomobject]@{From=$From;To=$To;Subject=$Subject;BodyHtml=$BodyHtml;MailPurpose=$MailPurpose}) | Out-Null
}

$script:OrgDomain = 'example.invalid'
$script:ScriptLocalConfig = @{To='reports@example.invalid';From='sender@example.invalid';EnableLicenseSummaryEmail=$true}
$script:Sampled = $false
$script:SentMail = [System.Collections.Generic.List[object]]::new()
$rows = @(
  [pscustomobject]@{TenantSkuPartNumber='M365_F1';TenantPrepaidEnabled=10;TenantConsumedUnits=8}
  [pscustomobject]@{TenantSkuPartNumber='M365_F1_COMM';TenantPrepaidEnabled=5;TenantConsumedUnits=4}
  [pscustomobject]@{TenantSkuPartNumber='SPE_F1';TenantPrepaidEnabled=20;TenantConsumedUnits=17}
  [pscustomobject]@{TenantSkuPartNumber='SPE_E3';TenantPrepaidEnabled=30;TenantConsumedUnits=28}
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

Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z'
Assert-Equal $script:SentMail.Count 1 'Sent mail count'
Assert-Equal $script:SentMail[0].MailPurpose 'Report' 'Mail purpose'
Assert-Equal $script:SentMail[0].To 'reports@example.invalid' 'Report recipient'
foreach ($product in @('F1','F3','E3','E5')) {
  if ($script:SentMail[0].BodyHtml -notmatch ('>' + $product + '</td>')) { throw "Missing product '$product' in mail body." }
}
if ($script:SentMail[0].BodyHtml -like '*OTHER_SKU*' -or $script:SentMail[0].BodyHtml -like '*99*') {
  throw 'Unrelated SKU leaked into the focused summary email.'
}

$script:Sampled = $true
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z'
Assert-Equal $script:SentMail.Count 1 'Sampled run sends no mail'
$script:Sampled = $false
$script:ScriptLocalConfig.EnableLicenseSummaryEmail = $false
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z'
Assert-Equal $script:SentMail.Count 1 'Disabled summary sends no mail'
$script:Sampled = $true
Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc '2026-10-06T10:00:00Z' -Manual
Assert-Equal $script:SentMail.Count 2 'Explicit email-only mode sends mail'
if ($script:SentMail[1].Subject -notlike '*existing CSV*' -or $script:SentMail[1].BodyHtml -notlike '*no new inventory*') {
  throw 'Email-only mode did not identify the existing CSV snapshot.'
}
$script:Sampled = $false

$badRows = @([pscustomobject]@{TenantSkuPartNumber='SPE_E5';TenantPrepaidEnabled=$null;TenantConsumedUnits=1})
$threw = $false
try { Get-LicensesFocusedSummaryRows -TenantRows $badRows | Out-Null } catch { $threw = $true }
Assert-Equal $threw $true 'Missing count is rejected'

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
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_E5'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u5';SkuPartNumber='M365_F1_COMM'}
    [pscustomobject]@{TenantKey='prod';UserId='u6';SkuPartNumber='SPE_E3'}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_Users.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';'Object Id'='u1';'User principal name'='u1@example.invalid';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u2';'User principal name'='u2@example.invalid';AccountEnabled='True';OnPremisesImmutableId='a2';LastSuccessfulSignInDateTime=$old}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u3';'User principal name'='u3@example.invalid';AccountEnabled='True';OnPremisesImmutableId='a3';LastSuccessfulSignInDateTime=$recent}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u4';'User principal name'='u4@example.invalid';AccountEnabled='True';OnPremisesImmutableId='a4';LastSuccessfulSignInDateTime=$old}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u6';'User principal name'='u6@example.invalid';AccountEnabled='True';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=$recent}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Users_Active.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a2';UserPrincipalName='u2@example.invalid';LastLogonDate=$adOld}
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a3';UserPrincipalName='u3@example.invalid';LastLogonDate=$adRecent}
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a4';UserPrincipalName='u4@example.invalid';LastLogonDate=$adOld}
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
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u6';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
  ) | Export-Csv -LiteralPath $sharedPath -NoTypeInformation

  $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $byProduct = @{}
  foreach ($item in $usage.Rows) { $byProduct[$item.Product] = $item.Counts }
  Assert-Equal $byProduct['Microsoft 365 F1'].Assigned 2 'F1 distinct users'
  Assert-Equal $byProduct['Microsoft 365 F1'].Multiple 1 'F1 variants count as multiple distinct SKUs'
  Assert-Equal $byProduct['Microsoft 365 F1'].Disabled 1 'F1 disabled'
  Assert-Equal $byProduct['Microsoft 365 F1'].DisabledUnknown 1 'F1 unmatched account unknown'
  Assert-Equal $byProduct['Microsoft 365 F1'].SharedEligible 1 'F1 disabled shared mailbox under 50 GB'
  Assert-Equal $byProduct['Microsoft 365 F3'].Assigned 2 'F3 duplicate assignment path deduplicated'
  Assert-Equal $byProduct['Microsoft 365 F3'].Multiple 1 'F3 multi-SKU user'
  Assert-Equal $byProduct['Microsoft 365 F3'].MultipleAll 2 'F3 users with any second SKU'
  Assert-Equal $byProduct['Microsoft 365 F3'].AdEntraInactive 2 'F3 AD and Entra inactive'
  Assert-Equal $byProduct['Microsoft 365 F3'].MailboxInactive 2 'F3 mailbox inactive'
  Assert-Equal $byProduct['Microsoft 365 F3'].M365Inactive 2 'F3 M365 inactive'
  Assert-Equal $byProduct['Microsoft 365 F3'].SharedUnder50 1 'F3 licensed shared mailbox under 50 GB'
  Assert-Equal $byProduct['Microsoft 365 F3'].SharedEligible 0 'F3 archived shared mailbox excluded'
  Assert-Equal $byProduct['Microsoft 365 F3'].RecoveryCandidates 1 'Archived shared mailbox excluded from recovery union'
  Assert-Equal $byProduct['Microsoft 365 E3'].LocalAppsInactive 1 'E3 local Apps inactive'
  Assert-Equal $byProduct['Microsoft 365 E3'].MailboxUnknown 1 'E3 missing mailbox row unknown'
  Assert-Equal $byProduct['Microsoft 365 E3'].M365Inactive 0 'Recent Apps activity prevents M365 false inactivity'
  Assert-Equal $byProduct['Microsoft 365 E3'].SharedEligible 1 'E3 active shared mailbox is a recovery candidate'
  Assert-Equal $byProduct['Microsoft 365 E3'].RecoveryCandidates 1 'E3 shared recovery candidate counted once'
  Assert-Equal $byProduct['Microsoft 365 E5'].Multiple 1 'E5 multi-SKU user'
  Assert-Equal $byProduct['Microsoft 365 E5'].LocalAppsInactive 1 'E5 local Apps inactive'
  Assert-Equal $byProduct['Microsoft 365 E5'].RecoveryCandidates 0 'E5 archived shared mailbox is not recoverable'
  Assert-Equal ((ConvertTo-LicensesActivityDate '01/09/2026 20:00:00').ToString('yyyy-MM-dd')) '2026-09-01' 'AD day/month parsing'
  Assert-Equal (ConvertTo-LicensesMailboxSizeGb '49,99') ([decimal]49.99) 'French mailbox size parsing'
  Assert-Equal (ConvertTo-LicensesMailboxSizeGb '50,00') ([decimal]50) '50 GB boundary parsing'

  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -Manual
  Assert-Equal $script:SentMail.Count 1 'Enriched email sent'
  if ($script:SentMail[0].BodyHtml -notlike '*Multiple assigned SKUs*' -or $script:SentMail[0].BodyHtml -notlike '*Multiple target SKUs*' -or $script:SentMail[0].BodyHtml -notlike '*Recovery candidates*' -or $script:SentMail[0].BodyHtml -notlike '*Removal candidates after archive and hold checks*') {
    throw 'Enriched KPI headers are missing from email.'
  }
  if ($script:SentMail[0].BodyHtml -notmatch 'License recovery overview' -or $script:SentMail[0].BodyHtml -notmatch 'Recovery candidate assignments' -or $script:SentMail[0].BodyHtml -notmatch 'Banner totals count license assignments') {
    throw 'KPI banner and assignment-grain note are missing.'
  }
  if ($script:SentMail[0].BodyHtml -notmatch '3 \(N/D: 1\)</div><div[^>]*>Recovery candidate assignments' -or
      $script:SentMail[0].BodyHtml -notmatch '1 \(N/D: 1\)</div><div[^>]*>Disabled account assignments' -or
      $script:SentMail[0].BodyHtml -notmatch '3 \(N/D: 1\)</div><div[^>]*>No M365 activity \(90d\)' -or
      $script:SentMail[0].BodyHtml -notmatch '>2</div><div[^>]*>Shared mailbox candidate assignments' -or
      $script:SentMail[0].BodyHtml -notmatch '>3</strong> &nbsp; Assignments held by users with multiple target SKUs') {
    throw 'KPI banner totals do not match qualified license assignments.'
  }
  if ($script:SentMail[0].BodyHtml -match '<table border="1"' -or $script:SentMail[0].BodyHtml -match '<th>License</th>') {
    throw 'Legacy unstyled table remains in email.'
  }
  if ($script:SentMail[0].BodyHtml -notmatch '<td style="[^"]*">F3</td><td style="[^"]*">20</td><td style="[^"]*">17</td><td style="[^"]*">2</td><td style="[^"]*">Subscribed</td>') {
    throw 'Enriched F3 metrics are missing from email.'
  }

  $misaligned = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -LicenseSnapshotUtc ([datetimeoffset]::UtcNow.AddDays(-3)) -AsOfUtc $today
  $f3Misaligned = @($misaligned.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $f3Misaligned.Available $false 'Misaligned license snapshots are not qualified'
  Assert-Equal $misaligned.Sources[0].Ready $false 'Misaligned license user source is not ready'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.AddDays(-3).ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -Manual
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
  Assert-Equal $boundaryE3.Counts.RecoveryCandidates 0 '50 GB boundary is excluded from recovery union'

  $u3Shared.TotalItemSizeGB = '49,99'
  $u4Shared.TotalItemSizeGB = '12,0'
  $u4Shared.ArchiveStatus = 'Active'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $exoManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-EXO-Mailboxes-Inventory.current.json.txt'
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='Exchange_EXO_Mailboxes_AllDomains.csv';Status='Failed';IsPartialInventory=$true})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $exoManifestPath
  $partialExo = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $partialExo.SharedSourceReady $false 'Partial EXO file receipt is rejected'
  $partialE3 = @($partialExo.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Assert-Equal $partialE3.Counts.RecoveryCandidates 0 'Unavailable EXO source cannot yield recovery candidates'
  Assert-Equal $partialE3.Counts.RecoveryUnknown 2 'Unavailable EXO source makes recovery unknown'
  Remove-Item -LiteralPath $exoManifestPath -Force

  $adManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'
  @{Status='Failed';IsPartialInventory=$true} | ConvertTo-Json | Set-Content -LiteralPath $adManifestPath
  $partialAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $f3PartialAd = @($partialAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $f3PartialAd.Counts.AdEntraInactive 0 'Partial AD export cannot prove inactivity'
  Assert-Equal $f3PartialAd.Counts.AdEntraUnknown 2 'Partial AD export is unknown'
  Assert-Equal @($partialAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'Failed AD receipt is rejected'
  $forcedAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -ForceAdCsvAnalysis
  $forcedAdSource = @($forcedAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0]
  $forcedF3 = @($forcedAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $forcedAdSource.Ready $true 'Fresh AD CSV is accepted with explicit override'
  Assert-Equal $forcedAdSource.Forced $true 'AD override is recorded in source metadata'
  Assert-Equal $forcedAd.AdSourceForced $true 'AD override is reported to the email builder'
  Assert-Equal $forcedF3.Counts.AdEntraInactive 2 'Forced AD CSV contributes observed inactivity'
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -Manual -ForceAdCsvAnalysis
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
  Assert-Equal $e5.Counts.LocalAppsUnknown 1 'Stale Apps report is unknown'
}
finally {
  if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
'PASS: Focused license summary email offline checks.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJMIrp/ybfWrNS
# BZwZVPwljyezT27yqQTHcBMbn2FS4aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINl+Bt0OQBnR/pWbE/SM5ZgZdiqqKt9IFeBUnQVf//cLMA0GCSqG
# SIb3DQEBAQUABIIBgIU/NX8DcZHrWUyBVaXLs+QCb9P6zImjYj0b/6MCc6ttPG3d
# sgloDFLukOgj2nXZY+UYGPgNsBye06eNWoBe2x/CW9rydQ7Ogo8dKmmZFJiXQK19
# FArICS0cs5ckfNTFYKjaS31r98o5bT6Wkc8HvCtCvTZVa/lZaVHRc1iT2pHZz3Ue
# Laq/eSZq5zlyC5d4sSJe8Sr304aP/cicRb8eBIvn5YmiC/CpjSQpWA1rRXoWuJJY
# h+Jy3iKgXfl6LYiCdnQ5q+NMoMDFiYw56TP7706CeXcOaRlTDblrdqeO7OrLDuSw
# Nddh3JPvvr5Q5BR7X+jELG9f6sIDFb0TNELIq2qJIR+Kg9Bvm1Q91beFIruhCkLX
# aZUANSlChM6PbBX15Ls16PNOUiP6V5ki9kje2W+0b65aI9RFLK3wPzb7dJLtLYjG
# TuEBc0TCKS5t0Pu0tey49kOpJ7ne/DT2MDzIHq+yQTW7/ROHpxd366AXOsf939X3
# 3AHw57SVjA1tW17V8aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDYxMjA2
# MzVaMC8GCSqGSIb3DQEJBDEiBCAv/nBe/iVK5d0QMQaGA11Ol9qjRzWIYzRG2tao
# 0ovzMjANBgkqhkiG9w0BAQEFAASCAgAqwmUyT79LCc2MsXFfEU2+sOjeDPkGR91A
# xJ56Ic8OdxoBf1jtmz0n74eL+shMellD88TOmVA2OqiFXO6kJ7c6EXYGlsWaXF4q
# uFxmATJ7GxTO2xe/mvpDlg4Fg8Vsyp6YOF5KyREtKQ0Y2joSk4vq3qnhX0cFi/lg
# LHF7abqo7GJiEmyX8bd5KuNHLFKGsEbYr2sfnmM3OAZQtyXajw101434M4cLA/M5
# FZ+OflXHIoXKYcPB52OPty8/Uv067s6fpOD2TEAI0NLQhmZB90OjXnwZNVw/5qBg
# HZ/byRHngbrlSt6f3XhFN+8mjQjdRDq0orknt2aB3BTS7vLc4vDYqainxQFIcO/r
# ah+jLHXBnCc0DgH9S3fazJL8SLPIPA7a2KM/yTC4kPltpJmhA0lVOzHmUH8bGYBK
# k6LhoESz94d3DNxl5FMw9/+w3Yib8yVRaUbCD79BC/edeNJowTK8/6XNJGCCFM4M
# QkT8ddpdMOFDFGFZ/Y0Fkv3o5Lf1kGmWxc0nNNrzTKvIiJIRSwKilL6PALccRkSz
# gmTvNZ6lVEoJFFqAj5X3tLvL73rDlqe/wmDBO4REqYiDffa7wJ/aRhFiUrKv5fwO
# RwGS03T2u9cSELcuMTqXxMHMO0BhHJNlOX18L9poq4k+GowvyOtNVYeD5gqjnZSB
# UNwc73dB3Q==
# SIG # End signature block
