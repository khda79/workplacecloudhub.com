<#
.SYNOPSIS
  Offline checks for the license views by AD domain and country and the per-domain license summary emails.
.VERSION
1.4.3
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
Import-Module -Name $coreModulePath -MinimumVersion '1.0.79' -Force -ErrorAction Stop | Out-Null
# Workbook checks may load the installed module, but never repair/install it.
function Initialize-SmartM365ImportExcel {
  SmartM365.Core\Initialize-SmartM365ImportExcel -CheckOnly | Out-Null
}
$names = @(
  'Get-LicensesSuiteProduct', 'Get-LicensesOptionalProperty', 'Get-LicensesFocusedSummaryRows', 'Get-LicensesAdditionalOverviewRows', 'Get-LicensesOtherProduct', 'New-LicensesOverviewCardHtml',
  'ConvertTo-LicensesActivityDate', 'ConvertTo-LicensesMailboxSizeGb', 'Get-LicensesQualifiedPartialAdCoverage', 'Get-LicensesCsvSource', 'Get-LicensesPublicationPath',
  'Import-LicensesSourceCsv', 'Import-LicensesSourceCsvCached', 'Read-LicensesIndexedSource', 'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary',
  'Get-LicensesUserAdScopes', 'Get-LicensesUserScope', 'Get-LicensesScopeBreakdown', 'Get-LicensesFocusedUsageRows', 'Format-LicensesMetric',
  'Get-LicensesAlertSummary', 'New-LicensesRecoveryWorkbook', 'Publish-LicensesReportCsv', 'Get-LicensesStandalonePlan2Stock', 'Get-LicensesEquivalentPlan2Stock', 'Get-LicensesSharedMigrationEvidence', 'Publish-LicensesReportSnapshot', 'Write-LicensesDailyMailState', 'Enter-LicensesDailyMailGate',
  'Read-LicensesTenantSnapshot', 'ConvertFrom-LicensesJsonElement', 'ConvertTo-LicensesCountTable', 'Read-LicensesPublishedReport', 'Get-LicensesDomainPublishedReport',
  'New-LicensesScopeTablesHtml', 'New-LicensesSummaryMailContent', 'ConvertTo-LicensesMailSchedule', 'Get-LicensesMailSchedule', 'Get-LicensesDomainMailSettings', 'Get-LicensesDomainMailStatePath', 'Send-LicensesSummaryMailScope',
  'Send-LicensesFocusedSummaryEmail'
)
$definitions = @($ast.FindAll({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
}, $true))
if ($definitions.Count -ne $names.Count) { throw 'AD domain scope functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }

$script:Checks = 0
function Check {
  param([bool]$Condition, [string]$Label)
  if (-not $Condition) { throw "FAIL: $Label" }
  $script:Checks++
}
function Get-ScriptLocalConfigValue {
  param($Config, [string]$Name, $DefaultValue)
  if ($Config.ContainsKey($Name)) { return $Config[$Name] }
  return $DefaultValue
}
function Test-SmartM365MaxItemsMode { return $false }
$script:Warnings = [System.Collections.Generic.List[string]]::new()
function WriteLog { param([string]$Message, [string]$Level) if ($Level -eq 'WARNING') { $script:Warnings.Add($Message) } }
# Records the pause between the tenant email and the domain emails instead of waiting.
$script:Sleeps = [System.Collections.Generic.List[int]]::new()
function Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
function New-SmartM365EmailBody {
  param([string]$Title, [string]$Category, [string]$HostName, [string]$GeneratedAt, [string]$BodyHtml, [string]$Footer)
  return "<html><body><header>$Title</header><main>$BodyHtml</main><footer>$Footer</footer></body></html>"
}
$script:SentMail = [System.Collections.Generic.List[object]]::new()
function Send-SmartM365Mail {
  param([string]$From, [string]$To, [string]$Cc, [string]$Subject, [string]$BodyHtml, [string]$MailPurpose, [string[]]$Attachments, [switch]$AllowAttachments, [switch]$SuppressAttachmentLinks)
  $candidates = @()
  $downgrade = @()
  $book = $null
  if ($Attachments) {
    if (-not $AllowAttachments -or -not $SuppressAttachmentLinks -or $Attachments.Count -ne 1) { throw 'Recovery workbook was not attached privately.' }
    $candidates = @(Import-Excel -Path $Attachments[0] -WorksheetName 'Recovery candidates' -WarningAction SilentlyContinue)
    $downgrade = @(Import-Excel -Path $Attachments[0] -WorksheetName 'E3 to F3 review' -WarningAction SilentlyContinue)
    $package = Open-ExcelPackage -Path $Attachments[0] -ErrorAction Stop
    try {
      $sheet = $package.Workbook.Worksheets['Recovery candidates']
      $overview = $package.Workbook.Worksheets['Overview']
      $book = [pscustomobject]@{
        Name = [IO.Path]::GetFileName($Attachments[0]); Folder = [IO.Path]::GetDirectoryName($Attachments[0])
        Sheets = @($package.Workbook.Worksheets | ForEach-Object Name)
        Headers = @(for ($c = 1; $c -le $sheet.Dimension.End.Column; $c++) { [string]$sheet.Cells[1, $c].Text })
        Rules = @($sheet.ConditionalFormatting).Count
        Links = @($overview.Cells | Where-Object { $_.Hyperlink }).Count
        Title = [string]$overview.Cells['B3'].Text; Scope = [string]$overview.Cells['B4'].Text
        HeaderFill = [string]$sheet.Cells[1, 1].Style.Fill.BackgroundColor.Rgb
        RepeatRows = [string]$sheet.PrinterSettings.RepeatRows.Address; Landscape = $sheet.PrinterSettings.Orientation -eq [OfficeOpenXml.eOrientation]::Landscape
      }
    }
    finally { $package.Dispose() }
  }
  $script:SentMail.Add([pscustomobject]@{ From=$From; To=$To; Cc=$Cc; CcBound=$PSBoundParameters.ContainsKey('Cc'); Subject=$Subject; BodyHtml=$BodyHtml
    MailPurpose=$MailPurpose; HasAttachment=[bool]$Attachments; Candidates=$candidates; Downgrade=$downgrade; Book=$book }) | Out-Null
}
function Update-TestLicensesReceipt {
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string[]]$CsvPaths)
  # Records the report CSVs as optional files of the fixture licensing receipt, as Complete-SmartM365CmdbSourceReceipt does.
  $path = Join-Path $Folder 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  $receipt = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable
  $fileNames = @($CsvPaths | ForEach-Object { [IO.Path]::GetFileName($_) })
  $files = @(@($receipt.Files) | Where-Object { $_ -and [string]$_.File -notin $fileNames })
  foreach ($csv in $CsvPaths) {
    $files += @{File=[IO.Path]::GetFileName($csv);Status='Success';IsPartialInventory=$false;Required=$false;SHA256=(Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash}
  }
  $receipt.Files = $files
  $receipt.TenantKey = 'prod'
  $receipt.Producer = 'SmartM365-Licences-Inventory.ps1'
  $receipt | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $path
}
function Publish-TestLicenseReport {
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][object[]]$TenantRows, [Parameter(Mandatory)][string]$CollectedAtUtc, [switch]$WithoutUsage)
  $summaryRows = @(Get-LicensesFocusedSummaryRows -TenantRows $TenantRows)
  $additionalRows = @(Get-LicensesAdditionalOverviewRows -TenantRows $TenantRows)
  $usage = if ($WithoutUsage) { $null } else { Get-LicensesFocusedUsageRows -CsvFolderPath $Folder -ExpectedTenantKey 'prod' -AsOfUtc ([datetime]::UtcNow) }
  $snapshot = Publish-LicensesReportSnapshot -Folder $Folder -TenantKey 'prod' -CollectedAtUtc $CollectedAtUtc -SummaryRows $summaryRows -AdditionalRows $additionalRows -Usage $usage -AssignmentIssues $null
  Update-TestLicensesReceipt -Folder $Folder -CsvPaths $snapshot.CsvPaths
  return [pscustomobject]@{ Snapshot=$snapshot; Usage=$usage }
}
function Write-Fixture {
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][object[]]$Rows)
  $path = Join-Path $Folder $Name
  $Rows | Export-Csv -LiteralPath $path -NoTypeInformation
  return $path
}

$script:OrgDomain = 'example.invalid'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-LicenseDomainScopes-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
  $alpha = 'alpha.example.invalid'
  $beta = 'beta.example.invalid'
  $collectedAt = [datetimeoffset]::UtcNow.ToString('o')
  $tenantRows = @(
    [pscustomobject]@{TenantSkuPartNumber='M365_F1';TenantPrepaidEnabled=10;TenantConsumedUnits=2}
    [pscustomobject]@{TenantSkuPartNumber='SPE_F1';TenantPrepaidEnabled=10;TenantConsumedUnits=2}
    [pscustomobject]@{TenantSkuPartNumber='SPE_E3';TenantPrepaidEnabled=10;TenantConsumedUnits=3}
    [pscustomobject]@{TenantSkuPartNumber='SPE_E5';TenantPrepaidEnabled=10;TenantConsumedUnits=2}
    [pscustomobject]@{TenantSkuPartNumber='Microsoft_365_Copilot';TenantPrepaidEnabled=5;TenantConsumedUnits=1}
  )
  $tenantRows | ForEach-Object {
    [pscustomobject]@{ TenantKey='prod'; TenantSkuPartNumber=$_.TenantSkuPartNumber; TenantPrepaidEnabled=$_.TenantPrepaidEnabled; TenantConsumedUnits=$_.TenantConsumedUnits; CollectedAtUtc=$collectedAt }
  } | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_Tenant.csv') -NoTypeInformation

  $licenseUsers = Write-Fixture $testRoot 'M365_Licenses_Users.csv' @(
    [pscustomobject]@{TenantKey='prod';UserId='a1';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='a2';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='a2';SkuPartNumber='Microsoft_365_Copilot'}
    [pscustomobject]@{TenantKey='prod';UserId='a3';SkuPartNumber='SPE_E5'}
    [pscustomobject]@{TenantKey='prod';UserId='b1';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='b2';SkuPartNumber='M365_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='c1';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='n1';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='m1';SkuPartNumber='SPE_E5'}
  )
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='M365_Licenses_Users.csv';Status='Success';IsPartialInventory=$false;SHA256=(Get-FileHash -LiteralPath $licenseUsers -Algorithm SHA256).Hash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $testRoot 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt')
  $recent = [datetime]::UtcNow.AddDays(-2).ToString('o')
  $active = { param($Id, $Immutable, $Enabled) [pscustomobject]@{TenantKey='prod';'Object Id'=$Id;'User principal name'="$Id@example.invalid";'Display name'="User $Id";UserType='Member';AccountEnabled=$Enabled;OnPremisesImmutableId=$Immutable;LastSuccessfulSignInDateTime=$recent} }
  Write-Fixture $testRoot 'M365_Users_Active.csv' @(
    (& $active 'a1' 'ia1' 'False'), (& $active 'a2' 'ia2' 'True'), (& $active 'a3' 'ia3' 'True'), (& $active 'b1' 'ib1' 'False'), (& $active 'b2' 'ib2' 'True'),
    (& $active 'c1' '' 'False'), (& $active 'n1' 'in1' 'True'), (& $active 'm1' 'im1' 'True'), (& $active 'g1' 'ig1' 'True')
  ) | Out-Null
  $adUser = { param($Immutable, $Upn, $Domain, $Country, $CountryName) [pscustomobject]@{TenantKey='prod';DomainName=$Domain;ImmutableId_AD=$Immutable;UserPrincipalName=$Upn;Enabled='True';LastLogonDate='';Country=$Country;CountryName=$CountryName} }
  Write-Fixture $testRoot 'AD_Users_AllDomains.csv' @(
    (& $adUser 'ia1' 'a1@example.invalid' 'ALPHA.example.invalid' 'FR' 'France'),
    (& $adUser 'ia2' 'a2@example.invalid' $alpha 'FR' 'France'),
    (& $adUser 'ia3' 'a3@example.invalid' $alpha '' 'Unknown'),
    (& $adUser 'ib1' 'b1@example.invalid' $beta 'ES' 'Spain'),
    (& $adUser 'ib2' 'b2@example.invalid' $beta 'DE' ''),
    (& $adUser 'ix1' 'm1@example.invalid' $alpha 'FR' 'France'),
    (& $adUser 'ix2' 'm1@example.invalid' $beta 'FR' 'France'),
    (& $adUser 'ig1' 'g1@example.invalid' 'gamma.example.invalid' 'IT' 'Italy')
  ) | Out-Null
  $mailbox = { param($Id, $Kind) [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId=$Id;RecipientTypeDetails=$Kind;NativeIdentityStatus='Observed';TotalItemSizeGB='1';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'} }
  Write-Fixture $testRoot 'Exchange_EXO_Mailboxes_AllDomains.csv' @(
    (& $mailbox 'a1' 'UserMailbox'), (& $mailbox 'a2' 'UserMailbox'), (& $mailbox 'a3' 'UserMailbox'), (& $mailbox 'b1' 'UserMailbox'), (& $mailbox 'b2' 'UserMailbox'),
    (& $mailbox 'c1' 'UserMailbox'), (& $mailbox 'n1' 'UserMailbox'), (& $mailbox 'm1' 'UserMailbox'), (& $mailbox 'g1' 'UserMailbox')
  ) | Out-Null

  # Attribution: ImmutableId first, then a unique UPN; cloud-only, unmatched and ambiguous accounts keep explicit labels.
  $activeSource = Get-LicensesCsvSource -Folder $testRoot -FileName 'M365_Users_Active.csv' -Columns @('TenantKey','Object Id') -AsOfUtc ([datetime]::UtcNow)
  $adSource = Get-LicensesCsvSource -Folder $testRoot -FileName 'AD_Users_AllDomains.csv' -Columns @('TenantKey','ImmutableId_AD') -AsOfUtc ([datetime]::UtcNow)
  $scopes = Get-LicensesUserAdScopes -ActiveSource $activeSource -AdSource $adSource -ExpectedTenantKey 'prod'
  Check $scopes.Available 'AD scopes are available'
  Check ($scopes.Users['a1'].Domain -ceq $alpha -and $scopes.Users['a1'].Country -eq 'France') 'Domain is normalized to lower case and country uses CountryName'
  Check ($scopes.Users['a3'].Status -eq 'Matched' -and $scopes.Users['a3'].Country -eq 'Country not set') 'Unknown CountryName without code is not a country'
  Check ($scopes.Users['b2'].Country -eq 'DE') 'Country code is used when CountryName is empty'
  Check ($scopes.Users['c1'].Status -eq 'CloudOnly' -and $scopes.Users['c1'].Domain -eq 'Cloud-only (no AD account)') 'Cloud-only account'
  Check ($scopes.Users['n1'].Status -eq 'NoMatch' -and $scopes.Users['n1'].Domain -eq 'No AD match') 'Synchronized account without AD match'
  Check ($scopes.Users['m1'].Status -eq 'Ambiguous' -and $scopes.Users['m1'].Domain -eq 'Ambiguous AD match') 'Duplicate AD UPN without ImmutableId match is ambiguous, never guessed'
  $upnOnly = $adSource | Select-Object *
  $upnOnlyPath = Write-Fixture $testRoot 'AD_UpnOnly.csv' @((& $adUser '' 'a2@example.invalid' $beta 'FR' 'France'))
  $upnOnly.Path = $upnOnlyPath; $upnOnly.SHA256 = (Get-FileHash -LiteralPath $upnOnlyPath).Hash
  Check ((Get-LicensesUserAdScopes -ActiveSource $activeSource -AdSource $upnOnly -ExpectedTenantKey 'prod').Users['a2'].Domain -ceq $beta) 'Unique UPN fallback attributes the domain'
  $noDomain = $adSource | Select-Object *
  $noDomainPath = Write-Fixture $testRoot 'AD_NoDomain.csv' @([pscustomobject]@{TenantKey='prod';ImmutableId_AD='ia1';UserPrincipalName='a1@example.invalid'})
  $noDomain.Path = $noDomainPath; $noDomain.SHA256 = (Get-FileHash -LiteralPath $noDomainPath).Hash
  $missingColumn = Get-LicensesUserAdScopes -ActiveSource $activeSource -AdSource $noDomain -ExpectedTenantKey 'prod'
  Check (-not $missingColumn.Available -and $missingColumn.Reason -like '*DomainName*') 'An AD export without DomainName leaves the views N/D'

  $published = Publish-TestLicenseReport -Folder $testRoot -TenantRows $tenantRows -CollectedAtUtc $collectedAt
  $usage = $published.Usage
  Check $usage.AdScope.Available ("Usage carries the AD scope views: " + $usage.AdScope.Reason + " | " + ($script:Warnings -join " | "))
  $domainRow = @{}
  foreach ($row in $usage.AdScope.ByDomain) { $domainRow[$row.Scope] = $row }
  Check ($domainRow[$alpha].E3 -eq 1 -and $domainRow[$alpha].F3 -eq 1 -and $domainRow[$alpha].E5 -eq 1 -and $domainRow[$alpha].Copilot -eq 1 -and $domainRow[$alpha].LicensedUsers -eq 3 -and $domainRow[$alpha].RecoveryCandidates -eq 1) 'Alpha domain breakdown'
  Check ($domainRow[$beta].E3 -eq 1 -and $domainRow[$beta].F1 -eq 1 -and $domainRow[$beta].LicensedUsers -eq 2 -and $domainRow[$beta].RecoveryCandidates -eq 1) 'Beta domain breakdown'
  Check ($domainRow['Cloud-only (no AD account)'].E3 -eq 1 -and $domainRow['Cloud-only (no AD account)'].ScopeKind -eq 'Unattributed') 'Cloud-only row is unattributed'
  Check ($domainRow['No AD match'].F3 -eq 1 -and $domainRow['Ambiguous AD match'].E5 -eq 1) 'Unmatched and ambiguous users stay visible'
  Check (-not $domainRow.ContainsKey('gamma.example.invalid')) 'A domain without licensed users has no breakdown row'
  $licensedTotal = 0; foreach ($row in $usage.AdScope.ByDomain) { $licensedTotal += $row.LicensedUsers }
  Check ($licensedTotal -eq 8) 'Domain rows add up to the licensed tenant users'
  Check ($usage.AdScope.ByDomain[0].ScopeKind -eq 'Attributed' -and $usage.AdScope.ByDomain[0].Scope -ceq $alpha) 'Attributed domains come first, largest first'
  $countryRow = @{}
  foreach ($row in $usage.AdScope.ByCountry) { $countryRow[$row.Scope] = $row }
  Check ($countryRow['France'].LicensedUsers -eq 2 -and $countryRow['Spain'].LicensedUsers -eq 1 -and $countryRow['DE'].LicensedUsers -eq 1) 'Country breakdown'
  Check ($countryRow['Country not set'].ScopeKind -eq 'Unattributed' -and $countryRow['Country not set'].LicensedUsers -eq 1) 'Missing AD country is an unattributed row'
  Check (@($usage.AdScope.Domains | ForEach-Object Domain) -join ',' -ceq "$alpha,$beta,gamma.example.invalid") 'Each attributed AD domain is measured'
  $alphaScope = @($usage.AdScope.Domains | Where-Object Domain -ceq $alpha)[0]
  $alphaE3 = @($alphaScope.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Check ($alphaE3.Counts.Assigned -eq 1 -and $alphaE3.Counts.RecoveryCandidates -eq 1 -and $alphaE3.Counts.Disabled -eq 1) 'Alpha E3 is measured on alpha users only'
  $gammaScope = @($usage.AdScope.Domains | Where-Object Domain -ceq 'gamma.example.invalid')[0]
  Check ($gammaScope.MailboxGap.UserMailboxes.Available -and $gammaScope.MailboxGap.UserMailboxes.Total -eq 1 -and $gammaScope.MailboxGap.UserMailboxes.Universe -eq 1) 'Section 06 is scoped to the domain mailboxes'
  Check ($usage.MailboxGap.UserMailboxes.Total -eq 1 -and $usage.MailboxGap.UserMailboxes.Universe -eq 9) 'Tenant section 06 is unchanged by the domain views'

  $snapshot = Get-Content -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
  Check ($snapshot.AdScope.Available -and @($snapshot.AdScope.Domains).Count -eq 3) 'Snapshot publishes the AD scope views'
  Check (@($snapshot.AdScope.Users | ForEach-Object UserId | Sort-Object) -join ',' -eq 'a1,b1,c1') 'Snapshot lists only users referenced by the report'
  $byDomainCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportByDomain.csv'))
  $byCountryCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportByCountry.csv'))
  Check ($byDomainCsv.Count -eq @($usage.AdScope.ByDomain).Count -and @($byDomainCsv | Where-Object { $_.SnapshotId -ne $snapshot.SnapshotId -or $_.Dimension -ne 'Domain' }).Count -eq 0) 'Domain CSV carries the snapshot identity'
  Check ($byCountryCsv.Count -eq @($usage.AdScope.ByCountry).Count -and ($byCountryCsv | Measure-Object -Property LicensedUsers -Sum).Sum -eq 8) 'Country CSV adds up to the licensed users'
  $candidateHeader = (Get-Content -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportCandidates.csv') -TotalCount 1)
  Check ($candidateHeader -notmatch 'AdDomain|AdCountry') 'Existing report CSV schemas are unchanged'

  # Mails: tenant, two configured domains, one configured domain without data; a disabled entry is ignored.
  $statePath = Join-Path $testRoot 'M365_Licenses_SummaryEmail_SendState.json.txt'
  $script:ScriptLocalConfig = @{ To='reports@example.invalid'; From='sender@example.invalid'; EnableLicenseSummaryEmail=$true
    LicenseSummaryDomainMails=@(
      @{ Domain='Alpha.Example.Invalid'; To='alpha@example.invalid'; Cc='alpha-cc@example.invalid' },
      [pscustomobject]@{ Domain=$beta; To='beta@example.invalid' },
      @{ Domain='missing.example.invalid'; To='missing@example.invalid' },
      @{ Domain='off.example.invalid'; To='off@example.invalid'; Enabled=$false }
    ) }
  Check ((Get-LicensesDomainMailStatePath -Path $statePath -Domain $alpha) -like '*M365_Licenses_SummaryEmail_SendState.alpha.example.invalid.json.txt') 'Domain send state sits next to the tenant state'
  $sent = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath
  Check ($sent -and $script:SentMail.Count -eq 4) 'Tenant mail plus three enabled domain mails'
  Check (($script:Sleeps -join ',') -eq '30' -and $script:SentMail[0].To -eq 'reports@example.invalid') 'Tenant mail first, then a 30 second pause before the domain mails'
  $tenantMail = $script:SentMail[0]
  Check ($tenantMail.To -eq 'reports@example.invalid' -and -not $tenantMail.CcBound -and $tenantMail.Subject -eq 'Microsoft 365 license overview and recovery') 'Tenant mail keeps its recipients and the global Cc'
  $tablesAt = $tenantMail.BodyHtml.IndexOf('Licenses by AD domain and country')
  Check ($tablesAt -gt 0 -and $tablesAt -lt $tenantMail.BodyHtml.IndexOf('01 &nbsp; License capacity')) 'Tenant mail shows the domain and country tables before section 01'
  Check ($tenantMail.BodyHtml -match [regex]::Escape($alpha) -and $tenantMail.BodyHtml -match 'Licenses by country' -and $tenantMail.BodyHtml -match 'Cloud-only \(no AD account\)') 'Tenant tables list domains, countries and unattributed users'
  Check ($tenantMail.BodyHtml -match '>Total</td>(<td[^>]*>\d+</td>){5}<td[^>]*>8</td><td[^>]*>3</td>') 'Total row equals the tenant licensed users and recovery candidates'
  Check ($tenantMail.Candidates.Count -eq 3) 'Tenant workbook keeps every recovery candidate'
  $alphaMail = @($script:SentMail | Where-Object To -eq 'alpha@example.invalid')[0]
  Check ($alphaMail.CcBound -and $alphaMail.Cc -eq 'alpha-cc@example.invalid' -and $alphaMail.Subject -ceq "Microsoft 365 license overview and recovery - $alpha") 'Alpha mail uses its own recipients and subject'
  Check ($alphaMail.BodyHtml -match "Scope: AD domain $([regex]::Escape($alpha))" -and $alphaMail.BodyHtml -match 'Assigned users in domain' -and $alphaMail.BodyHtml -match 'Tenant enabled units') 'Alpha mail states its scope and tenant-wide units'
  Check ($alphaMail.BodyHtml -notmatch 'Licenses by AD domain and country') 'Domain tables are in the tenant mail only'
  Check ($alphaMail.Candidates.Count -eq 1 -and $alphaMail.Candidates[0].'User ID' -eq 'a1') 'Alpha workbook lists only alpha recovery candidates'
  Check ($alphaMail.BodyHtml -notmatch 'b1@example\.invalid|c1@example\.invalid' -and @($alphaMail.Candidates | Where-Object { $_.'User principal name' -notlike 'a*' }).Count -eq 0) 'No user of another domain reaches the alpha mail'
  $betaMail = @($script:SentMail | Where-Object To -eq 'beta@example.invalid')[0]
  Check ($betaMail.CcBound -and $betaMail.Cc -eq '' -and $betaMail.Candidates.Count -eq 1 -and $betaMail.Candidates[0].'User ID' -eq 'b1') 'Beta mail and workbook are limited to beta'
  $missingMail = @($script:SentMail | Where-Object To -eq 'missing@example.invalid')[0]
  Check ($missingMail.Subject -like '* - missing.example.invalid - N/D' -and -not $missingMail.HasAttachment -and $missingMail.BodyHtml -match 'no licensed or mailbox user') 'A configured domain without data receives an explicit N/D notice'
  Check (@($script:SentMail | Where-Object To -eq 'off@example.invalid').Count -eq 0) 'Disabled domain entry sends nothing'
  Check (Test-Path -LiteralPath (Get-LicensesDomainMailStatePath -Path $statePath -Domain $alpha)) 'Alpha send state is written'

  # Workbook in the email style: Overview first, readable headers, scope columns, colors, links, print setup, readable name.
  $today = Get-Date -Format 'yyyyMMdd'
  $tenantBook = $tenantMail.Book
  Check ($tenantBook.Name -eq "M365_Licenses_Recovery_prod_$today.xlsx" -and $alphaMail.Book.Name -eq "M365_Licenses_Recovery_$($alpha)_$today.xlsx") 'Attachment names carry the tenant or the domain and the date'
  Check (-not (Test-Path -LiteralPath $tenantBook.Folder)) 'Temporary attachment folder is removed after sending'
  Check (($tenantBook.Sheets -join '|') -eq 'Overview|Recovery candidates|E3 to F3 review|Awaiting licenses|Licensing errors') 'Overview is the first of five sheets'
  Check ($tenantBook.Title -eq 'License recovery workbook' -and $tenantBook.Scope -like 'example.invalid*Snapshot*' -and $alphaMail.Book.Scope -like "AD domain $alpha*") 'Overview header names the scope and the snapshot'
  Check ($tenantBook.Headers -contains 'AD domain' -and $tenantBook.Headers -contains 'Country' -and $tenantBook.Headers -contains 'Last AD logon') 'Tenant workbook has readable headers with AD domain and country'
  Check ($alphaMail.Book.Headers -notcontains 'AD domain' -and $alphaMail.Book.Headers -contains 'Country') 'Domain workbook keeps the country but not its own domain'
  Check (@($tenantMail.Candidates | Where-Object { $_.'User ID' -eq 'c1' })[0].'AD domain' -eq 'Cloud-only (no AD account)' -and @($tenantMail.Candidates | Where-Object { $_.'User ID' -eq 'b1' })[0].Country -eq 'Spain') 'Candidate rows carry their AD domain and country'
  Check ($tenantBook.Rules -ge 6 -and $tenantBook.Links -eq 4) 'Recovery reasons, Intune PC, old dates and banding are colored; Overview links every sheet'
  Check ($tenantBook.HeaderFill -match 'EAF1F8$' -and $tenantBook.Landscape -and $tenantBook.RepeatRows -match '1:\$?1') 'Headers use the email color and pages print landscape with repeated headers'

  $sentAgain = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath
  Check (-not $sentAgain -and $script:SentMail.Count -eq 4) 'Each mail is sent once per Paris week by default'
  Check ($script:Sleeps.Count -eq 1) 'No pause when the tenant mail was not sent'
  Remove-Item -LiteralPath (Get-LicensesDomainMailStatePath -Path $statePath -Domain $beta)
  [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath)
  Check ($script:SentMail.Count -eq 5 -and $script:SentMail[4].To -eq 'beta@example.invalid') 'Domain send states are independent of each other and of the tenant'
  Remove-Item -LiteralPath $statePath
  [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath)
  Check ($script:SentMail.Count -eq 6 -and $script:SentMail[5].To -eq 'reports@example.invalid' -and $script:Sleeps.Count -eq 1) 'No pause when the tenant mail is sent but no domain mail is due'
  $only = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath -Manual -ForceLicenseSummaryEmail -Domain 'ALPHA.example.invalid'
  Check ($only -and $script:SentMail.Count -eq 7 -and $script:SentMail[6].To -eq 'alpha@example.invalid') 'Email-only domain option resends one domain'
  $tenantOnly = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath -Manual -ForceLicenseSummaryEmail -TenantOnly
  Check ($tenantOnly -and $script:SentMail.Count -eq 8 -and $script:SentMail[7].To -eq 'reports@example.invalid') 'Email-only mode without a domain sends the tenant mail only'
  $allDomains = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath -Manual -ForceLicenseSummaryEmail -Domain 'All' -TenantOnly:$false
  Check ($script:Sleeps.Count -eq 1) 'No pause for a single domain, the tenant mail only or All domains'
  $script:ScriptLocalConfig.LicenseSummaryDomainMailDelaySeconds = 601
  $threw = $false
  try { Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath -Manual -Domain $beta | Out-Null } catch { $threw = $_.Exception.Message -like '*between 0 and 600*' }
  Check $threw 'The pause is bounded to 600 seconds'
  $script:ScriptLocalConfig.Remove('LicenseSummaryDomainMailDelaySeconds')
  Check ($allDomains -and $script:SentMail.Count -eq 11 -and @($script:SentMail[8..10] | Where-Object To -eq 'reports@example.invalid').Count -eq 0) 'Domain All sends every enabled domain without the tenant mail'
  $threw = $false
  try { Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath -Manual -Domain 'off.example.invalid' | Out-Null } catch { $threw = $_.Exception.Message -like '*not an enabled entry*' }
  Check $threw 'A disabled or unknown domain cannot be resent'

  # Configuration validation: manual mode fails, an automatic run keeps the tenant mail.
  foreach ($bad in @(
      @{ Label='duplicate'; Items=@(@{Domain=$alpha;To='x@example.invalid'}, @{Domain='ALPHA.example.invalid';To='y@example.invalid'}); Message='*more than once*' },
      @{ Label='invalid'; Items=@(@{Domain='not a domain';To='x@example.invalid'}); Message='*invalid AD domain*' },
      @{ Label='missing To'; Items=@(@{Domain=$alpha}); Message='*needs a To recipient*' })) {
    $failed = $false
    try { [void](Get-LicensesDomainMailSettings -Config @{ LicenseSummaryDomainMails=$bad.Items }) } catch { $failed = $_.Exception.Message -like $bad.Message }
    Check $failed "Domain mail settings reject a $($bad.Label) entry"
  }
  $script:ScriptLocalConfig.LicenseSummaryDomainMails = @(@{ Domain=$alpha })
  Remove-Item -LiteralPath $statePath
  $script:Warnings.Clear()
  $tenantMail2 = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath
  Check ($tenantMail2 -and $script:SentMail.Count -eq 12 -and @($script:Warnings | Where-Object { $_ -like 'License summary emails by AD domain skipped*' }).Count -eq 1) 'Invalid domain settings never block the tenant mail'

  # Country table: all countries up to 25, otherwise the 20 largest plus one aggregated row.
  $manyCountries = [pscustomobject]@{ Available=$true; Provisional=$false; Reason=''; ByDomain=@()
    ByCountry=@(1..30 | ForEach-Object { [pscustomobject]@{ Scope=('Country {0:00}' -f $_); ScopeKind='Attributed'; F1=0; F3=0; E3=1; E5=0; Copilot=0; LicensedUsers=(100 - $_); RecoveryCandidates=0 } }) }
  $html = New-LicensesScopeTablesHtml -AdScope $manyCountries -TableStyle 'x' -HeadStyle 'x'
  Check ($html -match 'Country 20' -and $html -notmatch 'Country 21' -and $html -match 'Other countries \(10\)') 'Country table keeps 20 rows plus Other countries above 25'
  $fewCountries = $manyCountries | Select-Object *
  $fewCountries.ByCountry = @($manyCountries.ByCountry | Select-Object -First 25)
  Check ((New-LicensesScopeTablesHtml -AdScope $fewCountries -TableStyle 'x' -HeadStyle 'x') -match 'Country 25' -and (New-LicensesScopeTablesHtml -AdScope $fewCountries -TableStyle 'x' -HeadStyle 'x') -notmatch 'Other countries') 'Up to 25 countries are all listed'

  # A report published before the AD views: no tables, and a domain mail is an explicit N/D notice.
  $legacyRoot = Join-Path $testRoot 'legacy'
  New-Item -ItemType Directory -Path $legacyRoot -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $testRoot 'M365_Licenses_Tenant.csv') -Destination $legacyRoot
  @{Status='Completed';IsPartialInventory=$false;Files=@()} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $legacyRoot 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt')
  Publish-TestLicenseReport -Folder $legacyRoot -TenantRows $tenantRows -CollectedAtUtc $collectedAt -WithoutUsage | Out-Null
  Check (-not (Test-Path -LiteralPath (Join-Path $legacyRoot 'M365_Licenses_ReportByDomain.csv'))) 'No usage, no AD scope CSV'
  $script:ScriptLocalConfig.LicenseSummaryDomainMails = @(@{ Domain=$alpha; To='alpha@example.invalid' })
  $legacyState = Join-Path $legacyRoot 'M365_Licenses_SummaryEmail_SendState.json.txt'
  [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $legacyRoot -ExpectedTenantKey 'prod' -MailStatePath $legacyState -Manual)
  $legacyTenant = $script:SentMail[$script:SentMail.Count - 2]
  $legacyDomain = $script:SentMail[$script:SentMail.Count - 1]
  Check ($legacyTenant.BodyHtml -notmatch 'Licenses by AD domain' -and $legacyDomain.Subject -like '* - N/D' -and $legacyDomain.BodyHtml -match 'predates the AD domain views') 'Legacy report: tenant mail unchanged, domain mail N/D'
}
finally {
  Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: License AD domain and country views and per-domain emails ($script:Checks checks)."

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCbaPF5QTc5YXif
# cj7Ad2D5RKPxo+tsfhpx0FJMkxyh7qCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg8qqJTaKOGLwU
# ff5HXjfMj1qLh5iNeigz5VclhNb5PDMwDQYJKoZIhvcNAQEBBQAEggGAa8cVQCPC
# p3uuQCo2R2vRg9AUZOzfNMg+RPOdkz9ZFvprazqpLQt9EtjQATyndd6J4z4aINuL
# gbfTYICsjbrAtOeO7cwFg8pKPwgPV0gQbewR87VQ8fG8s9sXnF+BudYmvAwnWN7d
# kn7djjADGq+49QAgW40kq7g4ALOe2206mNzsMpz/+Z0OQMAsPBmySZpyXFU49HTa
# r0covTqS0CUBfJ41k0rX9OjeIyM1fAWTEFrakCCPZcz6uFMo88xf/H5YY3LETy43
# mV+2Pm3rR+FAQVDEB24B4IoyhkhGgWQEL/vNClq9oaoyWfLMpSTOdZG7ru1ayiKQ
# 8wa4FPh5u01qTJ5k2+ZZbMn5Pan/1TW8pLxFZzCKLiLepQAhIM0nHvC7mM7R/EOH
# Wogoi3GoUTA2cmpz6BSyMbPnomrPFatVkXJgSjF/Xy59FsrYApq7UGKU/yHBpuxN
# H51/2Lhwci1J1j1sce8jEbEDpdc+DqMtOXp25EQUkrgce7qua6Qa6oeqoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTA0MTYxNDExWjA/BgkqhkiG9w0BCQQxMgQwkFa6y/MVdRyIB0LWXqe6UrvV
# PVkOmEev4UE/YcnVlfjvAytIIj0caqk3q4TxxkyeMA0GCSqGSIb3DQEBAQUABIIC
# AAwwZVT2GEq5xX+zCU8L1dLUZ+SKhC1tx52X1BQ+vEcUDtKlG03oFKn3UcyrPNvN
# 4v5BBeeP6DZGIVF0oiLfd4CQdkewZR4lWq/nW7sMu7yR6np1SRB3NAWRVZUxmIdm
# 02wEhZ6tXhNCJBdhGbywRLwQhV9oZeksSf230uTOvUWLW8AfIQrHJz+PXqn4Fc5R
# F4t2jzQQFmJv36TFL/7QeZy1ZuvKJbpG4VyK7rCoa8dTeov1HTD27YOl7i1hK9aa
# 4zgB4r86RHEcZigDPl3wWG1Lzwl8h/8TdeBvIRH5mRI6pntFQMPQs87MFof/gaEw
# 9uUKAU0HjemcS6Ty7Ynn3mJRGyLXtzSUO+YOXp1gf7vIBL53nWnEautshwJwlXB5
# cqf/Y/yTWVsDmgir9OarGunhl/jdh/H4BwexxndBWJKpoC5NQD6JNWTX5q18AnlJ
# DPsPRdAbg0wjEWp/OjIAYC3gpDKf8TAcCLzc+VJncQd/DyHLoaEF9p5oyIl/eYES
# u80HYtlOgpeb9ccIgdG6k0iIESkaAVIHAreQgUahXXsNvWFdeZU+45cnlQe9MNeC
# VbXFICANSqT3kFhl2uKGxvCImfa3uHvHcu0r5G01P73eSTjfXIMDLVaKF4zHW1i9
# lZoN642iFDHZsDNE9JQtHu5nVAroNXAmCf9CjZ6TWmXv
# SIG # End signature block
