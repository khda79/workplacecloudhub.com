<#
.SYNOPSIS
  Offline checks for the license views by AD domain and country and the per-domain license summary emails.
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
  'Get-LicensesAlertSummary', 'New-LicensesRecoveryWorkbook', 'Publish-LicensesReportCsv', 'Publish-LicensesReportSnapshot', 'Write-LicensesDailyMailState', 'Enter-LicensesDailyMailGate',
  'Read-LicensesTenantSnapshot', 'ConvertFrom-LicensesJsonElement', 'ConvertTo-LicensesCountTable', 'Read-LicensesPublishedReport', 'Get-LicensesDomainPublishedReport',
  'New-LicensesScopeTablesHtml', 'New-LicensesSummaryMailContent', 'Get-LicensesDomainMailSettings', 'Get-LicensesDomainMailStatePath', 'Send-LicensesSummaryMailScope',
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
function New-SmartM365EmailBody {
  param([string]$Title, [string]$Category, [string]$HostName, [string]$GeneratedAt, [string]$BodyHtml, [string]$Footer)
  return "<html><body><header>$Title</header><main>$BodyHtml</main><footer>$Footer</footer></body></html>"
}
$script:SentMail = [System.Collections.Generic.List[object]]::new()
function Send-SmartM365Mail {
  param([string]$From, [string]$To, [string]$Cc, [string]$Subject, [string]$BodyHtml, [string]$MailPurpose, [string[]]$Attachments, [switch]$AllowAttachments, [switch]$SuppressAttachmentLinks)
  $candidates = @()
  $downgrade = @()
  if ($Attachments) {
    if (-not $AllowAttachments -or -not $SuppressAttachmentLinks -or $Attachments.Count -ne 1) { throw 'Recovery workbook was not attached privately.' }
    $candidates = @(Import-Excel -Path $Attachments[0] -WorksheetName 'Recovery candidates' -WarningAction SilentlyContinue)
    $downgrade = @(Import-Excel -Path $Attachments[0] -WorksheetName 'E3 to F3 review' -WarningAction SilentlyContinue)
  }
  $script:SentMail.Add([pscustomobject]@{ From=$From; To=$To; Cc=$Cc; CcBound=$PSBoundParameters.ContainsKey('Cc'); Subject=$Subject; BodyHtml=$BodyHtml
    MailPurpose=$MailPurpose; HasAttachment=[bool]$Attachments; Candidates=$candidates; Downgrade=$downgrade }) | Out-Null
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
  Check ($alphaMail.Candidates.Count -eq 1 -and $alphaMail.Candidates[0].UserId -eq 'a1') 'Alpha workbook lists only alpha recovery candidates'
  Check ($alphaMail.BodyHtml -notmatch 'b1@example\.invalid|c1@example\.invalid' -and @($alphaMail.Candidates | Where-Object { $_.UserPrincipalName -notlike 'a*' }).Count -eq 0) 'No user of another domain reaches the alpha mail'
  $betaMail = @($script:SentMail | Where-Object To -eq 'beta@example.invalid')[0]
  Check ($betaMail.CcBound -and $betaMail.Cc -eq '' -and $betaMail.Candidates.Count -eq 1 -and $betaMail.Candidates[0].UserId -eq 'b1') 'Beta mail and workbook are limited to beta'
  $missingMail = @($script:SentMail | Where-Object To -eq 'missing@example.invalid')[0]
  Check ($missingMail.Subject -like '* - missing.example.invalid - N/D' -and -not $missingMail.HasAttachment -and $missingMail.BodyHtml -match 'no licensed or mailbox user') 'A configured domain without data receives an explicit N/D notice'
  Check (@($script:SentMail | Where-Object To -eq 'off@example.invalid').Count -eq 0) 'Disabled domain entry sends nothing'
  Check (Test-Path -LiteralPath (Get-LicensesDomainMailStatePath -Path $statePath -Domain $alpha)) 'Alpha send state is written'

  $sentAgain = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath
  Check (-not $sentAgain -and $script:SentMail.Count -eq 4) 'Each mail is sent once per Paris day'
  Remove-Item -LiteralPath (Get-LicensesDomainMailStatePath -Path $statePath -Domain $beta)
  [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath)
  Check ($script:SentMail.Count -eq 5 -and $script:SentMail[4].To -eq 'beta@example.invalid') 'Domain send states are independent of each other and of the tenant'
  $only = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath -Manual -ForceLicenseSummaryEmail -Domain 'ALPHA.example.invalid'
  Check ($only -and $script:SentMail.Count -eq 6 -and $script:SentMail[5].To -eq 'alpha@example.invalid') 'Email-only domain option resends one domain'
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
  $tenantOnly = Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $statePath
  Check ($tenantOnly -and $script:SentMail.Count -eq 7 -and @($script:Warnings | Where-Object { $_ -like 'License summary emails by AD domain skipped*' }).Count -eq 1) 'Invalid domain settings never block the tenant mail'

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
