#Requires -Version 7.0
# Version 1.1.0
[CmdletBinding()]
param()
# Synthetic contoso fixtures only. No tenant initialization, collectors, production files or Power BI.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { $script:checks++; if (-not $Pass) { throw $Message } }

$root = Join-Path ([IO.Path]::GetTempPath()) ('HybridIdentityEvidence-' + [guid]::NewGuid().ToString('N'))
$last = Join-Path $root 'DATA-LAST'
New-Item -ItemType Directory -Path $last -Force | Out-Null
$producer = Join-Path $PSScriptRoot 'New-HybridIdentityEvidence.ps1'
$tenant = 'synthetic-test'
$routing = 'contoso.mail.onmicrosoft.com'
$rules = Join-Path $root 'AccountClassification.json.txt'
[IO.File]::WriteAllText($rules, '{"SchemaVersion":"1.0","RuleVersion":"SYNTHETIC-1","Population":{"Human":["Named Account"],"NonHuman":["Service Account","System Account"],"ReviewRequired":["Admin Account","Generic Account","Unclassified Account"]}}')

$adRows = [Collections.Generic.List[object]]::new(); $local = [Collections.Generic.List[object]]::new(); $remote = [Collections.Generic.List[object]]::new()
$exoRows = [Collections.Generic.List[object]]::new(); $statsRows = [Collections.Generic.List[object]]::new(); $m365 = [Collections.Generic.List[object]]::new()
$licenses = [Collections.Generic.List[object]]::new(); $workforce = [Collections.Generic.List[object]]::new(); $expected = @{}
$script:sequence = 0

# One account: AD row, optional mailbox rows, Entra user, license and Workforce population.
# Defaults describe a clean remote user mailbox with no finding.
function Add-Account {
    param([string]$Name, [string[]]$Expect = @(), [hashtable]$Ad = @{}, [string]$Hosting = 'Remote', [hashtable]$Mailbox = @{}, [string]$RecipientType,
        [AllowNull()][string]$Sku = 'SPE_E3', [string]$Population = 'Human', [switch]$NoEntra, [hashtable]$Exo, [hashtable]$Stats, [string]$Guid)
    $script:sequence++
    if (-not $Guid) { $Guid = '00000000-0000-0000-0000-{0:d12}' -f $script:sequence }
    $upn = "$Name@contoso.com"; $primary = "$Name@contoso.com"; $immutable = "imm-$Name"; $entra = "entra-$Name"
    $row = [ordered]@{TenantKey = $tenant; ObjectGUID = $Guid; UserPrincipalName = $upn; ImmutableId_AD = $immutable; DomainAndSam = "CONTOSO\$Name"; Enabled = 'True'
        EmailAddress = $primary; ProxyAddresses = "SMTP:$primary;smtp:$Name@$routing"; TargetAddress = $(if ($Hosting -eq 'Remote') { "SMTP:$Name@$routing" } else { '' })
        AccountType = 'Named Account'; ObjectSID = "S-1-5-21-$($script:sequence)"; PrimarySmtpAddress = '' }
    foreach ($key in $Ad.Keys) { $row[$key] = $Ad[$key] }
    $adRows.Add([pscustomobject]$row)
    $mailboxRow = [ordered]@{}
    if ($Hosting -eq 'Remote') {
        $mailboxRow = [ordered]@{ObjectGuid = $Guid; UserPrincipalName = $row.UserPrincipalName; PrimarySmtpAddress = $row.EmailAddress; RecipientTypeDetails = $(if ($RecipientType) { $RecipientType } else { 'RemoteUserMailbox' }); ExchangeGuid = "xg-$Name" }
        foreach ($key in $Mailbox.Keys) { $mailboxRow[$key] = $Mailbox[$key] }; $remote.Add([pscustomobject]$mailboxRow)
    } elseif ($Hosting -eq 'Local') {
        $mailboxRow = [ordered]@{ObjectGUID = $Guid; UserPrincipalName = $row.UserPrincipalName; PrimarySMTPaddress = $row.EmailAddress; DomainAndSam = $row.DomainAndSam; RecipientType = $(if ($RecipientType) { $RecipientType } else { 'UserMailbox' }); ExchangeGuid = "xg-$Name"; ItemCount = '10'; 'TotalItemSize-In-MB' = '10' }
        foreach ($key in $Mailbox.Keys) { $mailboxRow[$key] = $Mailbox[$key] }; $local.Add([pscustomobject]$mailboxRow)
    }
    if ($Exo) { $e = [ordered]@{UserPrincipalName = $row.UserPrincipalName; PrimarySmtpAddress = $row.EmailAddress; RecipientTypeDetails = 'UserMailbox'; MailboxGuid = "mg-$Name"; ItemCount = '10'; TotalItemSizeGB = '0,01'; OnPremisesImmutableId = $immutable }; foreach ($key in $Exo.Keys) { $e[$key] = $Exo[$key] }; $exoRows.Add([pscustomobject]$e) }
    if ($Stats) { $s = [ordered]@{UserPrincipalName = $row.UserPrincipalName; ItemCount = ''; TotalItemSizeGB = '' }; foreach ($key in $Stats.Keys) { $s[$key] = $Stats[$key] }; $statsRows.Add([pscustomobject]$s) }
    if (-not $NoEntra) {
        $m365.Add([pscustomobject]@{TenantKey = $tenant; 'User principal name' = $row.UserPrincipalName; 'Object Id' = $entra; OnPremisesImmutableId = $immutable; OnPremisesSecurityIdentifier = $row.ObjectSID })
        if ($Sku) { $licenses.Add([pscustomobject]@{TenantKey = $tenant; 'User principal name' = $row.UserPrincipalName; UserId = $entra; 'SKU name' = $Sku }) }
        if ($Population) { $workforce.Add([pscustomobject]@{TenantKey = $tenant; 'User Source ID' = $entra; 'Account Population' = $Population }) }
    }
    $script:expected[$Guid] = @($Expect | Sort-Object)
    return $Guid
}

# Clean and identity sync.
Add-Account 'clean' | Out-Null
Add-Account 'notinentra' @('HYB-001') -NoEntra | Out-Null
Add-Account 'dupupn' @('HYB-009') | Out-Null
Add-Account 'badupn' @('HYB-039') -Ad @{UserPrincipalName = 'bad#upn@contoso.com'} | Out-Null
Add-Account 'unverified' @('HYB-035') -Ad @{UserPrincipalName = 'unverified@unverified.example'} | Out-Null
# Item count (rule 2): above, at the limit, unknown on-premises with known EXO statistics, all unknown.
Add-Account 'items' @('HYB-002') -Hosting Local -Mailbox @{ItemCount = '1000000'} | Out-Null
Add-Account 'itemsedge' -Hosting Local -Mailbox @{ItemCount = '999999'} | Out-Null
Add-Account 'itemsstats' @('HYB-002') -Hosting Local -Mailbox @{ItemCount = 'Error'} -Stats @{ItemCount = '1000001'} | Out-Null
# A migrated remote mailbox without statistics uses the EXO mailbox export.
Add-Account 'itemsexo' @('HYB-002') -Exo @{ItemCount = '1000002'} | Out-Null
Add-Account 'itemsunknown' -Hosting Local -Mailbox @{ItemCount = 'Error'} | Out-Null
# Mailbox size (rule 6): 2000 MB is the legacy threshold, never 2048.
Add-Account 'f3below' -Hosting Local -Mailbox @{'TotalItemSize-In-MB' = '1999'} -Sku 'SPE_F3' | Out-Null
Add-Account 'f3limit' @('HYB-006') -Hosting Local -Mailbox @{'TotalItemSize-In-MB' = '2000'} -Sku 'SPE_F3' | Out-Null
Add-Account 'f3under2048' @('HYB-006') -Hosting Local -Mailbox @{'TotalItemSize-In-MB' = '2047'} -Sku 'SPE_F1' | Out-Null
Add-Account 'f3exogb' @('HYB-006') -Hosting Local -Mailbox @{'TotalItemSize-In-MB' = 'Error'} -Stats @{TotalItemSizeGB = '1,953125'} -Sku 'SPE_F3' | Out-Null
Add-Account 'e3large' -Hosting Local -Mailbox @{'TotalItemSize-In-MB' = '5000'} | Out-Null
# Addressing.
Add-Account 'norouting' @('HYB-003') -Hosting Local -Ad @{ProxyAddresses = 'SMTP:norouting@contoso.com'} | Out-Null
Add-Account 'dupsmtp' @('HYB-016') -Ad @{ProxyAddresses = "SMTP:dupsmtp@contoso.com;smtp:dupsmtp@$routing;smtp:shared@contoso.com"} | Out-Null
Add-Account 'noprimary' @('HYB-024') -Hosting Local -Ad @{ProxyAddresses = "smtp:noprimary@contoso.com;smtp:noprimary@$routing"} | Out-Null
Add-Account 'twoprimary' @('HYB-025') -Hosting Local -Ad @{ProxyAddresses = "SMTP:twoprimary@contoso.com;SMTP:second@contoso.com;smtp:twoprimary@$routing"} | Out-Null
Add-Account 'primarymissing' @('HYB-026') -Hosting Local -Ad @{ProxyAddresses = "SMTP:alias@contoso.com;smtp:primarymissing@$routing"} | Out-Null
Add-Account 'caseconflict' @('HYB-051') -Hosting Local -Ad @{ProxyAddresses = "SMTP:caseconflict@contoso.com;smtp:CaseConflict@contoso.com;smtp:caseconflict@$routing"} | Out-Null
Add-Account 'sip' -Hosting Local -Ad @{ProxyAddresses = "SMTP:sip@contoso.com;SIP:sip@contoso.com;smtp:xsip@contoso.com;smtp:sip@$routing"} | Out-Null
Add-Account 'notaccepted' @('HYB-030') -Hosting Local -Ad @{EmailAddress = 'notaccepted@notaccepted.example'; ProxyAddresses = "SMTP:notaccepted@notaccepted.example;smtp:notaccepted@$routing"} | Out-Null
Add-Account 'initial' @('HYB-031') -Hosting Local -Ad @{EmailAddress = 'initial@contoso.onmicrosoft.com'; ProxyAddresses = "SMTP:initial@contoso.onmicrosoft.com;smtp:initial@$routing"} | Out-Null
Add-Account 'relay' @('HYB-032') -Hosting Local -Ad @{EmailAddress = 'relay@relay.contoso.com'; ProxyAddresses = "SMTP:relay@relay.contoso.com;smtp:relay@$routing"} | Out-Null
Add-Account 'badsmtp' @('HYB-040') -Hosting Local -Ad @{EmailAddress = 'bad+smtp@contoso.com'; ProxyAddresses = "SMTP:bad+smtp@contoso.com;smtp:badsmtp@$routing"} | Out-Null
# Routing (rules 19-23): remote mailboxes only, smtp: prefix removed, case-insensitive.
Add-Account 'notarget' @('HYB-019') -Ad @{TargetAddress = ''} | Out-Null
Add-Account 'badtarget' @('HYB-020') -Ad @{TargetAddress = 'SMTP:badtarget-at-contoso'; ProxyAddresses = "SMTP:badtarget@contoso.com;smtp:badtarget@$routing;smtp:badtarget-at-contoso"} | Out-Null
Add-Account 'targetcase' -Ad @{TargetAddress = "smtp:TargetCase@CONTOSO.MAIL.ONMICROSOFT.COM"; ProxyAddresses = "SMTP:targetcase@contoso.com;smtp:targetcase@$routing"} | Out-Null
Add-Account 'disallowed' @('HYB-021') -Ad @{TargetAddress = 'SMTP:disallowed@other.example'; ProxyAddresses = "SMTP:disallowed@contoso.com;smtp:disallowed@$routing;smtp:disallowed@other.example"} | Out-Null
Add-Account 'allowed' -Ad @{TargetAddress = 'SMTP:allowed@partner.example'; ProxyAddresses = "SMTP:allowed@contoso.com;smtp:allowed@$routing;smtp:allowed@partner.example"} | Out-Null
Add-Account 'targetproxy' @('HYB-022') -Ad @{ProxyAddresses = "SMTP:targetproxy@contoso.com;smtp:elsewhere@$routing"} | Out-Null
$sharedTarget = "shared-target@$routing"
Add-Account 'duptarget1' @('HYB-023') -Ad @{TargetAddress = "SMTP:$sharedTarget"; ProxyAddresses = "SMTP:duptarget1@contoso.com;smtp:$sharedTarget"} | Out-Null
Add-Account 'duptarget2' @('HYB-023') -Ad @{TargetAddress = $sharedTarget.ToUpperInvariant(); ProxyAddresses = "SMTP:duptarget2@contoso.com;smtp:$sharedTarget"} | Out-Null
Add-Account 'localtarget' -Hosting Local -Ad @{TargetAddress = ''} | Out-Null
# The same account exported twice is one account, not a duplicate TargetAddress.
Add-Account 'repeated' | Out-Null
$adRows.Add($adRows[$adRows.Count - 1])
# Mailbox state.
Add-Account 'linked' @('HYB-029') -Hosting Local -RecipientType 'LinkedMailbox' | Out-Null
Add-Account 'splitbrain' @('HYB-038') -Hosting Local -Exo @{} | Out-Null
Add-Account 'exoonly' @('HYB-054') -Hosting None -Ad @{ProxyAddresses = ''; EmailAddress = ''} -Exo @{} | Out-Null
Add-Account 'nomailbox' -Hosting None -Ad @{ProxyAddresses = ''; EmailAddress = ''} | Out-Null
# License governance (rule 58): governed Non-human population only; uncertain stays Review required.
Add-Account 'service' @('HYB-058') -Ad @{AccountType = 'Service Account'} -Population '' | Out-Null
Add-Account 'servicereview' -Ad @{AccountType = 'Service Account'} -Population 'Workforce - review classification' | Out-Null
Add-Account 'generic' -Ad @{AccountType = 'Generic Account'} -Population '' | Out-Null
Add-Account 'admin' -Ad @{AccountType = 'Admin Account'} -Population '' | Out-Null
Add-Account 'serviceunlicensed' -Ad @{AccountType = 'Service Account'} -Population '' -Sku $null | Out-Null
Add-Account 'serviceshared' -Ad @{AccountType = 'Service Account'} -Population '' -RecipientType 'RemoteSharedMailbox' | Out-Null
$conflict = Add-Account 'conflict' -Ad @{AccountType = 'Service Account'} -Population ''
$adRows.Add([pscustomobject]@{TenantKey = $tenant; ObjectGUID = '00000000-0000-0000-0000-999999999999'; UserPrincipalName = 'conflict-named@contoso.com'; ImmutableId_AD = ''; DomainAndSam = 'CONTOSO\conflict-named'; Enabled = 'True'; EmailAddress = ''; ProxyAddresses = ''; TargetAddress = ''; AccountType = 'Named Account'; ObjectSID = ($adRows | Where-Object ObjectGUID -EQ $conflict | Select-Object -First 1).ObjectSID; PrimarySmtpAddress = '' })

function Write-Csv([string]$Name, [object[]]$Rows, [string]$Header) {
    $path = Join-Path $last $Name
    if ($Rows.Count) { $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8NoBOM } else { [IO.File]::WriteAllText($path, $Header + [Environment]::NewLine) }
}
Write-Csv 'AD_Users_AllDomains.csv' $adRows.ToArray()
Write-Csv 'Exchange_OnPrem_Mailboxes_AllDomains.csv' $local.ToArray()
Write-Csv 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv' $remote.ToArray()
Write-Csv 'Exchange_EXO_Mailboxes_AllDomains.csv' $exoRows.ToArray()
Write-Csv 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv' $statsRows.ToArray()
Write-Csv 'M365_Users_Active.csv' $m365.ToArray()
Write-Csv 'M365_Licenses_Users.csv' $licenses.ToArray()
Write-Csv 'AD_Users_DuplicateUPN.csv' @([pscustomobject]@{TenantKey = $tenant; UserPrincipalName = 'dupupn@contoso.com'})
Write-Csv 'AD_Users_DuplicateSMTP.csv' @([pscustomobject]@{TenantKey = $tenant; SmtpAddress = 'shared@contoso.com'})
Write-Csv 'Exchange_EXO_AcceptedDomains.csv' @(
    [pscustomobject]@{TenantKey = $tenant; DomainName = 'contoso.com'; DomainType = 'Authoritative'}
    [pscustomobject]@{TenantKey = $tenant; DomainName = 'contoso.onmicrosoft.com'; DomainType = 'Authoritative'}
    [pscustomobject]@{TenantKey = $tenant; DomainName = $routing; DomainType = 'Authoritative'}
    [pscustomobject]@{TenantKey = $tenant; DomainName = 'relay.contoso.com'; DomainType = 'InternalRelay'})
Write-Csv 'M365_Entra_VerifiedDomains.csv' @([pscustomobject]@{TenantKey = $tenant; Id = 'contoso.com'}, [pscustomobject]@{TenantKey = $tenant; Id = $routing})
$userEvidence = Join-Path $root 'UserInventoryEvidence.csv'
$workforce | Export-Csv -LiteralPath $userEvidence -NoTypeInformation -Encoding utf8NoBOM

function Invoke-Producer([string]$Name, [string[]]$Allowed = @('partner.example')) {
    $output = Join-Path $root "$Name.csv"; $audit = Join-Path $root "$Name.audit.json.txt"
    & $producer -DataRoot $last -UserEvidencePath $userEvidence -OutputPath $output -AuditPath $audit -AccountClassificationConfigPath $rules -AllowedTargetDomains $Allowed -AsOfUtc '2026-10-25T08:00:00Z' 6>$null | Out-Null
    [pscustomobject]@{Path = $output; Rows = @(Import-Csv -LiteralPath $output); Audit = (Get-Content -LiteralPath $audit -Raw | ConvertFrom-Json)}
}

try {
    $run = Invoke-Producer 'main'
    $byAccount = @{}
    foreach ($row in $run.Rows) { $key = $row.'AD Object GUID'; if (-not $byAccount.ContainsKey($key)) { $byAccount[$key] = [Collections.Generic.List[string]]::new() }; $byAccount[$key].Add($row.'Rule ID') }
    foreach ($guid in $expected.Keys) {
        $actual = if ($byAccount.ContainsKey($guid)) { @($byAccount[$guid] | Sort-Object) } else { @() }
        Check (($actual -join ',') -ceq ($expected[$guid] -join ',')) "Account $guid expected [$($expected[$guid] -join ',')] got [$($actual -join ',')]"
    }
    Check (-not $byAccount.ContainsKey('00000000-0000-0000-0000-999999999999')) 'An account without mailbox produced findings'
    Check (@($run.Rows | Group-Object 'Finding Key' | Where-Object Count -GT 1).Count -eq 0) 'Finding keys are not unique'
    $first = $run.Rows | Where-Object 'Rule ID' -EQ 'HYB-001' | Select-Object -First 1
    Check ($first.Severity -eq 'Blocker' -and $first.'Is Blocker' -eq 'True' -and $first.'Legacy Issue Number' -eq '1' -and $first.TenantKey -eq $tenant -and $first.'Evidence Date' -eq '2026-10-25') 'Finding metadata changed'
    $service = $run.Rows | Where-Object 'Rule ID' -EQ 'HYB-058'
    Check ($service.Severity -eq 'Review' -and $service.'Is Blocker' -eq 'False' -and $service.'Account Population' -eq 'Non-human' -and $service.'Account Population Source' -eq 'AD AccountType' -and $service.'Account Classification Rule Version' -eq 'SYNTHETIC-1') 'Rule 58 population metadata changed'
    $dup = @($run.Rows | Where-Object 'Rule ID' -EQ 'HYB-023')
    Check ($dup.Count -eq 2 -and @($dup | Where-Object { $_.'Observed Value' -match '\(2 accounts\)' }).Count -eq 2) 'Duplicate TargetAddress must count distinct accounts'
    Check ((@($run.Rows | Where-Object 'Rule ID' -EQ 'HYB-006' | ForEach-Object 'Observed Value') -join '|') -match '2000 MB' ) 'Rule 6 observed size lost its unit'
    Check ($run.Audit.Rules.'HYB-002'.NotAssessed -ge 1 -and -not $byAccount.ContainsKey(($adRows | Where-Object UserPrincipalName -EQ 'itemsunknown@contoso.com').ObjectGUID)) 'Unknown item count must be not assessed, never zero'
    Check ($run.Audit.RoutingDomainDecision -eq 'Single accepted mail.onmicrosoft.com domain' -and $run.Audit.AllowedTargetDomainCount -eq 2) 'Routing domain decision changed'
    Check (($run.Audit | ConvertTo-Json -Depth 6) -notmatch 'contoso\.com|partner\.example') 'Audit exposes addresses or domains'

    # Deterministic: same inputs, same bytes.
    $again = Invoke-Producer 'again'
    Check ((Get-FileHash $run.Path).Hash -eq (Get-FileHash $again.Path).Hash) 'Output is not deterministic'

    # A comma-separated value from a -File child process is split into domains.
    $comma = Invoke-Producer 'comma' @('partner.example,other.example')
    Check (@($comma.Rows | Where-Object 'Rule ID' -EQ 'HYB-021').Count -eq 0) 'Comma-separated allowed domains were not split'

    # Without accepted domains: routing, domain and relay rules are not assessed instead of guessed.
    Move-Item -LiteralPath (Join-Path $last 'Exchange_EXO_AcceptedDomains.csv') -Destination (Join-Path $root 'accepted.bak')
    # Configured domains are never an exhaustive list: without a routing domain HYB-021 stays not assessed.
    $noDomains = Invoke-Producer 'nodomains' @('partner.example')
    foreach ($id in 'HYB-003', 'HYB-021', 'HYB-030', 'HYB-032') { Check (@($noDomains.Rows | Where-Object 'Rule ID' -EQ $id).Count -eq 0 -and $noDomains.Audit.Rules.$id.NotAssessed -gt 0) "$id was evaluated without accepted domains" }
    Check ($noDomains.Audit.RoutingDomainDecision -eq 'Accepted domains not available') 'Missing accepted domains decision changed'
    Move-Item -LiteralPath (Join-Path $root 'accepted.bak') -Destination (Join-Path $last 'Exchange_EXO_AcceptedDomains.csv')
    $acceptedPath = Join-Path $last 'Exchange_EXO_AcceptedDomains.csv'
    $acceptedBytes = [IO.File]::ReadAllBytes($acceptedPath)
    $accepted = @(Import-Csv -LiteralPath $acceptedPath)
    foreach ($case in @(
        @{Name = 'norouting'; Rows = @($accepted | Where-Object DomainName -NE $routing); Decision = 'No accepted mail.onmicrosoft.com domain'}
        @{Name = 'tworouting'; Rows = @($accepted) + @([pscustomobject]@{TenantKey = $tenant; DomainName = 'fabrikam.mail.onmicrosoft.com'; DomainType = 'Authoritative'}); Decision = 'Several accepted mail.onmicrosoft.com domains'})) {
        $case.Rows | Export-Csv -LiteralPath $acceptedPath -NoTypeInformation -Encoding utf8NoBOM
        $result = Invoke-Producer $case.Name @('partner.example', 'other.example')
        Check ($result.Audit.RoutingDomainDecision -eq $case.Decision) "$($case.Name): routing decision changed"
        foreach ($id in 'HYB-003', 'HYB-021') { Check (@($result.Rows | Where-Object 'Rule ID' -EQ $id).Count -eq 0 -and $result.Audit.Rules.$id.NotAssessed -gt 0 -and $result.Audit.Rules.$id.Evaluated -eq 0) "$($case.Name): $id evaluated without a single routing domain" }
        Check (@($result.Rows | Where-Object 'Rule ID' -EQ 'HYB-030').Count -ge 1) "$($case.Name): accepted-domain rules stopped"
    }
    [IO.File]::WriteAllBytes($acceptedPath, $acceptedBytes)

    # Without the duplicate SMTP source, rule 16 is not assessed.
    Move-Item -LiteralPath (Join-Path $last 'AD_Users_DuplicateSMTP.csv') -Destination (Join-Path $root 'dupsmtp.bak')
    $noDuplicateSmtp = Invoke-Producer 'nodupsmtp'
    Check (@($noDuplicateSmtp.Rows | Where-Object 'Rule ID' -EQ 'HYB-016').Count -eq 0 -and $noDuplicateSmtp.Audit.Rules.'HYB-016'.NotAssessed -gt 0) 'Rule 16 evaluated without its source'
    Move-Item -LiteralPath (Join-Path $root 'dupsmtp.bak') -Destination (Join-Path $last 'AD_Users_DuplicateSMTP.csv')

    # No finding at all still writes the declared header.
    $emptyLast = Join-Path $root 'empty/DATA-LAST'; New-Item -ItemType Directory -Path $emptyLast -Force | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath $last -Filter '*.csv') { [IO.File]::WriteAllText((Join-Path $emptyLast $file.Name), ((Get-Content -LiteralPath $file.FullName -TotalCount 1) + [Environment]::NewLine)) }
    $emptyOutput = Join-Path $root 'empty.csv'
    & $producer -DataRoot $emptyLast -UserEvidencePath $userEvidence -OutputPath $emptyOutput -AccountClassificationConfigPath $rules -AsOfUtc '2026-10-25T08:00:00Z' 6>$null | Out-Null
    $header = (Get-Content -LiteralPath $emptyOutput -TotalCount 1)
    $contract = (Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json).tables | Where-Object file -EQ 'HybridIdentityFindings.csv'
    Check ($header -ceq (($contract.columns.name | ForEach-Object { '"' + $_ + '"' }) -join ',')) 'Empty output header differs from the contract'
    Check ((Get-Content -LiteralPath $run.Path -TotalCount 1) -ceq $header) 'Output header differs from the contract'

    # The missing staged Workforce output stops the producer.
    $failed = $false
    try { & $producer -DataRoot $last -UserEvidencePath (Join-Path $root 'missing.csv') -OutputPath (Join-Path $root 'x.csv') -AccountClassificationConfigPath $rules 6>$null | Out-Null } catch { $failed = $_.Exception.Message -match 'Staged Workforce output' }
    Check $failed 'Missing Workforce output was accepted'

    # The shared population mapping keeps Workforce semantics.
    Import-Module (Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/Common/SmartM365.AccountClassification.psd1') -MinimumVersion '1.2.0' -Force
    $population = @{Human = @('Named Account'); NonHuman = @('Service Account')}
    Check ((Get-SmartM365AccountPopulation -PopulationRules $population -AccountType ' Service Account ') -eq 'Non-human') 'Trimmed Non-human type not mapped'
    Check ((Get-SmartM365AccountPopulation -PopulationRules $population -AccountType 'named account') -eq 'Human') 'Case-insensitive Human type not mapped'
    Check ((Get-SmartM365AccountPopulation -PopulationRules $population -AccountType 'Generic Account') -eq 'Review required') 'Unlisted type was not Review required'
    Check ((Get-SmartM365AccountPopulation -PopulationRules $population -AccountType $null) -eq 'Review required') 'Blank type was not Review required'
    Write-Host "PASS: $script:checks hybrid identity checks (all retained rules, boundaries, unknowns, routing, population, determinism). Synthetic fixtures only."
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
