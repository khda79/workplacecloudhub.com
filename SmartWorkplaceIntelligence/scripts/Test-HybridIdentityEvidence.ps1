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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+NSWH5GhzyovX
# qBVv7ila33pfvRzRBE95D72OAbArm6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB62eEBLGQFKW5rJzCYyh2YYaJPozZIOgfAH4841qMFnMA0GCSqG
# SIb3DQEBAQUABIIBgI4e0pltV6NNxGb2qoK8sqWdvqrzMhGl+SbYeQZ4amo7fXDJ
# t1v3L6E3rYcz3gvMGmQdTFE9i5/c204orXju0BmprY6zsmG3l1iOpmYsYBRxwCEo
# S60+DhD9EuCE/KgqE/fujq3M7BnL1ghVQBkwGyLKQb2lCoOJRHdQzo19RRUttwh1
# 3RBe5UgIiz4aGl2lkW0gwnMtNWEEEGM5euIHPyDII31KMOSJWduiETtZS1U9cBfF
# I9/68lQnCFV9uRff2WQ6R9uxyJoL8MaPwZXqE4JsiAEH+zKiehi2BHEfWlKXIg4w
# xHHNKFB6l5kE38ezk5J+sNItH9C/hnaUyZRdVBHiIGGmoO7V57MawpB23uIEq7HH
# 5CAVCH5n8u3EQ7vsRQq48MQEOqHGzSphV0FXIVAsa7AdqJx2MdlNHBAalnw3fXmG
# DeAoZnv6+niDHxP7gDumvUIs5ar8RhAqDXnrLhHkHIIQcAxhqQAMHBFe05EZz5ZY
# ghUCgvl2WSEmzrr4g6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMjU3
# NTZaMC8GCSqGSIb3DQEJBDEiBCCu5sQ/klqB99Wlm92UPEjwv1L/l1aNZ8KQPBgM
# IjWtsTANBgkqhkiG9w0BAQEFAASCAgCNen+Zth9bmoMeL3MQ1pgi0LFsImE5FkIB
# c7UZlL+DgSngl3s8j7rAO0V90dtzuIT+tNj5fljHzP90Mx0lddoHsmfuZvWlNSB3
# jn/GkCwFkcGOku3RA5pshF85HVqjZ404JVqXasNe5Wg5cBy5hCQVqfI6uWfQ1QhZ
# E8DfJmXJE8inq4gYchT/u6w0zaMG7wcjIs9CO2D9ZvojDNfrej954emXPd7NoLSm
# BNJPAp1JrId+nhWQ+0Umf856g3KaXIu4KUA7D8Bnn/mADf3B1zCRAutF2aQO1G7A
# ZJkrvgxG75I5oPbQgCrE8Qe7pR8+KKbXzv+/xHhFRQstUgPFH+0ku0qc38QfMwR+
# 8ane1UJHZ6uVKLH2TAWSVjfEqpR+J8gZmd4TYQU/ZgM/4ln/dTWYcnuO7l/JFm9V
# D/BbT9CgO0VYMp1owny+jtuF48RKEPI375zaTqmV3kr5mOIzBkUzuVIbQOlXSkAw
# ZBPA87sSdrkHTbK7aKc9O/bghqsHh6dIvP/XcrvJMA8i8bI0Lv/qYJ7/dyarwfoq
# qyd4izNRT+si8GEXt2Ou660aIyUEu2Ouvo3nhaIa21TckBpAplhkoqCHdx2EDRbY
# XGvHTc8YxXeEZZ9PiQfY3FqsNznfTP7xmmEKK9H7apYtShbKGrTt9kUzwCJzf1Xt
# i1S5USxu8Q==
# SIG # End signature block
