# Version 1.0.1
# Hybrid identity findings for the retained legacy hybrid identity rules (see
# docs/HYBRID-IDENTITY-DEVICE-READINESS-FINDINGS.md). Offline: reads DATA-LAST and the
# staged Workforce output; no collection, upload or legacy producer change.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [Parameter(Mandatory = $true)]
    [string]$UserEvidencePath,

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\HybridIdentityFindings.csv'),

    [string]$AccountClassificationConfigPath = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'SmartM365\SmartInventory\Config\AccountClassification.local.json'),

    # TargetAddress domains accepted in addition to the remote routing domain (private tenant setting);
    # entries may also be comma-separated.
    [string[]]$AllowedTargetDomains = @(),

    [string]$AuditPath,

    # Optional reference clock for reproducible offline preparation; defaults to the current time.
    [Nullable[datetimeoffset]]$AsOfUtc
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ruleVersion = 'HybridIdentityFindings/1.1'
Import-Module (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'SmartM365\SmartInventory\Common\SmartM365.AccountClassification.psd1') -MinimumVersion '1.2.0' -ErrorAction Stop
$classification = Read-SmartM365AccountClassification -Path $AccountClassificationConfigPath
$populationRules = $classification.Population
$classificationRuleVersion = [string]$classification.RuleVersion
$reference = if ($null -ne $AsOfUtc) { $AsOfUtc.ToUniversalTime() } else { [datetimeoffset]::UtcNow }
$evidenceDate = $reference.UtcDateTime.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)

function T([AllowNull()][object]$Value) { if ($null -eq $Value) { '' } else { ([string]$Value).Trim() } }
function K([AllowNull()][object]$Value) { (T $Value).ToLowerInvariant() }
function P([AllowNull()][object]$Row, [string[]]$Names) {
    if ($null -eq $Row) { return $null }
    foreach ($name in $Names) { $property = $Row.PSObject.Properties[$name]; if ($property -and $null -ne $property.Value) { return $property.Value } }
    return $null
}
function B([AllowNull()][object]$Value) { (K $Value) -in @('true', '1', 'yes', 'y', 'enabled') }
function Get-SourceCsv([string]$Name, [switch]$Optional) {
    $path = Join-Path $DataRoot $Name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        if ($Optional) { return $null }
        throw "Required source CSV was not found: $path"
    }
    $script:sourceFiles[$Name] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    return @(Import-Csv -LiteralPath $path)
}
function Add-First([hashtable]$Map, [AllowNull()][object]$Key, [object]$Value) { $k = K $Key; if ($k -and -not $Map.ContainsKey($k)) { $Map[$k] = $Value } }
# The comma keeps an empty set from being unrolled to $null on output.
function New-Set { , [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
# Unknown is $null, never zero: 'Error', blanks and unparsable values stay unknown.
# Decimal commas are accepted ('0,43'), as in the legacy producer.
function ConvertTo-KnownNumber([AllowNull()][object]$Value) {
    $text = (T $Value) -replace ',', '.'
    $number = 0.0
    if ($text -and [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number)) { return $number }
    return $null
}
function Get-MaxKnown([object[]]$Values) {
    $known = @($Values | Where-Object { $null -ne $_ })
    if ($known.Count -eq 0) { return $null }
    return ($known | Measure-Object -Maximum).Maximum
}
function Split-Addresses([AllowNull()][object]$Value) {
    $text = T $Value
    if (-not $text) { return @() }
    return @($text -split '[;|,]' | ForEach-Object { (T $_) -replace '^(smtp|SMTP):', '' } | Where-Object { $_ })
}
function Get-DomainPart([AllowNull()][object]$Value) { $s = (T $Value) -replace '^(smtp|SMTP):', ''; if ($s -match '@') { $s.Split('@')[-1].ToLowerInvariant() } else { '' } }
function Get-LocalPart([AllowNull()][object]$Value) { $s = (T $Value) -replace '^(smtp|SMTP):', ''; if ($s -match '@') { $s.Split('@')[0] } else { $s } }
# Character sets unchanged from the legacy producer.
function Test-InvalidUpnCharacter([string]$Value) {
    $s = T $Value
    if (-not $s) { return $false }
    foreach ($character in @(' ', '#', '%', '&', '~', '+', '/', '=', '?', '*', '{', '}', '|', '<', '>', '(', ')', ';', ':', ',', '[', ']', '"', '\', '`')) { if ($s.Contains($character)) { return $true } }
    return $false
}
function Test-InvalidSmtpLocalCharacter([string]$Value) { (Get-LocalPart $Value) -match '[\\/\[\]:;\|=,+*?<>\s]' }
function Get-PrimaryProxyCount([string]$ProxyText) { @((T $ProxyText) -split '[;|,]' | Where-Object { $_ -cmatch '^SMTP:' }).Count }
function Test-ProxyContains([string]$ProxyText, [string]$Address) {
    $a = (T $Address) -replace '^(smtp|SMTP):', ''
    foreach ($entry in Split-Addresses $ProxyText) { if ($entry -ieq $a) { return $true } }
    return $false
}
# Legacy license grouping, unchanged (rule 6 F3/F1 and rule 58 E3/F3/F1).
function Get-LicenseGroup([string]$RecipientType, [object[]]$Licenses, [bool]$InM365) {
    if ((T $RecipientType) -in @('SharedMailbox', 'RemoteSharedMailbox')) { return 'No License (Shared)' }
    if ((T $RecipientType) -in @('RoomMailbox', 'RemoteRoomMailbox', 'EquipmentMailbox', 'RemoteEquipmentMailbox')) { return 'No License (Resource Mailbox)' }
    if (-not $InM365) { return 'Not in M365' }
    $text = ((@($Licenses) | ForEach-Object { T $_ }) -join ';')
    if ([string]::IsNullOrWhiteSpace($text)) { return 'No License (Unlicensed)' }
    if ($text -match 'Microsoft 365 E3|ENTERPRISEPACK|SPE_E3') { return 'Microsoft 365 E3' }
    if ($text -match 'Microsoft 365 F3|SPE_F1|SPE_F3|M365_F3|O365_F3') { return 'Microsoft 365 F3' }
    if ($text -match 'Microsoft 365 F1|M365_F1|O365_F1') { return 'Microsoft 365 F1' }
    return 'No License (Unlicensed)'
}

$script:sourceFiles = [ordered]@{}
$ad = Get-SourceCsv 'AD_Users_AllDomains.csv'
$duplicateUpn = Get-SourceCsv 'AD_Users_DuplicateUPN.csv'
$duplicateSmtp = Get-SourceCsv 'AD_Users_DuplicateSMTP.csv' -Optional
$local = Get-SourceCsv 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
$remote = Get-SourceCsv 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv'
$exo = Get-SourceCsv 'Exchange_EXO_Mailboxes_AllDomains.csv'
$stats = Get-SourceCsv 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv' -Optional
$m365 = Get-SourceCsv 'M365_Users_Active.csv'
$licenses = Get-SourceCsv 'M365_Licenses_Users.csv'
$domains = Get-SourceCsv 'Exchange_EXO_AcceptedDomains.csv' -Optional
$verified = Get-SourceCsv 'M365_Entra_VerifiedDomains.csv' -Optional
if (-not (Test-Path -LiteralPath $UserEvidencePath -PathType Leaf)) { throw "Staged Workforce output was not found: $UserEvidencePath" }
$userEvidence = @(Import-Csv -LiteralPath $UserEvidencePath)

$localByGuid = @{}; $localByUpn = @{}; $localBySmtp = @{}; $localByDomainSam = @{}
foreach ($r in $local) { Add-First $localByGuid (P $r 'ObjectGUID') $r; Add-First $localByUpn (P $r 'UserPrincipalName') $r; Add-First $localBySmtp (P $r 'PrimarySMTPaddress', 'PrimarySmtpAddress') $r; Add-First $localByDomainSam (P $r 'DomainAndSam') $r }
$remoteByGuid = @{}; $remoteByUpn = @{}; $remoteBySmtp = @{}
foreach ($r in $remote) { Add-First $remoteByGuid (P $r 'ObjectGuid', 'ObjectGUID') $r; Add-First $remoteByUpn (P $r 'UserPrincipalName') $r; Add-First $remoteBySmtp (P $r 'PrimarySmtpAddress', 'PrimarySMTPaddress') $r }
$exoByUpn = @{}; $exoBySmtp = @{}; $exoByImmutable = @{}
foreach ($r in $exo) { Add-First $exoByUpn (P $r 'UserPrincipalName') $r; Add-First $exoBySmtp (P $r 'PrimarySmtpAddress', 'PrimarySMTPaddress') $r; Add-First $exoByImmutable (P $r 'OnPremisesImmutableId', 'ImmutableId') $r }
$statsByUpn = @{}
foreach ($r in @($stats)) { if ($null -ne $r) { Add-First $statsByUpn (P $r 'UserPrincipalName') $r } }
$m365ByUpn = @{}; $m365ByImmutable = @{}; $m365BySid = @{}
foreach ($r in $m365) { Add-First $m365ByUpn (P $r 'User principal name', 'UserPrincipalName') $r; Add-First $m365ByImmutable (P $r 'OnPremisesImmutableId') $r; Add-First $m365BySid (P $r 'OnPremisesSecurityIdentifier') $r }
$licensesByUpn = @{}; $licensesById = @{}
foreach ($r in $licenses) {
    $upnKey = K (P $r 'User principal name', 'UserPrincipalName'); $idKey = K (P $r 'UserId', 'Object Id'); $sku = T (P $r 'SKU name', 'SkuPartNumber')
    if (-not $sku) { continue }
    if ($upnKey) { if (-not $licensesByUpn.ContainsKey($upnKey)) { $licensesByUpn[$upnKey] = [Collections.Generic.List[string]]::new() }; $licensesByUpn[$upnKey].Add($sku) }
    if ($idKey) { if (-not $licensesById.ContainsKey($idKey)) { $licensesById[$idKey] = [Collections.Generic.List[string]]::new() }; $licensesById[$idKey].Add($sku) }
}
$duplicateUpnSet = New-Set; foreach ($r in $duplicateUpn) { $v = T (P $r 'UPN_norm', 'UserPrincipalName'); if ($v) { [void]$duplicateUpnSet.Add($v) } }
$duplicateSmtpSet = New-Set; foreach ($r in @($duplicateSmtp)) { if ($null -ne $r) { $v = T (P $r 'PrimarySmtp', 'SmtpAddress'); if ($v) { [void]$duplicateSmtpSet.Add($v) } } }
$domainSet = New-Set; $domainType = @{}
foreach ($r in @($domains)) { if ($null -eq $r) { continue }; $name = T (P $r 'DomainName', 'Name'); if (-not $name) { continue }; [void]$domainSet.Add($name); $domainType[$name.ToLowerInvariant()] = T (P $r 'DomainType') }
$verifiedSet = New-Set; foreach ($r in @($verified)) { if ($null -ne $r) { $v = T (P $r 'Id', 'Name', 'DomainName'); if ($v) { [void]$verifiedSet.Add($v) } } }

# The remote routing domain is the single accepted *.mail.onmicrosoft.com domain; anything
# else leaves the routing rules unassessed instead of guessing.
$routingCandidates = @($domainSet | Where-Object { $_ -match '^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+mail\.onmicrosoft\.com$' } | Sort-Object -Unique)
$routingDomain = if ($routingCandidates.Count -eq 1) { $routingCandidates[0].ToLowerInvariant() } else { '' }
$routingDecision = if ($null -eq $domains) { 'Accepted domains not available' } elseif ($routingCandidates.Count -eq 1) { 'Single accepted mail.onmicrosoft.com domain' } elseif ($routingCandidates.Count -eq 0) { 'No accepted mail.onmicrosoft.com domain' } else { 'Several accepted mail.onmicrosoft.com domains' }
$allowedTargetSet = New-Set
if ($routingDomain) { [void]$allowedTargetSet.Add($routingDomain) }
# A child process started with -File receives one comma-separated string.
foreach ($domain in @($AllowedTargetDomains | ForEach-Object { ([string]$_) -split ',' })) { $value = (T $domain).TrimStart('@').ToLowerInvariant(); if ($value) { [void]$allowedTargetSet.Add($value) } }

# Duplicate TargetAddress counts distinct accounts (ObjectGUID, else UPN), not repeated rows.
$targetAccounts = @{}
foreach ($r in $ad) {
    $target = K ((T (P $r 'TargetAddress', 'targetAddress')) -replace '(?i)^smtp:', '')
    if (-not $target) { continue }
    $account = K (P $r 'ObjectGUID'); if (-not $account) { $account = 'upn:' + (K (P $r 'UserPrincipalName')) }
    if (-not $targetAccounts.ContainsKey($target)) { $targetAccounts[$target] = New-Set }
    [void]$targetAccounts[$target].Add($account)
}

# Account population. Workforce values map as Human -> Human and
# 'Workforce - review classification' -> Review required. Accounts absent from Workforce use the
# governed AccountType mapping through the same strong keys as Workforce (SID, immutable ID, UPN);
# conflicting AD classifications for one Entra identity remain Review required.
$workforcePopulation = @{}
foreach ($row in $userEvidence) {
    $id = K (P $row 'User Source ID'); if (-not $id) { continue }
    $workforcePopulation[$id] = if ((T (P $row 'Account Population')) -ceq 'Human') { 'Human' } else { 'Review required' }
}
$populationByType = @{}
function Get-TypePopulation([AllowNull()][object]$AccountType) {
    $type = T $AccountType; if (-not $type) { $type = 'Unclassified Account' }
    if (-not $script:populationByType.ContainsKey($type)) { $script:populationByType[$type] = Get-SmartM365AccountPopulation -PopulationRules $populationRules -AccountType $type }
    return $script:populationByType[$type]
}
$adPopulationByEntraId = @{}
foreach ($r in $ad) {
    $population = Get-TypePopulation (P $r 'AccountType')
    foreach ($candidate in @($m365BySid[(K (P $r 'ObjectSID'))], $m365ByImmutable[(K (P $r 'ImmutableId_AD'))], $m365ByUpn[(K (P $r 'UserPrincipalName'))])) {
        if ($null -eq $candidate) { continue }
        $id = K (P $candidate 'Object Id'); if (-not $id) { continue }
        if ($adPopulationByEntraId.ContainsKey($id) -and $adPopulationByEntraId[$id] -ne $population) { $adPopulationByEntraId[$id] = 'Review required' }
        elseif (-not $adPopulationByEntraId.ContainsKey($id)) { $adPopulationByEntraId[$id] = $population }
    }
}

$userMailboxTypes = New-Set; @('UserMailbox', 'RemoteUserMailbox') | ForEach-Object { [void]$userMailboxTypes.Add($_) }
$remoteMailboxTypes = New-Set; @('RemoteUserMailbox', 'RemoteSharedMailbox', 'RemoteRoomMailbox') | ForEach-Object { [void]$remoteMailboxTypes.Add($_) }
$targetRecipientTypes = New-Set; @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'RemoteUserMailbox', 'RemoteSharedMailbox', 'RemoteRoomMailbox') | ForEach-Object { [void]$targetRecipientTypes.Add($_) }

# Plain hashtable: an [ordered] dictionary would treat an integer index as a position.
$rules = @{
    1  = @('HYB-001', 'Identity sync', 'Blocker', 'Sync user account with Azure Entra')
    2  = @('HYB-002', 'Migration capacity', 'High', 'Reduce mailbox item count')
    3  = @('HYB-003', 'Addressing', 'Blocker', 'Add required proxy address')
    6  = @('HYB-006', 'Migration capacity', 'High', 'Review F3/F1 license or reduce mailbox size')
    9  = @('HYB-009', 'Identity sync', 'Blocker', 'Resolve UPN conflict before AAD Connect sync')
    16 = @('HYB-016', 'Addressing', 'Blocker', 'Resolve proxy address conflicts')
    19 = @('HYB-019', 'Routing', 'Blocker', 'Set TargetAddress for mail-enabled user')
    20 = @('HYB-020', 'Routing', 'High', 'Fix TargetAddress format to a valid SMTP address')
    21 = @('HYB-021', 'Routing', 'Blocker', 'Update TargetAddress to an allowed routing domain')
    22 = @('HYB-022', 'Routing', 'Medium', 'Add TargetAddress as a proxy entry in ProxyAddresses')
    23 = @('HYB-023', 'Routing', 'Blocker', 'Resolve duplicate TargetAddress routing')
    24 = @('HYB-024', 'Addressing', 'Blocker', 'Add a primary SMTP entry (SMTP:) in proxyAddresses')
    25 = @('HYB-025', 'Addressing', 'Blocker', 'Keep only one primary SMTP (SMTP:) in proxyAddresses')
    26 = @('HYB-026', 'Addressing', 'High', 'Ensure proxyAddresses contains the PrimarySMTPaddress')
    29 = @('HYB-029', 'Mailbox state', 'Medium', 'Review mailbox type and migration scope for this recipient')
    30 = @('HYB-030', 'Addressing', 'Blocker', 'Fix Primary SMTP domain or add the domain to EXO Accepted Domains')
    31 = @('HYB-031', 'Addressing', 'Medium', 'Move Primary SMTP from onmicrosoft.com to corporate accepted domain')
    32 = @('HYB-032', 'Addressing', 'Info', 'Review why an InternalRelay domain is used as Primary SMTP (routing/coexistence)')
    35 = @('HYB-035', 'Identity sync', 'Blocker', 'Add and verify UPN domain in Azure AD or update user UPN')
    38 = @('HYB-038', 'Mailbox state', 'Blocker', 'Remediate split-brain: ensure a single authoritative mailbox')
    39 = @('HYB-039', 'Identity sync', 'Blocker', 'Remove or replace invalid characters in UPN before AAD Connect sync')
    40 = @('HYB-040', 'Addressing', 'Blocker', 'Remove or replace invalid characters in Primary SMTP address')
    51 = @('HYB-051', 'Addressing', 'High', 'Remove duplicate proxy entries keeping only one instance per address')
    54 = @('HYB-054', 'Mailbox state', 'Blocker', 'Verify if migration is complete or if RemoteMailbox conversion was skipped')
    58 = @('HYB-058', 'License governance', 'Review', 'Review license assignment: non-human accounts should not hold E3, F3 or F1 licenses')
}
$counters = [ordered]@{}
foreach ($number in @($rules.Keys | Sort-Object)) { $counters[$rules[$number][0]] = [ordered]@{LegacyIssueNumber = $number; Evaluated = 0; Triggered = 0; NotAssessed = 0 } }
function Add-Evaluated([int]$Number) { $script:counters[$rules[$Number][0]].Evaluated++ }
function Add-NotAssessed([int]$Number) { $script:counters[$rules[$Number][0]].NotAssessed++ }

$findings = [Collections.Generic.List[object]]::new()
$findingKeys = New-Set
function Add-Finding([int]$Number, [hashtable]$Context, [string]$Observed) {
    $rule = $rules[$Number]
    $accountKey = if ($Context.AdGuid) { $Context.AdGuid } elseif ($Context.Upn) { 'upn:' + $Context.Upn.ToLowerInvariant() } else { 'smtp:' + $Context.Primary.ToLowerInvariant() }
    $key = $rule[0] + '|' + $accountKey
    if (-not $script:findingKeys.Add($key)) { return }
    $script:counters[$rule[0]].Triggered++
    $script:findings.Add([pscustomobject][ordered]@{
        'TenantKey' = $Context.TenantKey
        'Finding Key' = $key
        'AD Object GUID' = $Context.AdGuid
        'Entra Object ID' = $Context.EntraId
        'Exchange GUID' = $Context.ExchangeGuid
        'User Principal Name' = $Context.Upn
        'Primary SMTP Address' = $Context.Primary
        'Recipient Type' = $Context.RecipientType
        'Mailbox Hosting' = $Context.Hosting
        'Rule ID' = $rule[0]
        'Legacy Issue Number' = $Number
        'Rule Family' = $rule[1]
        'Severity' = $rule[2]
        'Is Blocker' = ($rule[2] -eq 'Blocker')
        'Observed Value' = $Observed
        'Recommended Action' = $rule[3]
        'Account Population' = $Context.Population
        'Account Population Source' = $Context.PopulationSource
        'Account Classification Rule Version' = $classificationRuleVersion
        'Rule Version' = $ruleVersion
        'Evidence Date' = $evidenceDate
    })
}

$adTenant = ''
foreach ($u in $ad) {
    $guid = T (P $u 'ObjectGUID'); $upn = T (P $u 'UserPrincipalName'); $upnKey = $upn.ToLowerInvariant(); $immutableId = T (P $u 'ImmutableId_AD'); $domainAndSam = T (P $u 'DomainAndSam')
    if (-not $adTenant) { $adTenant = T (P $u 'TenantKey') }
    $primary = T (P $u 'EmailAddress', 'PrimarySmtpAddress', 'PrimarySMTPaddress')
    if (-not $primary) { foreach ($address in Split-Addresses (P $u 'ProxyAddresses')) { if ($address -match '@') { $primary = $address; break } } }
    $target = T (P $u 'TargetAddress'); $proxyText = T (P $u 'ProxyAddresses')
    $lm = if ($domainAndSam -and $localByDomainSam.ContainsKey((K $domainAndSam))) { $localByDomainSam[(K $domainAndSam)] } elseif ($localByGuid.ContainsKey((K $guid))) { $localByGuid[(K $guid)] } elseif ($localByUpn.ContainsKey($upnKey)) { $localByUpn[$upnKey] } elseif ($localBySmtp.ContainsKey((K $primary))) { $localBySmtp[(K $primary)] } else { $null }
    $rm = if ($remoteByGuid.ContainsKey((K $guid))) { $remoteByGuid[(K $guid)] } elseif ($remoteByUpn.ContainsKey($upnKey)) { $remoteByUpn[$upnKey] } elseif ($remoteBySmtp.ContainsKey((K $primary))) { $remoteBySmtp[(K $primary)] } else { $null }
    $em = if ($exoByImmutable.ContainsKey((K $immutableId))) { $exoByImmutable[(K $immutableId)] } elseif ($exoByUpn.ContainsKey($upnKey)) { $exoByUpn[$upnKey] } elseif ($exoBySmtp.ContainsKey((K $primary))) { $exoBySmtp[(K $primary)] } else { $null }
    $st = if ($statsByUpn.ContainsKey($upnKey)) { $statsByUpn[$upnKey] } else { $null }
    # Recipient type: primary SMTP match on-premises, remote, EXO, then the linked mailbox rows.
    $smtpKey = K $primary
    $recipientType = if ($smtpKey -and $localBySmtp.ContainsKey($smtpKey)) { T (P $localBySmtp[$smtpKey] 'RecipientType', 'RecipientTypeDetails') }
        elseif ($smtpKey -and $remoteBySmtp.ContainsKey($smtpKey)) { T (P $remoteBySmtp[$smtpKey] 'RecipientTypeDetails', 'RecipientType') }
        elseif ($smtpKey -and $exoBySmtp.ContainsKey($smtpKey)) { T (P $exoBySmtp[$smtpKey] 'RecipientTypeDetails', 'MailboxType') }
        elseif (T (P $lm 'RecipientType', 'RecipientTypeDetails')) { T (P $lm 'RecipientType', 'RecipientTypeDetails') }
        elseif (T (P $rm 'RecipientTypeDetails', 'RecipientType')) { T (P $rm 'RecipientTypeDetails', 'RecipientType') }
        elseif (T (P $em 'RecipientTypeDetails', 'MailboxType')) { T (P $em 'RecipientTypeDetails', 'MailboxType') }
        else { 'NoMailboxes' }
    $isMailbox = ($recipientType -ne 'NoMailboxes' -or $lm -or $rm -or $em)
    if (-not ($isMailbox -and $primary)) { continue }

    $m365User = if ($immutableId -and $m365ByImmutable.ContainsKey((K $immutableId))) { $m365ByImmutable[(K $immutableId)] } elseif ($upnKey -and $m365ByUpn.ContainsKey($upnKey)) { $m365ByUpn[$upnKey] } else { $null }
    $inM365 = $null -ne $m365User
    $entraId = T (P $m365User 'Object Id', 'Id', 'UserId')
    $userLicenses = if ($entraId -and $licensesById.ContainsKey((K $entraId))) { @($licensesById[(K $entraId)]) } elseif ($upnKey -and $licensesByUpn.ContainsKey($upnKey)) { @($licensesByUpn[$upnKey]) } else { @() }
    $licenseGroup = Get-LicenseGroup $recipientType $userLicenses $inM365
    $hasAllowedLicense = $licenseGroup -in @('Microsoft 365 E3', 'Microsoft 365 F3', 'Microsoft 365 F1')
    $population = $null; $populationSource = ''
    if ($entraId -and $workforcePopulation.ContainsKey((K $entraId))) { $population = $workforcePopulation[(K $entraId)]; $populationSource = 'Workforce' }
    elseif ($entraId -and $adPopulationByEntraId.ContainsKey((K $entraId))) { $population = $adPopulationByEntraId[(K $entraId)]; $populationSource = 'AD AccountType' }
    else { $population = Get-TypePopulation (P $u 'AccountType'); $populationSource = 'AD AccountType' }
    $hosting = if ($lm) { 'On-premises' } elseif ($rm) { 'Remote' } elseif ($em) { 'Exchange Online' } else { 'None' }
    $exchangeGuid = T (P $lm 'ExchangeGuid'); if (-not $exchangeGuid) { $exchangeGuid = T (P $rm 'ExchangeGuid') }; if (-not $exchangeGuid) { $exchangeGuid = T (P $em 'MailboxGuid') }
    $context = @{TenantKey = $(if (T (P $u 'TenantKey')) { T (P $u 'TenantKey') } else { $adTenant }); AdGuid = $guid; EntraId = $entraId; ExchangeGuid = $exchangeGuid; Upn = $upn; Primary = $primary; RecipientType = $recipientType; Hosting = $hosting; Population = $population; PopulationSource = $populationSource }

    Add-Evaluated 1; if (-not $inM365) { Add-Finding 1 $context 'No matching Entra user (immutable ID or UPN)' }

    # Rule 2: item count from the on-premises mailbox, and from EXO statistics, else the EXO mailbox export.
    $localItems = ConvertTo-KnownNumber (P $lm 'ItemCount', 'TotalItemCount')
    $exoItems = ConvertTo-KnownNumber (P $st 'ItemCount', 'TotalItemCount'); if ($null -eq $exoItems) { $exoItems = ConvertTo-KnownNumber (P $em 'ItemCount') }
    $itemCount = Get-MaxKnown @($localItems, $exoItems)
    if ($null -eq $itemCount) { Add-NotAssessed 2 } else { Add-Evaluated 2; if ($itemCount -gt 999999) { Add-Finding 2 $context ('{0} items' -f [long]$itemCount) } }

    if (-not $routingDomain) { Add-NotAssessed 3 } else { Add-Evaluated 3; if ($proxyText -notmatch ('@' + [regex]::Escape($routingDomain))) { Add-Finding 3 $context ('No proxy address in ' + $routingDomain) } }

    # Rule 6: mailbox size in MB (on-premises MB; EXO GB x 1024), threshold 2000 MB as in the legacy rule.
    if ($licenseGroup -in @('Microsoft 365 F3', 'Microsoft 365 F1')) {
        $exoGb = Get-MaxKnown @((ConvertTo-KnownNumber (P $em 'TotalItemSizeGB')), (ConvertTo-KnownNumber (P $st 'TotalItemSizeGB')))
        $sizeMb = Get-MaxKnown @((ConvertTo-KnownNumber (P $lm 'TotalItemSize-In-MB')), $(if ($null -ne $exoGb) { $exoGb * 1024.0 } else { $null }))
        if ($null -eq $sizeMb) { Add-NotAssessed 6 } else { Add-Evaluated 6; if ($sizeMb -ge 2000) { Add-Finding 6 $context ('{0} MB with {1}' -f [math]::Round($sizeMb, 0), $licenseGroup) } }
    }

    Add-Evaluated 9; if ($upn -and $duplicateUpnSet.Contains($upn)) { Add-Finding 9 $context $upn }

    if ($null -eq $duplicateSmtp) { Add-NotAssessed 16 } else {
        Add-Evaluated 16
        foreach ($address in @(@(Split-Addresses $proxyText) + @($primary))) { if ($duplicateSmtpSet.Contains($address)) { Add-Finding 16 $context $address; break } }
    }

    if ($remoteMailboxTypes.Contains($recipientType)) {
        $normalizedTarget = (T $target) -replace '(?i)^smtp:', ''
        Add-Evaluated 19
        if (-not $normalizedTarget) { Add-Finding 19 $context 'TargetAddress is empty' }
        else {
            Add-Evaluated 20; Add-Evaluated 22; Add-Evaluated 23
            if ($normalizedTarget -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { Add-Finding 20 $context $normalizedTarget }
            $targetDomain = Get-DomainPart $normalizedTarget
            # Configured domains only extend the routing domain; they are not an exhaustive list.
            if (-not $routingDomain) { Add-NotAssessed 21 } else { Add-Evaluated 21; if ($targetDomain -and -not $allowedTargetSet.Contains($targetDomain)) { Add-Finding 21 $context $targetDomain } }
            if (-not $proxyText -or -not (Test-ProxyContains $proxyText $normalizedTarget)) { Add-Finding 22 $context $normalizedTarget }
            $targetKey = $normalizedTarget.ToLowerInvariant()
            if ($targetAccounts.ContainsKey($targetKey) -and $targetAccounts[$targetKey].Count -gt 1) { Add-Finding 23 $context ('{0} ({1} accounts)' -f $normalizedTarget, $targetAccounts[$targetKey].Count) }
        }
    }

    if ($targetRecipientTypes.Contains($recipientType)) {
        $primaryProxyCount = Get-PrimaryProxyCount $proxyText
        Add-Evaluated 24; Add-Evaluated 25; Add-Evaluated 26; Add-Evaluated 51
        if ($primaryProxyCount -eq 0) { Add-Finding 24 $context 'No SMTP: entry' }
        if ($primaryProxyCount -gt 1) { Add-Finding 25 $context ('{0} SMTP: entries' -f $primaryProxyCount) }
        if ($proxyText -and -not (Test-ProxyContains $proxyText $primary)) { Add-Finding 26 $context $primary }
        # Whole proxy entries are compared, so an address is never counted inside a longer one.
        $sameAddress = @(Split-Addresses $proxyText | Where-Object { $_ -ieq $primary }).Count
        if ($sameAddress -gt 1) { Add-Finding 51 $context ('{0} ({1} entries)' -f $primary, $sameAddress) }
    } else {
        Add-Evaluated 29; Add-Finding 29 $context $recipientType
    }

    $primaryDomain = Get-DomainPart $primary
    if ($domainSet.Count -eq 0) { Add-NotAssessed 30; Add-NotAssessed 32 } else {
        Add-Evaluated 30; if ($primaryDomain -and -not $domainSet.Contains($primaryDomain)) { Add-Finding 30 $context $primaryDomain }
        if ($userMailboxTypes.Contains($recipientType)) { Add-Evaluated 32; if ($primaryDomain -and $domainType.ContainsKey($primaryDomain) -and $domainType[$primaryDomain] -eq 'InternalRelay') { Add-Finding 32 $context $primaryDomain } }
    }
    Add-Evaluated 31; if ($primaryDomain -like '*.onmicrosoft.com' -or $primaryDomain -eq 'onmicrosoft.com') { Add-Finding 31 $context $primaryDomain }
    $upnDomain = Get-DomainPart $upn
    if ($verifiedSet.Count -eq 0) { Add-NotAssessed 35 } else { Add-Evaluated 35; if ($upnDomain -and -not $verifiedSet.Contains($upnDomain)) { Add-Finding 35 $context $upnDomain } }
    Add-Evaluated 38; if ($lm -and $em -and $recipientType -eq 'UserMailbox') { Add-Finding 38 $context 'On-premises UserMailbox and EXO mailbox' }
    Add-Evaluated 39; if ($upn -and (Test-InvalidUpnCharacter $upn)) { Add-Finding 39 $context $upn }
    Add-Evaluated 40; if (Test-InvalidSmtpLocalCharacter $primary) { Add-Finding 40 $context $primary }
    if ($hasAllowedLicense) { Add-Evaluated 58; if ($population -eq 'Non-human') { Add-Finding 58 $context ('{0}; {1}' -f (T (P $u 'AccountType')), $licenseGroup) } }
}

# Rule 54: EXO mailbox with an AD account but no on-premises mailbox or remote mailbox.
$adByUpn = @{}; $adBySmtp = @{}; $adByImmutable = @{}
foreach ($r in $ad) { Add-First $adByUpn (P $r 'UserPrincipalName') $r; Add-First $adBySmtp (P $r 'EmailAddress', 'PrimarySmtpAddress', 'PrimarySMTPaddress') $r; Add-First $adByImmutable (P $r 'ImmutableId_AD') $r }
foreach ($e in $exo) {
    $upn = T (P $e 'UserPrincipalName'); $smtp = T (P $e 'PrimarySmtpAddress', 'PrimarySMTPaddress'); $immutableId = K (P $e 'OnPremisesImmutableId', 'ImmutableId')
    Add-Evaluated 54
    if ($localByUpn.ContainsKey((K $upn)) -or $remoteByUpn.ContainsKey((K $upn)) -or $localBySmtp.ContainsKey((K $smtp)) -or $remoteBySmtp.ContainsKey((K $smtp))) { continue }
    $adRow = if ($immutableId -and $adByImmutable.ContainsKey($immutableId)) { $adByImmutable[$immutableId] } elseif ($adByUpn.ContainsKey((K $upn))) { $adByUpn[(K $upn)] } elseif ($adBySmtp.ContainsKey((K $smtp))) { $adBySmtp[(K $smtp)] } else { $null }
    if ($null -eq $adRow -or -not (T (P $adRow 'ObjectGUID'))) { continue }
    $m365User = if ($immutableId -and $m365ByImmutable.ContainsKey($immutableId)) { $m365ByImmutable[$immutableId] } elseif ($m365ByUpn.ContainsKey((K $upn))) { $m365ByUpn[(K $upn)] } else { $null }
    $context = @{TenantKey = $(if (T (P $adRow 'TenantKey')) { T (P $adRow 'TenantKey') } else { $adTenant }); AdGuid = T (P $adRow 'ObjectGUID'); EntraId = T (P $m365User 'Object Id'); ExchangeGuid = T (P $e 'MailboxGuid'); Upn = $upn; Primary = $smtp; RecipientType = T (P $e 'RecipientTypeDetails', 'MailboxType'); Hosting = 'Exchange Online'; Population = ''; PopulationSource = '' }
    Add-Finding 54 $context 'No on-premises mailbox or remote mailbox'
}

$columns = @('TenantKey', 'Finding Key', 'AD Object GUID', 'Entra Object ID', 'Exchange GUID', 'User Principal Name', 'Primary SMTP Address', 'Recipient Type', 'Mailbox Hosting', 'Rule ID', 'Legacy Issue Number', 'Rule Family', 'Severity', 'Is Blocker', 'Observed Value', 'Recommended Action', 'Account Population', 'Account Population Source', 'Account Classification Rule Version', 'Rule Version', 'Evidence Date')
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
if ($findings.Count -gt 0) {
    $findings | Sort-Object 'Rule ID', 'Finding Key' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM
} else {
    [IO.File]::WriteAllText($OutputPath, (($columns | ForEach-Object { '"' + $_ + '"' }) -join ',') + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
}
foreach ($id in $counters.Keys) { $c = $counters[$id]; Write-Host ('Hybrid identity {0} (legacy {1}): evaluated={2}; triggered={3}; not assessed={4}' -f $id, $c.LegacyIssueNumber, $c.Evaluated, $c.Triggered, $c.NotAssessed) }
Write-Host ('Hybrid identity routing domain: {0}; allowed TargetAddress domains: {1}.' -f $routingDecision, $allowedTargetSet.Count)
if ($AuditPath) {
    $audit = [ordered]@{
        RuleVersion = $ruleVersion; AccountClassificationRuleVersion = $classificationRuleVersion; EvidenceDate = $evidenceDate
        RoutingDomainDecision = $routingDecision; AllowedTargetDomainCount = $allowedTargetSet.Count
        SourceSHA256 = $script:sourceFiles; UserEvidenceSHA256 = (Get-FileHash -LiteralPath $UserEvidencePath -Algorithm SHA256).Hash
        Rows = $findings.Count; OutputSHA256 = (Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256).Hash; Rules = $counters
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent ([IO.Path]::GetFullPath($AuditPath))) -Force | Out-Null
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($AuditPath), ($audit | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCftFDrVQCLeN3t
# IYtRSSnt5aRItUdJV4sxhJc9PUjASqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINvGSShOZMIvyeXUI5POQmko8Xt7SmjbCa0wle+LipkGMA0GCSqG
# SIb3DQEBAQUABIIBgGmDzISFbWQ/U0h1CXdo21YY9Zcn4HRUt7T2ZSrZXcFeWQqc
# gZvW/g8rI49nKDV++smQjs7PnYFHuOXoDU9St7uIEsPxsYI5nkehz16bvrDSer44
# 12kV2SdHSuLUW0SuWqkSnMvHvrTSowgG4DEZbrEq2EyKsDv6JTkDxtG9VYRLeQIQ
# Ih83tpetL6cNGD3/ZbPx4RxLLVV4NqovJsUNi9DAa6JFVgU6cruv3F0rAFNtkXcf
# SCBJH/qQ/JGh2uBlzQVvNnuLJ4jUXY6iblVQ26T+TeElcXolmgwvz2SC1OUSRtY6
# /c8SbqhQumti1GgzdQxB3Z4GGzXjp8XIudimaN+5a9yhpYfzKkalgX3cLDY6kINT
# iahmdEaRRMoJNHMtKjhT94xnWQJj7QfT99TSn5xGWUjZy4EvQsG4kAFzgCQvtZBN
# fpUfdh8sY/zc7hKJl6uREnohjMF/InWXf1Rkls6Ykfb2+d2qE7GesUdB2Fn3qP++
# jHeCqwnN+MBEVwFe2qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMjU3
# NTRaMC8GCSqGSIb3DQEJBDEiBCCeXlt0CKxCSh9ixFdvyxRSPmyggx54YJTeOjD0
# EcJu4DANBgkqhkiG9w0BAQEFAASCAgCUBQk9UsgS0tOATW6nTFgxmYEgo5R2RUvq
# Uyi5mk/1tJUYbj4QP5MKUQCo2iJZafa99wfRmCg4bR59CrDFUaktDhZZxgr76SuB
# oZ2kET3XTNgOutqYKqdTmTFoArzbg4f5O72J105jFR0jVhhHtXbfwFvPsgvGtTpW
# H/SovmPAETJPJmsE21NTfAeQjopLaQQ5UqvPBJ8ngodnn16uMpSKTiDemRaQQRbZ
# RQkdNSTl4NvqkZO+VJZVE9TvWeypJ0a7KXH686QTLIoQpJVGyHxr8tO/WbpQ2D5p
# sQGNX5s5GxffNkzg+AN0GO42LCiHGFuVXCoi0BJ4iuQwVbIAGqtwqf4pgkHEAbcA
# 0hbfU5bSaFLOYxILYvHVSuDNXl870SF1AbO42CY6Gm3t6wzI9vgD+SDM1XVxGV+B
# 3i++p176pjFXQdalKlVrEWS/vgWBS1UlCmCbVOQg74HiLi6SyR0yBCa6lY5EBNeH
# XFOCJOkn0RLOWCA346BJD91sHjrU0SWv5ZW7DeEXbr2XCA+k3R6hVIGTfedlWUWU
# tx2pqvqdz8uEPDRIN8fpON23dFYTZ0KZIfneSs2btiOEqXesdi0VndCwCrsofhjM
# AGzNk3BoW1stOAwwHyGnjMiyAWd4hEnMcVHXruHGk3Flm/T6JZApFKhkQyBYnTr8
# oNLPAVSWQA==
# SIG # End signature block
