Set-StrictMode -Version 2.0

function ConvertTo-SmartM365CommunicationHashtable {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    $hash = @{}
    if ($null -eq $InputObject) { return $hash }

    foreach ($property in $InputObject.PSObject.Properties) {
        $hash[$property.Name] = $property.Value
    }
    return $hash
}

function Read-SmartM365CommunicationJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Required) { throw "Configuration file not found: $Path" }
        return @{}
    }

    try {
        return ConvertTo-SmartM365CommunicationHashtable -InputObject (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        throw ("Failed to read configuration file '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Merge-SmartM365CommunicationConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$BaseConfig,
        [Parameter(Mandatory)][hashtable]$OverlayConfig
    )

    $merged = @{}
    foreach ($key in $BaseConfig.Keys) { $merged[$key] = $BaseConfig[$key] }
    foreach ($key in $OverlayConfig.Keys) { $merged[$key] = $OverlayConfig[$key] }
    return $merged
}

function Get-SmartM365CommunicationConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null,
        [AllowNull()][hashtable]$FallbackConfig = $null
    )

    $value = $null
    $hasValue = $false
    if ($Config.Contains($Name)) {
        $value = $Config[$Name]
        $hasValue = $true
    }

    if ($hasValue) {
        $text = if ($null -eq $value) { '' } else { [string]$value }
        if ($text -ne '__USE_GLOBAL__' -and $text -ne 'USE_GLOBAL') {
            return $value
        }
    }

    if ($FallbackConfig -and $FallbackConfig.Contains($Name)) {
        return $FallbackConfig[$Name]
    }

    return $DefaultValue
}

function Resolve-SmartM365CommunicationIPv4Address {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$HostName)

    $addresses = [System.Net.Dns]::GetHostAddresses($HostName) | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork }
    $first = $addresses | Select-Object -First 1
    if (-not $first) { throw "No IPv4 address found for $HostName" }
    return $first.IPAddressToString
}

function Initialize-SmartM365CommunicationExchangeSnapIn {
    [CmdletBinding()]
    param(
        [bool]$Enabled = $false,
        [string]$SnapInName = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
        [bool]$ViewEntireForest = $true
    )

    if (-not $Enabled) {
        return [pscustomobject]@{
            Enabled = $false
            Available = $false
            ViewEntireForestRequested = $ViewEntireForest
            ViewEntireForestApplied = $false
            Status = 'Disabled'
            ErrorMessage = ''
            ForestErrorMessage = ''
        }
    }

    if (-not (Get-Command -Name Get-PSSnapin -ErrorAction SilentlyContinue) -or
        -not (Get-Command -Name Add-PSSnapin -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{
            Enabled = $true
            Available = $false
            ViewEntireForestRequested = $ViewEntireForest
            ViewEntireForestApplied = $false
            Status = 'PSSnapinCommandsUnavailable'
            ErrorMessage = 'PSSnapin commands are not available in this PowerShell session.'
            ForestErrorMessage = ''
        }
    }

    try {
        if (-not (Get-PSSnapin -Name $SnapInName -ErrorAction SilentlyContinue)) {
            Add-PSSnapin $SnapInName -ErrorAction Stop
        }

        $forestApplied = $false
        $forestError = ''
        if ($ViewEntireForest) {
            try {
                Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
                $forestApplied = $true
            }
            catch {
                $forestError = $_.Exception.Message
            }
        }

        return [pscustomobject]@{
            Enabled = $true
            Available = $true
            ViewEntireForestRequested = $ViewEntireForest
            ViewEntireForestApplied = $forestApplied
            Status = $(if ($forestError) { 'AvailableForestViewFailed' } else { 'Available' })
            ErrorMessage = ''
            ForestErrorMessage = $forestError
        }
    }
    catch {
        return [pscustomobject]@{
            Enabled = $true
            Available = $false
            ViewEntireForestRequested = $ViewEntireForest
            ViewEntireForestApplied = $false
            Status = 'Unavailable'
            ErrorMessage = $_.Exception.Message
            ForestErrorMessage = ''
        }
    }
}

function ConvertTo-SmartM365CommunicationLdapFilterSafe {
    [CmdletBinding()]
    param([AllowNull()][string]$InputString)

    if ($null -eq $InputString) { return '' }
    $safe = [string]$InputString
    $safe = $safe -replace '\\', '\5c'
    $safe = $safe -replace '\*', '\2a'
    $safe = $safe -replace '\(', '\28'
    $safe = $safe -replace '\)', '\29'
    $safe = $safe -replace "`0", '\00'
    return $safe
}

function Resolve-SmartM365CommunicationAdUserInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SmtpAddress,
        [string]$GcServer = '',
        [string]$DefaultLanguageTag = '',
        [switch]$ResolveDisplayName,
        [switch]$ResolveLanguage
    )

    $result = [pscustomobject]@{
        Found = $false
        DisplayName = ''
        GivenName = ''
        Surname = ''
        PreferredLanguage = ''
        Source = ''
        ErrorMessage = ''
    }

    if ([string]::IsNullOrWhiteSpace($SmtpAddress)) { return $result }

    $safe = ConvertTo-SmartM365CommunicationLdapFilterSafe -InputString $SmtpAddress
    $filter = "(|(mail=$safe)(userPrincipalName=$safe)(proxyAddresses=SMTP:$safe)(proxyAddresses=smtp:$safe))"

    if (Get-Command -Name Get-ADUser -ErrorAction SilentlyContinue) {
        try {
            $server = if ([string]::IsNullOrWhiteSpace($GcServer)) { $null } else { "$GcServer`:3268" }
            $properties = @('displayName', 'givenName', 'sn', 'preferredLanguage')
            $user = if ($server) {
                Get-ADUser -LDAPFilter $filter -Server $server -Properties $properties -ErrorAction Stop | Select-Object -First 1
            }
            else {
                Get-ADUser -LDAPFilter $filter -Properties $properties -ErrorAction Stop | Select-Object -First 1
            }

            if ($user) {
                $result.Found = $true
                $result.DisplayName = [string]$user.DisplayName
                $result.GivenName = [string]$user.GivenName
                $result.Surname = [string]$user.Surname
                $result.PreferredLanguage = [string]$user.PreferredLanguage
                $result.Source = 'Get-ADUser'
                return $result
            }
        }
        catch {
            $result.ErrorMessage = $_.Exception.Message
        }
    }

    try {
        $root = if ([string]::IsNullOrWhiteSpace($GcServer)) { 'GC://' } else { "GC://$GcServer" }
        $entry = New-Object System.DirectoryServices.DirectoryEntry($root)
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)$filter)"
        $searcher.PageSize = 1000
        foreach ($property in @('displayName', 'givenName', 'sn', 'preferredLanguage')) {
            [void]$searcher.PropertiesToLoad.Add($property)
        }

        $found = $searcher.FindOne()
        if ($found) {
            $result.Found = $true
            if ($found.Properties['displayname']) { $result.DisplayName = [string]$found.Properties['displayname'][0] }
            if ($found.Properties['givenname']) { $result.GivenName = [string]$found.Properties['givenname'][0] }
            if ($found.Properties['sn']) { $result.Surname = [string]$found.Properties['sn'][0] }
            if ($found.Properties['preferredlanguage']) { $result.PreferredLanguage = [string]$found.Properties['preferredlanguage'][0] }
            $result.Source = 'DirectorySearcher'
        }
    }
    catch {
        if ([string]::IsNullOrWhiteSpace($result.ErrorMessage)) { $result.ErrorMessage = $_.Exception.Message }
    }

    if ($ResolveLanguage -and [string]::IsNullOrWhiteSpace($result.PreferredLanguage)) {
        $result.PreferredLanguage = $DefaultLanguageTag
    }

    if ($ResolveDisplayName -and [string]::IsNullOrWhiteSpace($result.DisplayName)) {
        if (-not [string]::IsNullOrWhiteSpace(($result.GivenName + $result.Surname).Trim())) {
            $result.DisplayName = "$($result.GivenName) $($result.Surname)".Trim()
        }
        else {
            $result.DisplayName = ($SmtpAddress -split '@')[0]
        }
    }

    return $result
}

function Test-SmartM365CommunicationExchangeMailboxState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SmtpAddress)

    $state = [pscustomobject]@{
        Exists = $false
        IsRemoteMailbox = $false
        RecipientTypeDetails = ''
        DisplayName = ''
        PrimarySmtpAddress = ''
        Source = ''
        ErrorMessage = ''
    }

    function Set-ExchangeMailboxStateFromRecipient {
        param($Recipient, [string]$Source, [bool]$IsRemote)
        $state.Exists = $true
        $state.RecipientTypeDetails = [string]$Recipient.RecipientTypeDetails
        $state.DisplayName = [string]$Recipient.DisplayName
        try { $state.PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress } catch {}
        $state.IsRemoteMailbox = $IsRemote
        $state.Source = $Source
        return $state
    }

    foreach ($attempt in @(
        @{ Command = 'Get-Recipient'; Mode = 'Identity'; Remote = $null },
        @{ Command = 'Get-Recipient'; Mode = 'Filter'; Remote = $null },
        @{ Command = 'Get-RemoteMailbox'; Mode = 'Identity'; Remote = $true },
        @{ Command = 'Get-RemoteMailbox'; Mode = 'Filter'; Remote = $true },
        @{ Command = 'Get-Mailbox'; Mode = 'Identity'; Remote = $false }
    )) {
        if (-not (Get-Command -Name $attempt.Command -ErrorAction SilentlyContinue)) { continue }

        try {
            $recipient = $null
            if ($attempt.Mode -eq 'Identity') {
                $recipient = & $attempt.Command -Identity $SmtpAddress -ResultSize 1 -ReadFromDomainController:$true -ErrorAction Stop
            }
            else {
                $filter = "EmailAddresses -eq 'SMTP:$SmtpAddress' -or EmailAddresses -eq 'smtp:$SmtpAddress' -or PrimarySmtpAddress -eq '$SmtpAddress'"
                $recipient = & $attempt.Command -Filter $filter -ResultSize 1 -ReadFromDomainController:$true -ErrorAction Stop
            }

            if ($recipient) {
                $remote = if ($null -ne $attempt.Remote) {
                    [bool]$attempt.Remote
                }
                else {
                    ([string]$recipient.RecipientTypeDetails -like '*RemoteUserMailbox*' -or [string]$recipient.RecipientTypeDetails -like '*RemoteMailbox*')
                }
                return Set-ExchangeMailboxStateFromRecipient -Recipient $recipient -Source "$($attempt.Command)($($attempt.Mode))" -IsRemote $remote
            }
        }
        catch {
            $state.ErrorMessage = $_.Exception.Message
        }
    }

    return $state
}

function ConvertTo-SmartM365CommunicationBytes {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    try {
        if ($null -eq $Value) { return $null }
        if ($Value -is [int64] -or $Value -is [double] -or $Value -is [int]) { return [int64]$Value }
        if ($Value -is [string] -and $Value -match '\((?<bytes>[0-9,\.]+)\s*bytes\)') {
            return [int64](($Matches['bytes'] -replace '[,\.]', ''))
        }
        if ($null -ne $Value.Value -and ($Value.Value | Get-Member -Name ToBytes -ErrorAction SilentlyContinue)) {
            return [int64]$Value.Value.ToBytes()
        }
        if ($Value -and ($Value | Get-Member -Name ToBytes -ErrorAction SilentlyContinue)) {
            return [int64]$Value.ToBytes()
        }
    }
    catch {}
    return $null
}

function Get-SmartM365CommunicationMailboxUsageInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SmtpAddress,
        [double]$MailboxMaxGb = 100
    )

    $result = [pscustomobject]@{
        Available = $false
        TotalBytes = $null
        TotalGb = $null
        Percent = $null
        ErrorMessage = ''
    }

    if (-not (Get-Command -Name Get-MailboxStatistics -ErrorAction SilentlyContinue)) {
        $result.ErrorMessage = 'Get-MailboxStatistics is not available.'
        return $result
    }

    try {
        $stats = Get-MailboxStatistics -Identity $SmtpAddress -ErrorAction Stop
        $totalBytes = ConvertTo-SmartM365CommunicationBytes -Value $stats.TotalItemSize
        if ($null -eq $totalBytes) {
            $result.ErrorMessage = 'TotalItemSize could not be converted to bytes.'
            return $result
        }

        $totalGb = [math]::Round(($totalBytes / 1GB), 2)
        if ($MailboxMaxGb -le 0) { $MailboxMaxGb = 100 }

        $result.Available = $true
        $result.TotalBytes = [int64]$totalBytes
        $result.TotalGb = [double]$totalGb
        $result.Percent = [math]::Round((100.0 * $totalGb / [double]$MailboxMaxGb), 1)
        return $result
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        return $result
    }
}

function Resolve-SmartM365CommunicationExchangeLanguageTag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SmtpAddress,
        [string]$DefaultLanguageTag = '',
        [hashtable]$DomainLanguageMap = @{}
    )

    $result = [pscustomobject]@{
        LanguageTag = ''
        Source = ''
        DistinguishedName = ''
        DomainName = ''
        ErrorMessage = ''
    }

    if (Get-Command -Name Get-MailboxRegionalConfiguration -ErrorAction SilentlyContinue) {
        try {
            $regional = Get-MailboxRegionalConfiguration -Identity $SmtpAddress -ErrorAction Stop
            if ($regional -and $regional.Language) {
                $language = $regional.Language | Select-Object -First 1
                if ($language -and -not [string]::IsNullOrWhiteSpace([string]$language.Name)) {
                    $result.LanguageTag = [string]$language.Name
                    $result.Source = 'Get-MailboxRegionalConfiguration'
                    return $result
                }
            }
        }
        catch {
            $result.ErrorMessage = $_.Exception.Message
        }
    }

    $mailbox = $null
    foreach ($command in @('Get-Mailbox', 'Get-RemoteMailbox')) {
        if (-not (Get-Command -Name $command -ErrorAction SilentlyContinue)) { continue }
        try {
            $mailbox = & $command -Identity $SmtpAddress -ErrorAction Stop
            if ($mailbox) { break }
        }
        catch {
            if ([string]::IsNullOrWhiteSpace($result.ErrorMessage)) { $result.ErrorMessage = $_.Exception.Message }
        }
    }

    if ($mailbox -and $mailbox.DistinguishedName -match '(?i)(DC=.+)$') {
        $dn = [string]$mailbox.DistinguishedName
        $domainFromDn = ($Matches[1] -replace '(?i)DC=', '' -replace ',', '.').Trim('.')
        $result.DistinguishedName = $dn
        $result.DomainName = $domainFromDn

        try {
            $adLang = ''
            if (Get-Command -Name Get-ADUser -ErrorAction SilentlyContinue) {
                $user = Get-ADUser -Identity $dn -Server $domainFromDn -Properties preferredLanguage -ErrorAction Stop
                if ($user) { $adLang = [string]$user.preferredLanguage }
            }
            else {
                $entry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domainFromDn/$dn")
                $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
                $searcher.Filter = '(objectClass=user)'
                [void]$searcher.PropertiesToLoad.Add('preferredLanguage')
                $found = $searcher.FindOne()
                if ($found -and $found.Properties['preferredlanguage']) { $adLang = [string]$found.Properties['preferredlanguage'][0] }
            }

            if (-not [string]::IsNullOrWhiteSpace($adLang)) {
                $result.LanguageTag = $adLang
                $result.Source = 'MailboxDNPreferredLanguage'
                return $result
            }
        }
        catch {
            if ([string]::IsNullOrWhiteSpace($result.ErrorMessage)) { $result.ErrorMessage = $_.Exception.Message }
        }

        if ($domainFromDn -match '^([a-z]{2})\.') {
            $countryCode = $Matches[1].ToLowerInvariant()
            if ($DomainLanguageMap.Contains($countryCode)) {
                $result.LanguageTag = [string]$DomainLanguageMap[$countryCode]
                $result.Source = 'MailboxDNCountry'
                return $result
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($DefaultLanguageTag)) {
        $result.LanguageTag = $DefaultLanguageTag
        $result.Source = 'Default'
    }
    return $result
}

function Get-SmartM365CommunicationLinkedLogoTokens {
    [CmdletBinding()]
    param(
        [string]$LogoPath = '',
        [string]$LogoContentId = 'smartm365logo',
        [string]$LogoAltText = 'SmartM365'
    )

    $logoImgTag = if (-not [string]::IsNullOrWhiteSpace($LogoPath) -and (Test-Path -LiteralPath $LogoPath)) {
        "<img src=""cid:$LogoContentId"" alt=""$LogoAltText"" style=""height:40px; display:block;"" />"
    }
    else {
        ''
    }

    return @{
        LogoImgTag = $logoImgTag
        FooterLogoImgTag = $logoImgTag
    }
}

function Resolve-SmartM365CommunicationTokenizedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory)][hashtable]$Tokens
    )

    if ($null -eq $Value) { return $null }
    $resolved = [string]$Value
    foreach ($key in $Tokens.Keys) {
        $resolved = $resolved.Replace("{{$key}}", [string]$Tokens[$key])
    }

    return $resolved
}

function Import-SmartM365CommunicationCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string[]]$RequiredColumns = @()
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "CSV file not found: $Path" }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
    $raw = $raw -replace '^\uFEFF', ''

    $firstLine = ($raw -split "(`r`n|`n|`r)") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($firstLine)) { throw "CSV file is empty: $Path" }

    $semiCount = ([regex]::Matches($firstLine, ';')).Count
    $commaCount = ([regex]::Matches($firstLine, ',')).Count
    $delimiter = if ($semiCount -gt $commaCount) { ';' } else { ',' }

    $rows = @($raw | ConvertFrom-Csv -Delimiter $delimiter)
    if ($rows.Count -eq 0) { return @() }

    $headers = ($rows | Select-Object -First 1).PSObject.Properties.Name
    foreach ($column in $RequiredColumns) {
        if ($headers -notcontains $column) {
            throw "CSV '$Path' is missing required column '$column'."
        }
    }

    return $rows
}

function Get-SmartM365CommunicationProperty {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string[]]$Names,
        [AllowNull()]$DefaultValue = $null
    )

    if ($null -eq $InputObject) { return $DefaultValue }
    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return $property.Value
        }
    }
    return $DefaultValue
}

function Get-SmartM365CommunicationSmtpAddress {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Row)

    $value = Get-SmartM365CommunicationProperty -InputObject $Row -Names @('EmailAddress','PrimarySmtpAddress','PrimarySMTPAddress','primarysmtp','Mail','UserPrincipalName','UPN') -DefaultValue ''
    return ([string]$value).Trim().ToLowerInvariant()
}

function Get-SmartM365CommunicationTemplateContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TemplateRoot,
        [Parameter(Mandatory)][string]$TemplateBaseName,
        [Parameter(Mandatory)][string]$LanguageTag,
        [string]$DefaultLanguageTag = 'en'
    )

    $tag = if ([string]::IsNullOrWhiteSpace($LanguageTag)) { $DefaultLanguageTag } else { $LanguageTag }
    $neutral = ($tag -split '[-_]')[0].ToLowerInvariant()
    $defaultNeutral = ($DefaultLanguageTag -split '[-_]')[0].ToLowerInvariant()

    $candidates = @(
        (Join-Path -Path $TemplateRoot -ChildPath ("{0}.{1}.html" -f $TemplateBaseName, $tag)),
        (Join-Path -Path $TemplateRoot -ChildPath ("{0}.{1}.html" -f $TemplateBaseName, $neutral)),
        (Join-Path -Path $TemplateRoot -ChildPath ("{0}.{1}.html" -f $TemplateBaseName, $DefaultLanguageTag)),
        (Join-Path -Path $TemplateRoot -ChildPath ("{0}.{1}.html" -f $TemplateBaseName, $defaultNeutral)),
        (Join-Path -Path $TemplateRoot -ChildPath ("{0}.html" -f $TemplateBaseName))
    ) | Select-Object -Unique

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return [pscustomobject]@{
                Path    = $candidate
                Content = Get-Content -LiteralPath $candidate -Raw -Encoding UTF8 -ErrorAction Stop
            }
        }
    }

    throw "No template found for '$LanguageTag'. Tried: $($candidates -join ', ')"
}

function Expand-SmartM365CommunicationTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TemplateContent,
        [Parameter(Mandatory)][hashtable]$Tokens
    )

    $html = $TemplateContent
    foreach ($key in $Tokens.Keys) {
        $pattern = [regex]::Escape("{{$key}}")
        $value = [string]$Tokens[$key]
        $html = [regex]::Replace($html, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $value })
    }

    return $html
}

function Assert-SmartM365CommunicationNoUnresolvedToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Html)

    $matches = [regex]::Matches($Html, '\{\{[^}]+\}\}')
    if ($matches.Count -gt 0) {
        $tokens = @($matches | ForEach-Object { $_.Value } | Select-Object -Unique)
        throw "Unresolved template token(s): $($tokens -join ', ')"
    }
}

function Load-SmartM365CommunicationSentRegistry {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $map = @{}
    if (-not (Test-Path -LiteralPath $Path)) { return $map }

    foreach ($row in (Import-SmartM365CommunicationCsv -Path $Path)) {
        $email = ([string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('Email') -DefaultValue '')).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($email)) { continue }
        $count = 0
        [void][int]::TryParse([string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('Count') -DefaultValue 0), [ref]$count)
        $map[$email] = [pscustomobject]@{
            Email       = [string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('Email') -DefaultValue $email)
            FirstSentOn = [string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('FirstSentOn') -DefaultValue '')
            LastSentOn  = [string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('LastSentOn') -DefaultValue '')
            Count       = $count
        }
    }

    return $map
}

function Save-SmartM365CommunicationSentRegistry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Registry,
        [Parameter(Mandatory)][string]$Path
    )

    $directory = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $directory)) { New-Item -Path $directory -ItemType Directory -Force | Out-Null }

    $rows = foreach ($key in ($Registry.Keys | Sort-Object)) {
        $entry = $Registry[$key]
        [pscustomobject]@{
            Email       = $entry.Email
            FirstSentOn = $entry.FirstSentOn
            LastSentOn  = $entry.LastSentOn
            Count       = $entry.Count
        }
    }

    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Register-SmartM365CommunicationSentItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Registry,
        [Parameter(Mandatory)][string]$Email
    )

    $key = $Email.Trim().ToLowerInvariant()
    $now = [datetime]::UtcNow.ToString('o')
    if ($Registry.ContainsKey($key)) {
        $Registry[$key].LastSentOn = $now
        $Registry[$key].Count = [int]$Registry[$key].Count + 1
    }
    else {
        $Registry[$key] = [pscustomobject]@{
            Email       = $Email
            FirstSentOn = $now
            LastSentOn  = $now
            Count       = 1
        }
    }
}

function Resolve-SmartM365CommunicationLanguageTag {
    [CmdletBinding()]
    param(
        [AllowNull()]$Row,
        [string]$ForceLanguage = '',
        [string]$DefaultLanguageTag = 'en',
        [hashtable]$DomainLanguageMap = @{}
    )

    if (-not [string]::IsNullOrWhiteSpace($ForceLanguage)) { return $ForceLanguage }

    $rowLanguage = [string](Get-SmartM365CommunicationProperty -InputObject $Row -Names @('LanguageTag','PreferredLanguage','preferredLanguage','Language','Lang') -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($rowLanguage)) { return $rowLanguage }

    $email = Get-SmartM365CommunicationSmtpAddress -Row $Row
    if ($email -match '@(?<domain>.+)$') {
        $domain = $Matches['domain'].ToLowerInvariant()
        foreach ($suffix in $DomainLanguageMap.Keys) {
            if ($domain.EndsWith(([string]$suffix).ToLowerInvariant())) {
                return [string]$DomainLanguageMap[$suffix]
            }
        }
    }

    return $DefaultLanguageTag
}

function Get-SmartM365CommunicationSubject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$SubjectByLanguage,
        [Parameter(Mandatory)][string]$LanguageTag,
        [string]$DefaultSubject = 'SmartM365'
    )

    $tag = $LanguageTag.Trim()
    $neutral = ($tag -split '[-_]')[0].ToLowerInvariant()
    if ($SubjectByLanguage.Contains($tag)) { return [string]$SubjectByLanguage[$tag] }
    if ($SubjectByLanguage.Contains($neutral)) { return [string]$SubjectByLanguage[$neutral] }
    if ($SubjectByLanguage.Contains('default')) { return [string]$SubjectByLanguage['default'] }
    return $DefaultSubject
}

function Get-SmartM365CommunicationHotline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$HotlineByLanguageOrCountry,
        [Parameter(Mandatory)][string]$LanguageTag,
        [string]$DefaultHotline = ''
    )

    $tag = $LanguageTag.Trim()
    $neutral = ($tag -split '[-_]')[0].ToLowerInvariant()
    $country = $neutral
    if ($tag -match '^[a-zA-Z]{2}[-_](?<region>[a-zA-Z]{2})$') { $country = $Matches['region'].ToLowerInvariant() }

    foreach ($key in @($tag, $neutral, $country, 'default')) {
        if ($HotlineByLanguageOrCountry.Contains($key)) { return [string]$HotlineByLanguageOrCountry[$key] }
    }

    return $DefaultHotline
}

function Add-SmartM365CommunicationLogRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)][string]$Path,
        [string[]]$Columns = @()
    )

    $directory = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $directory)) { New-Item -Path $directory -ItemType Directory -Force | Out-Null }

    if ($Columns -and $Columns.Count -gt 0) {
        $Row | Select-Object -Property $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Append:(Test-Path -LiteralPath $Path) -Encoding UTF8
    }
    else {
        $Row | Export-Csv -LiteralPath $Path -NoTypeInformation -Append:(Test-Path -LiteralPath $Path) -Encoding UTF8
    }
}

function ConvertTo-SmartM365GraphEmailRecipients {
    [CmdletBinding()]
    param([string]$Recipients)

    $items = @()
    foreach ($recipient in ($Recipients -split '[;,]')) {
        $trimmed = $recipient.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        $items += @{ emailAddress = @{ address = $trimmed } }
    }
    return $items
}

function Send-SmartM365CommunicationMail {
    [CmdletBinding()]
    param(
        [string]$SmtpServer = '',
        [int]$SmtpPort = 25,
        [string]$From,
        [Parameter(Mandatory)][string]$To,
        [string]$Cc = '',
        [string]$Bcc = '',
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyHtml,
        [string]$AppId = '',
        [string]$TenantId = '',
        [string]$Thumbprint = '',
        [bool]$SmtpUseIntegratedAuth = $false,
        [bool]$SmtpEnableSsl = $false,
        [string]$LogoPath = '',
        [string]$LogoContentId = 'smartm365logo',
        [string]$LogoMediaType = 'image/x-icon',
        [int]$RetryCount = 1,
        [int]$RetryDelaySeconds = 2,
        [switch]$WhatIf
    )

    if ($WhatIf) { return [pscustomobject]@{ Sent = $false; Mode = 'DryRun' } }
    if ([string]::IsNullOrWhiteSpace($From)) { throw 'Mail sender (From) is required.' }

    if ([string]::IsNullOrWhiteSpace($SmtpServer)) {
        if (-not (Get-Command -Name Connect-SmartM365GraphAppOnly -ErrorAction SilentlyContinue)) {
            throw 'Connect-SmartM365GraphAppOnly is not available. Import SmartM365.Core first.'
        }
        if (-not (Connect-SmartM365GraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'Communications mail')) {
            throw 'Microsoft Graph app-only connection failed.'
        }

        $message = @{
            subject = $Subject
            body = @{
                contentType = 'HTML'
                content = $BodyHtml
            }
            toRecipients = @(ConvertTo-SmartM365GraphEmailRecipients -Recipients $To)
        }
        $ccRecipients = @(ConvertTo-SmartM365GraphEmailRecipients -Recipients $Cc)
        $bccRecipients = @(ConvertTo-SmartM365GraphEmailRecipients -Recipients $Bcc)
        if ($ccRecipients.Count -gt 0) { $message['ccRecipients'] = $ccRecipients }
        if ($bccRecipients.Count -gt 0) { $message['bccRecipients'] = $bccRecipients }

        $body = @{ message = $message; saveToSentItems = $false } | ConvertTo-Json -Depth 10
        $encodedFrom = [System.Uri]::EscapeDataString($From)
        Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $encodedFrom) -Body $body -ContentType 'application/json' | Out-Null
        return [pscustomobject]@{ Sent = $true; Mode = 'Graph' }
    }

    $mail = New-Object System.Net.Mail.MailMessage
    try {
        $mail.From = New-Object System.Net.Mail.MailAddress($From)
        foreach ($addr in ($To -split '[;,]')) { if (-not [string]::IsNullOrWhiteSpace($addr)) { $mail.To.Add($addr.Trim()) } }
        foreach ($addr in ($Cc -split '[;,]')) { if (-not [string]::IsNullOrWhiteSpace($addr)) { $mail.CC.Add($addr.Trim()) } }
        foreach ($addr in ($Bcc -split '[;,]')) { if (-not [string]::IsNullOrWhiteSpace($addr)) { $mail.Bcc.Add($addr.Trim()) } }
        $mail.Subject = $Subject
        $mail.Body = $BodyHtml
        $mail.IsBodyHtml = $true
        $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
        $mail.BodyEncoding = [System.Text.Encoding]::UTF8
        $mail.HeadersEncoding = [System.Text.Encoding]::UTF8
        $mail.BodyTransferEncoding = [System.Net.Mime.TransferEncoding]::QuotedPrintable

        $hasLogo = -not [string]::IsNullOrWhiteSpace($LogoPath) -and (Test-Path -LiteralPath $LogoPath)
        if ($hasLogo) {
            $alternateView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($BodyHtml, [System.Text.Encoding]::UTF8, 'text/html')
            $alternateView.ContentType.CharSet = 'utf-8'
            $alternateView.TransferEncoding = [System.Net.Mime.TransferEncoding]::QuotedPrintable
            $logo = New-Object System.Net.Mail.LinkedResource($LogoPath, $LogoMediaType)
            $logo.ContentId = $LogoContentId
            $logo.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
            [void]$alternateView.LinkedResources.Add($logo)
            $mail.AlternateViews.Clear()
            $mail.AlternateViews.Add($alternateView)
        }

        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        try {
            $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
            $smtp.Timeout = 30000
            $smtp.EnableSsl = $SmtpEnableSsl
            if ($SmtpUseIntegratedAuth) {
                $smtp.UseDefaultCredentials = $true
                $smtp.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            }
            $attempt = 0
            while ($true) {
                try {
                    $attempt++
                    $smtp.Send($mail)
                    break
                }
                catch {
                    if ($attempt -gt $RetryCount -or -not ($_.Exception -is [System.Net.Mail.SmtpException])) { throw }
                    if ($RetryDelaySeconds -gt 0) { Start-Sleep -Seconds $RetryDelaySeconds }
                }
            }
        }
        finally {
            if ($smtp) { $smtp.Dispose() }
        }
        return [pscustomobject]@{ Sent = $true; Mode = 'SMTP' }
    }
    finally {
        $mail.Dispose()
    }
}

function New-SmartM365CommunicationSummaryHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][hashtable]$Facts,
        [object[]]$Items = @()
    )

    $factRows = foreach ($key in ($Facts.Keys | Sort-Object)) {
        "<tr><td><b>$([System.Net.WebUtility]::HtmlEncode([string]$key))</b></td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Facts[$key]))</td></tr>"
    }
    $itemRows = foreach ($item in ($Items | Select-Object -First 500)) {
        $email = [System.Net.WebUtility]::HtmlEncode([string](Get-SmartM365CommunicationProperty -InputObject $item -Names @('Email') -DefaultValue ''))
        $language = [System.Net.WebUtility]::HtmlEncode([string](Get-SmartM365CommunicationProperty -InputObject $item -Names @('LanguageTag') -DefaultValue ''))
        $status = [System.Net.WebUtility]::HtmlEncode([string](Get-SmartM365CommunicationProperty -InputObject $item -Names @('Status') -DefaultValue ''))
        "<tr><td>$email</td><td>$language</td><td>$status</td></tr>"
    }

    return @"
<html>
  <body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#222;">
    <h2 style="margin:0 0 12px 0;">$([System.Net.WebUtility]::HtmlEncode($Title))</h2>
    <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;">
      $($factRows -join "`r`n")
    </table>
    <h3 style="margin:16px 0 8px 0;">Processed items</h3>
    <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;width:100%;">
      <tr><th align="left">Email</th><th align="left">Language</th><th align="left">Status</th></tr>
      $($itemRows -join "`r`n")
    </table>
  </body>
</html>
"@
}

Export-ModuleMember -Function `
    ConvertTo-SmartM365CommunicationHashtable, Read-SmartM365CommunicationJson, Merge-SmartM365CommunicationConfig, Get-SmartM365CommunicationConfigValue, `
    Resolve-SmartM365CommunicationIPv4Address, `
    Resolve-SmartM365CommunicationTokenizedValue, Import-SmartM365CommunicationCsv, Get-SmartM365CommunicationProperty, `
    Get-SmartM365CommunicationSmtpAddress, Get-SmartM365CommunicationTemplateContent, Expand-SmartM365CommunicationTemplate, `
    Assert-SmartM365CommunicationNoUnresolvedToken, Load-SmartM365CommunicationSentRegistry, Save-SmartM365CommunicationSentRegistry, `
    Register-SmartM365CommunicationSentItem, Resolve-SmartM365CommunicationLanguageTag, Get-SmartM365CommunicationSubject, `
    Get-SmartM365CommunicationHotline, Add-SmartM365CommunicationLogRow, Send-SmartM365CommunicationMail, `
    Initialize-SmartM365CommunicationExchangeSnapIn, ConvertTo-SmartM365CommunicationLdapFilterSafe, `
    Resolve-SmartM365CommunicationAdUserInfo, Test-SmartM365CommunicationExchangeMailboxState, `
    ConvertTo-SmartM365CommunicationBytes, Get-SmartM365CommunicationMailboxUsageInfo, Resolve-SmartM365CommunicationExchangeLanguageTag, `
    Get-SmartM365CommunicationLinkedLogoTokens, New-SmartM365CommunicationSummaryHtml
