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

function Initialize-SmartM365CommunicationLocalJsonFromTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$TemplatePath,
        [string]$ConfigDescription = 'communication local configuration'
    )

    if (Test-Path -LiteralPath $Path) { return $false }
    if ([string]::IsNullOrWhiteSpace($TemplatePath)) { $TemplatePath = '{0}.template' -f $Path }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        throw ((@(
            "Local JSON not found: $Path",
            "Template to copy is missing: $TemplatePath",
            'Create the missing local JSON from the matching template, then run the script again.'
        )) -join [Environment]::NewLine)
    }

    try {
        Copy-Item -LiteralPath $TemplatePath -Destination $Path -ErrorAction Stop
    }
    catch {
        throw ("Failed to create {0} '{1}' from template '{2}': {3}" -f $ConfigDescription, $Path, $TemplatePath, $_.Exception.Message)
    }

    Write-Host ((@(
        "Created $ConfigDescription from template.",
        "Local JSON: $Path",
        "Template: $TemplatePath",
        'Review the generated local JSON values; continuing with default template values unless edited before next run.'
    )) -join [Environment]::NewLine) -ForegroundColor Yellow

    return $true
}

function Read-SmartM365CommunicationJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Path -like '*.local.json') {
            Initialize-SmartM365CommunicationLocalJsonFromTemplate -Path $Path -TemplatePath ('{0}.template' -f $Path) | Out-Null
        }
        elseif ($Required) {
            throw "Configuration file not found: $Path"
        }
        else {
            return @{}
        }
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

function Initialize-SmartM365CommunicationExchangeOnline {
    [CmdletBinding()]
    param(
        [bool]$Enabled = $true,
        [string]$AppId = '',
        [string]$TenantId = '',
        [string]$Thumbprint = '',
        [string]$Organization = ''
    )

    if (-not $Enabled) {
        return [pscustomobject]@{
            Enabled = $false
            Available = $false
            Source = 'ExchangeOnline'
            Status = 'Disabled'
            ErrorMessage = ''
            ViewEntireForestRequested = $false
            ViewEntireForestApplied = $false
            ForestErrorMessage = ''
        }
    }

    if (-not (Get-Command -Name Connect-SmartM365CloudSession -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{
            Enabled = $true
            Available = $false
            Source = 'ExchangeOnline'
            Status = 'CloudSessionFunctionUnavailable'
            ErrorMessage = 'Connect-SmartM365CloudSession is not available. Import SmartM365.Core first.'
            ViewEntireForestRequested = $false
            ViewEntireForestApplied = $false
            ForestErrorMessage = ''
        }
    }

    try {
        $connectResult = Connect-SmartM365CloudSession `
            -AppId $AppId `
            -TenantId $TenantId `
            -Thumbprint $Thumbprint `
            -Organization $Organization `
            -ExchangeOnline $true `
            -Graph $false

        if (-not [bool]$connectResult.ExchangeOnlineConnected) {
            return [pscustomobject]@{
                Enabled = $true
                Available = $false
                Source = 'ExchangeOnline'
                Status = 'Unavailable'
                ErrorMessage = 'Exchange Online connection failed.'
                ViewEntireForestRequested = $false
                ViewEntireForestApplied = $false
                ForestErrorMessage = ''
            }
        }

        if (-not (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue) -and
            -not (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue)) {
            return [pscustomobject]@{
                Enabled = $true
                Available = $false
                Source = 'ExchangeOnline'
                Status = 'ExchangeCommandsUnavailable'
                ErrorMessage = 'Exchange Online connected, but mailbox commands are not available.'
                ViewEntireForestRequested = $false
                ViewEntireForestApplied = $false
                ForestErrorMessage = ''
            }
        }

        return [pscustomobject]@{
            Enabled = $true
            Available = $true
            Source = 'ExchangeOnline'
            Status = 'Available'
            ErrorMessage = ''
            ViewEntireForestRequested = $false
            ViewEntireForestApplied = $false
            ForestErrorMessage = ''
        }
    }
    catch {
        return [pscustomobject]@{
            Enabled = $true
            Available = $false
            Source = 'ExchangeOnline'
            Status = 'Unavailable'
            ErrorMessage = $_.Exception.Message
            ViewEntireForestRequested = $false
            ViewEntireForestApplied = $false
            ForestErrorMessage = ''
        }
    }
}

function Initialize-SmartM365CommunicationExchangeManagement {
    [CmdletBinding()]
    param(
        [string]$Mode = 'Auto',
        [bool]$Required = $true,
        [bool]$EnableExchange2016Fallback = $true,
        [string]$AppId = '',
        [string]$TenantId = '',
        [string]$Thumbprint = '',
        [string]$Organization = '',
        [string]$SnapInName = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
        [bool]$ViewEntireForest = $true
    )

    $normalizedMode = if ([string]::IsNullOrWhiteSpace($Mode)) { 'Auto' } else { $Mode.Trim() }
    $attempts = New-Object System.Collections.Generic.List[object]

    switch -Regex ($normalizedMode.ToLowerInvariant()) {
        '^(disabled|none|off)$' {
            return [pscustomobject]@{
                Enabled = $false
                Available = $false
                Source = 'Disabled'
                Status = 'Disabled'
                ErrorMessage = ''
                Attempts = @()
                ViewEntireForestRequested = $ViewEntireForest
                ViewEntireForestApplied = $false
                ForestErrorMessage = ''
            }
        }
        '^(exo|exchangeonline|online)$' {
            $attempts.Add('ExchangeOnline') | Out-Null
        }
        '^(exchange2016|snapin|onprem|onpremises)$' {
            $attempts.Add('Exchange2016') | Out-Null
        }
        '^(auto|preferexo|exchangeonlinewithfallback)$' {
            $attempts.Add('ExchangeOnline') | Out-Null
            if ($EnableExchange2016Fallback) { $attempts.Add('Exchange2016') | Out-Null }
        }
        default {
            throw "Unsupported ExchangeManagementMode '$Mode'. Use Auto, ExchangeOnline, Exchange2016, or Disabled."
        }
    }

    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($attempt in $attempts) {
        if ($attempt -eq 'ExchangeOnline') {
            $state = Initialize-SmartM365CommunicationExchangeOnline `
                -Enabled $true `
                -AppId $AppId `
                -TenantId $TenantId `
                -Thumbprint $Thumbprint `
                -Organization $Organization
        }
        else {
            $state = Initialize-SmartM365CommunicationExchangeSnapIn `
                -Enabled $true `
                -SnapInName $SnapInName `
                -ViewEntireForest $ViewEntireForest
            if ($state) { $state | Add-Member -NotePropertyName Source -NotePropertyValue 'Exchange2016' -Force }
        }

        if ($state -and $state.Available) {
            $state | Add-Member -NotePropertyName Attempts -NotePropertyValue @($attempts) -Force
            return $state
        }

        if ($state) {
            $errors.Add(("{0}: {1} ({2})" -f $attempt, $state.Status, $state.ErrorMessage)) | Out-Null
        }
    }

    return [pscustomobject]@{
        Enabled = $true
        Available = $false
        Source = $normalizedMode
        Status = $(if ($Required) { 'UnavailableRequired' } else { 'Unavailable' })
        ErrorMessage = ($errors -join ' | ')
        Attempts = @($attempts)
        ViewEntireForestRequested = $ViewEntireForest
        ViewEntireForestApplied = $false
        ForestErrorMessage = ''
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
        @{ Command = 'Get-EXORecipient'; Mode = 'Identity'; Remote = $null },
        @{ Command = 'Get-EXORecipient'; Mode = 'Filter'; Remote = $null },
        @{ Command = 'Get-RemoteMailbox'; Mode = 'Identity'; Remote = $true },
        @{ Command = 'Get-RemoteMailbox'; Mode = 'Filter'; Remote = $true },
        @{ Command = 'Get-Mailbox'; Mode = 'Identity'; Remote = $false },
        @{ Command = 'Get-EXOMailbox'; Mode = 'Identity'; Remote = $false }
    )) {
        $commandInfo = Get-Command -Name $attempt.Command -ErrorAction SilentlyContinue
        if (-not $commandInfo) { continue }

        try {
            $recipient = $null
            if ($attempt.Mode -eq 'Identity') {
                $parameters = @{ Identity = $SmtpAddress; ErrorAction = 'Stop' }
                if ($commandInfo.Parameters.ContainsKey('ResultSize')) { $parameters['ResultSize'] = 1 }
                if ($commandInfo.Parameters.ContainsKey('ReadFromDomainController')) { $parameters['ReadFromDomainController'] = $true }
                $recipient = & $attempt.Command @parameters
            }
            else {
                $filter = "EmailAddresses -eq 'SMTP:$SmtpAddress' -or EmailAddresses -eq 'smtp:$SmtpAddress' -or PrimarySmtpAddress -eq '$SmtpAddress'"
                $parameters = @{ Filter = $filter; ErrorAction = 'Stop' }
                if ($commandInfo.Parameters.ContainsKey('ResultSize')) { $parameters['ResultSize'] = 1 }
                if ($commandInfo.Parameters.ContainsKey('ReadFromDomainController')) { $parameters['ReadFromDomainController'] = $true }
                $recipient = & $attempt.Command @parameters
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

    $statisticsCommand = Get-Command -Name Get-MailboxStatistics -ErrorAction SilentlyContinue
    if (-not $statisticsCommand) { $statisticsCommand = Get-Command -Name Get-EXOMailboxStatistics -ErrorAction SilentlyContinue }
    if (-not $statisticsCommand) {
        $result.ErrorMessage = 'Get-MailboxStatistics/Get-EXOMailboxStatistics is not available.'
        return $result
    }

    try {
        $stats = & $statisticsCommand.Name -Identity $SmtpAddress -ErrorAction Stop
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
        [Parameter(Mandatory)][string[]]$TemplateRoot,
        [Parameter(Mandatory)][string]$TemplateBaseName,
        [Parameter(Mandatory)][string]$LanguageTag,
        [string]$DefaultLanguageTag = 'en'
    )

    $tag = if ([string]::IsNullOrWhiteSpace($LanguageTag)) { $DefaultLanguageTag } else { $LanguageTag }
    $neutral = ($tag -split '[-_]')[0].ToLowerInvariant()
    $defaultNeutral = ($DefaultLanguageTag -split '[-_]')[0].ToLowerInvariant()
    $templateRoots = @($TemplateRoot | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)

    $candidateNames = @(
        ("{0}.{1}.html" -f $TemplateBaseName, $tag),
        ("{0}.{1}.html" -f $TemplateBaseName, $neutral),
        ("{0}.{1}.html" -f $TemplateBaseName, $DefaultLanguageTag),
        ("{0}.{1}.html" -f $TemplateBaseName, $defaultNeutral),
        ("{0}.html" -f $TemplateBaseName)
    ) | Select-Object -Unique

    $candidates = foreach ($candidateName in $candidateNames) {
        foreach ($root in $templateRoots) {
            Join-Path -Path $root -ChildPath $candidateName
        }
    }

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

function Get-SmartM365CommunicationMailMode {
    [CmdletBinding()]
    param(
        [string]$MailSendMode = 'Auto',
        [string]$SmtpServer = ''
    )

    $mode = if ([string]::IsNullOrWhiteSpace($MailSendMode) -or $MailSendMode -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { 'Auto' } else { $MailSendMode.Trim() }
    switch ($mode.ToLowerInvariant()) {
        'auto' { if ([string]::IsNullOrWhiteSpace($SmtpServer)) { return 'Graph' }; return 'SmtpRelay' }
        'graph' { return 'Graph' }
        'smtprelay' { return 'SmtpRelay' }
        'smtp' { return 'SmtpRelay' }
        'disabled' { return 'Disabled' }
        default { throw "Unsupported MailSendMode '$MailSendMode'. Use Auto, Graph, SmtpRelay, or Disabled." }
    }
}

function Send-SmartM365CommunicationMail {
    [CmdletBinding()]
    param(
        [string]$MailSendMode = 'Auto',
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

    $effectiveMailMode = Get-SmartM365CommunicationMailMode -MailSendMode $MailSendMode -SmtpServer $SmtpServer
    if ($effectiveMailMode -eq 'Disabled') { return [pscustomobject]@{ Sent = $false; Mode = 'Disabled' } }

    if ($effectiveMailMode -eq 'Graph') {
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

        $hasLogo = -not [string]::IsNullOrWhiteSpace($LogoPath) -and (Test-Path -LiteralPath $LogoPath)
        if ($hasLogo) {
            $message['attachments'] = @(
                @{
                    '@odata.type' = '#microsoft.graph.fileAttachment'
                    name = [System.IO.Path]::GetFileName($LogoPath)
                    contentType = $LogoMediaType
                    contentBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($LogoPath))
                    isInline = $true
                    contentId = $LogoContentId
                }
            )
        }

        $body = @{ message = $message; saveToSentItems = $false } | ConvertTo-Json -Depth 10
        $encodedFrom = [System.Uri]::EscapeDataString($From)
        $attempt = 0
        while ($true) {
            try {
                $attempt++
                Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $encodedFrom) -Body $body -ContentType 'application/json' | Out-Null
                break
            }
            catch {
                if ($attempt -gt $RetryCount) { throw }
                if ($RetryDelaySeconds -gt 0) { Start-Sleep -Seconds $RetryDelaySeconds }
            }
        }
        return [pscustomobject]@{ Sent = $true; Mode = 'Graph' }
    }

    if ([string]::IsNullOrWhiteSpace($SmtpServer)) {
        throw 'SmtpServer is required when MailSendMode is SmtpRelay.'
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

function Get-SmartM365CommunicationTeamsUserMode {
    [CmdletBinding()]
    param([string]$TeamsUserMessageMode = 'Disabled')

    $mode = if ([string]::IsNullOrWhiteSpace($TeamsUserMessageMode) -or $TeamsUserMessageMode -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { 'Disabled' } else { $TeamsUserMessageMode.Trim() }
    switch ($mode.ToLowerInvariant()) {
        'disabled' { return 'Disabled' }
        'off' { return 'Disabled' }
        'none' { return 'Disabled' }
        'graph' { return 'GraphDelegated' }
        'graphdelegated' { return 'GraphDelegated' }
        'delegated' { return 'GraphDelegated' }
        default { throw "Unsupported TeamsUserMessageMode '$TeamsUserMessageMode'. Use Disabled or GraphDelegated." }
    }
}

function ConvertTo-SmartM365GraphODataStringLiteral {
    [CmdletBinding()]
    param([string]$Value)

    return ([string]$Value).Replace("'", "''")
}

function Connect-SmartM365CommunicationTeamsUserGraph {
    [CmdletBinding()]
    param(
        [string]$TenantId = '',
        [string[]]$Scopes = @('User.Read', 'Chat.Create', 'ChatMessage.Send')
    )

    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($context -and $context.AuthType -eq 'Delegated') {
        $scopeMap = @{}
        foreach ($scope in @($context.Scopes)) { $scopeMap[[string]$scope] = $true }
        $missingScope = $false
        foreach ($scope in $Scopes) {
            if (-not $scopeMap.ContainsKey($scope)) { $missingScope = $true; break }
        }
        if (-not $missingScope) { return $true }
    }

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose ("Microsoft Graph disconnect before Teams delegated connection failed: {0}" -f $_.Exception.Message) }
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -in @('__USE_GLOBAL__', 'USE_GLOBAL')) {
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
    }
    else {
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
    }

    return $true
}

function Resolve-SmartM365CommunicationGraphUserId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$User)

    $trimmed = $User.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return '' }

    $escapedPath = [System.Uri]::EscapeDataString($trimmed)
    try {
        $direct = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}?`$select=id" -f $escapedPath) -ErrorAction Stop
        if ($direct -and $direct.id) { return [string]$direct.id }
    }
    catch {
        Write-Verbose ("Direct Microsoft Graph user lookup failed for Teams target '{0}': {1}" -f $trimmed, $_.Exception.Message)
    }

    $literal = ConvertTo-SmartM365GraphODataStringLiteral -Value $trimmed
    foreach ($filter in @(
        "mail eq '$literal' or userPrincipalName eq '$literal'",
        "proxyAddresses/any(p:p eq 'SMTP:$literal') or proxyAddresses/any(p:p eq 'smtp:$literal')"
    )) {
        try {
            $encodedFilter = [System.Uri]::EscapeDataString($filter)
            $result = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users?`$select=id&`$top=1&`$filter={0}" -f $encodedFilter) -ErrorAction Stop
            if ($result -and $result.value -and $result.value.Count -gt 0 -and $result.value[0].id) {
                return [string]$result.value[0].id
            }
        }
        catch {
            Write-Verbose ("Filtered Microsoft Graph user lookup failed for Teams target '{0}': {1}" -f $trimmed, $_.Exception.Message)
        }
    }

    throw "Microsoft Graph user not found for Teams message target '$User'."
}

function Send-SmartM365CommunicationTeamsUserMessage {
    [CmdletBinding()]
    param(
        [string]$TeamsUserMessageMode = 'Disabled',
        [Parameter(Mandatory)][string]$To,
        [Parameter(Mandatory)][string]$MessageText,
        [string]$TenantId = '',
        [string]$SenderUserId = '',
        [int]$RetryCount = 1,
        [int]$RetryDelaySeconds = 2,
        [switch]$WhatIf
    )

    $effectiveMode = Get-SmartM365CommunicationTeamsUserMode -TeamsUserMessageMode $TeamsUserMessageMode
    if ($WhatIf) { return [pscustomobject]@{ Sent = $false; Mode = 'DryRun'; Error = '' } }
    if ($effectiveMode -eq 'Disabled') { return [pscustomobject]@{ Sent = $false; Mode = 'Disabled'; Error = '' } }
    if ([string]::IsNullOrWhiteSpace($To)) { throw 'Teams message recipient is required.' }
    if ([string]::IsNullOrWhiteSpace($MessageText)) { throw 'Teams message text is required.' }

    Connect-SmartM365CommunicationTeamsUserGraph -TenantId $TenantId | Out-Null

    if ([string]::IsNullOrWhiteSpace($SenderUserId)) {
        $me = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id' -ErrorAction Stop
        $SenderUserId = [string]$me.id
    }
    if ([string]::IsNullOrWhiteSpace($SenderUserId)) { throw 'Unable to resolve the delegated Teams sender user id.' }

    $targetUserId = Resolve-SmartM365CommunicationGraphUserId -User $To
    $senderBind = "https://graph.microsoft.com/v1.0/users('$SenderUserId')"
    $targetBind = "https://graph.microsoft.com/v1.0/users('$targetUserId')"
    $chatBody = @{
        chatType = 'oneOnOne'
        members = @(
            @{
                '@odata.type' = '#microsoft.graph.aadUserConversationMember'
                roles = @('owner')
                'user@odata.bind' = $senderBind
            },
            @{
                '@odata.type' = '#microsoft.graph.aadUserConversationMember'
                roles = @('owner')
                'user@odata.bind' = $targetBind
            }
        )
    } | ConvertTo-Json -Depth 8

    $attempt = 0
    while ($true) {
        try {
            $attempt++
            $chat = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/chats' -Body $chatBody -ContentType 'application/json' -ErrorAction Stop
            if (-not $chat -or [string]::IsNullOrWhiteSpace([string]$chat.id)) { throw 'Microsoft Graph did not return a chat id.' }

            $messageBody = @{
                body = @{
                    contentType = 'text'
                    content = $MessageText
                }
            } | ConvertTo-Json -Depth 6
            $chatId = [System.Uri]::EscapeDataString([string]$chat.id)
            Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/chats/{0}/messages" -f $chatId) -Body $messageBody -ContentType 'application/json' -ErrorAction Stop | Out-Null
            return [pscustomobject]@{ Sent = $true; Mode = 'GraphDelegated'; Error = '' }
        }
        catch {
            if ($attempt -gt $RetryCount) {
                return [pscustomobject]@{ Sent = $false; Mode = 'GraphDelegated'; Error = $_.Exception.Message }
            }
            if ($RetryDelaySeconds -gt 0) { Start-Sleep -Seconds $RetryDelaySeconds }
        }
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
    Get-SmartM365CommunicationHotline, Add-SmartM365CommunicationLogRow, Get-SmartM365CommunicationMailMode, Send-SmartM365CommunicationMail, `
    Get-SmartM365CommunicationTeamsUserMode, Send-SmartM365CommunicationTeamsUserMessage, `
    Initialize-SmartM365CommunicationExchangeSnapIn, Initialize-SmartM365CommunicationExchangeOnline, Initialize-SmartM365CommunicationExchangeManagement, ConvertTo-SmartM365CommunicationLdapFilterSafe, `
    Resolve-SmartM365CommunicationAdUserInfo, Test-SmartM365CommunicationExchangeMailboxState, `
    ConvertTo-SmartM365CommunicationBytes, Get-SmartM365CommunicationMailboxUsageInfo, Resolve-SmartM365CommunicationExchangeLanguageTag, `
    Get-SmartM365CommunicationLinkedLogoTokens, New-SmartM365CommunicationSummaryHtml

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCNeIdgYvNYv+yM
# XIGNkkfsh8LMm/igXBhVwF5bkPuDr6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIIke1kQ/SAZgC/N42G8D3WLLMT3L+cEausL9mf7PWgSqMA0GCSqG
# SIb3DQEBAQUABIIBgIICmu1WhMxOLV1GNOvUrE0y/Nx8VkGo638/U3Nf9zHa+jvs
# jbLTK9KPhsE8aBjQcwLdzcfkr0BE90SZ++48gJK+3cjqrgUQJjvTRPXHjat2aPLt
# iq8AsXnwJPu4+EZgO3bfT5uENzXG+1cgSdqsfXlsmdWFcYwFAh2VcB1DjxbBp8xT
# tmeIFXPt/TEoQF9aBk140JNQsAe/9mUceh1eEtE+b1rxtbBMaEs41JbMVUNBE34G
# 6K0Wjef6/kqzxv8WpNUX9BncG62nOT46vRLJjf9LVJKqjccj99jrPSr0PXpwQU1m
# q7TbYSSTUZZiRfI3D/9OQyDf1p8g9fvaSkdoByZgC1Gnoysa8uENNbsuJRRjiSJC
# JtetWET0YSjPsJhXy7kV2easRupAxOd5+LcPawKXCj20QxCfO7+XSi2RougHfZCE
# nxa2Ry4KOJ+VwMIaLLuYVHQjr8ZZfyB0wakz378O45MY4naumpjiKqtcbqEEKIEM
# +MF2s2LIupLJ68tCuKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MjVaMC8GCSqGSIb3DQEJBDEiBCBvbdo3pCvFkyPx4LjOJX8GZY45znxgNHIg+Zy7
# 3NqmADANBgkqhkiG9w0BAQEFAASCAgBsxaJoLuCq5kEHWXdCkKrbwqnuo4jikjkH
# TARFGNiwdQDrRREpRMZkXoG9nRlpuImaxqUEBkBRg93kf8+rfQuvnZytCY6z3DQP
# FsuLsMYtY3BKRnXwGG1l5Ro8DyI0IffyGrwIUqXdDsbD2KL8MbnTNcb/yzWkdRCO
# 8JmZg8yX92pHAV9TKZJuIDwvDDIqI/9MSrYo2beL0gOA8oX1x7JQB0sgGXi2k5Ls
# Cp7Dv63UBcx4PtcI3ahz6ODAq61eTX0XY+oPL7mbjRTAIi8/QzoUbk3w+10sa0Nm
# AaSnOIO9KBw8J3sQ4Wi1CAqwEBsmGE4YnVj93A1EHXabqx5E64MRhum5oY6H0/Nq
# PvTW3s7VOQsf55/xBaRZjh3N9mY9IbziVffuKpzQK3ohzlBtDmQxmthQqV5vAbJr
# ohoPCUvMwf9aSpFyTeC/kgI1Nfk4VxPi7UjA8kLX7LArj813lnWrRvu1HGVnfib7
# p78FN01WQKEv/zx0QHXt+VTyACG63gCoLpopSIbJYsV/vZI1ieUB8svulQmS6U07
# 3u8t9Y4ZaHnAJowdxfIReK8PZ314scqMu/U5bJCG6ptR4WyI85TJxGSlhs9UetnX
# hA7J95Qa+pV5Jq/CmamWeP7x/9H+tE+bkHu0nD+hyhG20VY1VyMb8p/2OddZgc5m
# oTx0fqjedw==
# SIG # End signature block
