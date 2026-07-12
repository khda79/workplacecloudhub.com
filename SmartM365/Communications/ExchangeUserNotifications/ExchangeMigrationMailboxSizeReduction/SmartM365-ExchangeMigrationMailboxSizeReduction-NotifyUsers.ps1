<#
.SYNOPSIS
Notifies selected Exchange mailbox users that they must reduce mailbox size before migration.

.DESCRIPTION
Loads SmartM365 tenant context, the shared Communications configuration, and the
ExchangeMigrationMailboxSizeReduction campaign configuration, then sends localized HTML notification
emails to users whose Exchange mailbox size is above the quota associated with their target
migration license.

In Inventory mode, the script reads the SmartM365 license CSV export to select users matching the
configured target SKUs. It maps each user's SKU to MailboxQuotaBySkuPartNumber, for example E3 at
100 GB and F1/F3 at 2 GB. If the license CSV is missing or older than LicenseCsvMaxAgeHours and
EnableLiveLicenseLookupWhenCsvMissing is enabled, it resolves the license live from Microsoft Graph:
for FromList recipients only in FromList mode, or by enumerating Microsoft Graph users in Inventory
mode. Mailbox inventory CSVs are used only as optional enrichment. Exchange management uses Exchange
Online first by default, with Exchange 2016 snap-in fallback when enabled, and the effective mailbox
state and mailbox size threshold are verified live before a notification is sent.

The script resolves language from the email domain or ForceLanguage, injects mailbox size, quota,
hotline, and the correct webmail URL for UserMailbox or RemoteUserMailbox, prevents duplicate sends
through the sent registry, excludes mailboxes that do not exist in Exchange management or are RemoteMailbox
objects by default, writes a per-recipient log, prompts before live sends unless SkipConfirmation is
used, and can send a summary email and Teams notification after a live run.

Expected list columns include an email column such as PrimarySmtpAddress, EmailAddress, Mail,
UserPrincipalName, or Recipient. Inventory column aliases are handled for license SKU, mailbox size
in MB, and mailbox type.

.PARAMETER Tenant
SmartM365 tenant context key used to resolve root paths, mail credentials, Teams settings, and defaults.

.PARAMETER FromList
Switches from inventory selection to an operator-provided CSV list of target recipients.

.PARAMETER ListCsvPath
CSV file used with FromList. When omitted, the path comes from the campaign JSON.

.PARAMETER CampaignConfigPath
Optional campaign JSON override. When omitted, the script uses Config/Campaigns/ExchangeMigrationMailboxSizeReduction.local.json.

.PARAMETER ForceLanguage
Forces every email to use one language tag, such as fr or en, regardless of domain detection.

.PARAMETER ForceSend
Bypasses duplicate and threshold skip checks for operator-controlled replays.

.PARAMETER NoSummaryEmail
Suppresses the end-of-run summary email. Logs and console summary are still produced.

.PARAMETER TeamsUserMessageMode
Optional user-facing Teams message mode. Disabled by default. Use GraphDelegated to send a one-on-one
Teams chat message with Microsoft Graph delegated permissions.

.PARAMETER SkipConfirmation
Skips the live-run YES prompt. Intended for scheduled or already-approved execution only.

.PARAMETER WhatIf
Performs a dry run: resolves candidates/templates and writes logs, but does not send user, summary, or Teams notifications.

.EXAMPLE
.\SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.ps1 -Tenant prod -WhatIf

Builds candidates from the license inventory CSV and performs a dry run using the quota mapped to each target SKU.

.EXAMPLE
.\SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.ps1 -Tenant prod -FromList -ListCsvPath .\Recipients\pilot.csv -WhatIf

Performs a dry run for a manually supplied pilot recipient list.

.VERSION
1.3
#>
[CmdletBinding(DefaultParameterSetName = 'Inventory')]
param(
    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [string]$Tenant = 'test',

    [Parameter(ParameterSetName = 'FromList', Mandatory)]
    [switch]$FromList,

    [Parameter(ParameterSetName = 'FromList')]
    [string]$ListCsvPath = '',

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [string]$CampaignConfigPath = '',

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [string]$ForceLanguage = '',

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [string]$MailSendMode = '',

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [string]$ExchangeManagementMode = '',

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [string]$TeamsUserMessageMode = '',

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [switch]$ForceSend,

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [switch]$NoSummaryEmail,

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [switch]$SkipConfirmation,

    [Parameter(ParameterSetName = 'Inventory')]
    [Parameter(ParameterSetName = 'FromList')]
    [switch]$WhatIf
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Find-UpwardFile {
    param([Parameter(Mandatory)][string]$FileName, [string]$StartPath = $PSScriptRoot)
    $current = $StartPath
    while ($current) {
        $candidate = Join-Path -Path $current -ChildPath $FileName
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    throw "Unable to find '$FileName' from '$StartPath'."
}

function Convert-ToHash {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return @{} }
    if ($Value -is [hashtable]) { return $Value }
    return ConvertTo-SmartM365CommunicationHashtable -InputObject $Value
}

function Resolve-CommPath {
    param([string]$Value)
    return Resolve-SmartM365CommunicationTokenizedValue -Value $Value -Tokens $pathTokens
}

function ConvertTo-DoubleValue {
    param([AllowNull()]$Value, [double]$DefaultValue = -1)
    if ($null -eq $Value) { return $DefaultValue }
    $text = ([string]$Value).Trim().Replace(',', '.')
    $parsed = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) { return $parsed }
    return $DefaultValue
}

function ConvertTo-StringArray {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) {
        return @($Value -split '[,;|]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $items = New-Object System.Collections.ArrayList
    foreach ($item in @($Value)) {
        if ($null -eq $item) { continue }
        $text = ([string]$item).Trim()
        if ($text) { [void]$items.Add($text) }
    }
    return @($items)
}

function Get-LocalizedTextValue {
    param(
        [hashtable]$TextByLanguage,
        [string]$LanguageTag,
        [string]$DefaultLanguageTag = 'en'
    )

    if (-not $TextByLanguage -or $TextByLanguage.Count -eq 0) { return '' }

    $candidates = New-Object System.Collections.ArrayList
    foreach ($candidate in @($LanguageTag, $DefaultLanguageTag, 'default')) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) { [void]$candidates.Add(([string]$candidate).Trim()) }
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and ([string]$candidate).Contains('-')) {
            [void]$candidates.Add((([string]$candidate).Split('-')[0]).Trim())
        }
    }

    foreach ($candidate in @($candidates | Select-Object -Unique)) {
        if ($TextByLanguage.ContainsKey($candidate)) { return [string]$TextByLanguage[$candidate] }
    }

    return ''
}

function ConvertTo-FreeTextBlock {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return [pscustomobject]@{
            Style = 'display:none;'
            TextHtml = ''
        }
    }

    $encoded = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $encoded = $encoded -replace "(`r`n|`n|`r)", '<br>'

    return [pscustomobject]@{
        Style = ''
        TextHtml = $encoded
    }
}

function ConvertTo-MailboxQuotaMap {
    param([hashtable]$Config)

    $map = @{}
    $configuredMap = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $Config -Name 'MailboxQuotaBySkuPartNumber' -DefaultValue @{})

    foreach ($key in $configuredMap.Keys) {
        $sku = ([string]$key).Trim().ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($sku)) { continue }

        $entry = Convert-ToHash $configuredMap[$key]
        $maxGb = ConvertTo-DoubleValue -Value (Get-SmartM365CommunicationConfigValue -Config $entry -Name 'MaxGb' -DefaultValue -1) -DefaultValue -1
        if ($maxGb -le 0) { continue }
        $thresholdMb = [math]::Round($maxGb * 1024, 2)

        $quotaLabel = [string](Get-SmartM365CommunicationConfigValue -Config $entry -Name 'QuotaLabel' -DefaultValue ("{0} GB" -f $maxGb))
        $map[$sku] = [pscustomobject]@{
            SkuPartNumber = $sku
            ThresholdMB = [double]$thresholdMb
            MaxGb = [double]$maxGb
            QuotaLabel = $quotaLabel
        }
    }

    return $map
}

function Get-PrimarySmtpAddressFromGraphUser {
    param([AllowNull()]$User)

    if ($null -eq $User) { return '' }
    foreach ($proxy in @($User.ProxyAddresses)) {
        if ([string]$proxy -cmatch '^SMTP:(.+)$') { return $Matches[1].Trim() }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$User.Mail)) { return ([string]$User.Mail).Trim() }
    if (-not [string]::IsNullOrWhiteSpace([string]$User.UserPrincipalName)) { return ([string]$User.UserPrincipalName).Trim() }
    return ''
}

function ConvertTo-GraphODataStringLiteral {
    param([string]$Value)
    return ([string]$Value).Replace("'", "''")
}

function Ensure-GraphLicenseLookup {
    param([hashtable]$TenantConfig)

    if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($context) { return }
        }
        catch {}
    }

    $connectResult = Connect-SmartM365CloudSession `
        -AppId ([string](Get-SmartM365CommunicationConfigValue -Config $TenantConfig -Name 'AppId' -DefaultValue '')) `
        -TenantId ([string](Get-SmartM365CommunicationConfigValue -Config $TenantConfig -Name 'TenantId' -DefaultValue '')) `
        -Thumbprint ([string](Get-SmartM365CommunicationConfigValue -Config $TenantConfig -Name 'Thumbprint' -DefaultValue (Get-SmartM365CommunicationConfigValue -Config $TenantConfig -Name 'Thumb' -DefaultValue ''))) `
        -Organization ([string](Get-SmartM365CommunicationConfigValue -Config $TenantConfig -Name 'OrgDomain' -DefaultValue '')) `
        -ExchangeOnline:$false `
        -Graph:$true `
        -GraphScopes @('User.Read.All', 'Directory.Read.All')

    if (-not $connectResult.GraphConnected) {
        throw 'Microsoft Graph live license lookup failed to connect.'
    }
}

function Get-GraphUserBySmtpAddress {
    param([Parameter(Mandatory)][string]$SmtpAddress)

    $smtp = $SmtpAddress.Trim()
    if ([string]::IsNullOrWhiteSpace($smtp)) { return $null }

    try {
        $user = Get-MgUser -UserId $smtp -Property 'DisplayName','UserPrincipalName','Id','Mail','ProxyAddresses' -ErrorAction Stop
        if ($user) { return $user }
    }
    catch {}

    $safe = ConvertTo-GraphODataStringLiteral -Value $smtp
    foreach ($filter in @(
        "mail eq '$safe' or userPrincipalName eq '$safe'",
        "proxyAddresses/any(p:p eq 'SMTP:$safe') or proxyAddresses/any(p:p eq 'smtp:$safe')"
    )) {
        try {
            $user = Get-MgUser -Filter $filter -Property 'DisplayName','UserPrincipalName','Id','Mail','ProxyAddresses' -ConsistencyLevel eventual -Top 1 -ErrorAction Stop | Select-Object -First 1
            if ($user) { return $user }
        }
        catch {}
    }

    return $null
}

function Get-LiveLicenseQuotaInfoForGraphUser {
    param(
        [Parameter(Mandatory)]$User,
        [hashtable]$QuotaBySku
    )

    if ($null -eq $User -or -not $User.Id) { return $null }
    $licenseDetails = @(Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop)
    $best = $null
    foreach ($detail in $licenseDetails) {
        $sku = ([string]$detail.SkuPartNumber).Trim().ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($sku) -or -not $QuotaBySku.ContainsKey($sku)) { continue }
        $quota = $QuotaBySku[$sku]
        if ($null -eq $best -or [double]$quota.ThresholdMB -gt [double]$best.ThresholdMB) {
            $smtp = Get-PrimarySmtpAddressFromGraphUser -User $User
            $best = [pscustomobject]@{
                SkuPartNumber = $quota.SkuPartNumber
                ThresholdMB = [double]$quota.ThresholdMB
                MaxGb = [double]$quota.MaxGb
                QuotaLabel = [string]$quota.QuotaLabel
                LicenseRow = [pscustomobject]@{
                    primarysmtp = $smtp
                    'User principal name' = [string]$User.UserPrincipalName
                    'Display name' = [string]$User.DisplayName
                    SkuPartNumber = [string]$quota.SkuPartNumber
                }
            }
        }
    }
    return $best
}

function Resolve-LiveLicenseQuotaInfo {
    param(
        [Parameter(Mandatory)][string]$SmtpAddress,
        [hashtable]$QuotaBySku
    )

    $user = Get-GraphUserBySmtpAddress -SmtpAddress $SmtpAddress
    if (-not $user) { return $null }
    return Get-LiveLicenseQuotaInfoForGraphUser -User $user -QuotaBySku $QuotaBySku
}

function Find-ConfiguredCsv {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$PrimaryName,
        [AllowNull()]$FallbackNames,
        [string]$ConfiguredPath = '',
        [bool]$Required = $true
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredPath)) {
        $resolvedConfiguredPath = Resolve-CommPath $ConfiguredPath
        if (Test-Path -LiteralPath $resolvedConfiguredPath) { return $resolvedConfiguredPath }
    }

    $names = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace($PrimaryName)) { [void]$names.Add($PrimaryName) }
    foreach ($name in @($FallbackNames)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$name)) { [void]$names.Add([string]$name) }
    }

    foreach ($name in $names) {
        $path = Join-Path -Path $Folder -ChildPath $name
        if (Test-Path -LiteralPath $path) { return $path }
    }

    if ($Required) { throw "None of the configured CSV files were found in '$Folder': $($names -join ', ')" }
    return ''
}

function Build-MailboxIndex {
    param([object[]]$ExoRows, [object[]]$OnPremRows)

    $index = @{}
    foreach ($row in $OnPremRows) {
        $smtp = ([string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('PrimarySMTPaddress','PrimarySmtpAddress','EmailAddress') -DefaultValue '')).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($smtp)) { continue }
        $sizeMb = ConvertTo-DoubleValue -Value (Get-SmartM365CommunicationProperty -InputObject $row -Names @('TotalItemSize-In-MB','TotalItemSizeMB_Integer','MailboxSizeMB') -DefaultValue -1)
        $type = [string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('RecipientType','RecipientTypeDetails','MailboxType') -DefaultValue 'UserMailbox')
        if ($type -like '*RemoteUserMailbox*') { $type = 'RemoteUserMailbox' }
        elseif ($type -like '*UserMailbox*') { $type = 'UserMailbox' }
        $index[$smtp] = [pscustomobject]@{ Email = $smtp; SizeMB = $sizeMb; MailboxType = $type; Source = 'OnPrem' }
    }

    foreach ($row in $ExoRows) {
        $smtp = ([string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('PrimarySmtpAddress','PrimarySMTPaddress','EmailAddress') -DefaultValue '')).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($smtp)) { continue }
        $sizeMb = ConvertTo-DoubleValue -Value (Get-SmartM365CommunicationProperty -InputObject $row -Names @('TotalItemSizeMB_Integer','TotalItemSize-In-MB','MailboxSizeMB','TotalItemSizeMB') -DefaultValue -1)
        $index[$smtp] = [pscustomobject]@{ Email = $smtp; SizeMB = $sizeMb; MailboxType = 'RemoteUserMailbox'; Source = 'EXO' }
    }
    return $index
}

function Build-Candidates {
    param(
        [string[]]$SmtpList,
        [hashtable]$MailboxIndex,
        [hashtable]$LicenseMap,
        [hashtable]$Registry,
        [bool]$RequireMailboxInventory = $false
    )

    $candidates = New-Object System.Collections.ArrayList
    $pre = @{ EmptySmtp = 0; DuplicateInRun = 0; NotFoundInIndex = 0; NoTargetLicense = 0; AlreadySent = 0 }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($raw in $SmtpList) {
        $smtp = ([string]$raw).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($smtp)) { $pre.EmptySmtp++; continue }
        if (-not $seen.Add($smtp) -and -not $ForceSend) { $pre.DuplicateInRun++; continue }
        if (-not $MailboxIndex.ContainsKey($smtp)) {
            if ($RequireMailboxInventory) { $pre.NotFoundInIndex++; continue }
            $mailbox = [pscustomobject]@{ Email = $smtp; SizeMB = -1; MailboxType = ''; Source = 'PendingExchange' }
        }
        else {
            $mailbox = $MailboxIndex[$smtp]
        }

        if ($Registry.ContainsKey($smtp) -and -not $ForceSend) { $pre.AlreadySent++; continue }
        if (-not $LicenseMap -or -not $LicenseMap.ContainsKey($smtp)) { $pre.NoTargetLicense++; continue }

        $quotaInfo = $LicenseMap[$smtp]

        [void]$candidates.Add([pscustomobject]@{
            Email = $smtp
            MailboxSizeMB = [double]$mailbox.SizeMB
            MailboxType = [string]$mailbox.MailboxType
            MailboxSource = [string]$mailbox.Source
            TargetSkuPartNumber = [string]$quotaInfo.SkuPartNumber
            MailboxThresholdMB = [double]$quotaInfo.ThresholdMB
            MailboxMaxGb = [double]$quotaInfo.MaxGb
            MailboxQuotaLabel = [string]$quotaInfo.QuotaLabel
        })
    }

    return [pscustomobject]@{ Candidates = @($candidates); Preflight = $pre }
}

$tenantContextPath = Find-UpwardFile -FileName 'Config\SmartM365-TenantContext.ps1'
. $tenantContextPath
$effectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

$coreModulePath = Find-UpwardFile -FileName 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module $coreModulePath -MinimumVersion '1.0.24' -Force -ErrorAction Stop

$commModulePath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'SmartM365.Communications.psm1'
Import-Module $commModulePath -Force -DisableNameChecking

$tenantConfig = Convert-ToHash $effectiveConfig
$global:AppId = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'AppId' -DefaultValue '')
$global:TenantId = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'TenantId' -DefaultValue '')
$global:Thumb = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'Thumb' -DefaultValue '')
$global:Thumbprint = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'Thumbprint' -DefaultValue $global:Thumb)

$pathTokens = @{
    TenantKey = $Tenant
    CampaignRootPath = $PSScriptRoot
    SmartM365RootPath = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'SmartM365RootPath' -DefaultValue (Split-Path -Path $tenantContextPath -Parent))
    WorkspaceRootPath = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'WorkspaceRootPath' -DefaultValue '')
    DataAllRootPath = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'DataAllRootPath' -DefaultValue '')
    LatestCsvFolderPath = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'LatestCsvFolderPath' -DefaultValue '')
    LogAllRootPath = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'LogAllRootPath' -DefaultValue '')
}
$pathTokens.WorkspaceRootPath = Resolve-SmartM365CommunicationTokenizedValue -Value $pathTokens.WorkspaceRootPath -Tokens $pathTokens
$pathTokens.DataAllRootPath = Resolve-SmartM365CommunicationTokenizedValue -Value $pathTokens.DataAllRootPath -Tokens $pathTokens
$pathTokens.LatestCsvFolderPath = Resolve-SmartM365CommunicationTokenizedValue -Value $pathTokens.LatestCsvFolderPath -Tokens $pathTokens
$pathTokens.LogAllRootPath = Resolve-SmartM365CommunicationTokenizedValue -Value $pathTokens.LogAllRootPath -Tokens $pathTokens

$commRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$baseConfig = Read-SmartM365CommunicationJson -Path (Join-Path $commRoot 'ExchangeUserNotifications\Config\Communications.local.json.template')
$baseConfig = Merge-SmartM365CommunicationConfig -BaseConfig $baseConfig -OverlayConfig (Read-SmartM365CommunicationJson -Path (Join-Path $commRoot 'ExchangeUserNotifications\Config\Communications.local.json'))

$campaignTemplatePath = Join-Path $commRoot 'ExchangeUserNotifications\Config\Campaigns\ExchangeMigrationMailboxSizeReduction.local.json.template'
if ([string]::IsNullOrWhiteSpace($CampaignConfigPath)) {
    $CampaignConfigPath = Join-Path $commRoot 'ExchangeUserNotifications\Config\Campaigns\ExchangeMigrationMailboxSizeReduction.local.json'
}
$campaignConfig = Merge-SmartM365CommunicationConfig -BaseConfig (Read-SmartM365CommunicationJson -Path $campaignTemplatePath) -OverlayConfig (Read-SmartM365CommunicationJson -Path $CampaignConfigPath)
$config = Merge-SmartM365CommunicationConfig -BaseConfig $baseConfig -OverlayConfig $campaignConfig

$campaignName = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'CampaignName' -DefaultValue 'ExchangeMigrationMailboxSizeReduction')
$runId = [guid]::NewGuid().ToString()
$runOutputPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RunOutputPath' -DefaultValue ''))
$logOutputPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogOutputPath' -DefaultValue $runOutputPath))
$sentRegistryPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SentRegistryPath' -DefaultValue (Join-Path $runOutputPath 'Data\SentRegistry.csv')))
$logPath = Join-Path -Path $logOutputPath -ChildPath ("SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers-Log_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
foreach ($folder in @($runOutputPath, $logOutputPath, (Split-Path -Path $sentRegistryPath -Parent))) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
}

Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $runOutputPath -LatestOutputRoot $pathTokens.LatestCsvFolderPath -LogPath (Join-Path $logOutputPath 'SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.log')

$latestFolder = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LatestCsvFolderPath' -DefaultValue $pathTokens.LatestCsvFolderPath))
if ([string]::IsNullOrWhiteSpace($ListCsvPath)) {
    $ListCsvPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ListCsvPath' -DefaultValue ''))
}
$defaultTemplateRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Templates'
$configuredTemplateRoot = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TemplateRootPath' -DefaultValue $defaultTemplateRoot))
$templateRoot = @($configuredTemplateRoot, $defaultTemplateRoot) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique
$templateBaseName = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TemplateBaseName' -DefaultValue 'SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers-Template')
$logoPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogoPath' -DefaultValue (Join-Path $pathTokens.SmartM365RootPath 'WorkplaceCloudHub.ico')))
$logoContentId = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogoContentId' -DefaultValue 'smartm365logo')
$logoMediaType = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogoMediaType' -DefaultValue 'image/x-icon')
$defaultLanguage = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultLanguageTag' -DefaultValue 'en' -FallbackConfig $baseConfig)
$from = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'From' -DefaultValue '' -FallbackConfig $tenantConfig)
$configuredMailSendMode = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'MailSendMode' -DefaultValue 'Auto' -FallbackConfig $baseConfig)
if ([string]::IsNullOrWhiteSpace($MailSendMode)) { $mailSendMode = $configuredMailSendMode } else { $mailSendMode = $MailSendMode }
$smtpServer = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpServer' -DefaultValue '' -FallbackConfig $tenantConfig)
$relayIp = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RelayIp' -DefaultValue '' -FallbackConfig $tenantConfig)
$smtpResolveIPv4 = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpResolveIPv4' -DefaultValue $true -FallbackConfig $baseConfig)
if ((Get-SmartM365CommunicationMailMode -MailSendMode $mailSendMode -SmtpServer $smtpServer) -eq 'SmtpRelay') {
    if (-not [string]::IsNullOrWhiteSpace($relayIp)) { $smtpServer = $relayIp }
    elseif ($smtpResolveIPv4 -and -not [string]::IsNullOrWhiteSpace($smtpServer)) { $smtpServer = Resolve-SmartM365CommunicationIPv4Address -HostName $smtpServer }
}
$smtpPort = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpPort' -DefaultValue 25 -FallbackConfig $tenantConfig)
$smtpUseIntegratedAuth = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpUseIntegratedAuth' -DefaultValue $false -FallbackConfig $tenantConfig)
$smtpEnableSsl = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpEnableSsl' -DefaultValue $false -FallbackConfig $tenantConfig)
$smtpRetryCount = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpRetryCount' -DefaultValue 1 -FallbackConfig $baseConfig)
$smtpRetryDelaySeconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpRetryDelaySeconds' -DefaultValue 2 -FallbackConfig $baseConfig)
$summaryTo = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryTo' -DefaultValue (Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'ErrorMailTo' -DefaultValue ''))
$summaryBcc = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryBcc' -DefaultValue '')
$bccAll = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'BccAll' -DefaultValue '')
$preventResend = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'PreventResendAcrossRuns' -DefaultValue $true)
$mailboxQuotaBySku = ConvertTo-MailboxQuotaMap -Config $config
if ($mailboxQuotaBySku.Count -eq 0) { throw 'MailboxQuotaBySkuPartNumber is mandatory and must contain at least one usable SKU quota entry.' }
$targetSkuList = ConvertTo-StringArray (Get-SmartM365CommunicationConfigValue -Config $config -Name 'TargetSkuPartNumbers' -DefaultValue @())
if ($targetSkuList.Count -eq 0 -and $mailboxQuotaBySku.Count -gt 0) { $targetSkuList = @($mailboxQuotaBySku.Keys) }
$targetSkuSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($sku in $targetSkuList) {
    if (-not [string]::IsNullOrWhiteSpace($sku)) { [void]$targetSkuSet.Add($sku.Trim().ToUpperInvariant()) }
}
if ($targetSkuSet.Count -eq 0) { throw 'TargetSkuPartNumbers is empty and no SKU could be inferred from MailboxQuotaBySkuPartNumber.' }
$subjectByLanguage = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'SubjectByLanguage' -DefaultValue @{})
$freeTextByLanguage = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'FreeTextByLanguage' -DefaultValue @{})
$defaultSubject = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultSubject' -DefaultValue (Get-SmartM365CommunicationConfigValue -Config $subjectByLanguage -Name 'default' -DefaultValue $campaignName))
$summaryTitle = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryTitle' -DefaultValue $campaignName)
$summarySubject = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummarySubject' -DefaultValue $campaignName)
$teamsSuccessTitle = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsSuccessTitle' -DefaultValue $campaignName)
$teamsFailureTitle = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsFailureTitle' -DefaultValue $campaignName)
$configuredTeamsUserMessageMode = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsUserMessageMode' -DefaultValue 'Disabled' -FallbackConfig $baseConfig)
if ([string]::IsNullOrWhiteSpace($TeamsUserMessageMode)) { $teamsUserMessageMode = $configuredTeamsUserMessageMode } else { $teamsUserMessageMode = $TeamsUserMessageMode }
$teamsUserMessageByLanguage = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsUserMessageByLanguage' -DefaultValue @{})
$teamsUserSenderUserId = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsUserSenderUserId' -DefaultValue '' -FallbackConfig $baseConfig)
$teamsUserRetryCount = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsUserRetryCount' -DefaultValue 1 -FallbackConfig $baseConfig)
$teamsUserRetryDelaySeconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsUserRetryDelaySeconds' -DefaultValue 2 -FallbackConfig $baseConfig)
$hotlineMap = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'HotlineByLanguageOrCountry' -DefaultValue @{} -FallbackConfig $baseConfig)
$domainLanguageMap = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'DomainLanguageMap' -DefaultValue @{} -FallbackConfig $baseConfig)
$webmailByType = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'WebmailUrlByMailboxType' -DefaultValue @{})
$enableAdLookup = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableAdLookupForLanguage' -DefaultValue $true)
$forestGcServer = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ForestGcServer' -DefaultValue '')
$enableExchangeMailboxStateCheck = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchangeMailboxStateCheck' -DefaultValue $true)
$enableExchangeMailboxUsageCheck = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchangeMailboxUsageCheck' -DefaultValue $true)
$skipMailboxNotFound = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipMailboxNotFoundWhenExchangeCheckAvailable' -DefaultValue $true)
$skipRemoteMailbox = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipRemoteMailboxWhenExchangeCheckAvailable' -DefaultValue $true)
$skipMailboxUsageUnavailable = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipMailboxUsageUnavailableWhenExchangeCheckAvailable' -DefaultValue $true)
$enableLiveLicenseLookupWhenCsvMissing = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableLiveLicenseLookupWhenCsvMissing' -DefaultValue $true)
$licenseCsvMaxAgeHours = [double](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LicenseCsvMaxAgeHours' -DefaultValue 12)
$batchSize = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'BatchSize' -DefaultValue 50 -FallbackConfig $baseConfig)
$batchPauseSeconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'BatchPauseSeconds' -DefaultValue 3 -FallbackConfig $baseConfig)
$intraEmailDelayMilliseconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'IntraEmailDelayMilliseconds' -DefaultValue 150 -FallbackConfig $baseConfig)
$extendedPauseEvery = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExtendedPauseEvery' -DefaultValue 0 -FallbackConfig $baseConfig)
$extendedPauseSeconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExtendedPauseSeconds' -DefaultValue 30 -FallbackConfig $baseConfig)
if ($batchSize -lt 1) { $batchSize = 50 }

$configuredExchangeManagementMode = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeManagementMode' -DefaultValue 'Auto')
if ([string]::IsNullOrWhiteSpace($ExchangeManagementMode)) { $exchangeManagementMode = $configuredExchangeManagementMode } else { $exchangeManagementMode = $ExchangeManagementMode }
$requireExchangeManagement = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RequireExchangeManagement' -DefaultValue ([bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RequireExchange2016SnapIn' -DefaultValue $true)))
$enableExchange2016Fallback = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchange2016Fallback' -DefaultValue ([bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchange2016SnapIn' -DefaultValue $true)))
$exchangeManagementState = Initialize-SmartM365CommunicationExchangeManagement `
    -Mode $exchangeManagementMode `
    -Required $requireExchangeManagement `
    -EnableExchange2016Fallback $enableExchange2016Fallback `
    -AppId $global:AppId `
    -TenantId $global:TenantId `
    -Thumbprint $global:Thumbprint `
    -Organization ([string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'OrgDomain' -DefaultValue '')) `
    -SnapInName ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInName' -DefaultValue 'Microsoft.Exchange.Management.PowerShell.SnapIn')) `
    -ViewEntireForest ([bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'Exchange2016ViewEntireForest' -DefaultValue ([bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInViewEntireForest' -DefaultValue $true))))
$exchangeSnapInState = $exchangeManagementState
$exchangeScope = if ($exchangeManagementState.Source -eq 'ExchangeOnline') { 'ExchangeOnline' } else { 'OnPremExchange' }
$exchangeMailboxSource = if ($exchangeManagementState.Source -eq 'ExchangeOnline') { 'ExchangeOnline' } else { 'Exchange2016' }
if ($exchangeManagementState.Enabled) {
    if ($exchangeManagementState.Available) {
        Write-Host ("Exchange management available via {0}." -f $exchangeManagementState.Source) -ForegroundColor DarkCyan
        if ($exchangeManagementState.ViewEntireForestApplied) { Write-Host 'Exchange ADServerSettings: ViewEntireForest = True' -ForegroundColor DarkCyan }
        elseif (-not [string]::IsNullOrWhiteSpace($exchangeManagementState.ForestErrorMessage)) { Write-Host ("WARNING: Failed to set ViewEntireForest=True: {0}" -f $exchangeManagementState.ForestErrorMessage) -ForegroundColor Yellow }
    }
    else {
        Write-Host ("Exchange management unavailable. Status={0}; Error={1}" -f $exchangeManagementState.Status, $exchangeManagementState.ErrorMessage) -ForegroundColor DarkYellow
    }
}
if ($requireExchangeManagement -and $exchangeManagementState.Enabled -and -not $exchangeManagementState.Available) {
    throw ("Exchange management is mandatory for this campaign and is not available. Status={0}; Error={1}" -f $exchangeManagementState.Status, $exchangeManagementState.ErrorMessage)
}

$columns = @('Timestamp','RunId','Campaign','Email','LanguageTag','MailboxSizeGB','MailboxType','MailboxSource','TargetSkuPartNumber','Subject','Status','TeamsStatus','TeamsError','SkipReason','ErrorMessage','RecipientTypeDetails','ExchangeSource')
$counters = @{ Candidates = 0; Sent = 0; DryRun = 0; Failed = 0; Skipped = 0; MailboxNotFound = 0; RemoteMailbox = 0; MailboxUsageUnavailable = 0; MailboxUsageBelowThreshold = 0; TeamsSent = 0; TeamsFailed = 0; TeamsDryRun = 0 }
$processedItems = New-Object System.Collections.ArrayList
$registry = Load-SmartM365CommunicationSentRegistry -Path $sentRegistryPath

try {
    $exoPath = Find-ConfiguredCsv -Folder $latestFolder -PrimaryName ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExoMailboxCsvName' -DefaultValue 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv')) -FallbackNames (Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExoMailboxCsvFallbackNames' -DefaultValue @()) -ConfiguredPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExoMailboxCsvPath' -DefaultValue '')) -Required:$false
    $onPremPath = Find-ConfiguredCsv -Folder $latestFolder -PrimaryName ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'OnPremMailboxCsvName' -DefaultValue 'Exchange_OnPrem_Mailboxes_AllDomains.csv')) -FallbackNames (Get-SmartM365CommunicationConfigValue -Config $config -Name 'OnPremMailboxCsvFallbackNames' -DefaultValue @()) -ConfiguredPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'OnPremMailboxCsvPath' -DefaultValue '')) -Required:$false
    $exoRows = if ([string]::IsNullOrWhiteSpace($exoPath)) { @() } else { @(Import-SmartM365CommunicationCsv -Path $exoPath) }
    $onPremRows = if ([string]::IsNullOrWhiteSpace($onPremPath)) { @() } else { @(Import-SmartM365CommunicationCsv -Path $onPremPath) }
    $mailboxIndex = Build-MailboxIndex -ExoRows $exoRows -OnPremRows $onPremRows

    $licenseMap = @{}
    $smtpList = @()
    $licensedSmtpList = New-Object System.Collections.ArrayList
    $licensePath = Find-ConfiguredCsv -Folder $latestFolder -PrimaryName ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LicenseCsvName' -DefaultValue 'M365_Licenses_Users.csv')) -FallbackNames (Get-SmartM365CommunicationConfigValue -Config $config -Name 'LicenseCsvFallbackNames' -DefaultValue @()) -ConfiguredPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LicenseCsvPath' -DefaultValue '')) -Required:$false
    $licenseCsvStale = $false
    if (-not [string]::IsNullOrWhiteSpace($licensePath) -and $licenseCsvMaxAgeHours -gt 0) {
        $licenseItem = Get-Item -LiteralPath $licensePath -ErrorAction Stop
        $licenseAgeHours = ((Get-Date) - $licenseItem.LastWriteTime).TotalHours
        if ($licenseAgeHours -gt $licenseCsvMaxAgeHours) {
            $licenseCsvStale = $true
            Write-Host ("License source: CSV is stale ({0:n1}h old, max {1}h): {2}" -f $licenseAgeHours, $licenseCsvMaxAgeHours, $licensePath) -ForegroundColor DarkYellow
            $licensePath = ''
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($licensePath)) {
        Write-Host ("License source: CSV ({0})" -f $licensePath) -ForegroundColor DarkCyan
        $licenseRows = @(Import-SmartM365CommunicationCsv -Path $licensePath)
        foreach ($row in $licenseRows) {
            $sku = ([string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('SkuPartNumber','skuPartNumber','SKU') -DefaultValue '')).Trim().ToUpperInvariant()
            if ([string]::IsNullOrWhiteSpace($sku)) { continue }
            if (-not $targetSkuSet.Contains($sku)) { continue }
            if (-not $mailboxQuotaBySku.ContainsKey($sku)) { continue }
            $smtp = Get-SmartM365CommunicationSmtpAddress -Row $row
            if ([string]::IsNullOrWhiteSpace($smtp)) { continue }
            $smtp = $smtp.Trim().ToLowerInvariant()
            $quota = $mailboxQuotaBySku[$sku]
            $quotaInfo = [pscustomobject]@{
                SkuPartNumber = $quota.SkuPartNumber
                ThresholdMB = [double]$quota.ThresholdMB
                MaxGb = [double]$quota.MaxGb
                QuotaLabel = [string]$quota.QuotaLabel
                LicenseRow = $row
            }
            if (-not $licenseMap.ContainsKey($smtp) -or [double]$quotaInfo.ThresholdMB -gt [double]$licenseMap[$smtp].ThresholdMB) {
                $licenseMap[$smtp] = $quotaInfo
            }
            [void]$licensedSmtpList.Add($smtp)
        }
    }
    elseif ($enableLiveLicenseLookupWhenCsvMissing) {
        $reason = if ($licenseCsvStale) { 'license CSV is stale' } else { 'license CSV not found' }
        Write-Host ("License source: Microsoft Graph live lookup ({0})." -f $reason) -ForegroundColor DarkYellow
        Ensure-GraphLicenseLookup -TenantConfig $tenantConfig
    }
    else {
        $reason = if ($licenseCsvStale) { 'older than LicenseCsvMaxAgeHours' } else { 'not found' }
        throw ("License CSV is {0} and EnableLiveLicenseLookupWhenCsvMissing is disabled." -f $reason)
    }

    if ($FromList) {
        $listRows = @(Import-SmartM365CommunicationCsv -Path $ListCsvPath)
        $smtpList = @($listRows | ForEach-Object { Get-SmartM365CommunicationSmtpAddress -Row $_ })
        if ([string]::IsNullOrWhiteSpace($licensePath)) {
            foreach ($smtpRaw in $smtpList) {
                $smtp = ([string]$smtpRaw).Trim().ToLowerInvariant()
                if ([string]::IsNullOrWhiteSpace($smtp) -or $licenseMap.ContainsKey($smtp)) { continue }
                $quotaInfo = Resolve-LiveLicenseQuotaInfo -SmtpAddress $smtp -QuotaBySku $mailboxQuotaBySku
                if ($quotaInfo) { $licenseMap[$smtp] = $quotaInfo }
            }
        }
    }
    else {
        if ([string]::IsNullOrWhiteSpace($licensePath)) {
            $graphUsers = @(Get-MgUser -All -Property 'DisplayName','UserPrincipalName','Id','Mail','ProxyAddresses' -ErrorAction Stop)
            foreach ($user in $graphUsers) {
                $smtp = (Get-PrimarySmtpAddressFromGraphUser -User $user).Trim().ToLowerInvariant()
                if ([string]::IsNullOrWhiteSpace($smtp)) { continue }
                $quotaInfo = Get-LiveLicenseQuotaInfoForGraphUser -User $user -QuotaBySku $mailboxQuotaBySku
                if ($quotaInfo) {
                    $licenseMap[$smtp] = $quotaInfo
                    [void]$licensedSmtpList.Add($smtp)
                }
            }
        }
        $smtpList = @($licensedSmtpList)
    }

    $candidateResult = Build-Candidates -SmtpList $smtpList -MailboxIndex $mailboxIndex -LicenseMap $licenseMap -Registry $registry -RequireMailboxInventory:$false
    $candidates = @($candidateResult.Candidates)
    $counters.Candidates = $candidates.Count
    Write-Host ("Candidates: {0}" -f $candidates.Count)
    Write-Host ("Preflight: Empty={0}; Duplicate={1}; NotFound={2}; NoTargetLicense={3}; AlreadySent={4}" -f $candidateResult.Preflight.EmptySmtp, $candidateResult.Preflight.DuplicateInRun, $candidateResult.Preflight.NotFoundInIndex, $candidateResult.Preflight.NoTargetLicense, $candidateResult.Preflight.AlreadySent)

    if (-not $WhatIf -and -not $SkipConfirmation) {
        $answer = Read-Host "Send Exchange migration mailbox size reduction notifications to $($candidates.Count) candidate(s)? Type YES to continue"
        if ($answer -ne 'YES') { Write-Host 'Cancelled by operator.'; return }
    }

    $totalCount = $candidates.Count
    $batchCount = [math]::Ceiling($totalCount / [double]$batchSize)
    for ($batchIndex = 0; $batchIndex -lt $batchCount; $batchIndex++) {
        $batchStart = $batchIndex * $batchSize
        $batchEnd = [math]::Min($batchStart + $batchSize - 1, $totalCount - 1)
        $batch = @($candidates[$batchStart..$batchEnd])
        Write-Host ("--- Batch {0}/{1} ({2} email(s)) ---" -f ($batchIndex + 1), $batchCount, $batch.Count) -ForegroundColor Cyan

    foreach ($candidate in $batch) {
        $email = [string]$candidate.Email
        $language = Resolve-SmartM365CommunicationLanguageTag -Row ([pscustomobject]@{ EmailAddress = $email }) -ForceLanguage $ForceLanguage -DefaultLanguageTag $defaultLanguage -DomainLanguageMap $domainLanguageMap
        if ($enableAdLookup -and [string]::IsNullOrWhiteSpace($ForceLanguage)) {
            $adInfo = Resolve-SmartM365CommunicationAdUserInfo -SmtpAddress $email -GcServer $forestGcServer -DefaultLanguageTag $defaultLanguage -ResolveLanguage
            if ($adInfo.Found) {
                if (-not [string]::IsNullOrWhiteSpace($adInfo.PreferredLanguage)) { $language = $adInfo.PreferredLanguage }
            }
        }
        $subject = Get-SmartM365CommunicationSubject -SubjectByLanguage $subjectByLanguage -LanguageTag $language -DefaultSubject $defaultSubject
        $candidateTargetSku = [string]$candidate.TargetSkuPartNumber
        $candidateThresholdMb = ConvertTo-DoubleValue -Value $candidate.MailboxThresholdMB -DefaultValue -1
        $candidateMaxGb = ConvertTo-DoubleValue -Value $candidate.MailboxMaxGb -DefaultValue -1
        $candidateQuotaLabel = [string]$candidate.MailboxQuotaLabel
        if ([string]::IsNullOrWhiteSpace($candidateTargetSku) -or $candidateThresholdMb -lt 0 -or $candidateMaxGb -le 0 -or [string]::IsNullOrWhiteSpace($candidateQuotaLabel)) {
            throw "Candidate '$email' has no resolved target license quota. Check LicenseCsvPath, TargetSkuPartNumbers and MailboxQuotaBySkuPartNumber."
        }

        $recipientTypeDetails = ''
        $exchangeSource = ''
        if ($enableExchangeMailboxStateCheck -and $exchangeSnapInState.Available) {
            $exchangeState = Test-SmartM365CommunicationExchangeMailboxState -SmtpAddress $email
            $recipientTypeDetails = [string]$exchangeState.RecipientTypeDetails
            $exchangeSource = [string]$exchangeState.Source

            if (-not $exchangeState.Exists -and $skipMailboxNotFound -and -not $ForceSend) {
                $counters.Skipped++
                $counters.MailboxNotFound++
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:MailboxNotFound(Exchange)' })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; Email = $email; LanguageTag = $language; MailboxSizeGB = ''; MailboxType = $candidate.MailboxType; MailboxSource = $candidate.MailboxSource; TargetSkuPartNumber = $candidateTargetSku; Subject = $subject; Status = 'Skipped'; SkipReason = 'MailboxNotFound(Exchange)'; ErrorMessage = $exchangeState.ErrorMessage; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                continue
            }

            if ($exchangeState.IsRemoteMailbox -and $skipRemoteMailbox -and -not $ForceSend) {
                $counters.Skipped++
                $counters.RemoteMailbox++
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:IsRemoteMailbox(Exchange)' })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; Email = $email; LanguageTag = $language; MailboxSizeGB = ''; MailboxType = 'RemoteUserMailbox'; MailboxSource = $exchangeMailboxSource; TargetSkuPartNumber = $candidateTargetSku; Subject = $subject; Status = 'Skipped'; SkipReason = 'IsRemoteMailbox(Exchange)'; ErrorMessage = ''; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($recipientTypeDetails)) {
                if ($recipientTypeDetails -like '*Remote*') { $candidate.MailboxType = 'RemoteUserMailbox' }
                else { $candidate.MailboxType = 'UserMailbox' }
                $candidate.MailboxSource = $exchangeMailboxSource
            }
        }

        if ($enableExchangeMailboxUsageCheck -and $exchangeSnapInState.Available) {
            $usageInfo = Get-SmartM365CommunicationMailboxUsageInfo -SmtpAddress $email -MailboxMaxGb $candidateMaxGb
            if ($usageInfo.Available) {
                $candidate.MailboxSizeMB = [math]::Round(([double]$usageInfo.TotalBytes / 1MB), 2)
                $candidate.MailboxSource = $exchangeMailboxSource
            }
            elseif ($skipMailboxUsageUnavailable -and -not $ForceSend) {
                $counters.Skipped++
                $counters.MailboxUsageUnavailable++
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:MailboxUsageUnavailable' })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; Email = $email; LanguageTag = $language; MailboxSizeGB = ''; MailboxType = $candidate.MailboxType; MailboxSource = $candidate.MailboxSource; TargetSkuPartNumber = $candidateTargetSku; Subject = $subject; Status = 'Skipped'; SkipReason = 'MailboxUsageUnavailable'; ErrorMessage = $usageInfo.ErrorMessage; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                continue
            }
        }

        if ([double]$candidate.MailboxSizeMB -ge 0 -and [double]$candidate.MailboxSizeMB -le $candidateThresholdMb -and -not $ForceSend) {
            $counters.Skipped++
            $counters.MailboxUsageBelowThreshold++
            $sizeGbForLog = [math]::Round(([double]$candidate.MailboxSizeMB / 1024), 2)
            [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:MailboxUsageBelowThreshold' })
            Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; Email = $email; LanguageTag = $language; MailboxSizeGB = ("{0} GB" -f $sizeGbForLog); MailboxType = $candidate.MailboxType; MailboxSource = $candidate.MailboxSource; TargetSkuPartNumber = $candidateTargetSku; Subject = $subject; Status = 'Skipped'; SkipReason = ("MailboxUsageBelowThreshold({0}MB<={1}MB)" -f $candidate.MailboxSizeMB, $candidateThresholdMb); ErrorMessage = ''; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
            continue
        }

        $sizeGb = [math]::Round(([double]$candidate.MailboxSizeMB / 1024), 2)
        $sizeLabel = ("{0} GB" -f $sizeGb)
        $webmailUrl = if ($webmailByType.Contains($candidate.MailboxType)) { [string]$webmailByType[$candidate.MailboxType] } elseif ($webmailByType.Contains('default')) { [string]$webmailByType['default'] } else { 'https://outlook.office365.com' }

        try {
            $template = Get-SmartM365CommunicationTemplateContent -TemplateRoot $templateRoot -TemplateBaseName $templateBaseName -LanguageTag $language -DefaultLanguageTag $defaultLanguage
            $logoTokens = Get-SmartM365CommunicationLinkedLogoTokens -LogoPath $logoPath -LogoContentId $logoContentId
            $freeTextBlock = ConvertTo-FreeTextBlock -Text (Get-LocalizedTextValue -TextByLanguage $freeTextByLanguage -LanguageTag $language -DefaultLanguageTag $defaultLanguage)
            $tokens = @{
                MailboxSizeGB = $sizeLabel
                MailboxQuotaGB = $candidateQuotaLabel
                FreeTextBlockStyle = [string]$freeTextBlock.Style
                FreeTextBlockTextHtml = [string]$freeTextBlock.TextHtml
                WebmailUrl = $webmailUrl
                Hotline = Get-SmartM365CommunicationHotline -HotlineByLanguageOrCountry $hotlineMap -LanguageTag $language -DefaultHotline ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultHotline' -DefaultValue ''))
                LogoImgTag = $logoTokens.LogoImgTag
                FooterLogoImgTag = $logoTokens.FooterLogoImgTag
            }
            $html = Expand-SmartM365CommunicationTemplate -TemplateContent $template.Content -Tokens $tokens
            Assert-SmartM365CommunicationNoUnresolvedToken -Html $html

            $result = Send-SmartM365CommunicationMail -MailSendMode $mailSendMode -SmtpServer $smtpServer -SmtpPort $smtpPort -From $from -To $email -Bcc $bccAll -Subject $subject -BodyHtml $html -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $global:Thumbprint -SmtpUseIntegratedAuth $smtpUseIntegratedAuth -SmtpEnableSsl $smtpEnableSsl -LogoPath $logoPath -LogoContentId $logoContentId -LogoMediaType $logoMediaType -RetryCount $smtpRetryCount -RetryDelaySeconds $smtpRetryDelaySeconds -WhatIf:$WhatIf
            $teamsStatus = 'Disabled'
            $teamsError = ''
            try {
                if ((Get-SmartM365CommunicationTeamsUserMode -TeamsUserMessageMode $teamsUserMessageMode) -ne 'Disabled' -and ($WhatIf -or $result.Sent)) {
                    $teamsMessageTemplate = Get-SmartM365CommunicationSubject -SubjectByLanguage $teamsUserMessageByLanguage -LanguageTag $language -DefaultSubject ''
                    $teamsMessageText = Expand-SmartM365CommunicationTemplate -TemplateContent $teamsMessageTemplate -Tokens $tokens
                    $teamsResult = Send-SmartM365CommunicationTeamsUserMessage -TeamsUserMessageMode $teamsUserMessageMode -To $email -MessageText $teamsMessageText -TenantId $global:TenantId -SenderUserId $teamsUserSenderUserId -RetryCount $teamsUserRetryCount -RetryDelaySeconds $teamsUserRetryDelaySeconds -WhatIf:$WhatIf
                    if ($WhatIf) { $counters.TeamsDryRun++; $teamsStatus = 'DryRun' }
                    elseif ($teamsResult.Sent) { $counters.TeamsSent++; $teamsStatus = 'Success' }
                    else { $counters.TeamsFailed++; $teamsStatus = $teamsResult.Mode; $teamsError = [string]$teamsResult.Error }
                }
            }
            catch {
                $counters.TeamsFailed++
                $teamsStatus = 'Failed'
                $teamsError = $_.Exception.Message
            }
            if (-not $WhatIf -and $intraEmailDelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $intraEmailDelayMilliseconds }
            if ($WhatIf) { $counters.DryRun++ }
            elseif ($result.Sent) { $counters.Sent++; if ($preventResend) { Register-SmartM365CommunicationSentItem -Registry $registry -Email $email } }
            else { $counters.Skipped++ }
            $status = if ($WhatIf) { 'DryRun' } elseif ($result.Sent) { 'Success' } else { $result.Mode }
            [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = $status })
            Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; Email = $email; LanguageTag = $language; MailboxSizeGB = $sizeLabel; MailboxType = $candidate.MailboxType; MailboxSource = $candidate.MailboxSource; TargetSkuPartNumber = $candidateTargetSku; Subject = $subject; Status = $status; TeamsStatus = $teamsStatus; TeamsError = $teamsError; SkipReason = ''; ErrorMessage = $result.Mode; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
        }
        catch {
            $counters.Failed++
            [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Failed' })
            Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; Email = $email; LanguageTag = $language; MailboxSizeGB = $sizeLabel; MailboxType = $candidate.MailboxType; MailboxSource = $candidate.MailboxSource; TargetSkuPartNumber = $candidateTargetSku; Subject = $subject; Status = 'Failed'; SkipReason = ''; ErrorMessage = $_.Exception.Message; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
        }
    }

        if (-not $WhatIf -and ($batchIndex + 1) -lt $batchCount) {
            if ($extendedPauseEvery -gt 0 -and (($batchIndex + 1) % $extendedPauseEvery) -eq 0) {
                Write-Host ("--- Batch {0}/{1} complete. Extended pause: {2}s ---" -f ($batchIndex + 1), $batchCount, $extendedPauseSeconds) -ForegroundColor DarkYellow
                if ($extendedPauseSeconds -gt 0) { Start-Sleep -Seconds $extendedPauseSeconds }
            }
            elseif ($batchPauseSeconds -gt 0) {
                Write-Host ("--- Batch {0}/{1} complete. Pausing {2}s before next batch... ---" -f ($batchIndex + 1), $batchCount, $batchPauseSeconds) -ForegroundColor DarkCyan
                Start-Sleep -Seconds $batchPauseSeconds
            }
        }
    }

    if ($preventResend -and -not $WhatIf) { Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $sentRegistryPath }

    $summary = "Candidates=$($counters.Candidates); Sent=$($counters.Sent); DryRun=$($counters.DryRun); Failed=$($counters.Failed); Skipped=$($counters.Skipped); MailboxNotFound=$($counters.MailboxNotFound); RemoteMailbox=$($counters.RemoteMailbox); MailboxUsageUnavailable=$($counters.MailboxUsageUnavailable); MailboxUsageBelowThreshold=$($counters.MailboxUsageBelowThreshold); TeamsSent=$($counters.TeamsSent); TeamsDryRun=$($counters.TeamsDryRun); TeamsFailed=$($counters.TeamsFailed)."
    if (-not $WhatIf -and -not $NoSummaryEmail -and -not [string]::IsNullOrWhiteSpace($summaryTo)) {
        $summaryHtml = New-SmartM365CommunicationSummaryHtml -Title $summaryTitle -Facts @{
            RunId = $runId; Tenant = $Tenant; Mode = $(if ($WhatIf) { 'DryRun' } else { $(if ($FromList) { 'FromList' } else { 'Inventory' }) }); Summary = $summary; LogPath = $logPath; SentRegistryPath = $sentRegistryPath; ExchangeManagement = ("{0}:{1}" -f $exchangeManagementState.Source, $exchangeManagementState.Status)
        } -Items @($processedItems)
        Send-SmartM365CommunicationMail -MailSendMode $mailSendMode -SmtpServer $smtpServer -SmtpPort $smtpPort -From $from -To $summaryTo -Bcc $summaryBcc -Subject $summarySubject -BodyHtml $summaryHtml -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $global:Thumbprint -SmtpUseIntegratedAuth $smtpUseIntegratedAuth -SmtpEnableSsl $smtpEnableSsl -RetryCount $smtpRetryCount -RetryDelaySeconds $smtpRetryDelaySeconds | Out-Null
    }

    if (-not $WhatIf) {
        Send-SmartM365TeamsNotification -Level SUCCESS -Title $teamsSuccessTitle -Message $summary -ResultSummary $summary -Facts @{ Tenant = $Tenant; RunId = $runId; LogPath = $logPath; ExchangeManagement = ("{0}:{1}" -f $exchangeManagementState.Source, $exchangeManagementState.Status) } | Out-Null
    }
    Write-Host $summary
}
catch {
    $message = $_.Exception.Message
    if (-not $WhatIf) {
        try { Send-SmartM365TeamsNotification -Level ERROR -Title $teamsFailureTitle -Message $message -Facts @{ Tenant = $Tenant; RunId = $runId; LogPath = $logPath } | Out-Null } catch {}
    }
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfV2+iBLus91QE
# YDY7bsyZJP6SZsd62zVAiu21eoXyaKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAz+OflSOA1c9Q6f/dXXahOYLBYe4gi94Eg7P+YjKs/pTANBgkqhkiG9w0B
# AQEFAASCAYBG5tNwNJcKVwu3f7Y/+Pu6vPlousjSeyUGC1feJzYIf0ema2WVV0J8
# OX6Gy4m3S7hT7o6k4ywXZmYDxUZhXGH9VB9WdTHv5eKoGSv5UZP26dIkCV1rRL1u
# Gc8B8j6oPuJxiw12F02jMmm3Rh5zeyqnp+yUNRka8SmqesElJDc+0NB8qUuxPEyE
# 1LfShTjs8BmOHQ8b3E1xyNDtsC7xT11HIuojPMLIVkbCaN4PZS49qbpO7nY4rCRv
# aBPa/3nCqnP4D/FuUyyg3yzeAS6hHMacPrqcgUc7HCDGapAkrNzFsbvIXXxpn027
# cBk6fbRb+3M+C6OCsCIn93mWlfGg1N44Ct1WXdn7FdyEVy8FiRrmfAwkCAhUWEsQ
# 1EGWyct1CZYm/UjFACpA7s9FHRw50xc+IeFxEw8McCKhZYPpoWM4ymxSyeRUzTTQ
# EC9QxdTXyl5vftraG66v3Y5UBl53Vj6DAq7303tsozUcXcxd6bKARbK/fNlZ1c8H
# 4l3XUS1nxjU=
# SIG # End signature block
