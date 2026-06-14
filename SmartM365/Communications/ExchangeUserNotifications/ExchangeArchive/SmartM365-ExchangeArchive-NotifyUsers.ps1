<#
.SYNOPSIS
Notifies users whose primary mailbox is approaching the configured limit that Online Archive will be enabled.

.DESCRIPTION
Loads SmartM365 tenant context, the shared Communications configuration, and the ExchangeArchive
campaign configuration, then sends localized HTML notification emails to recipients listed in one CSV
file or in every CSV file from a recipients folder.

The script evaluates mailbox usage from each recipient row, skips users below the configured size
threshold unless ForceSend is used, resolves the language from the row or email domain, injects the
archive date and localized mailbox usage text, prevents duplicate sends through the sent registry,
writes a per-recipient log, and can send a summary email and Teams notification after a live run.

Expected recipient columns include an email column such as PrimarySmtpAddress, EmailAddress, Mail,
UserPrincipalName, or Recipient, plus optional UserName, DisplayName, Name, LanguageTag,
PreferredLanguage, EffectiveDate, Date, ArchiveDate, MailboxTotalGb, TotalGb, TotalItemSizeGB,
MailboxSizeGB, MailboxUsagePercent, UsagePercent, or Percent.

.PARAMETER Tenant
SmartM365 tenant context key used to resolve root paths, mail credentials, Teams settings, and defaults.

.PARAMETER RecipientsPath
CSV file or folder containing recipient CSV files. When omitted, the path comes from the campaign JSON.

.PARAMETER CampaignConfigPath
Optional campaign JSON override. When omitted, the script uses Config/Campaigns/ExchangeArchive.local.json.

.PARAMETER ForceLanguage
Forces every email to use one language tag, such as fr-FR or en, regardless of CSV/domain detection.

.PARAMETER ForceSend
Bypasses duplicate, past-date, and mailbox-size skip checks for operator-controlled replays.

.PARAMETER NoSummaryEmail
Suppresses the end-of-run summary email. Logs and console summary are still produced.

.PARAMETER WhatIf
Performs a dry run: resolves recipients/templates and writes logs, but does not send user, summary, or Teams notifications.

.EXAMPLE
.\SmartM365-ExchangeArchive-NotifyUsers.ps1 -Tenant prod -WhatIf

Runs the configured archive campaign in dry-run mode using the recipients path and thresholds from JSON.

.EXAMPLE
.\SmartM365-ExchangeArchive-NotifyUsers.ps1 -Tenant prod -RecipientsPath .\Recipients\archive-pilot.csv -ForceSend

Sends the archive notification to a pilot CSV even when rows are below the configured mailbox-size threshold.
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$RecipientsPath = '',
    [string]$CampaignConfigPath = '',
    [string]$ForceLanguage = '',
    [switch]$ForceSend,
    [switch]$NoSummaryEmail,
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

function Get-RecipientFiles {
    param([Parameter(Mandatory)][string]$Path)
    if (Test-Path -LiteralPath $Path -PathType Leaf) { return @(Get-Item -LiteralPath $Path) }
    if (Test-Path -LiteralPath $Path -PathType Container) { return @(Get-ChildItem -LiteralPath $Path -Filter '*.csv' -File | Sort-Object Name) }
    throw "Recipients path not found: $Path"
}

function Resolve-EffectiveDate {
    param($Row, [string]$DefaultEffectiveDate)
    $raw = [string](Get-SmartM365CommunicationProperty -InputObject $Row -Names @('EffectiveDate','Date','ArchiveDate') -DefaultValue $DefaultEffectiveDate)
    if ([string]::IsNullOrWhiteSpace($raw)) { return [pscustomobject]@{ Text = ''; Date = $null } }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($raw, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed) -or
        [datetime]::TryParse($raw, [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed) -or
        [datetime]::TryParse($raw, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        return [pscustomobject]@{ Text = $parsed.ToString('yyyy-MM-dd'); Date = $parsed }
    }
    return [pscustomobject]@{ Text = $raw; Date = $null }
}

function ConvertTo-DoubleValue {
    param([AllowNull()]$Value, [double]$DefaultValue = -1)
    if ($null -eq $Value) { return $DefaultValue }
    $text = ([string]$Value).Trim().Replace(',', '.')
    $parsed = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) { return $parsed }
    return $DefaultValue
}

$tenantContextPath = Find-UpwardFile -FileName 'Config\SmartM365-TenantContext.ps1'
. $tenantContextPath
$effectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

$coreModulePath = Find-UpwardFile -FileName 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module $coreModulePath -Force

$commModulePath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'SmartM365.Communications.psm1'
Import-Module $commModulePath -Force -DisableNameChecking

$tenantConfig = Convert-ToHash $effectiveConfig
$global:AppId = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'AppId' -DefaultValue '')
$global:TenantId = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'TenantId' -DefaultValue '')
$global:Thumb = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'Thumb' -DefaultValue '')
$global:Thumbprint = [string](Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'Thumbprint' -DefaultValue $global:Thumb)

$pathTokens = @{
    TenantKey = $Tenant
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

$campaignTemplatePath = Join-Path $commRoot 'ExchangeUserNotifications\Config\Campaigns\ExchangeArchive.local.json.template'
if ([string]::IsNullOrWhiteSpace($CampaignConfigPath)) {
    $CampaignConfigPath = Join-Path $commRoot 'ExchangeUserNotifications\Config\Campaigns\ExchangeArchive.local.json'
}
$campaignConfig = Merge-SmartM365CommunicationConfig -BaseConfig (Read-SmartM365CommunicationJson -Path $campaignTemplatePath) -OverlayConfig (Read-SmartM365CommunicationJson -Path $CampaignConfigPath)
$config = Merge-SmartM365CommunicationConfig -BaseConfig $baseConfig -OverlayConfig $campaignConfig

$campaignName = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'CampaignName' -DefaultValue 'ExchangeArchive')
$runId = [guid]::NewGuid().ToString()
$runOutputPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RunOutputPath' -DefaultValue ''))
$logOutputPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogOutputPath' -DefaultValue $runOutputPath))
$sentRegistryPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SentRegistryPath' -DefaultValue (Join-Path $runOutputPath 'Data\SentRegistry.csv')))
$logPath = Join-Path -Path $logOutputPath -ChildPath ("SmartM365-ExchangeArchive-NotifyUsers-Log_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
foreach ($folder in @($runOutputPath, $logOutputPath, (Split-Path -Path $sentRegistryPath -Parent))) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
}

Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $runOutputPath -LatestOutputRoot $pathTokens.LatestCsvFolderPath -LogPath (Join-Path $logOutputPath 'SmartM365-ExchangeArchive-NotifyUsers.log')

if ([string]::IsNullOrWhiteSpace($RecipientsPath)) {
    $RecipientsPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RecipientsPath' -DefaultValue ''))
}
$templateRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Templates'
$templateBaseName = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TemplateBaseName' -DefaultValue 'SmartM365-ExchangeArchive-NotifyUsers-Template')
$logoPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogoPath' -DefaultValue (Join-Path $pathTokens.SmartM365RootPath 'SmartM365-logo.ico')))
$logoContentId = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogoContentId' -DefaultValue 'smartm365logo')
$logoMediaType = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogoMediaType' -DefaultValue 'image/x-icon')
$defaultLanguage = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultLanguageTag' -DefaultValue 'en' -FallbackConfig $baseConfig)
$from = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'From' -DefaultValue '' -FallbackConfig $tenantConfig)
$smtpServer = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpServer' -DefaultValue '' -FallbackConfig $tenantConfig)
$relayIp = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RelayIp' -DefaultValue '' -FallbackConfig $tenantConfig)
$smtpResolveIPv4 = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpResolveIPv4' -DefaultValue $true -FallbackConfig $baseConfig)
if (-not [string]::IsNullOrWhiteSpace($relayIp)) { $smtpServer = $relayIp }
elseif ($smtpResolveIPv4 -and -not [string]::IsNullOrWhiteSpace($smtpServer)) { $smtpServer = Resolve-SmartM365CommunicationIPv4Address -HostName $smtpServer }
$smtpPort = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpPort' -DefaultValue 25 -FallbackConfig $tenantConfig)
$smtpUseIntegratedAuth = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpUseIntegratedAuth' -DefaultValue $false -FallbackConfig $tenantConfig)
$smtpEnableSsl = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpEnableSsl' -DefaultValue $false -FallbackConfig $tenantConfig)
$smtpRetryCount = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpRetryCount' -DefaultValue 1 -FallbackConfig $baseConfig)
$smtpRetryDelaySeconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SmtpRetryDelaySeconds' -DefaultValue 2 -FallbackConfig $baseConfig)
$intraEmailDelayMilliseconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'IntraEmailDelayMilliseconds' -DefaultValue 150 -FallbackConfig $baseConfig)
$summaryTo = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryTo' -DefaultValue (Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'ErrorMailTo' -DefaultValue ''))
$summaryBcc = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryBcc' -DefaultValue '')
$bccAll = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'BccAll' -DefaultValue '')
$defaultEffectiveDate = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultEffectiveDate' -DefaultValue '')
$skipPast = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipWhenEffectiveDateIsPast' -DefaultValue $true)
$preventResend = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'PreventResendAcrossRuns' -DefaultValue $true)
$mailboxMaxGb = [double](Get-SmartM365CommunicationConfigValue -Config $config -Name 'MailboxMaxGb' -DefaultValue 100)
$mailboxMinGbToNotify = [double](Get-SmartM365CommunicationConfigValue -Config $config -Name 'MailboxMinGbToNotify' -DefaultValue 70)
$subjectByLanguage = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'SubjectByLanguage' -DefaultValue @{})
$usageTextByLanguage = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'MailboxUsageTextByLanguage' -DefaultValue @{})
$defaultSubject = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultSubject' -DefaultValue (Get-SmartM365CommunicationConfigValue -Config $subjectByLanguage -Name 'default' -DefaultValue $campaignName))
$defaultUsageText = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultMailboxUsageText' -DefaultValue (Get-SmartM365CommunicationConfigValue -Config $usageTextByLanguage -Name 'default' -DefaultValue ''))
$unknownUsageLabel = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'UnknownUsageLabel' -DefaultValue '')
$summaryTitle = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryTitle' -DefaultValue $campaignName)
$summarySubject = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummarySubject' -DefaultValue $campaignName)
$teamsSuccessTitle = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsSuccessTitle' -DefaultValue $campaignName)
$teamsFailureTitle = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TeamsFailureTitle' -DefaultValue $campaignName)
$hotlineMap = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'HotlineByLanguageOrCountry' -DefaultValue @{} -FallbackConfig $baseConfig)
$domainLanguageMap = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'DomainLanguageMap' -DefaultValue @{} -FallbackConfig $baseConfig)
$enableAdLookup = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableAdLookupForLanguageAndName' -DefaultValue $true)
$forestGcServer = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ForestGcServer' -DefaultValue '')
$enableExchangeMailboxStateCheck = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchangeMailboxStateCheck' -DefaultValue $true)
$enableExchangeMailboxUsageCheck = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchangeMailboxUsageCheck' -DefaultValue $true)
$skipMailboxNotFound = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipMailboxNotFoundWhenExchangeCheckAvailable' -DefaultValue $true)
$skipRemoteMailbox = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipRemoteMailboxWhenExchangeCheckAvailable' -DefaultValue $true)
$skipMailboxUsageUnavailable = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipMailboxUsageUnavailableWhenExchangeCheckAvailable' -DefaultValue $true)
$requireExchange2016SnapIn = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RequireExchange2016SnapIn' -DefaultValue $true)
$enableExchange2016SnapIn = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchange2016SnapIn' -DefaultValue $true)
if ($requireExchange2016SnapIn) { $enableExchange2016SnapIn = $true }
$exchangeSnapInState = Initialize-SmartM365CommunicationExchangeSnapIn `
    -Enabled $enableExchange2016SnapIn `
    -SnapInName ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInName' -DefaultValue 'Microsoft.Exchange.Management.PowerShell.SnapIn')) `
    -ViewEntireForest ([bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInViewEntireForest' -DefaultValue $true))
if ($exchangeSnapInState.Enabled) {
    if ($exchangeSnapInState.Available) {
        $message = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInAvailableMessage' -DefaultValue '')
        if (-not [string]::IsNullOrWhiteSpace($message)) { Write-Host $message -ForegroundColor DarkCyan }
        if ($exchangeSnapInState.ViewEntireForestApplied) {
            $message = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInForestViewMessage' -DefaultValue '')
            if (-not [string]::IsNullOrWhiteSpace($message)) { Write-Host $message -ForegroundColor DarkCyan }
        }
        elseif (-not [string]::IsNullOrWhiteSpace($exchangeSnapInState.ForestErrorMessage)) {
            $format = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInForestViewWarningFormat' -DefaultValue '')
            if (-not [string]::IsNullOrWhiteSpace($format)) { Write-Host ($format -f $exchangeSnapInState.ForestErrorMessage) -ForegroundColor Yellow }
        }
    }
    else {
        $message = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ExchangeSnapInUnavailableMessage' -DefaultValue '')
        if (-not [string]::IsNullOrWhiteSpace($message)) { Write-Host $message -ForegroundColor DarkYellow }
    }
}
if ($requireExchange2016SnapIn -and -not $exchangeSnapInState.Available) {
    throw ("Exchange 2016 snap-in is mandatory for this campaign and is not available. Status={0}; Error={1}" -f $exchangeSnapInState.Status, $exchangeSnapInState.ErrorMessage)
}

$columns = @('Timestamp','RunId','Campaign','BatchFile','Email','UserName','NameSource','LanguageTag','Scope','MailboxTotalGb','MailboxUsagePercent','Subject','Status','SkipReason','ErrorMessage','RecipientTypeDetails','ExchangeSource')
$counters = @{ Files = 0; Rows = 0; Sent = 0; DryRun = 0; Failed = 0; Skipped = 0; AlreadySent = 0; MailboxNotFound = 0; RemoteMailbox = 0; MailboxUsageUnavailable = 0; MailboxUsageBelowThreshold = 0 }
$processedItems = New-Object System.Collections.ArrayList
$seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$registry = Load-SmartM365CommunicationSentRegistry -Path $sentRegistryPath

try {
    foreach ($file in (Get-RecipientFiles -Path $RecipientsPath)) {
        $counters.Files++
        $rows = @(Import-SmartM365CommunicationCsv -Path $file.FullName -RequiredColumns @())
        foreach ($row in $rows) {
            $counters.Rows++
            $email = Get-SmartM365CommunicationSmtpAddress -Row $row
            if ([string]::IsNullOrWhiteSpace($email)) { $counters.Skipped++; continue }

            $totalGb = ConvertTo-DoubleValue -Value (Get-SmartM365CommunicationProperty -InputObject $row -Names @('MailboxTotalGb','TotalGb','TotalItemSizeGB','MailboxSizeGB') -DefaultValue -1)
            $usagePercent = ConvertTo-DoubleValue -Value (Get-SmartM365CommunicationProperty -InputObject $row -Names @('MailboxUsagePercent','UsagePercent','Percent') -DefaultValue -1)
            if ($usagePercent -lt 0 -and $totalGb -ge 0 -and $mailboxMaxGb -gt 0) { $usagePercent = [math]::Round((100.0 * $totalGb / $mailboxMaxGb), 1) }

            $skipReason = ''
            if (-not $seen.Add($email) -and -not $ForceSend) { $skipReason = 'DuplicateInRun' }
            elseif ($preventResend -and $registry.ContainsKey($email) -and -not $ForceSend) { $skipReason = 'AlreadySent' }

            $effectiveDate = Resolve-EffectiveDate -Row $row -DefaultEffectiveDate $defaultEffectiveDate
            if ([string]::IsNullOrWhiteSpace($skipReason) -and $skipPast -and $effectiveDate.Date -and $effectiveDate.Date.Date -lt (Get-Date).Date -and -not $ForceSend) {
                $skipReason = 'EffectiveDateInPast'
            }

            $userName = [string](Get-SmartM365CommunicationProperty -InputObject $row -Names @('UserName','DisplayName','Name') -DefaultValue (($email -split '@')[0]))
            $nameSource = 'CSV'
            $language = Resolve-SmartM365CommunicationLanguageTag -Row $row -ForceLanguage $ForceLanguage -DefaultLanguageTag $defaultLanguage -DomainLanguageMap $domainLanguageMap
            $exchangeLanguageSource = ''
            if ([string]::IsNullOrWhiteSpace($ForceLanguage) -and $exchangeSnapInState.Available) {
                $exchangeLanguage = Resolve-SmartM365CommunicationExchangeLanguageTag -SmtpAddress $email -DefaultLanguageTag $defaultLanguage -DomainLanguageMap $domainLanguageMap
                if (-not [string]::IsNullOrWhiteSpace($exchangeLanguage.LanguageTag)) { $language = $exchangeLanguage.LanguageTag }
                if ($exchangeLanguage.Source -ne 'Default') { $exchangeLanguageSource = [string]$exchangeLanguage.Source }
            }
            if ($enableAdLookup) {
                $adInfo = Resolve-SmartM365CommunicationAdUserInfo -SmtpAddress $email -GcServer $forestGcServer -DefaultLanguageTag $defaultLanguage -ResolveDisplayName -ResolveLanguage
                if ($adInfo.Found) {
                    if (-not [string]::IsNullOrWhiteSpace($adInfo.DisplayName)) { $userName = $adInfo.DisplayName; $nameSource = $adInfo.Source }
                    if ([string]::IsNullOrWhiteSpace($ForceLanguage) -and [string]::IsNullOrWhiteSpace($exchangeLanguageSource) -and -not [string]::IsNullOrWhiteSpace($adInfo.PreferredLanguage)) { $language = $adInfo.PreferredLanguage }
                }
            }
            $subject = Get-SmartM365CommunicationSubject -SubjectByLanguage $subjectByLanguage -LanguageTag $language -DefaultSubject $defaultSubject

            if (-not [string]::IsNullOrWhiteSpace($skipReason)) {
                if ($skipReason -eq 'AlreadySent') { $counters.AlreadySent++ } else { $counters.Skipped++ }
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = "Skipped:$skipReason" })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; UserName = $userName; NameSource = $nameSource; LanguageTag = $language; Scope = 'Preflight'; MailboxTotalGb = $totalGb; MailboxUsagePercent = $usagePercent; Subject = $subject; Status = 'Skipped'; SkipReason = $skipReason; ErrorMessage = ''; RecipientTypeDetails = ''; ExchangeSource = '' })
                continue
            }

            $recipientTypeDetails = ''
            $exchangeSource = ''
            if ($enableExchangeMailboxStateCheck -and $exchangeSnapInState.Available) {
                $exchangeState = Test-SmartM365CommunicationExchangeMailboxState -SmtpAddress $email
                $recipientTypeDetails = [string]$exchangeState.RecipientTypeDetails
                $exchangeSource = [string]$exchangeState.Source

                if (-not $exchangeState.Exists -and $skipMailboxNotFound -and -not $ForceSend) {
                    $counters.MailboxNotFound++
                    [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:MailboxNotFound(OnPrem)' })
                    Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; UserName = $userName; NameSource = $nameSource; LanguageTag = $language; Scope = 'OnPremExchange'; MailboxTotalGb = $totalGb; MailboxUsagePercent = $usagePercent; Subject = $subject; Status = 'Skipped'; SkipReason = 'MailboxNotFound(OnPrem)'; ErrorMessage = $exchangeState.ErrorMessage; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                    continue
                }

                if ($exchangeState.IsRemoteMailbox -and $skipRemoteMailbox -and -not $ForceSend) {
                    $counters.RemoteMailbox++
                    [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:IsRemoteMailbox(OnPrem)' })
                    Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; UserName = $userName; NameSource = $nameSource; LanguageTag = $language; Scope = 'OnPremExchange'; MailboxTotalGb = $totalGb; MailboxUsagePercent = $usagePercent; Subject = $subject; Status = 'Skipped'; SkipReason = 'IsRemoteMailbox(OnPrem)'; ErrorMessage = ''; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                    continue
                }

                if ([string]::IsNullOrWhiteSpace($userName) -or $nameSource -eq 'CSV') {
                    if (-not [string]::IsNullOrWhiteSpace($exchangeState.DisplayName)) { $userName = $exchangeState.DisplayName; $nameSource = 'ExchangeRecipient' }
                }
            }

            if ($enableExchangeMailboxUsageCheck -and $exchangeSnapInState.Available) {
                $usageInfo = Get-SmartM365CommunicationMailboxUsageInfo -SmtpAddress $email -MailboxMaxGb $mailboxMaxGb
                if ($usageInfo.Available) {
                    $totalGb = $usageInfo.TotalGb
                    $usagePercent = $usageInfo.Percent
                }
                elseif ($skipMailboxUsageUnavailable -and -not $ForceSend) {
                    $counters.MailboxUsageUnavailable++
                    [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:MailboxUsageUnavailable' })
                    Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; UserName = $userName; NameSource = $nameSource; LanguageTag = $language; Scope = 'OnPremExchange'; MailboxTotalGb = $totalGb; MailboxUsagePercent = $usagePercent; Subject = $subject; Status = 'Skipped'; SkipReason = 'MailboxUsageUnavailable'; ErrorMessage = $usageInfo.ErrorMessage; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                    continue
                }
            }

            if ($totalGb -ge 0 -and $totalGb -lt $mailboxMinGbToNotify -and -not $ForceSend) {
                $counters.MailboxUsageBelowThreshold++
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:MailboxUsageBelowThreshold' })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; UserName = $userName; NameSource = $nameSource; LanguageTag = $language; Scope = 'OnPremExchange'; MailboxTotalGb = $totalGb; MailboxUsagePercent = $usagePercent; Subject = $subject; Status = 'Skipped'; SkipReason = ("MailboxUsageBelowThreshold({0}GB<{1}GB)" -f $totalGb, $mailboxMinGbToNotify); ErrorMessage = ''; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                continue
            }

            try {
                $template = Get-SmartM365CommunicationTemplateContent -TemplateRoot $templateRoot -TemplateBaseName $templateBaseName -LanguageTag $language -DefaultLanguageTag $defaultLanguage
                $usageTemplate = Get-SmartM365CommunicationSubject -SubjectByLanguage $usageTextByLanguage -LanguageTag $language -DefaultSubject $defaultUsageText
                $usageText = (Expand-SmartM365CommunicationTemplate -TemplateContent $usageTemplate -Tokens @{ MailboxUsagePercent = $(if ($usagePercent -ge 0) { $usagePercent } else { $unknownUsageLabel }); MailboxMaxGb = $mailboxMaxGb })
                $logoTokens = Get-SmartM365CommunicationLinkedLogoTokens -LogoPath $logoPath -LogoContentId $logoContentId
                $tokens = @{
                    UserName = $userName
                    Date = $(if ($effectiveDate.Text) { $effectiveDate.Text } else { (Get-Date).ToString('yyyy-MM-dd') })
                    Hotline = Get-SmartM365CommunicationHotline -HotlineByLanguageOrCountry $hotlineMap -LanguageTag $language -DefaultHotline ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultHotline' -DefaultValue ''))
                    MailboxUsageText = $usageText
                    LogoImgTag = $logoTokens.LogoImgTag
                    FooterLogoImgTag = $logoTokens.FooterLogoImgTag
                }
                $html = Expand-SmartM365CommunicationTemplate -TemplateContent $template.Content -Tokens $tokens
                Assert-SmartM365CommunicationNoUnresolvedToken -Html $html

                $result = Send-SmartM365CommunicationMail -SmtpServer $smtpServer -SmtpPort $smtpPort -From $from -To $email -Bcc $bccAll -Subject $subject -BodyHtml $html -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $global:Thumbprint -SmtpUseIntegratedAuth $smtpUseIntegratedAuth -SmtpEnableSsl $smtpEnableSsl -LogoPath $logoPath -LogoContentId $logoContentId -LogoMediaType $logoMediaType -RetryCount $smtpRetryCount -RetryDelaySeconds $smtpRetryDelaySeconds -WhatIf:$WhatIf
                if (-not $WhatIf -and $intraEmailDelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $intraEmailDelayMilliseconds }
                if ($WhatIf) { $counters.DryRun++ } else { $counters.Sent++; if ($preventResend) { Register-SmartM365CommunicationSentItem -Registry $registry -Email $email } }
                $status = if ($WhatIf) { 'DryRun' } else { 'Success' }
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = $status })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; UserName = $userName; NameSource = $nameSource; LanguageTag = $language; Scope = 'Domain'; MailboxTotalGb = $totalGb; MailboxUsagePercent = $usagePercent; Subject = $subject; Status = $status; SkipReason = ''; ErrorMessage = $result.Mode; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
            }
            catch {
                $counters.Failed++
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Failed' })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; UserName = $userName; NameSource = $nameSource; LanguageTag = $language; Scope = 'Domain'; MailboxTotalGb = $totalGb; MailboxUsagePercent = $usagePercent; Subject = $subject; Status = 'Failed'; SkipReason = ''; ErrorMessage = $_.Exception.Message; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
            }
        }
    }

    if ($preventResend -and -not $WhatIf) { Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $sentRegistryPath }

    $summary = "Files=$($counters.Files); Rows=$($counters.Rows); Sent=$($counters.Sent); DryRun=$($counters.DryRun); Failed=$($counters.Failed); Skipped=$($counters.Skipped); AlreadySent=$($counters.AlreadySent); MailboxNotFound=$($counters.MailboxNotFound); RemoteMailbox=$($counters.RemoteMailbox); MailboxUsageUnavailable=$($counters.MailboxUsageUnavailable); MailboxUsageBelowThreshold=$($counters.MailboxUsageBelowThreshold)."
    if (-not $WhatIf -and -not $NoSummaryEmail -and -not [string]::IsNullOrWhiteSpace($summaryTo)) {
        $summaryHtml = New-SmartM365CommunicationSummaryHtml -Title $summaryTitle -Facts @{
            RunId = $runId; Tenant = $Tenant; Mode = $(if ($WhatIf) { 'DryRun' } else { 'Live' }); Summary = $summary; LogPath = $logPath; SentRegistryPath = $sentRegistryPath; ExchangeSnapIn = $exchangeSnapInState.Status
        } -Items @($processedItems)
        Send-SmartM365CommunicationMail -SmtpServer $smtpServer -SmtpPort $smtpPort -From $from -To $summaryTo -Bcc $summaryBcc -Subject $summarySubject -BodyHtml $summaryHtml -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $global:Thumbprint -SmtpUseIntegratedAuth $smtpUseIntegratedAuth -SmtpEnableSsl $smtpEnableSsl -RetryCount $smtpRetryCount -RetryDelaySeconds $smtpRetryDelaySeconds | Out-Null
    }

    if (-not $WhatIf) {
        Send-SmartM365TeamsNotification -Level SUCCESS -Title $teamsSuccessTitle -Message $summary -ResultSummary $summary -Facts @{ Tenant = $Tenant; RunId = $runId; LogPath = $logPath } | Out-Null
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
