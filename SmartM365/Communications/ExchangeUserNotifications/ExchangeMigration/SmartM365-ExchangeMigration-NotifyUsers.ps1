<#
.SYNOPSIS
Notifies users that their mailbox is moving from the legacy webmail platform to Exchange Online.

.DESCRIPTION
Loads SmartM365 tenant context, the shared Communications configuration, and the ExchangeMigration
campaign configuration, then sends localized HTML notification emails to recipients listed in one CSV
file or in every CSV file from a recipients folder.

The script resolves language from the recipient row or email domain, applies the configured subject,
hotline, old webmail URL, new webmail URL, and effective migration date, prevents duplicate sends by
using the campaign sent registry, writes a per-recipient log, and can send a summary email and Teams
notification at the end of a live run.

Expected recipient columns include an email column such as PrimarySmtpAddress, EmailAddress, Mail,
UserPrincipalName, or Recipient, plus optional LanguageTag, PreferredLanguage, EffectiveDate, Date,
or MigrationDate. Dates must use yyyy-MM-dd when provided in CSV.

.PARAMETER Tenant
SmartM365 tenant context key used to resolve root paths, mail credentials, Teams settings, and defaults.

.PARAMETER RecipientsPath
CSV file or folder containing recipient CSV files. When omitted, the path comes from the campaign JSON.

.PARAMETER CampaignConfigPath
Optional campaign JSON override. When omitted, the script uses Config/Campaigns/ExchangeMigration.local.json.

.PARAMETER ForceLanguage
Forces every email to use one language tag, such as fr-FR or en, regardless of CSV/domain detection.

.PARAMETER ForceEffectiveDate
Forces every recipient to use one migration date, overriding EffectiveDate, Date, and MigrationDate CSV values.
Use yyyy-MM-dd; the GUI date picker sends this format.

.PARAMETER EffectiveDate
Backward-compatible alias for ForceEffectiveDate.

.PARAMETER ForceSend
Bypasses duplicate and past-date skip checks for operator-controlled replays.

.PARAMETER NoSummaryEmail
Suppresses the end-of-run summary email. Logs and console summary are still produced.

.PARAMETER TeamsUserMessageMode
Optional user-facing Teams message mode. Disabled by default. Use GraphDelegated to send a one-on-one
Teams chat message with Microsoft Graph delegated permissions.

.PARAMETER WhatIf
Performs a dry run: resolves recipients/templates and writes logs, but does not send user, summary, or Teams notifications.

.EXAMPLE
.\SmartM365-ExchangeMigration-NotifyUsers.ps1 -Tenant prod -WhatIf

Runs the configured migration campaign in dry-run mode using the recipients path from JSON.

.EXAMPLE
.\SmartM365-ExchangeMigration-NotifyUsers.ps1 -Tenant prod -RecipientsPath .\Recipients\pilot.csv -ForceLanguage fr-FR

Sends the migration notification to a pilot CSV and forces the French template/subject.

.VERSION
1.3
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$RecipientsPath = '',
    [string]$CampaignConfigPath = '',
    [string]$ForceLanguage = '',
    [string]$ForceEffectiveDate = '',
    [string]$EffectiveDate = '',
    [string]$MailSendMode = '',
    [string]$ExchangeManagementMode = '',
    [string]$TeamsUserMessageMode = '',
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
    param($Row, [string]$DefaultEffectiveDate, [string]$ForceEffectiveDate)
    $dateSource = 'EffectiveDate'
    if (-not [string]::IsNullOrWhiteSpace($ForceEffectiveDate)) {
        $raw = $ForceEffectiveDate
        $dateSource = 'ForceEffectiveDate'
    }
    else {
        $raw = [string](Get-SmartM365CommunicationProperty -InputObject $Row -Names @('EffectiveDate','Date','MigrationDate') -DefaultValue $DefaultEffectiveDate)
    }
    if ([string]::IsNullOrWhiteSpace($raw)) { return [pscustomobject]@{ Text = ''; Date = $null; Error = '' } }
    $dateText = $raw.Trim()
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact($dateText, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return [pscustomobject]@{ Text = $parsed.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture); Date = $parsed.Date; Error = '' }
    }
    return [pscustomobject]@{ Text = ''; Date = $null; Error = ("Invalid {0} '{1}'. Use yyyy-MM-dd, for example 2026-07-01." -f $dateSource, $dateText) }
}

function Get-EffectiveDateField {
    param([AllowNull()]$EffectiveDate, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $EffectiveDate) { return $null }
    $property = $EffectiveDate.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
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
$localBasePath = Join-Path $commRoot 'ExchangeUserNotifications\Config\Communications.local.json'
$baseConfig = Merge-SmartM365CommunicationConfig -BaseConfig $baseConfig -OverlayConfig (Read-SmartM365CommunicationJson -Path $localBasePath)

$campaignTemplatePath = Join-Path $commRoot 'ExchangeUserNotifications\Config\Campaigns\ExchangeMigration.local.json.template'
if ([string]::IsNullOrWhiteSpace($CampaignConfigPath)) {
    $CampaignConfigPath = Join-Path $commRoot 'ExchangeUserNotifications\Config\Campaigns\ExchangeMigration.local.json'
}
$campaignConfig = Merge-SmartM365CommunicationConfig -BaseConfig (Read-SmartM365CommunicationJson -Path $campaignTemplatePath) -OverlayConfig (Read-SmartM365CommunicationJson -Path $CampaignConfigPath)
$config = Merge-SmartM365CommunicationConfig -BaseConfig $baseConfig -OverlayConfig $campaignConfig

$campaignName = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'CampaignName' -DefaultValue 'ExchangeMigration')
$runId = [guid]::NewGuid().ToString()
$runOutputPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RunOutputPath' -DefaultValue ''))
$logOutputPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'LogOutputPath' -DefaultValue $runOutputPath))
$sentRegistryPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SentRegistryPath' -DefaultValue (Join-Path $runOutputPath 'Data\SentRegistry.csv')))
$logPath = Join-Path -Path $logOutputPath -ChildPath ("SmartM365-ExchangeMigration-NotifyUsers-Log_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
foreach ($folder in @($runOutputPath, $logOutputPath, (Split-Path -Path $sentRegistryPath -Parent))) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
}

Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $runOutputPath -LatestOutputRoot $pathTokens.LatestCsvFolderPath -LogPath (Join-Path $logOutputPath 'SmartM365-ExchangeMigration-NotifyUsers.log')

if ([string]::IsNullOrWhiteSpace($RecipientsPath)) {
    $RecipientsPath = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'RecipientsPath' -DefaultValue ''))
}
$defaultTemplateRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Templates'
$configuredTemplateRoot = Resolve-CommPath ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TemplateRootPath' -DefaultValue $defaultTemplateRoot))
$templateRoot = @($configuredTemplateRoot, $defaultTemplateRoot) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique
$templateBaseName = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TemplateBaseName' -DefaultValue 'SmartM365-ExchangeMigration-NotifyUsers-Template')
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
$intraEmailDelayMilliseconds = [int](Get-SmartM365CommunicationConfigValue -Config $config -Name 'IntraEmailDelayMilliseconds' -DefaultValue 150 -FallbackConfig $baseConfig)
$summaryTo = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryTo' -DefaultValue (Get-SmartM365CommunicationConfigValue -Config $tenantConfig -Name 'ErrorMailTo' -DefaultValue ''))
$summaryBcc = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SummaryBcc' -DefaultValue '')
$bccAll = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'BccAll' -DefaultValue '')
$defaultEffectiveDate = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultEffectiveDate' -DefaultValue '')
$effectiveDateOverride = $ForceEffectiveDate
if ([string]::IsNullOrWhiteSpace($effectiveDateOverride) -and -not [string]::IsNullOrWhiteSpace($EffectiveDate)) { $effectiveDateOverride = $EffectiveDate }
$skipPast = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipWhenEffectiveDateIsPast' -DefaultValue $true)
$preventResend = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'PreventResendAcrossRuns' -DefaultValue $true)
$subjectByLanguage = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'SubjectByLanguage' -DefaultValue @{})
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
$termsPortalUrl = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'TermsPortalUrl' -DefaultValue '')
$enableTermsPortalBlock = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableTermsPortalBlock' -DefaultValue $false)
$termsPortalBlockStyle = if ($enableTermsPortalBlock -and -not [string]::IsNullOrWhiteSpace($termsPortalUrl)) { '' } else { 'display:none;' }
$hotlineMap = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'HotlineByLanguageOrCountry' -DefaultValue @{} -FallbackConfig $baseConfig)
$domainLanguageMap = Convert-ToHash (Get-SmartM365CommunicationConfigValue -Config $config -Name 'DomainLanguageMap' -DefaultValue @{} -FallbackConfig $baseConfig)
$enableAdLookup = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableAdLookupForLanguage' -DefaultValue $true)
$forestGcServer = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'ForestGcServer' -DefaultValue '')
$enableExchangeMailboxStateCheck = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'EnableExchangeMailboxStateCheck' -DefaultValue $true)
$skipMailboxNotFound = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipMailboxNotFoundWhenExchangeCheckAvailable' -DefaultValue $true)
$skipRemoteMailbox = [bool](Get-SmartM365CommunicationConfigValue -Config $config -Name 'SkipRemoteMailboxWhenExchangeCheckAvailable' -DefaultValue $true)
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

$columns = @('Timestamp','RunId','Campaign','BatchFile','Email','LanguageTag','Scope','Subject','Status','TeamsStatus','TeamsError','SkipReason','ErrorMessage','RecipientTypeDetails','ExchangeSource')
$counters = @{ Files = 0; Rows = 0; Sent = 0; DryRun = 0; Failed = 0; Skipped = 0; AlreadySent = 0; MailboxNotFound = 0; RemoteMailbox = 0; TeamsSent = 0; TeamsFailed = 0; TeamsDryRun = 0 }
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
            if ([string]::IsNullOrWhiteSpace($email)) {
                $counters.Skipped++
                continue
            }

            $skipReason = ''
            $preflightErrorMessage = ''
            if (-not $seen.Add($email) -and -not $ForceSend) { $skipReason = 'DuplicateInRun' }
            elseif ($preventResend -and $registry.ContainsKey($email) -and -not $ForceSend) { $skipReason = 'AlreadySent' }

            $effectiveDate = Resolve-EffectiveDate -Row $row -DefaultEffectiveDate $defaultEffectiveDate -ForceEffectiveDate $effectiveDateOverride
            $effectiveDateError = [string](Get-EffectiveDateField -EffectiveDate $effectiveDate -Name 'Error')
            $effectiveDateText = [string](Get-EffectiveDateField -EffectiveDate $effectiveDate -Name 'Text')
            $effectiveDateValue = Get-EffectiveDateField -EffectiveDate $effectiveDate -Name 'Date'
            if ([string]::IsNullOrWhiteSpace($skipReason) -and -not [string]::IsNullOrWhiteSpace($effectiveDateError)) {
                $skipReason = 'InvalidEffectiveDate'
                $preflightErrorMessage = $effectiveDateError
            }
            elseif ([string]::IsNullOrWhiteSpace($skipReason) -and $skipPast -and $effectiveDateValue -and $effectiveDateValue.Date -lt (Get-Date).Date -and -not $ForceSend) {
                $skipReason = 'EffectiveDateInPast'
            }

            $language = Resolve-SmartM365CommunicationLanguageTag -Row $row -ForceLanguage $ForceLanguage -DefaultLanguageTag $defaultLanguage -DomainLanguageMap $domainLanguageMap
            $exchangeLanguageSource = ''
            if ([string]::IsNullOrWhiteSpace($ForceLanguage) -and $exchangeSnapInState.Available) {
                $exchangeLanguage = Resolve-SmartM365CommunicationExchangeLanguageTag -SmtpAddress $email -DefaultLanguageTag $defaultLanguage -DomainLanguageMap $domainLanguageMap
                if (-not [string]::IsNullOrWhiteSpace($exchangeLanguage.LanguageTag)) { $language = $exchangeLanguage.LanguageTag }
                if ($exchangeLanguage.Source -ne 'Default') { $exchangeLanguageSource = [string]$exchangeLanguage.Source }
            }
            if ($enableAdLookup) {
                $adInfo = Resolve-SmartM365CommunicationAdUserInfo -SmtpAddress $email -GcServer $forestGcServer -DefaultLanguageTag $defaultLanguage -ResolveLanguage
                if ($adInfo.Found) {
                    if ([string]::IsNullOrWhiteSpace($ForceLanguage) -and [string]::IsNullOrWhiteSpace($exchangeLanguageSource) -and -not [string]::IsNullOrWhiteSpace($adInfo.PreferredLanguage)) { $language = $adInfo.PreferredLanguage }
                }
            }
            $subject = Get-SmartM365CommunicationSubject -SubjectByLanguage $subjectByLanguage -LanguageTag $language -DefaultSubject $defaultSubject

            if (-not [string]::IsNullOrWhiteSpace($skipReason)) {
                if ($skipReason -eq 'AlreadySent') { $counters.AlreadySent++ } else { $counters.Skipped++ }
                $item = [pscustomobject]@{ Email = $email; LanguageTag = $language; Status = "Skipped:$skipReason" }
                [void]$processedItems.Add($item)
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; LanguageTag = $language; Scope = 'Preflight'; Subject = $subject; Status = 'Skipped'; SkipReason = $skipReason; ErrorMessage = $preflightErrorMessage; RecipientTypeDetails = ''; ExchangeSource = '' })
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
                    [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:MailboxNotFound(Exchange)' })
                    Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; LanguageTag = $language; Scope = $exchangeScope; Subject = $subject; Status = 'Skipped'; SkipReason = 'MailboxNotFound(Exchange)'; ErrorMessage = $exchangeState.ErrorMessage; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                    continue
                }

                if ($exchangeState.IsRemoteMailbox -and $skipRemoteMailbox -and -not $ForceSend) {
                    $counters.RemoteMailbox++
                    [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Skipped:IsRemoteMailbox(Exchange)' })
                    Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; LanguageTag = $language; Scope = $exchangeScope; Subject = $subject; Status = 'Skipped'; SkipReason = 'IsRemoteMailbox(Exchange)'; ErrorMessage = ''; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
                    continue
                }
            }

            try {
                $template = Get-SmartM365CommunicationTemplateContent -TemplateRoot $templateRoot -TemplateBaseName $templateBaseName -LanguageTag $language -DefaultLanguageTag $defaultLanguage
                $logoTokens = Get-SmartM365CommunicationLinkedLogoTokens -LogoPath $logoPath -LogoContentId $logoContentId
                $tokens = @{
                    Date = $(if ($effectiveDateText) { $effectiveDateText } else { (Get-Date).ToString('yyyy-MM-dd') })
                    Hotline = Get-SmartM365CommunicationHotline -HotlineByLanguageOrCountry $hotlineMap -LanguageTag $language -DefaultHotline ([string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'DefaultHotline' -DefaultValue ''))
                    OldWebmailUrl = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'OldWebmailUrl' -DefaultValue '')
                    NewWebmailUrl = [string](Get-SmartM365CommunicationConfigValue -Config $config -Name 'NewWebmailUrl' -DefaultValue 'https://outlook.office365.com')
                    TermsPortalBlockStyle = $termsPortalBlockStyle
                    TermsPortalUrl = $termsPortalUrl
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
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; LanguageTag = $language; Scope = 'Domain'; Subject = $subject; Status = $status; TeamsStatus = $teamsStatus; TeamsError = $teamsError; SkipReason = ''; ErrorMessage = $result.Mode; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
            }
            catch {
                $counters.Failed++
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Failed' })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; LanguageTag = $language; Scope = 'Domain'; Subject = $subject; Status = 'Failed'; SkipReason = ''; ErrorMessage = $_.Exception.Message; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
            }
        }
    }

    if ($preventResend -and -not $WhatIf) { Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $sentRegistryPath }

    $summary = "Files=$($counters.Files); Rows=$($counters.Rows); Sent=$($counters.Sent); DryRun=$($counters.DryRun); Failed=$($counters.Failed); Skipped=$($counters.Skipped); AlreadySent=$($counters.AlreadySent); MailboxNotFound=$($counters.MailboxNotFound); RemoteMailbox=$($counters.RemoteMailbox); TeamsSent=$($counters.TeamsSent); TeamsDryRun=$($counters.TeamsDryRun); TeamsFailed=$($counters.TeamsFailed)."
    if (-not $WhatIf -and -not $NoSummaryEmail -and -not [string]::IsNullOrWhiteSpace($summaryTo)) {
        $summaryHtml = New-SmartM365CommunicationSummaryHtml -Title $summaryTitle -Facts @{
            RunId = $runId; Tenant = $Tenant; Mode = $(if ($WhatIf) { 'DryRun' } else { 'Live' }); Summary = $summary; LogPath = $logPath; SentRegistryPath = $sentRegistryPath; ExchangeManagement = ("{0}:{1}" -f $exchangeManagementState.Source, $exchangeManagementState.Status)
        } -Items @($processedItems)
        Send-SmartM365CommunicationMail -MailSendMode $mailSendMode -SmtpServer $smtpServer -SmtpPort $smtpPort -From $from -To $summaryTo -Bcc $summaryBcc -Subject $summarySubject -BodyHtml $summaryHtml -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $global:Thumbprint -SmtpUseIntegratedAuth $smtpUseIntegratedAuth -SmtpEnableSsl $smtpEnableSsl -RetryCount $smtpRetryCount -RetryDelaySeconds $smtpRetryDelaySeconds | Out-Null
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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATIEmQYXp4KkOy
# 5KpBy7TshN590txrtnP+m5Dod4zphaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBb5pDAWODnLWjbaccHk6YcB1ZHf+7WQHD45FDLtaJTYDANBgkqhkiG9w0B
# AQEFAASCAYArIEG9hatfL7de+9fsvS2smjcI+4/5NF9uX2qUWI5akA2VMun9NC7n
# JO+cX8MgYtI2DZiQODHFb2wpXt+iaAWUKikDraZYvEEdyv0OVettZ9MUo7rVuMSt
# pbh4+mf2K5CiyJY/a4mLLxRZ1fVYzyQu4aXgRGTiGMl9SmxU+OyRPJsgt4HgB1ru
# VcDS5gYMHPwQq0+z622i5IZdOl7TL/xusghQ9Zo4cNtZE590OabeoR9MeifxRTxb
# C1IynzGsOdyTMKG820lBqaLoCYgWMhqQ0qMG7hVc1rLUcOwcN5gsJacVT89Q/Ot/
# fcd0/sfGt5dcrGNVMQkS+o1PGC4+dEP6jhejFu36FFho+i5xnNqHM75R6Xj9lXT4
# Nc2SsMRMfGAjMLCkKecaoi/U02KyskirDfEGjHcR5/6b5nl9q7y3TetNT0OZ4Oc7
# epuvEvpe/Zvmo0V/o7a5q+xDxGTGtZpLmNHXrXH3LUMJ4mf/cCX8+O3wPNaLEFd5
# au+IlMtsfHE=
# SIG # End signature block
