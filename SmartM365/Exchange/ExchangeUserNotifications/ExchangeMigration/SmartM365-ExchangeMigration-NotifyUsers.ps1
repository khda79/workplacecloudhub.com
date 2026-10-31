<#
.SYNOPSIS
Notifies users that their mailbox is moving from the legacy webmail platform to Exchange Online.

.DESCRIPTION
Loads SmartM365 tenant context, the shared ExchangeUserNotifications configuration, and the ExchangeMigration
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
1.7
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
    [switch]$MigrationBatchSelection,
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

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartM365.ExchangeMigrationSelection.psm1') -Force

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
$baseConfig = Read-SmartM365CommunicationJson -Path (Join-Path $commRoot 'ExchangeUserNotifications\Config\ExchangeUserNotifications.local.json.template')
$localBasePath = Join-Path $commRoot 'ExchangeUserNotifications\Config\ExchangeUserNotifications.local.json'
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
if ($MigrationBatchSelection) {
    if ($effectiveDateOverride) { throw 'Migration selection keeps per-recipient dates; a global ForceEffectiveDate is not allowed' }
    $selectionFiles = @(Get-RecipientFiles -Path $RecipientsPath)
    if ($selectionFiles.Count -ne 1 -or -not (Test-Path -LiteralPath $RecipientsPath -PathType Leaf)) { throw 'Migration selection requires one reviewed CSV file' }
    $migrationSelectionRows = @(Import-SmartM365CommunicationCsv -Path $RecipientsPath)
    Assert-ExoMigrationSelectionCsv -Rows $migrationSelectionRows -TenantId $global:TenantId -TenantKey $Tenant
    $exchangeManagementMode = 'Disabled'
    $enableExchangeMailboxStateCheck = $false
    $requireExchangeManagement = $false
}

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
$registryLock = $null
$registryCheckpointPending = $false

try {
    if ($preventResend -and -not $WhatIf) {
        $registryLock = Enter-SmartM365CommunicationSentRegistryLock -Path $sentRegistryPath
    }
    $registry = Load-SmartM365CommunicationSentRegistry -Path $sentRegistryPath
    foreach ($file in (Get-RecipientFiles -Path $RecipientsPath)) {
        $counters.Files++
        $rows = if ($MigrationBatchSelection) { $migrationSelectionRows } else { @(Import-SmartM365CommunicationCsv -Path $file.FullName -RequiredColumns @()) }
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
                if (-not $WhatIf -and $result.Sent) {
                    $counters.Sent++
                    if ($preventResend) {
                        $registryCheckpointPending = $true
                        Register-SmartM365CommunicationSentItem -Registry $registry -Email $email
                        Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $sentRegistryPath
                        $registryCheckpointPending = $false
                    }
                }
                $teamsStatus = 'Disabled'
                $teamsError = ''
                try {
                    if ((Get-SmartM365CommunicationTeamsUserMode -TeamsUserMessageMode $teamsUserMessageMode) -ne 'Disabled' -and ($WhatIf -or $result.Sent -or $result.Mode -eq 'AdministratorsOnly')) {
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
                elseif (-not $result.Sent) { $counters.Skipped++ }
                $status = if ($WhatIf) { 'DryRun' } elseif ($result.Sent) { 'Success' } else { $result.Mode }
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = $status })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; LanguageTag = $language; Scope = 'Domain'; Subject = $subject; Status = $status; TeamsStatus = $teamsStatus; TeamsError = $teamsError; SkipReason = ''; ErrorMessage = $result.Mode; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
            }
            catch {
                if ($registryCheckpointPending) { throw }
                $counters.Failed++
                [void]$processedItems.Add([pscustomobject]@{ Email = $email; LanguageTag = $language; Status = 'Failed' })
                Add-SmartM365CommunicationLogRow -Path $logPath -Columns $columns -Row ([pscustomobject]@{ Timestamp = Get-Date; RunId = $runId; Campaign = $campaignName; BatchFile = $file.Name; Email = $email; LanguageTag = $language; Scope = 'Domain'; Subject = $subject; Status = 'Failed'; SkipReason = ''; ErrorMessage = $_.Exception.Message; RecipientTypeDetails = $recipientTypeDetails; ExchangeSource = $exchangeSource })
            }
        }
    }

    $summary = "Files=$($counters.Files); Rows=$($counters.Rows); Sent=$($counters.Sent); DryRun=$($counters.DryRun); Failed=$($counters.Failed); Skipped=$($counters.Skipped); AlreadySent=$($counters.AlreadySent); MailboxNotFound=$($counters.MailboxNotFound); RemoteMailbox=$($counters.RemoteMailbox); TeamsSent=$($counters.TeamsSent); TeamsDryRun=$($counters.TeamsDryRun); TeamsFailed=$($counters.TeamsFailed)."
    if (-not $WhatIf -and -not $NoSummaryEmail -and -not [string]::IsNullOrWhiteSpace($summaryTo)) {
        $summaryHtml = New-SmartM365CommunicationSummaryHtml -Title $summaryTitle -Facts @{
            RunId = $runId; Tenant = $Tenant; Mode = $(if ($WhatIf) { 'DryRun' } else { 'Live' }); Summary = $summary; LogPath = $logPath; SentRegistryPath = $sentRegistryPath; ExchangeManagement = ("{0}:{1}" -f $exchangeManagementState.Source, $exchangeManagementState.Status)
        } -Items @($processedItems)
        Send-SmartM365CommunicationMail -MailSendMode $mailSendMode -SmtpServer $smtpServer -SmtpPort $smtpPort -From $from -To $summaryTo -Bcc $summaryBcc -Subject $summarySubject -BodyHtml $summaryHtml -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $global:Thumbprint -SmtpUseIntegratedAuth $smtpUseIntegratedAuth -SmtpEnableSsl $smtpEnableSsl -RetryCount $smtpRetryCount -RetryDelaySeconds $smtpRetryDelaySeconds | Out-Null
    }

    if ($counters.Failed -gt 0 -or $counters.TeamsFailed -gt 0) {
        Write-Host $summary
        throw ("Campaign completed with delivery failures. {0}" -f $summary)
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

finally {
    if ($null -ne $registryLock) { $registryLock.Dispose() }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCytmrGFp6VXWPY
# CmnHW933LaTbT6bDbOp6Q6MYVoSqRKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMHPHh6oji56SxpV0kNB4gV4wYT0N0cZhkBvqvzO9eCnMA0GCSqG
# SIb3DQEBAQUABIIBgLGpJXkANXya2PA12B970pU+lPMIJU0gBwFgozg9tlG/U/Dg
# nrcOD2wRDPqRvjrsWdfjjnUW9LuP3xHBvMCL5i75Z9p0PKeovX4Kl9/dyKpZ1k0Q
# PjcjFkP6bTUsLBEB2FTCNhNsB/IGUeppteu/f4jNc66JKiW5ovov/nG2LJvmbVyR
# kjLVds/0zSU+TDcsy4eRFzufUvZyCpcdCcE1W19W60XMDHMg5M7xpyTUoI70Ni1I
# nLlrdzJ3pNfWhyjCv5dWdUu9Y462058AVPgvCCFbFNVRQ2BYWFy4ai9pvFUm3NDA
# nmeaxJFILPjGQOguQR+1WoRUiV4PtXBwIm9jVLR3k+ow9qXG46uq4VS7URTuNW56
# 1p/29YhyvryER4l3nMRGfNMQXahUl4k4ojfm6e+qaM/+gXVFeldWHl2VHmjlwNXZ
# KY4aqVCNpyBtp98HX7XL62OFLikqXRbfggrw+gptNL4FOVhcfPz+Avi7L397uvPA
# 4HqyTYqQP2pJ+YtKhqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExNjE4
# MjNaMC8GCSqGSIb3DQEJBDEiBCDGHAAVi6HhNdL+Y5qmO2mhfLimTpnXisDsbvQY
# +0JO/DANBgkqhkiG9w0BAQEFAASCAgBaEmcco2erYdcX0f5kKz4IFxbXPal3+Ckp
# gqJpQefio9qdMAJuWYYix/TBQf5LLwN4Ws2Kc8SM65siFkKoifXAv62bNF8NyvEI
# DX4paDAI4Nap0shDdtAq5P2D9DDw8OcXmjF1uPnqAMgRo2WBOuI1MfIRroD+qQWa
# 0O59ne4Ee1hLxUEpnsDaVIby34CRtNn7vCg0C2BV0pByZ+gn/EV810EdTkO6P8Jc
# zZsOWj+zybzz9FXO+htjZ8AkAriRyLqKhxfA1z+uoaNxrUGLOh36WwFoBP3cU1b9
# ToD+rT3efhEWvOF4X9clM0fCG6kTk2QMT8eH7cS9eTqQrJ+wylYKSDFviFzEM4cJ
# V2OvHE827yUHtRdSYIKqkihb+3K79/oOkn575RcXLiHSTlseXiEkbczjGVD/0rqz
# QVC1kpMY5SmHTYLvIOJLIfZITJ1Yq00/rnnJe+Z/C4FqBNJ5Lhj3HR1OCLz2xR8h
# AmYWaOvEcRG5IpHxjZSa0rCW1x2cgJttqZAxzd3+0z+d0DszVx1w2qrfGtV7y6qt
# a7rbrrZ1aV1VqHQMsb5PCwf02sdANy30CI7afqXzFBLkTHcANROBM6PbA3KFBiG4
# LUtiBA+Z9+D4tZhPeq3wDLZyTRod0uuVu4SWxc8+D4R7uZjRDy+K3QRZ8SPiZBYY
# JobG7Uogew==
# SIG # End signature block
