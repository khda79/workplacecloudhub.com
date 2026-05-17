#Requires -Version 7.0

<#
.SYNOPSIS
    Registers and tests the Teams Workflows webhook URL used by SmartM365 notifications.

.DESCRIPTION
    Validates that the provided URL is an HTTPS Teams Workflows / Power Automate webhook URL,
    writes it to SmartM365.global.local.json as TeamsWebhookUrl, then sends a test notification
    through Send-SmartM365TeamsNotification.

    This script intentionally rejects legacy Office 365 Connector incoming webhook URLs.

.PARAMETER WebhookUrl
    Teams Workflows / Power Automate webhook URL from a trigger such as
    "When a Teams webhook request is received".

.PARAMETER Channel
    SmartM365 notification channel to configure. Alerts writes TeamsAlertsWebhookUrl,
    Infos writes TeamsInfosWebhookUrl, and Default writes the legacy fallback TeamsWebhookUrl.

.PARAMETER ConfigPath
    Path to SmartM365.global.local.json. Defaults to the file next to this script.

.PARAMETER TestOnly
    Sends the test notification without writing TeamsWebhookUrl to SmartM365.global.local.json.

.PARAMETER SkipTest
    Writes TeamsWebhookUrl without sending a test notification.

.EXAMPLE
    .\SmartM365-Set-TeamsWebhook.ps1 -Channel Alerts -WebhookUrl "https://prod-00.westeurope.logic.azure.com/workflows/..."

.EXAMPLE
    .\SmartM365-Set-TeamsWebhook.ps1 -Channel Infos -WebhookUrl "https://prod-00.westeurope.logic.azure.com/workflows/..." -TestOnly
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$WebhookUrl,

    [Parameter()]
    [ValidateSet('Alerts','Infos','Default')]
    [string]$Channel = 'Default',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath = (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.global.local.json'),

    [Parameter()]
    [string]$Title = 'SmartM365 Teams webhook test',

    [Parameter()]
    [string]$Message = 'Teams notifications are configured for SmartM365.',

    [Parameter()]
    [switch]$TestOnly,

    [Parameter()]
    [switch]$SkipTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-SmartM365TeamsWorkflowWebhookUrl {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    $uri = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) {
        throw 'TeamsWebhookUrl must be an absolute HTTPS URL.'
    }

    if ($uri.Scheme -ne 'https') {
        throw 'TeamsWebhookUrl must use HTTPS.'
    }

    $legacyConnectorHosts = @(
        'outlook.office.com',
        'outlook.office365.com',
        'webhook.office.com',
        'webhook.office365.com',
        'connectors.office.com'
    )

    $hostName = $uri.Host.ToLowerInvariant()
    foreach ($legacyHost in $legacyConnectorHosts) {
        if ($hostName -eq $legacyHost -or $hostName.EndsWith(".$legacyHost")) {
            throw "Legacy Office 365 Connector webhook URLs are not supported for SmartM365. Create a Teams Workflows / Power Automate webhook instead."
        }
    }

    $knownWorkflowHost = (
        $hostName -like '*.logic.azure.com' -or
        $hostName -like '*.api.powerplatform.com' -or
        $hostName -like '*.environment.api.powerplatform.com'
    )

    if (-not $knownWorkflowHost) {
        Write-Warning "The URL host '$($uri.Host)' does not look like a typical Teams Workflows / Power Automate trigger URL. Continuing because it is HTTPS and not a legacy connector URL."
    }

    return $uri
}

function Read-SmartM365JsonConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        $templatePath = Join-Path -Path (Split-Path -Parent $Path) -ChildPath 'SmartM365.global.local.json.template'
        if (Test-Path -LiteralPath $templatePath) {
            return Get-Content -LiteralPath $templatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }

        return [pscustomobject]@{}
    }

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [pscustomobject]@{}
    }

    return $raw | ConvertFrom-Json -ErrorAction Stop
}

$uri = Test-SmartM365TeamsWorkflowWebhookUrl -Url $WebhookUrl
$resolvedConfigPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
$configPropertyName = switch ($Channel) {
    'Alerts' { 'TeamsAlertsWebhookUrl' }
    'Infos' { 'TeamsInfosWebhookUrl' }
    default { 'TeamsWebhookUrl' }
}
$testLevel = if ($Channel -eq 'Alerts') { 'ERROR' } else { 'SUCCESS' }
$testTitle = if ($PSBoundParameters.ContainsKey('Title')) { $Title } else { "SmartM365 Teams webhook test - $Channel" }
$notificationChannel = if ($Channel -eq 'Default') { 'Auto' } else { $Channel }

if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "SmartM365.Core module was not found at '$modulePath'."
}

Import-Module $modulePath -Force -ErrorAction Stop

if (-not $TestOnly) {
    $config = Read-SmartM365JsonConfig -Path $resolvedConfigPath
    if ($null -eq $config.PSObject.Properties[$configPropertyName]) {
        $config | Add-Member -NotePropertyName $configPropertyName -NotePropertyValue $uri.AbsoluteUri
    }
    else {
        $config.$configPropertyName = $uri.AbsoluteUri
    }

    $configFolder = Split-Path -Parent $resolvedConfigPath
    if (-not (Test-Path -LiteralPath $configFolder)) {
        if ($PSCmdlet.ShouldProcess($configFolder, 'Create SmartM365 config folder')) {
            New-Item -Path $configFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }

    if ($PSCmdlet.ShouldProcess($resolvedConfigPath, 'Write TeamsWebhookUrl')) {
        $config | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $resolvedConfigPath -Encoding UTF8
        Write-Information ("[OK] {0} written to {1}" -f $configPropertyName, $resolvedConfigPath) -InformationAction Continue
    }
}
else {
    Write-Information '[INFO] TestOnly was used; SmartM365.global.local.json was not modified.' -InformationAction Continue
}

if (-not $SkipTest) {
    $testResult = Send-SmartM365TeamsNotification `
        -WebhookUrl $uri.AbsoluteUri `
        -Title $testTitle `
        -Message $Message `
        -Level $testLevel `
        -Channel $notificationChannel `
        -Facts @{
            Script = 'SmartM365-Set-TeamsWebhook.ps1'
            Mode   = $(if ($TestOnly) { 'TestOnly' } else { 'RegisterAndTest' })
            Channel = $Channel
            Url    = $uri.Host
        } `
        -ThrowOnError

    if ($testResult) {
        Write-Information '[OK] Teams webhook test notification sent.' -InformationAction Continue
    }
}
else {
    Write-Information '[INFO] Teams webhook test skipped because -SkipTest was used.' -InformationAction Continue
}
