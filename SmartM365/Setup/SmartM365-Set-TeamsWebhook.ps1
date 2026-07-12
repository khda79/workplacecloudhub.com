#Requires -Version 7.0

<#
.SYNOPSIS
    Registers and tests the Teams Workflows webhook URL used by SmartM365 notifications.

.DESCRIPTION
    Validates that the provided URL is an HTTPS Teams Workflows / Power Automate webhook URL,
    writes it to the selected tenant local profile as TeamsWebhookUrl, then sends a test notification
    through Send-SmartM365TeamsNotification.

    This script intentionally rejects legacy Office 365 Connector incoming webhook URLs.

.PARAMETER WebhookUrl
    Teams Workflows / Power Automate webhook URL from a trigger such as
    "When a Teams webhook request is received".

.PARAMETER Channel
    SmartM365 notification channel to configure. Alerts writes TeamsAlertsWebhookUrl,
    Infos writes TeamsInfosWebhookUrl, and Default writes the legacy fallback TeamsWebhookUrl.

.PARAMETER Tenant
    Tenant profile key. Defaults to test.

.PARAMETER ConfigPath
    Optional explicit JSON config path. Defaults to Config\Tenants\<Tenant>.local.json.

.PARAMETER TestOnly
    Sends the test notification without writing the selected tenant profile.

.PARAMETER SkipTest
    Writes TeamsWebhookUrl without sending a test notification.

.EXAMPLE
    .\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Alerts -WebhookUrl "https://prod-00.westeurope.logic.azure.com/workflows/..."

.EXAMPLE
    .\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Infos -WebhookUrl "https://prod-00.westeurope.logic.azure.com/workflows/..." -TestOnly
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Tenant = 'test',

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$WebhookUrl,

    [Parameter()]
    [ValidateSet('Alerts','Infos','Default')]
    [string]$Channel = 'Default',

    [Parameter()]
    [string]$ConfigPath = '',

    [Parameter()]
    [string]$Title = 'SmartM365 Teams webhook test',

    [Parameter()]
    [string]$Message = 'Teams notifications are configured for SmartM365.',

    [Parameter()]
    [switch]$TestOnly,

    [Parameter()]
    [switch]$SkipTest
)
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365RootPath = Find-SmartM365Root -StartPath $PSScriptRoot
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

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
        $templatePath = Join-Path -Path (Split-Path -Parent $Path) -ChildPath 'tenant.local.json.template'
        if (-not (Test-Path -LiteralPath $templatePath)) {
            $templatePath = Join-Path -Path $script:SmartM365RootPath -ChildPath 'Config\SmartM365.global.local.json.template'
        }

        Initialize-SmartM365LocalJsonFromTemplate -Path $Path -TemplatePath $templatePath -ConfigDescription 'Teams webhook local configuration' | Out-Null
    }

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [pscustomobject]@{}
    }

    return $raw | ConvertFrom-Json -ErrorAction Stop
}

$uri = Test-SmartM365TeamsWorkflowWebhookUrl -Url $WebhookUrl
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path -Path $script:SmartM365RootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $Tenant)
}
$resolvedConfigPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)
$modulePath = Join-Path -Path $script:SmartM365RootPath -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
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

    if ($null -eq $config.PSObject.Properties['EnableTeamsNotifications']) {
        $config | Add-Member -NotePropertyName 'EnableTeamsNotifications' -NotePropertyValue $true
    }
    else {
        $config.EnableTeamsNotifications = $true
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
    Write-Information '[INFO] TestOnly was used; tenant local configuration was not modified.' -InformationAction Continue
}

if (-not $SkipTest) {
    $testResult = Send-SmartM365TeamsNotification `
        -WebhookUrl $uri.AbsoluteUri `
        -Title $testTitle `
        -Message $Message `
        -Level $testLevel `
        -Channel $notificationChannel `
        -ResultSummary ("Webhook test sent for channel {0} on host {1}." -f $Channel, $uri.Host) `
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


# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD5ZxU0XSceIhLH
# ryVahh780uYJG3u0NQZe8VoGsx7BfaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDGQGO7IeztHfgr/eS5
# I7ZXXjCnUdjJuaPPBIL915GCJTANBgkqhkiG9w0BAQEFAASCAYAt+1UfPOFDHnAf
# DZ6G0DXXeFReecfdYJQLPCfTiVWoDKN6RV0oDYZpGhJBUzWc3kKAnMvPQi86NgJI
# qyIdxb9QrHdSxBv+KrMs208dq7xLhsYjTTcpehqK5jAJjni3+tSlLpbbpK0adUaR
# bFQwq5kvCisk47LyUBUYymHx2vZrmvrla1DNRIHZw8lcKKPeiFPWpYthqtt9Ln5d
# eXpswHJN8gm3XS3F0uHurDB441mMTFN9IWnQ7jnoicHKKXWV/yJ8D1xNrYioFmFQ
# XysMEkV5qe3fTH0tVzYqd4NKl1rt3I7zPyKpVyhwTJSN/VckoKeWM4LpeOMWAYZ6
# Og5keObqrYg9DWW+WrXVd+pIIrampDI5RXgWt4ulS35e4pRXJhn+yQbY8pVZsNwL
# pE8zwds6Vb5irJrFwzS7uKhNF4/2XWpW8b29Bbk5/2EY9KKoAnoGZ7wjaACOjdqS
# 0IOPlANcLeA/9RO/d337UjlkMwksWuOVvlDEDTMaKcV/QfLHgm0=
# SIG # End signature block
