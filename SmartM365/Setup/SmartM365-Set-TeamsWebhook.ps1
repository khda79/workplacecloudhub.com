#Requires -Version 7.0

<#
.SYNOPSIS
    Registers and tests the Teams Workflows webhook URL used by SmartM365 notifications.

.DESCRIPTION
    Validates that the provided URL is an HTTPS Teams Workflows / Power Automate webhook URL,
    writes it to a private per-profile Teams JSON file, then sends a test notification
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
    Optional explicit legacy tenant JSON config path. When specified, the URL is written there.

.PARAMETER TestOnly
    Sends the test notification without writing the selected tenant profile.

.PARAMETER SkipTest
    Writes TeamsWebhookUrl without sending a test notification.

.EXAMPLE
    .\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Alerts -WebhookUrl "https://prod-00.westeurope.logic.azure.com/workflows/..."

.EXAMPLE
    .\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Infos -WebhookUrl "https://prod-00.westeurope.logic.azure.com/workflows/..." -TestOnly

.VERSION
1.5
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

    $Path=Resolve-SmartM365JsonConfigurationPath $Path
    if (-not (Test-Path -LiteralPath $Path)) {
        $templatePath = Join-Path -Path (Split-Path -Parent $Path) -ChildPath 'tenant.local.json.template'
        if (-not (Test-Path -LiteralPath $templatePath)) {
            $templatePath = Join-Path -Path $script:SmartM365RootPath -ChildPath 'Config\SmartM365.global.local.json.template'
        }

        Initialize-SmartM365LocalJsonFromTemplate -Path $Path -TemplatePath $templatePath -ConfigDescription 'Teams webhook local configuration' | Out-Null
    }

    return (Read-SmartM365JsonDocument $Path).Document
}

$uri = Test-SmartM365TeamsWorkflowWebhookUrl -Url $WebhookUrl
$explicitConfigPath = -not [string]::IsNullOrWhiteSpace($ConfigPath)
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path -Path $script:SmartM365RootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $Tenant)
}
$resolvedConfigPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)
$resolvedConfigPath = Resolve-SmartM365JsonConfigurationPath $resolvedConfigPath
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

Import-Module $modulePath -MinimumVersion '1.0.58' -Force -ErrorAction Stop

if (-not $TestOnly) {
    $config = Read-SmartM365JsonConfig -Path $resolvedConfigPath
    $configReceipt=Read-SmartM365JsonDocument $resolvedConfigPath
    $config=$configReceipt.Document
    $configHash=$configReceipt.SHA256
    if (-not $explicitConfigPath) {
        if ($Tenant -notmatch '^[A-Za-z0-9][A-Za-z0-9_-]*$') { throw 'Invalid Teams configuration profile key.' }
        $privatePath = Join-Path $script:SmartM365RootPath ("Config\Tenants\{0}.teams.local.json" -f $Tenant)
        $privatePath = Resolve-SmartM365JsonConfigurationPath $privatePath
        $privateHash = 'ABSENT'
        $privateConfig = [pscustomobject]@{}
        if (Test-Path -LiteralPath $privatePath) {
            $privateReceipt = Read-SmartM365JsonDocument $privatePath
            $privateConfig = $privateReceipt.Document
            $privateHash = $privateReceipt.SHA256
            if ($privateConfig -isnot [pscustomobject]) { throw 'Private Teams configuration must be an object.' }
        }
        if ($null -eq $privateConfig.PSObject.Properties[$configPropertyName]) { $privateConfig | Add-Member -NotePropertyName $configPropertyName -NotePropertyValue $uri.AbsoluteUri }
        else { $privateConfig.$configPropertyName = $uri.AbsoluteUri }
        if ($PSCmdlet.ShouldProcess($privatePath, 'Write private Teams webhook')) {
            Write-SmartM365JsonBytesAtomically -Path $privatePath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($privateConfig | ConvertTo-Json -Depth 20))) -ExpectedSHA256 $privateHash -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Configuration must be an object.'}} | Out-Null
        }
    }
    elseif ($null -eq $config.PSObject.Properties[$configPropertyName]) {
        $config | Add-Member -NotePropertyName $configPropertyName -NotePropertyValue $uri.AbsoluteUri
    }
    elseif ($explicitConfigPath) {
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
        Write-SmartM365JsonBytesAtomically -Path $resolvedConfigPath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($config | ConvertTo-Json -Depth 20))) -ExpectedSHA256 $configHash -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Configuration must be an object.'}} | Out-Null
        if ($explicitConfigPath) { Write-Information ("[OK] {0} written to the specified configuration; a private profile value takes precedence if present." -f $configPropertyName) -InformationAction Continue }
        else { Write-Information ("[OK] {0} written to private Teams configuration; tenant notifications enabled." -f $configPropertyName) -InformationAction Continue }
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDGu//ndOaUR4eA
# nhR2j7xvfjp1OrNHIC99PpMUaOkpgaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIF9l2l7x+AhSeOuSR/fm7BsvF5dxyq8v/UnW3e7bjvOfMA0GCSqG
# SIb3DQEBAQUABIIBgFvKz/vniXJI5L1oMjoQRxbv8kuC4jDSfpeAtuTNOGvCXSuV
# BZvVnoaOy57PMVO0hhYBdA5QpcTshcdYoJQN/P898+keoK34Gfg2HJO8tfnDf+3z
# R149I9G4eNzvTBRWKnorA7JwNDf5HkSw6Pbyoxy2DUiZ+xQaemcxsqowI+RHYsPS
# ghZy2Ib+DAldAE9LWYeP6eKrT75TOKnT3E9heeBx3NUVDB4UdztkEwLxEQC74YkZ
# UDjdqXiot6Ada72CR8mJ3dzKB9aDMD745HyRLvuywGZwnJjyv9T4ajn5kG5k+Y5R
# 5fTC35Y6AbA/h38YjNrXUtGkJDZziwQ0m9NcGufHOIV+rCsfXJAFeUfeSGJLXwuf
# gA71g/uqB/VDvdwa0b+izSJI71JkFq1Vg/ltL2UtLBLwbppS4ldL5De0B7am3LyB
# nkSvodaLW766uFTc3Edc/p3Qh1W1x/z7zjrF5r3X51fXh0LThJ9FgRYIM1ErsAt5
# y52ntsslMbk3p7D8jqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTAwODMz
# MDRaMC8GCSqGSIb3DQEJBDEiBCDkwMHqQqqdVpwbRthCLIwUuiqffb2aKIWo+IzL
# YpwYfjANBgkqhkiG9w0BAQEFAASCAgB7FMRdV5/rBGp56wKZRKbxghKpdgNbVAaX
# 1DAP0FNQXgpArR5varc8ehCxm4kTonusOoNe54Pk/YITZNKrYp6JmrC1KHbGiljF
# Kd66/JK5q/HQJROLo5sdtjRe7uIUNTV3lFMN0LYodYw0iRJxhwAUiwWj9xyVUifd
# xDZfFJvpIGqqvmRKUxzokXKm7T5pn+3iQsuxTbA3ZbFbdmJzdnc/tg5A+KquTkAO
# Tt+KGAQdwqssfStWgoluSFA8RU7TX2NGVCc/NPuNRjINsPtDmNdjaaM6JmJtoQqP
# xKJwxkQwL9v1vyKoJcZRQ/wA2LuHa6sINQ1CdxpCys2dyXDqXfnS1KUTX2SDqHaT
# NX+/sX2dLGxXcpmMuzl1slmX0tUrtWP9oISjkDcLg4OWnXSr1ra4c0ZrD8VUHqwJ
# YMTV9jm4k4jpQM6ydGM9HODoWL+tXE0isLggNpE0jVwVNu2tcdwiGcHhYXXyHXwj
# dxNPHMUwVUVV0ezJilyz+Oo21Vss4HwhCbsRnjeHQhMUKMTg+KLZ9+q8UHaoiUK0
# QpFDZpLMywn/tN+Q9kWMe+FuO3upXCQ8KUeWFFaMuqXF8qqthHOImk1Yqpv37YJI
# gbJogTbxkv8XVtX9fqwPbCAB2QVg5S6XJQCy0FGDNkyBrSkbYQFUtkkDOweu94Ax
# wIduycqALg==
# SIG # End signature block
