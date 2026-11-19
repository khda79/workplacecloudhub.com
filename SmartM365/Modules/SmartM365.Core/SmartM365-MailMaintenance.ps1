<#
.SYNOPSIS
Shared mail admission policy, dot-sourced by both Core editions. No I/O on load.
.VERSION
1.1.0
#>
function Get-SmartM365MaintenanceMailRoot {
    param([Parameter(Mandatory)]$Config)
    $context = Get-Variable SmartM365MaintenanceMailContext -Scope Global -ErrorAction SilentlyContinue
    if ($context -and $context.Value -and $Config.PSObject.Properties['TenantKey'] -and $Config.PSObject.Properties['ProfileKey'] -and
        $context.Value.PSObject.Properties['TenantKey'] -and $context.Value.PSObject.Properties['ProfileKey'] -and
        $context.Value.TenantKey -eq $Config.TenantKey -and $context.Value.ProfileKey -eq $Config.ProfileKey) {
        return [string]$context.Value.SharedDataFolderPath
    }
    if ($env:SMARTM365_ORCHESTRATOR_SHARED_DATA_FOLDER) {
        if (-not $Config.PSObject.Properties['ProfileKey'] -or -not $env:SMARTM365_ORCHESTRATOR_TENANT -or $env:SMARTM365_ORCHESTRATOR_TENANT -ne $Config.ProfileKey) {
            throw 'Maintenance mail context belongs to another tenant profile; mail was not sent.'
        }
        return [string]$env:SMARTM365_ORCHESTRATOR_SHARED_DATA_FOLDER
    }
    $property = $Config.PSObject.Properties['OrchestratorSharedDataFolderPath']
    if ($property -and $property.Value -and $property.Value -notin @('__USE_GLOBAL__','USE_GLOBAL')) {
        return [string](Resolve-SmartM365ConfigValue $property.Value)
    }
    $root = $Config.PSObject.Properties['DataAllRootPath']
    if (-not $root -or -not $root.Value) { throw 'DataAllRootPath is missing; maintenance mail admission cannot be verified.' }
    return Join-Path ([string](Resolve-SmartM365ConfigValue $root.Value)) 'Orchestrator'
}

function Get-SmartM365MaintenanceMailState {
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    # Prove the parent is accessible before treating an absent deployment as inactive.
    $null = Get-Item -LiteralPath (Split-Path $SharedDataFolderPath -Parent) -ErrorAction Stop
    $statePath = Join-Path $SharedDataFolderPath 'Config/Orchestrator-Maintenance.json.txt'
    $legacyPath = Join-Path $SharedDataFolderPath 'Config/Orchestrator-Maintenance.json'
    $guardPath = Join-Path $SharedDataFolderPath 'Config/Orchestrator-Maintenance.guard'
    if (-not (Test-Path -LiteralPath $statePath -ErrorAction Stop) -and -not (Test-Path -LiteralPath $legacyPath -ErrorAction Stop)) {
        if ((Test-Path -LiteralPath $guardPath -ErrorAction Stop) -or
            (Test-Path -LiteralPath (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Cluster.json.txt') -ErrorAction Stop) -or
            (Test-Path -LiteralPath (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Cluster.json') -ErrorAction Stop)) {
            throw 'Initialized maintenance control is missing; mail was not sent.'
        }
        return [pscustomobject]@{Enabled=$false;Revision=0}
    }
    $state = (Read-SmartM365JsonDocument -Path $statePath).Document
    foreach ($name in @('SchemaVersion','Enabled','Revision','ChangedAtUtc','ChangedBy','ChangedFromServer','Reason','ResumeAfterUtc')) {
        if (-not $state.PSObject.Properties[$name]) { throw "Maintenance mail control is missing $name; mail was not sent." }
    }
    if (-not $state.PSObject.Properties['SchemaVersion'] -or ($state.SchemaVersion -isnot [int] -and $state.SchemaVersion -isnot [long]) -or $state.SchemaVersion -ne 1 -or
        -not $state.PSObject.Properties['Enabled'] -or $state.Enabled -isnot [bool] -or
        -not $state.PSObject.Properties['Revision'] -or ($state.Revision -isnot [int] -and $state.Revision -isnot [long]) -or $state.Revision -lt 0) {
        throw 'Maintenance mail control is invalid; mail was not sent.'
    }
    foreach ($name in @('ChangedBy','ChangedFromServer','Reason')) {
        if ($state.$name -isnot [string]) { throw "Maintenance mail control $name is invalid; mail was not sent." }
    }
    foreach ($name in @('ChangedAtUtc','ResumeAfterUtc')) {
        if ($state.$name -is [datetime]) { $state.$name=$state.$name.ToUniversalTime().ToString('o') }
        elseif ($state.$name -is [datetimeoffset]) { $state.$name=$state.$name.UtcDateTime.ToString('o') }
        if ($state.$name -isnot [string]) { throw "Maintenance mail control $name is invalid; mail was not sent." }
        if ($state.$name) {
            $timestamp=[datetimeoffset]::MinValue
            if ($state.$name -notmatch '(Z|\+00:00)$' -or -not [datetimeoffset]::TryParse($state.$name,[ref]$timestamp)) { throw "Maintenance mail control $name is not UTC; mail was not sent." }
        }
    }
    if (($state.Enabled -and $state.Revision -eq 0) -or
        ($state.Revision -gt 0 -and (-not $state.ChangedAtUtc -or -not $state.ChangedBy -or -not $state.Reason)) -or
        (-not $state.Enabled -and $state.Revision -gt 0 -and $state.ResumeAfterUtc -ne $state.ChangedAtUtc) -or
        ($state.ResumeAfterUtc -and $state.ChangedAtUtc -and [datetimeoffset]::Parse($state.ResumeAfterUtc) -gt [datetimeoffset]::Parse($state.ChangedAtUtc))) {
        throw 'Maintenance mail transition evidence is invalid; mail was not sent.'
    }
    return $state
}

function Add-SmartM365AdminRecipients {
    param($Route,$Config,$Control)
    $property=$Config.PSObject.Properties['Smart365Administrators']
    $administrators=if($property){[string](Resolve-SmartM365ConfigValue $property.Value)}else{''}
    $result=Resolve-SmartM365AdminMailRoute -To $Route.To -Cc $Route.Cc -Smart365Administrators $administrators -AdministratorsOnly $Control.AdministratorsOnly -ExcludeAdministrators (Test-SmartM365MailExcludeAdministrators)
    return [pscustomobject]@{To=$result.To;Cc=$result.Cc;MailPurpose=$Route.MailPurpose;Restricted=($Route.Restricted -or $Control.AdministratorsOnly);AdministratorsOnly=[bool]$Control.AdministratorsOnly}
}

function Test-SmartM365MailInventoryCaller {
    # The delivery window applies to SmartInventory collectors and the orchestrator only;
    # other SmartM365 consumers (user communication campaigns, tools) keep their behavior.
    return [bool]@(Get-PSCallStack | Where-Object { $_.ScriptName -match '[\\/]SmartInventory[\\/]' }).Count
}

function Get-SmartM365MailDeliveryWindow {
    # Business-hours mode: ReportMailSchedule enabled with Days and/or NotAfter.
    # Without them the legacy daily NotBefore cutoff keeps its previous behavior.
    param($Config = (Get-SmartM365EffectiveModuleGlobalConfig))
    $property = $Config.PSObject.Properties['ReportMailSchedule']
    if (-not $property -or -not $property.Value -or $property.Value -is [string]) { return $null }
    $settings = $property.Value
    $read = { param($name) if ($settings -is [Collections.IDictionary]) { if ($settings.Contains($name)) { ,$settings[$name] } } else { $p = $settings.PSObject.Properties[$name]; if ($p) { ,$p.Value } } }
    $enabled = & $read 'Enabled'
    if ($enabled -isnot [bool] -or -not $enabled) { return $null }
    $days = & $read 'Days'
    $notAfter = & $read 'NotAfter'
    if ($null -eq $days -and $null -eq $notAfter) { return $null }
    if ($null -eq $days) { $days = @('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday') }
    if ($null -eq $notAfter) { $notAfter = '24:00' }
    return [pscustomobject]@{ Days = @($days | ForEach-Object { [string]$_ }); NotBefore = [string](& $read 'NotBefore'); NotAfter = [string]$notAfter; TimeZone = [string](& $read 'TimeZone') }
}

function Get-SmartM365MailDeliveryWindowState {
    param([Parameter(Mandatory)]$Window,[datetimeoffset]$Now = [datetimeoffset]::UtcNow)
    if (-not (Get-Command Get-SmartM365ReportMailWindowState -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ReportMailSchedule.psm1') -Global -Force -ErrorAction Stop
    }
    return Get-SmartM365ReportMailWindowState -Now $Now -Days $Window.Days -NotBefore $Window.NotBefore -NotAfter $Window.NotAfter -TimeZoneId $Window.TimeZone
}

function Test-SmartM365MailDeliveryWindowClosed {
    param($Config)
    if (Test-SmartM365MailExcludeAdministrators) { return $false }
    if (-not (Test-SmartM365MailInventoryCaller)) { return $false }
    try {
        $window = Get-SmartM365MailDeliveryWindow -Config $Config
        if (-not $window) { return $false }
        $now = if (Get-Command Get-SmartM365MailUtcNow -ErrorAction SilentlyContinue) { Get-SmartM365MailUtcNow } else { [datetimeoffset]::UtcNow }
        return -not (Get-SmartM365MailDeliveryWindowState -Window $window -Now $now).IsOpen
    } catch {
        # Never block administrators because of an invalid window: fall back to them only.
        $message = 'Mail delivery window is invalid ({0}); mail is restricted to Smart365Administrators.' -f $_.Exception.Message
        if (Get-Command WriteLog -ErrorAction SilentlyContinue) { WriteLog -Message $message -Level WARNING } else { Write-Warning $message }
        return $true
    }
}

function Resolve-SmartM365MaintenanceMailRecipientRoute {
    param([string]$To,[string]$Cc,
        [ValidateSet('Auto','Report','Error','Maintenance')][string]$MailPurpose='Auto')
    $config = Get-SmartM365EffectiveModuleGlobalConfig
    $adminControl = Get-SmartM365AdminMailControl -SharedDataFolderPath (Get-SmartM365MaintenanceMailRoot $config)
    if ($adminControl.AdministratorsOnly) {
        return Add-SmartM365AdminRecipients -Route ([pscustomobject]@{To=$To;Cc=$Cc;MailPurpose=$MailPurpose;Restricted=$true}) -Config $config -Control $adminControl
    }
    if (Test-SmartM365MailDeliveryWindowClosed -Config $config) {
        # Outside business hours only administrators receive mail; deferred reports
        # reach the other recipients through the queue when the window opens.
        return Add-SmartM365AdminRecipients -Route ([pscustomobject]@{To=$To;Cc=$Cc;MailPurpose=$MailPurpose;Restricted=$true}) -Config $config -Control ([pscustomobject]@{AdministratorsOnly=$true})
    }
    $state = Get-SmartM365MaintenanceMailState -SharedDataFolderPath (Get-SmartM365MaintenanceMailRoot $config)
    if ($MailPurpose -eq 'Auto') {
        $MailPurpose = 'Report'
        # Preserve explicit ErrorMailTo intent, never guess from subject or priority.
        $caller = Get-SmartM365CallerLocalConfig
        $errorTo = [string](Get-ModuleLocalConfigValue -Config $caller -Name ErrorMailTo -DefaultValue '')
        $normalTo = [string](Get-ModuleLocalConfigValue -Config $caller -Name To -DefaultValue '')
        $normalize = { param($value) (@(([string]$value -split '[;,]') | ForEach-Object {$_.Trim().ToLowerInvariant()} | Where-Object {$_} | Sort-Object -Unique) -join ';') }
        if ($errorTo -and (& $normalize $To) -eq (& $normalize $errorTo) -and (& $normalize $To) -ne (& $normalize $normalTo)) { $MailPurpose = 'Error' }
        # Existing collectors explicitly mark terminal failure before their legacy mail call.
        $failure = Get-Variable ScriptFailed -Scope Global -ErrorAction SilentlyContinue
        if ($failure -and $failure.Value -is [bool] -and $failure.Value) { $MailPurpose = 'Error' }
    }
    if (-not $state.Enabled -and $MailPurpose -ne 'Maintenance') {
        if ($MailPurpose -eq 'Error') {
            $moduleConfig = Get-ModuleLocalConfig
            $callerConfig = Get-SmartM365CallerLocalConfig
            $recipient = [string](Get-ModuleLocalConfigValue -Config $moduleConfig -Name ErrorMailTo -DefaultValue '')
            $recipient = [string](Get-ModuleLocalConfigValue -Config $callerConfig -Name ErrorMailTo -DefaultValue $recipient)
            if ([string]::IsNullOrWhiteSpace($recipient) -or $recipient.Trim() -in @('__USE_GLOBAL__','USE_GLOBAL') -or $recipient -match '\{\{') {
                throw 'ErrorMailTo is not configured; error mail was not sent.'
            }
            return Add-SmartM365AdminRecipients -Route ([pscustomobject]@{To=$recipient;Cc='';MailPurpose='Error';Restricted=$true}) -Config $config -Control $adminControl
        }
        return Add-SmartM365AdminRecipients -Route ([pscustomobject]@{To=$To;Cc=$Cc;MailPurpose=$MailPurpose;Restricted=$false}) -Config $config -Control $adminControl
    }
    $key = if ($MailPurpose -eq 'Error') {'ErrorMailTo'} else {'To'}
    $property = $config.PSObject.Properties[$key]
    $recipient = if ($property) { [string](Resolve-SmartM365ConfigValue $property.Value) } else { '' }
    if ([string]::IsNullOrWhiteSpace($recipient) -or $recipient -in @('__USE_GLOBAL__','USE_GLOBAL') -or $recipient -match '\{\{') {
        throw "Global $key is not configured; restricted mail was not sent."
    }
    return Add-SmartM365AdminRecipients -Route ([pscustomobject]@{To=$recipient;Cc='';MailPurpose=$MailPurpose;Restricted=$true}) -Config $config -Control $adminControl
}

function Set-SmartM365SmtpMaintenanceRecipient {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Only rewrites the in-memory SMTP parameter map; transport owns the send operation.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$MailParameters,
        [ValidateSet('Auto','Report','Error','Maintenance')][string]$MailPurpose='Auto')
    $route = Resolve-SmartM365UnifiedMailRoute -To ($MailParameters['To'] -join ';') -Cc ($MailParameters['Cc'] -join ';') -Bcc ($MailParameters['Bcc'] -join ';') -MailPurpose $MailPurpose
    $MailParameters.To = @(ConvertToRecipientArray $route.To)
    if ($route.Cc) { $MailParameters.Cc = @(ConvertToRecipientArray $route.Cc) } else { $MailParameters.Remove('Cc') }
    if ($route.Bcc) { $MailParameters.Bcc = @(ConvertToRecipientArray $route.Bcc) } else { $MailParameters.Remove('Bcc') }
}

function Resolve-SmartM365UnifiedMailRoute {
    [CmdletBinding()]
    param([string]$To,[string]$Cc,[string]$Bcc,[ValidateSet('Auto','Report','Error','Maintenance')][string]$MailPurpose='Report')
    $config=Get-SmartM365EffectiveModuleGlobalConfig
    $admin=[string](Resolve-SmartM365ConfigValue $config.Smart365Administrators)
    $replay=Get-Variable SmartM365ReportReplayAudience -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($replay -eq 'Administrators') {
        return Resolve-SmartM365AdminMailRoute -Smart365Administrators $admin -AdministratorsOnly $true
    }
    $base=Resolve-SmartM365MaintenanceMailRecipientRoute -To $To -Cc $Cc -MailPurpose $MailPurpose
    if ($replay -eq 'Configured' -and $base.Restricted) { throw 'Normal recipients are currently restricted by the mail policy.' }
    $hidden=if($base.Restricted){''}else{$Bcc}
    $visibleTo=$base.To
    if(-not $base.Restricted -and $hidden){
        $hiddenSet=@($hidden -split '[;,]' | ForEach-Object {$_.Trim().ToLowerInvariant()} | Where-Object {$_})
        $originalTo=@($To -split '[;,]' | ForEach-Object {$_.Trim().ToLowerInvariant()} | Where-Object {$_})
        $visibleTo=@($base.To -split '[;,]' | Where-Object {$_.Trim().ToLowerInvariant() -notin $hiddenSet -or $_.Trim().ToLowerInvariant() -in $originalTo}) -join ';'
    }
    $final=Resolve-SmartM365AdminMailRoute -To $visibleTo -Cc $base.Cc -Bcc $hidden -Smart365Administrators $admin -AdministratorsOnly $false -ExcludeAdministrators (Test-SmartM365MailExcludeAdministrators)
    $onlyProperty=$base.PSObject.Properties['AdministratorsOnly']
    return [pscustomobject]@{To=$final.To;Cc=$final.Cc;Bcc=$final.Bcc;AdministratorsOnly=($null -ne $onlyProperty -and [bool]$onlyProperty.Value)}
}


# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhW06VXO3OuZ18
# ux6iWDQh3nBluutvk3X1x1MKDmVSDaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgXCHOiAr6gGLV
# zFKESGYz6RrW7puMFwzfCK/PUA3PXfcwDQYJKoZIhvcNAQEBBQAEggGAEyS3vl+f
# MsJjCYKu7NM7hHOYajRlu4o2f92iiZ/RleRZGz5hGk7UMnxS1++KDS2QOkQ/pCNt
# CLaOWupFaZ8oNRXGuQm/2HquTsLDz7cYheIGtJvovZzdPCuyuvLty65Gbjujlu31
# lBBre4Ef2XGBcfsr2rMrSwi6ol+Yjn4ChGzbiOYXJ2ErWxgltunqdCT3QjCtaxEt
# 7mspTBKHn1g6Zb3Tmd+g/zpWmvSqbcjZf76DLbp+lofuv8c88uuIkSHIzXiA690F
# CSBAk3y5Gm/DupKAnfybwxW6VJSRLxAn6RMycAPe3j6w9SS6ZO7kxpwNi+aIfbBf
# H1v5PfcQlyA7HQ2s89skOHkrs90IrIwrdbnnTsJmmdHGGJYIR5fecG3gKlyNM13Y
# REtpIo+S9ZwnZUWNrN85UufBolIOvX9cKop4CmyEl/f8X+32EoVKpQMfJq6CrD18
# bCuHuU2uC3gEKK3sapKBtFQrlh/lcqsUfHUO37/4S4ycLv6iYyIesTN+oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDgxODAwWjA/BgkqhkiG9w0BCQQxMgQwnDNFPoDkyBBdjPzy2q1BmKTk
# UH/OikDcNYalobxf5J/bHD/XyHX6hG447mm3VOj/MA0GCSqGSIb3DQEBAQUABIIC
# ABUpR4jkb6asoQvpPWPh45doO+62Kgk/yC1iAbNk2VJT24kB0PBD9qKTFKj/pul3
# bZ9Ux2xEjwrT1PjXOBNj+xv9veRKLK++jIrb68F/6h8hiffQ+Dme2WMfwoOlt5KK
# dX4Zf9UvP0t6j2dGG1ga1WcxlgwQTJ3OqlK3kLjZJF8xRU4Wzdy8GHuamcy0dO9+
# 81VNeGBfi9g1LvsxEGkB3taLpRdZNxSWPpQszQ8X8+UcE1Bu89fjGVrwyc4SHJA3
# VcLL58tOieLDnbBRsI5Wd4AbEpCkcp5pA1N9E5rOViVCeqIrFx4bCm0/iITNiIew
# i/B+4BE0FHtRPy1PGAYwFuNJQXPOK5XwQed2uKC+6OmFAAotvMwOzqjh65KFp1bd
# /4pxe5W4UKg5POT3IsjhxVC5VEuTLS66n8VKytIEGHKROgI6O6wTYpne1srx1j4D
# /tC13+T4gwpcwk4/vc2SikOui01ArcEQkUTB7udwG9IkPvZ4DivVsNiSTcZ0UlsR
# ZlgwqDSHBhYXqcE+m0v8uTk4BPie/8bgH4vvgmZuarMTMLV5k1uLaq+uy3ddElVF
# cdgl3pQtg4nXmQgLfTT4SV7IwblNfumWQvy2C875a1p5qekouO9AbC+1IGk6dYbE
# up4u9f1XWH3f3AqZIIVD4oqSNDJKAkOcczn0Grf8m0tC
# SIG # End signature block
