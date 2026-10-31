<#
.SYNOPSIS
Shared mail admission policy, dot-sourced by both Core editions. No I/O on load.
.VERSION
1.0.2
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
    $result=Resolve-SmartM365AdminMailRoute -To $Route.To -Cc $Route.Cc -Smart365Administrators $administrators -AdministratorsOnly $Control.AdministratorsOnly
    return [pscustomobject]@{To=$result.To;Cc=$result.Cc;MailPurpose=$Route.MailPurpose;Restricted=($Route.Restricted -or $Control.AdministratorsOnly);AdministratorsOnly=[bool]$Control.AdministratorsOnly}
}

function Resolve-SmartM365MaintenanceMailRecipientRoute {
    param([string]$To,[string]$Cc,
        [ValidateSet('Auto','Report','Error','Maintenance')][string]$MailPurpose='Auto')
    $config = Get-SmartM365EffectiveModuleGlobalConfig
    $adminControl = Get-SmartM365AdminMailControl -SharedDataFolderPath (Get-SmartM365MaintenanceMailRoot $config)
    if ($adminControl.AdministratorsOnly) {
        return Add-SmartM365AdminRecipients -Route ([pscustomobject]@{To=$To;Cc=$Cc;MailPurpose=$MailPurpose;Restricted=$true}) -Config $config -Control $adminControl
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
    $base=Resolve-SmartM365MaintenanceMailRecipientRoute -To $To -Cc $Cc -MailPurpose $MailPurpose
    $config=Get-SmartM365EffectiveModuleGlobalConfig
    $admin=[string](Resolve-SmartM365ConfigValue $config.Smart365Administrators)
    $hidden=if($base.Restricted){''}else{$Bcc}
    $visibleTo=$base.To
    if(-not $base.Restricted -and $hidden){
        $hiddenSet=@($hidden -split '[;,]' | ForEach-Object {$_.Trim().ToLowerInvariant()} | Where-Object {$_})
        $originalTo=@($To -split '[;,]' | ForEach-Object {$_.Trim().ToLowerInvariant()} | Where-Object {$_})
        $visibleTo=@($base.To -split '[;,]' | Where-Object {$_.Trim().ToLowerInvariant() -notin $hiddenSet -or $_.Trim().ToLowerInvariant() -in $originalTo}) -join ';'
    }
    $final=Resolve-SmartM365AdminMailRoute -To $visibleTo -Cc $base.Cc -Bcc $hidden -Smart365Administrators $admin -AdministratorsOnly $false
    $onlyProperty=$base.PSObject.Properties['AdministratorsOnly']
    return [pscustomobject]@{To=$final.To;Cc=$final.Cc;Bcc=$final.Bcc;AdministratorsOnly=($null -ne $onlyProperty -and [bool]$onlyProperty.Value)}
}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBahipvw8QhsjcZ
# gQNnKCVp2biBsl+2L+5KFnv+jX5iu6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOTFPp322BDjvIW+NxTqciSoVKEhBawn3Znh515HqsHYMA0GCSqG
# SIb3DQEBAQUABIIBgAA/9pXGqSRyLxpgXVPmGrsO8OVv+aMuhjwgRVlYBdi5cZMc
# j8WorkPcqYvEBmpeHkJW4Fhrid1ubE4wGU0g6E+CaXaWEHd/DpUXor1cT+3WDgIF
# 14NxexBAd02C2hd+iNjFzpTXxIk85HZ6k2VXtAQoyBLKHGDiFUoU9D/sEhmxp669
# bNUJNy3KPSYvLFisyTSfOJFTxt3RHzb8/zKX4doAilN9fpdaHYADfr6IrYoc7uug
# qIFmy1kqC6NmRCT1Lu8+kweI9A9IqkOG7rgwjxWA/rePABu235SULA0SwWTjeUOy
# RqiaVO7vJmScVeSYISBAcH3YFUOWe3txQO4JOxdPJZkhogMQQUybULXer2K/H1Lj
# fSWb/CVYJxFyLquqNQwXYAWOgTJs1NjE/0O2cJnVuPU5m+AjhXEKjFBpHWUResbG
# w5B7cAKjmfubtNRsTDSphqmzRrkISwaYQ0W4FBbx8RLE7bjsI5VW4sO7QNbVcjzH
# CEwTjGGavTKmlrPInaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExNjE4
# MjRaMC8GCSqGSIb3DQEJBDEiBCAA+KhDatZU+u3xytmWjoA0zfyMsRpcl+jFt0T/
# xpXaTTANBgkqhkiG9w0BAQEFAASCAgAVLFAYeS6PcSqEPgm3fFVgaDMzYTtWpCsG
# f6bmEcP/0WtTB5Gw9N6ABDwv73iRac5nRsWW/BcmOSPT4a46U6UXOLdwtKUeNOzt
# d7w6REw67NE8Wrfhuansh3Q6ZKTwnL52UGhWLax9QCf6qPOBbFhZ6Si82R+2kmNT
# xVcKOARZ/XgK87vswrQZM2KHrFj2Nk8gVkWxRkUwUOrmmunIaTfNFcsmnC23s/VO
# zjEAAO+YVwd3gnrIL/6vsGrZuw3L38jsc8KcVS/O8fib+7F6rTx5XtlqDLkhSxOP
# kfmUBH3kQ9NQdDYYFU06sSgjPcua1CjFBimaKLduP3B/B5jh4c59U+GQll4OMhsK
# PbNvMtXIoOKi5+sxT5hKHMwfxDG64NpvOSYZyi0F+zRRRQ7QhEBYc6+l1ika9+Fv
# Cvyi+vNo0bgySuUnoxZY9noZnEQB/sWc0S77G3J6g8+NxRny134yQlCkwF2TQv7f
# RvxJoJYBP6v3BjSsHtPwqFfObvuCWQ9b6M7S/z8rC6Z/kn8GxKv6vbDGMfilcTK4
# OqFkttoVPo7dL7CdrcwcmU0NoqD6g/AWdXEaqj1zsQaZVYUOtxgLJFiWw1Nk6j53
# tqKXSSS0XkOU6AA8hZw00YMu9xs5iD/WwD1KMKTPmRAyzcJzgS41K8gJhSuCi/O9
# QfWDSX21ig==
# SIG # End signature block
