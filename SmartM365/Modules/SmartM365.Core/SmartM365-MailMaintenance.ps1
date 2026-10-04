<#
.SYNOPSIS
Shared mail admission policy, dot-sourced by both Core editions. No I/O on load.
.VERSION
1.0.0
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

function Resolve-SmartM365MaintenanceMailRecipientRoute {
    param([string]$To,[string]$Cc,
        [ValidateSet('Auto','Report','Error','Maintenance')][string]$MailPurpose='Auto')
    $config = Get-SmartM365EffectiveModuleGlobalConfig
    $state = Get-SmartM365MaintenanceMailState -SharedDataFolderPath (Get-SmartM365MaintenanceMailRoot $config)
    if (-not $state.Enabled -and $MailPurpose -ne 'Maintenance') {
        return [pscustomobject]@{To=$To;Cc=$Cc;MailPurpose=$MailPurpose;Restricted=$false}
    }
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
    $key = if ($MailPurpose -eq 'Error') {'ErrorMailTo'} else {'To'}
    $property = $config.PSObject.Properties[$key]
    $recipient = if ($property) { [string](Resolve-SmartM365ConfigValue $property.Value) } else { '' }
    if ([string]::IsNullOrWhiteSpace($recipient) -or $recipient -in @('__USE_GLOBAL__','USE_GLOBAL') -or $recipient -match '\{\{') {
        throw "Global $key is not configured; restricted mail was not sent."
    }
    return [pscustomobject]@{To=$recipient;Cc='';MailPurpose=$MailPurpose;Restricted=$true}
}

function Set-SmartM365SmtpMaintenanceRecipient {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Only rewrites the in-memory SMTP parameter map; transport owns the send operation.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$MailParameters,
        [ValidateSet('Auto','Report','Error','Maintenance')][string]$MailPurpose='Auto')
    $route = Resolve-SmartM365MaintenanceMailRecipientRoute -To ($MailParameters['To'] -join ';') -Cc ($MailParameters['Cc'] -join ';') -MailPurpose $MailPurpose
    $MailParameters.To = @(ConvertToRecipientArray $route.To)
    if ($route.Restricted) { $MailParameters.Remove('Cc'); $MailParameters.Remove('Bcc') }
    elseif ($route.Cc) { $MailParameters.Cc = @(ConvertToRecipientArray $route.Cc) }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAMtKehO0xuBFlN
# 5ofm9CCvus2PkJKsxKKuAmoOZACKU6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBKh9B5tapOnqGoD0F7UzvtYUPpW2z8KHM9trRXlWBr9MA0GCSqG
# SIb3DQEBAQUABIIBgJxNyWhr+wHURpCyovGyQPJfqtElwzLy4XhzfRkJs9TDlfnw
# CyUPmLnLaGb0jrxZG5iSHT7i0geRUFfL6HTValC7eREb98gnRbp6RQBJcs+sh+wM
# fOKMloVN6gTqQAvaIP7X6JXgPchPG4rjFtBrVP0q/6x3g9myChWv14jzZpB57EI8
# 3viGwctipUmes/T0zihMhwpItSdN3PLYy4WGK/Vbzk8yRe/CIu6yfTh4LRO54CTH
# jjVEF4f0dspj6PjFEczIsawHArlvtaF8f+ircUPFD3Mr/PHhQbYupIexSEEW/DOt
# FWhdBn3VacWGc+a4ofgTDH3hiXWP/lz+OqId6S1VERin6d82IXIvFjFmmpNjEVFn
# fPwc4PY99miusucLB9/FyakQ1PNgpuaMetM8/2T1hf429+o0S4Yr9uq/bYTeltCY
# nX8CPJ6Ak625TAWPbJC0jspAbuqyCCWE/T0P6uQsiP63SOjc1yo0bQ2i8e9Uv7Xr
# Fkhn7JhAeS0AMMsOYKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQxODM3
# NDVaMC8GCSqGSIb3DQEJBDEiBCAr+CEkyLZaSrOe/JmJlCfTF5l84RqgahqQiVIW
# GSdZZTANBgkqhkiG9w0BAQEFAASCAgBRsC4qTmid/NG5eRfMVUM1wtlPGUE54KRe
# dbjRpSWdtuFEkwCGA5WRi67G5Pxjk6Hnulx/LfELfokoeHL1edlQWs3g/gNWYfaV
# GMulYSPSHuw4YKj/9eGL4TKJd3oKN1xjpZlvKTYpvaaBETbQd+QxPql8K+4TZQIE
# r9VVS+XDorG9J0USNGyQcOjcPKuHyzMY6MPQgcu0QWL5EAxqjVyz1XENKhFq5i1V
# HjjsIt4pGrqPXrPXnJCxcKzSUZgHg3vUG4PG+sIyqt9+oXHk3qSJTQGY/rtHNpJN
# vHyKI73iswXBc9B9cKPk84o/VJ1rNTZc2iFVZO5thbHx3ERPfADGae6MuKs/ov0J
# 8nC5PWyWJ93U+vN8RqTEgJF3R5gD5k8MGj5z4jQlYF7Dmo73vO35vINx+u2bPa5c
# YppZ8zPaQisOU7dqXlJAWHX9zNXbF8BRKc6N07DaoaOdYAlzu3uk4fc4gKFL3opp
# KPYEZLaL5i7uIdwOqQ0Getx9Bzf44sj1oHchkG2+cM+UlcGhXZlydaLKQPfNcTTA
# PNMRM4wPqQRk18jFzq5MQ4Kc5dXtdELI5saM/yljbzyMAPpMgs0aYRfgFoHDLZXT
# rJVy4zh2LE56vgLLXuhG9M/OWhmcd0YF9p/4CaNJJz4wmMPw4Ujvbic31UoQb2np
# 4bCw16pUtQ==
# SIG # End signature block
