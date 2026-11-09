# SmartInventory configuration audit helpers, version 0.3.0
# Passive inspection only: never invoke the tenant configuration loader here.
function Read-SiPrivateJson {
    param([Parameter(Mandatory)][string]$BasePath)
    $plain = $BasePath + '.json'
    $text = $plain + '.txt'
    $found = @($plain, $text | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
    if ($found.Count -gt 1) { return [pscustomobject]@{ State='Conflict'; Value=$null } }
    if ($found.Count -eq 0) { return [pscustomobject]@{ State='Missing'; Value=$null } }
    try {
        $obj = Get-Content -LiteralPath $found[0] -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $obj -or $obj -is [array] -or $obj -isnot [pscustomobject]) { throw 'Expected JSON object' }
        return [pscustomobject]@{ State='Present'; Value=$obj }
    } catch { return [pscustomobject]@{ State='Unreadable'; Value=$null } }
}
function Test-SiConfiguredValue {
    param($Value)
    if ($null -eq $Value) { return $false }
    $s = ([string]$Value).Trim()
    return ($s -ne '' -and $s -notmatch '^(?i:__USE_GLOBAL__|USE_GLOBAL|0{8}-0{4}-0{4}-0{4}-0{12}|0{40})$' -and $s -notmatch '(?i)@example\.com$|^contoso\.')
}
function Get-SiLocalIdentityState {
    param($Profile)
    if ($null -eq $Profile) { return 'Not checked' }
    $app=[string]$Profile.AppId; $tenant=[string]$Profile.TenantId
    $thumb=([string]$Profile.Thumb).Replace(' ','').ToUpperInvariant()
    $thumbprint=([string]$Profile.Thumbprint).Replace(' ','').ToUpperInvariant()
    if ($app -in @('__USE_GLOBAL__','USE_GLOBAL') -or $tenant -in @('__USE_GLOBAL__','USE_GLOBAL') -or $thumb -in @('__USE_GLOBAL__','USE_GLOBAL') -or $thumbprint -in @('__USE_GLOBAL__','USE_GLOBAL')) { return 'Local values incomplete; inheritance not checked' }
    if (-not (Test-SiConfiguredValue $app) -or -not (Test-SiConfiguredValue $tenant) -or (-not (Test-SiConfiguredValue $thumb) -and -not (Test-SiConfiguredValue $thumbprint))) { return 'Local values incomplete; inheritance not checked' }
    $appGuid=[guid]::Empty; $tenantGuid=[guid]::Empty
    if (-not [guid]::TryParse($app,[ref]$appGuid) -or $appGuid -eq [guid]::Empty -or -not [guid]::TryParse($tenant,[ref]$tenantGuid) -or $tenantGuid -eq [guid]::Empty) { return 'Local identifiers need correction' }
    if ($thumb -and $thumbprint -and $thumb -ne $thumbprint) { return 'Local thumbprints differ' }
    $effective=if ($thumbprint) { $thumbprint } else { $thumb }
    if ($effective -notmatch '^[0-9A-F]{40}$' -or $effective -match '^0+$') { return 'Local thumbprint needs correction' }
    return 'Local format consistent; access not checked'
}
function Get-SiConfigurationAudit {
    param([Parameter(Mandatory)][string]$SmartM365Root, [Parameter(Mandatory)][string]$Tenant)
    if ($Tenant -notmatch '^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$') { throw 'Invalid tenant profile key.' }
    $global = Read-SiPrivateJson (Join-Path $SmartM365Root 'Config/SmartM365.global.local')
    $profile = Read-SiPrivateJson (Join-Path $SmartM365Root ("Config/Tenants/$Tenant.local"))
    $rows = [System.Collections.Generic.List[object]]::new()
    function Add-Row([string]$Step,[string]$State,[string]$Detail) {
        $rows.Add([pscustomobject]@{ Step=$Step; State=$State; Detail=$Detail })
    }
    Add-Row 'Global configuration' $global.State 'Passive read of the private file; no initialization or migration.'
    Add-Row 'Tenant profile' $profile.State 'The profile is local; its presence does not verify the remote tenant.'
    $p = $profile.Value
    $identity = $null -ne $p -and (Test-SiConfiguredValue $p.TenantId) -and (Test-SiConfiguredValue $p.AppId) -and ((Test-SiConfiguredValue $p.Thumbprint) -or (Test-SiConfiguredValue $p.Thumb))
    if ($null -eq $p) { $identityState='Unknown' } elseif ($identity) { $identityState='Review required' } else { $identityState='Input required' }
    Add-Row 'Identity and permissions' $identityState 'Identifiers and certificate: local presence only. Consent, key access, roles and Exchange RBAC require separate checks.'
    if ($null -eq $p) { $adminState='Unknown' } elseif (Test-SiConfiguredValue $p.Smart365Administrators) { $adminState='Review required' } else { $adminState='Input required' }
    Add-Row 'Smart365 administrators' $adminState 'Verify recipients and access; a local value does not prove validity.'
    Add-Row 'Prerequisites and hosts' 'Review required' 'Audit modules under the service account, PS7, PS5.1 and workload-specific dependencies.'
    Add-Row 'Shared folder and access' 'Review required' 'Qualify the effective path and service-account read/write rights on each server.'
    if ($null -ne $p -and $p.EnableTeamsNotifications -eq $true) { $teamsState='Review required' } else { $teamsState='Disabled or unknown' }
    Add-Row 'Teams and notifications' $teamsState 'Check both Alerts and Infos workflows; no webhook is displayed or tested.'
    $spState='Review required'
    Add-Row 'SharePoint' $spState 'Check effective category policies, the site, Sites.Selected and site write permission.'
    Add-Row 'Mail' 'Review required' 'Check transport, sender, recipients and Exchange App RBAC restriction. No mail sent.'
    Add-Row 'Collectors and dependencies' 'Review required' 'Review enabled jobs, dependencies, private configuration and AD/Exchange/SharePoint hosts.'
    Add-Row 'Orchestrator' 'Review required' 'Check the effective catalog, service account, shared access, scheduled task and each server.'
    return $rows.ToArray()
}

function Get-SiStepAudit {
    param([Parameter(Mandatory)][string]$SmartM365Root, [Parameter(Mandatory)][string]$Tenant)
    if ($Tenant -notmatch '^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$') { throw 'Invalid tenant profile key.' }
    $global = Read-SiPrivateJson (Join-Path $SmartM365Root 'Config/SmartM365.global.local')
    $profile = Read-SiPrivateJson (Join-Path $SmartM365Root ("Config/Tenants/$Tenant.local"))
    $p = $profile.Value
    $state = @{}
    foreach ($step in (Get-SiConfigurationSteps)) { $state[$step.Id] = 'Not checked' }
    $state.welcome = 'Not selected'
    if ($global.State -eq 'Present' -and $profile.State -eq 'Present') { $state.discovery = 'Files present' }
    elseif ($global.State -in @('Conflict','Unreadable') -or $profile.State -in @('Conflict','Unreadable')) { $state.discovery = 'Correction required' }
    else { $state.discovery = 'Files missing' }
    if ($null -ne $p) {
        $expected = ([string]$p.OrganizationKey).Trim() + '-' + ([string]$p.EnvironmentKey).Trim()
        if (([string]$p.ProfileKey -ieq $Tenant) -and (Test-SiConfiguredValue $p.OrganizationKey) -and (Test-SiConfiguredValue $p.EnvironmentKey) -and ([string]$p.TenantKey -ieq $expected)) { $state.tenant = 'Local consistency checked' }
        else { $state.tenant = 'Input or correction required' }
        $state.identity = Get-SiLocalIdentityState -Profile $p
        if (Test-SiConfiguredValue $p.Smart365Administrators) { $state.administrators = 'Value present' } else { $state.administrators = 'Input required' }
        if ($p.EnableTeamsNotifications -eq $false) { $state.teams = 'Disabled locally' }
        elseif ($p.EnableTeamsNotifications -eq $true) { $state.teams = 'Enabled; test required' }
    }
    foreach ($step in (Get-SiConfigurationSteps)) {
        [pscustomobject]@{ Id=$step.Id; Title=$step.Title; State=$state[$step.Id]; Detail=$step.Detail; Action=$step.Action }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCaM4p0/XZYmoTc
# LFMObZpFq3B6xuxuR6wCw3vjPRiSxqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFHJUIpT70RKSiVIaR2UC1tbajyoFUr7XwR4CAaEzVvgMA0GCSqG
# SIb3DQEBAQUABIIBgBHRAs96WcMcC9xTbppukmF+swo2W+p/c5eJra+ftgyo/87A
# Jthi3aJ21qXBQCvuh2eQqthmt0lvvbdhlJgoL+PERN6IDMdEJlJnCTIsfKyXZhzR
# 4+w7izc4LHtTD/LrFm9wpHJ/Fz3UG9l28wdgOrvfsJPcCmx2FGmtem83HABVizVs
# h9j4rPN7AErQGAAxdmwzvrCDATSWGb4o+iYB25SSGktXK/yYYpNch6SYbM64uUuZ
# kmU5Ugy32WJySfpK6Zorua0xApmUaybHOwDqZ7nXuWz5WbU6MUa5F/9/HAaFwtH0
# BobBVt4LPg1FnN5An51f11MzpDVHqt6lQjfs7zhhOHE8boYKChVjXTZZSfkdhiFy
# wjlXtdCe1sMmchao19y8z4BkEYWBPGkqaPF/a70BA8gRlO9Ivq63ojRMh+l5sH89
# /IxObnRs3mID11wGylBxk7H4bIFV72au2qJ2wgEUD6u4hqk6JfC2leUJ/Jupd/m5
# iciB47C+gnwd2t5SHqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjVaMC8GCSqGSIb3DQEJBDEiBCBZ7k/womGURfgmqeIzbtzLLma4PnDXXyOU5OOq
# GDFNdjANBgkqhkiG9w0BAQEFAASCAgATsRULuatZ9zQ5U2S1jTvrbOaOvJI3JnAn
# lWr+zl6WWt2PLLhDb18fSnpMyUjZEZlz0VVjCezJpN9Z111/jsA1aMDfdsVRwv8K
# VyEZ4hzXz8As5oK6QpmB2ewhrPpwb1Bvv0TW2xvXJzIcMCXKp4TmvooOEC8l+Ch3
# gyIzSMQXlaKqUPXUJC60QOnLUzNI2cucgvITL5mvx9u2CfH7QgEnvJBG0o0Gtdxq
# vcSuZat0vRif2NREYLzag+3+Loo74bCuMIKPIJ+tK9hnn4R8gDckL2rDWwP08GcA
# vRQYCNggWCWMqo88Ii46241UQIo/1Ti6tBsATXKsL59JaCQhdwp6/14YrXBDSnEr
# rlWkZNB3Z5c3HeloPLL0/RWvdtvkASuipILkFo91kGqf+jskB+/3V0a5pz0PDpqR
# AUS2p6HVxh3gPCyzsfnbS/MDRmhtBseQULeDHxlaKSPdA9NSUCPE6heKBznR4eVe
# HUubhdlpN5yfyszz0o01S6wzw5hkRyzhYoUo1g2FLUQajMQawLjnUc5mEZ9mEnTL
# ZLkMc8DZhs+TJ0zrIqp3fpPZyxcIhCg9M6TF/Rycwdyf7+iP/IIMSFzuWZLwRMrG
# VTuLjSdRrwAnnWxRYvr5P4M2YsPGfwfFvmcUNxzg7nmiYkYj3ngFryuhYHT+t8s6
# YFCO/uV5ZQ==
# SIG # End signature block
