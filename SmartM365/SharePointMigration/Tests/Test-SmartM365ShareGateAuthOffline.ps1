<#
.SYNOPSIS
    Verify ShareGate endpoint authentication routing and same-host boundaries offline.
.VERSION
    1.0.5
#>
#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ShareGateAuth-' + [guid]::NewGuid().ToString('N'))
try {
    $module = Join-Path $root 'ShareGate'
    New-Item -ItemType Directory -Path $module -Force | Out-Null
    "@{ RootModule='ShareGate.psm1'; ModuleVersion='99.0.0'; GUID='7a75c782-6f07-4a08-84ba-91e7a1037883'; FunctionsToExport=@('Connect-Site') }" | Set-Content -LiteralPath (Join-Path $module 'ShareGate.psd1') -Encoding UTF8
    @'
function Connect-Site { [CmdletBinding()] param([string]$Url,[switch]$Browser,[switch]$ModernAuth,$UseCredentialsFrom,$UseConnection) [pscustomobject]@{Url=$Url; Address=$Url; Browser=[bool]$Browser; ModernAuth=[bool]$ModernAuth; UseCredentialsFrom=$UseCredentialsFrom; UseConnection=$UseConnection} }
Export-ModuleMember -Function Connect-Site
'@ | Set-Content -LiteralPath (Join-Path $module 'ShareGate.psm1') -Encoding UTF8
    Import-Module -Name (Join-Path $module 'ShareGate.psd1') -Force
    . (Join-Path $PSScriptRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateAuth.ps1')
    foreach ($mode in @('Default','Browser','ModernAuth')) {
        $site = Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode $mode
        if ($site.Browser -ne ($mode -eq 'Browser') -or $site.ModernAuth -ne ($mode -eq 'ModernAuth')) { throw "Wrong switches for $mode" }
    }
    $sourceConnection=[pscustomobject]@{Address='https://example.test/sites/a/';Marker='source'}
    $targetConnection=[pscustomobject]@{Address='https://example.test/sites/b';Marker='target'}
    Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/a' -Mode Browser -Site $sourceConnection
    Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/b/' -Mode Browser -Site $targetConnection
    if (-not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode Browser),$sourceConnection)) { throw 'Source session was not reused.' }
    if (-not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/b' -Mode Browser),$targetConnection)) { throw 'Target session was not isolated.' }
    $differentMode=Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode ModernAuth
    if (-not $differentMode.ModernAuth) { throw 'A session was reused for a different authentication mode.' }
    $rejected=$false
    try { Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/b' -Mode Browser -Site $sourceConnection } catch { $rejected=$true }
    if (-not $rejected) { throw 'A connection with the wrong site address was accepted.' }
    $rejected=$false
    try { Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/a' -Mode Browser -Site ([pscustomobject]@{Address='https://example.test/sites/a'}) } catch { $rejected=$true }
    if (-not $rejected) { throw 'A registered connection was silently replaced.' }
    $sourceConnection.Address='https://example.test/sites/b'
    $rejected=$false
    try { Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode Browser | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'A changed session address was reused.' }
    Clear-SmartM365ShareGateSessionConnections
    $defaultAfterClear=Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode Browser
    if (-not $defaultAfterClear.Browser) { throw 'Clearing the session did not restore default behavior.' }
    $credentialSite=[pscustomobject]@{Address='https://tenant.sharepoint.com/sites/source'}
    $fromCredentials=Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'https://tenant.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite
    if (-not [object]::ReferenceEquals($fromCredentials.UseCredentialsFrom,$credentialSite) -or -not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url $fromCredentials.Address -Mode Browser),$fromCredentials)) { throw 'Explicit credential reuse was not registered.' }
    $blocked=$false
    try { Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'https://other.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite | Out-Null } catch { $blocked=$true }
    if(-not $blocked){throw 'Cross-host credential reuse was not explicitly guarded.'}
    $blocked=$false
    try { Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'http://tenant.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite -AllowDifferentHost | Out-Null } catch { $blocked=$true }
    if(-not $blocked){throw 'Protocol changes must not reuse credentials.'}
    $other=Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'https://other.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite -AllowDifferentHost
    if(-not [object]::ReferenceEquals($other.UseCredentialsFrom,$credentialSite)){throw 'Explicit cross-host reuse failed.'}
    $saved=[pscustomobject]@{Id='synthetic'}
    $fromSaved=Connect-SmartM365ShareGateEndpointWithSavedConnection -Url 'https://third.sharepoint.com/sites/target' -Mode Browser -SavedConnection $saved
    if(-not [object]::ReferenceEquals($fromSaved.UseConnection,$saved) -or -not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url $fromSaved.Address -Mode Browser),$fromSaved)){throw 'Explicit saved connection was not registered.'}
    Clear-SmartM365ShareGateSessionConnections
    & {
        function Get-Module { param([switch]$ListAvailable,[string]$Name) }
        function Get-ItemProperty { param([string]$Path,$ErrorAction) [pscustomobject]@{DisplayName='ShareGate Migrate';InstallLocation=$root} }
        $discovered = Get-SmartM365ShareGateModule
        if (-not $discovered -or $discovered.Path -ne (Join-Path $root 'ShareGate\Sharegate.psd1') -or $discovered.Version -ne [version]'99.0.0') {
            throw 'Registry fallback did not resolve and validate the installed module manifest.'
        }
    }
    & {
        function Get-Module { param([switch]$ListAvailable,[string]$Name) [pscustomobject]@{Name='ShareGate';Version=[version]'99.0.0';Path=(Join-Path $root 'ShareGate.exe')} }
        function Get-ItemProperty { param([string]$Path,$ErrorAction) [pscustomobject]@{DisplayName='ShareGate Migrate';InstallLocation=$root} }
        $discovered = Get-SmartM365ShareGateModule
        if (-not $discovered -or $discovered.Path -ne (Join-Path $root 'ShareGate\Sharegate.psd1')) { throw 'GUI executable must not mask a valid module manifest.' }
    }
    & {
        function Get-Module { param([switch]$ListAvailable,[string]$Name) }
        function Get-ItemProperty { param([string]$Path,$ErrorAction) [pscustomobject]@{DisplayName='ShareGate Migrate';InstallLocation=$root} }
        function Test-ModuleManifest { param([string]$Path,$ErrorAction,$WarningAction) throw 'Invalid manifest fixture' }
        $blocked=$false
        try { Get-SmartM365ShareGateModule | Out-Null } catch { $blocked=$_.Exception.Message -match 'manifest validation failed' }
        if(-not $blocked){throw 'Invalid registered module manifest must not be imported.'}
    }
    Assert-SmartM365ShareGateDistinctSites -SourceUrl 'https://tenant.sharepoint.com/sites/a' -DestinationUrl 'https://tenant.sharepoint.com/sites/b'
    Assert-SmartM365ShareGateDistinctSites -SourceUrl 'https://farm.example.test/sites/a' -DestinationUrl 'https://tenant.sharepoint.com/sites/b'
    foreach ($pair in @(
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/a'),
        @('https://tenant.sharepoint.com/sites/a/sub', 'https://tenant.sharepoint.com/sites/a/other'),
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/%61'),
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/a%2fb'),
        @('https://tenant.sharepoint.com/', 'https://tenant.sharepoint.com/sites/b'),
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/b?x=1')
    )) {
        $blocked = $false
        try { Assert-SmartM365ShareGateDistinctSites -SourceUrl $pair[0] -DestinationUrl $pair[1] } catch { $blocked = $true }
        if (-not $blocked) { throw 'Unsafe or ambiguous boundary was accepted.' }
    }
    # ShareGate 26.11 exports cmdlets from a nested module: -Module ShareGate
    # does not find them, while the parent-qualified command name does.
    Remove-Module ShareGate -ErrorAction SilentlyContinue
    $nested = Join-Path $root 'Nested\ShareGate'
    New-Item -ItemType Directory -Path $nested -Force | Out-Null
    "@{ NestedModules=@('Sharegate.Automation.psm1'); ModuleVersion='99.0.1'; GUID='1eb75763-57f8-4141-8a2c-7ddf4e632772'; FunctionsToExport=@('Connect-Site') }" | Set-Content -LiteralPath (Join-Path $nested 'ShareGate.psd1') -Encoding UTF8
    "function Connect-Site { param([string]`$Url) `$Url }; Export-ModuleMember -Function Connect-Site" | Set-Content -LiteralPath (Join-Path $nested 'Sharegate.Automation.psm1') -Encoding UTF8
    Import-Module -Name (Join-Path $nested 'ShareGate.psd1') -Force
    $qualified = @(Get-Command -Name 'ShareGate\Connect-Site' -ErrorAction SilentlyContinue)
    if ($qualified.Count -ne 1 -or $qualified[0].Source -ne 'Sharegate.Automation') { throw 'Nested ShareGate command was not found via parent-qualified name.' }
    'PASS: independent ShareGate modes and distinct-site guards; no tenant access.'
}
finally {
    Clear-SmartM365ShareGateSessionConnections
    Remove-Module ShareGate -ErrorAction SilentlyContinue
    $resolved=[IO.Path]::GetFullPath($root)
    $allowed=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if ($resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-ShareGateAuth-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPMDQjk8C41Aaj
# mJ2TyAasL+142KbUnCjIGiW1Nm7iE6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM/rtvUJ2iN1AahuFpcbqsu9rLcXetumRa7TGu5iP+bbMA0GCSqG
# SIb3DQEBAQUABIIBgGg5tFkldsGuo9bPnd8SdOm7bzSRsC7c5Vd7IO05UZbAktHD
# 3y9QI1V4UpfnFjtoFHnhlGxuxkY1sOttO5DQxdC9fgXYzG0a8euKjMah1hlcY6X1
# TCBJre5x4p8yc3jnGRevCuXY8We1dzm/l//3vcoZdzU3cZENk4NmHj2/fXJZSLdY
# yMFmc2wh+tt40nPT4I/pnpC90TvucVFAmRP+c1XRSuVykjZAQqhsaTVp1ealzuVH
# pK0adZeTwbNw39Nwu9W2LS9hG3JJXhSDSlUA2I42TgcK0jIKNoe5rZeXjQIxnG6e
# sNMPASloMgTwhMb+YictrWdRMnqEycLX8IYnlnFAW75rbKH5MCr1ot9yd2eY0Fm4
# nhNDyjOcpgDJuP5gAzayqRLiJP0g0a8EfuRnSkVLnGwpkuBkgaz9uS5RunMMiyCb
# aN+RXXj2NfID6NcOtbIOl71Y7tywnAlFnB59tpbRtKrMe2ZihLr1v2URt6U/yS28
# pNXGl5vocTcFgsFwz6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDQyMjM4
# MzBaMC8GCSqGSIb3DQEJBDEiBCCJ5JkbeXy2Gh3cEgmavMQKLP3HSwVvpj+h3uMO
# 2Yx6fTANBgkqhkiG9w0BAQEFAASCAgAaQ5LnGJGsrVyvkfbzqFtUtNjrVtAk45DR
# tItAOGWpdaE7GY4z10klCX7ty1e4rPQJSp31XkjfJAkcvNO+lzQUaHtE1s/0i3I/
# hzpJGdSlg4jd/lEB67PLzRCGCBPoFgSwIdI/pgA1aH1cd9k4+/a9qsxIRhFWbZWp
# ZLDk6dz/Bd3um7WC3Zu1gZ5CBT+LOqbYcKhlsVjgS/+Q4YDOKpo+4hhj2LQ0P9cm
# lbpaG6VbQDsNn3JLNeBAa2EUQs9X/uswiNAbyZfjY9w3Z+y7LCf3Ahh7pOARc01Z
# BaEyQCL6IQHptCwhlBCEBNJDNm2RIX9DrtnD58eaycqrFWGo4qnuzQgMsPYeMspd
# 6UFlQAOgFPRfbxKtFk2sBxMDWdYZh97QrFe3T6jrx6MNlRQM1KELXr7ubI83qZY1
# Vd8d+IluW9YyzJCv2maYB7/mEzM4mEr31xM6ZvLvhfIMz8BPQiPn0YtW43PFu/mZ
# SBNZ8aUfMxILx0zOvPKMhW0Xlcc+SmpiaQwPfbjXF2GC1qkDZp9c6H+1TZms6+UD
# LgO0Jsh8pSWlheqfSyhIV0h2J8zysfNJ0QIJ/3jkw+BUd6f5nS7/ZQ3HsREe5Hes
# hGh7PUxiEQCrpXEIeoQ/gyK5cOj0n5xKdGy3wowWbU7DabsFEkjkHHWA5G11CMVQ
# RWToXyxXlw==
# SIG # End signature block
