<#
.SYNOPSIS
    Connect one ShareGate endpoint with an explicitly selected authentication mode.
.VERSION
    1.0.4
#>
function Get-SmartM365ShareGateModule {
    # PowerShell can discover the installed GUI executable as a version 0 module.
    # It is not an importable PowerShell module and must not mask a valid manifest.
    $available = @(Get-Module -ListAvailable -Name ShareGate | Where-Object {
        $_.Path -and [IO.Path]::GetExtension([string]$_.Path) -in @('.psd1','.psm1','.dll')
    } | Sort-Object Version -Descending | Select-Object -First 1)
    if ($available.Count) { return $available[0] }
    foreach ($registryRoot in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
        $apps = @(Get-ItemProperty -Path (Join-Path $registryRoot '*') -ErrorAction SilentlyContinue | Where-Object {
            $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -match '^ShareGate( Migrate)?$' -and
            $_.PSObject.Properties['InstallLocation'] -and $_.InstallLocation
        })
        foreach ($app in $apps) {
            $candidate = Join-Path ([string]$app.InstallLocation) 'ShareGate\Sharegate.psd1'
            if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
            try {
                $manifest = Test-ModuleManifest -Path $candidate -ErrorAction Stop -WarningAction SilentlyContinue
                return [pscustomobject]@{ Version=$manifest.Version; Path=$candidate }
            } catch { throw "ShareGate module manifest validation failed: $($_.Exception.Message)" }
        }
    }
    $programFiles = [Environment]::GetFolderPath('ProgramFilesX86')
    if ($programFiles) {
        $candidate = Join-Path $programFiles 'Sharegate\ShareGate\Sharegate.psd1'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            try {
                $manifest = Test-ModuleManifest -Path $candidate -ErrorAction Stop -WarningAction SilentlyContinue
                return [pscustomobject]@{ Version=$manifest.Version; Path=$candidate }
            } catch { throw "ShareGate module manifest validation failed: $($_.Exception.Message)" }
        }
    }
}

function Get-SmartM365ShareGateSessionKey {
    param([Parameter(Mandatory = $true)][string]$Url, [Parameter(Mandatory = $true)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode)
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -notin @('http','https') -or $uri.UserInfo -or $uri.Query -or $uri.Fragment) {
        throw 'A session connection requires an absolute HTTP(S) site URL without credentials, query or fragment.'
    }
    return ('{0}|{1}' -f $Mode.ToLowerInvariant(), $uri.AbsoluteUri.TrimEnd('/').ToLowerInvariant())
}

function Register-SmartM365ShareGateSessionConnection {
    param([Parameter(Mandatory = $true)][string]$Url, [Parameter(Mandatory = $true)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode, [Parameter(Mandatory = $true)]$Site)
    $key = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    $address = $Site.PSObject.Properties['Address']
    if (-not $address -or -not $address.Value -or (Get-SmartM365ShareGateSessionKey -Url ([string]$address.Value) -Mode $Mode) -ne $key) {
        throw 'Cannot register a ShareGate connection whose site address differs from the requested URL.'
    }
    if (-not (Get-Variable -Name SmartM365ShareGateSessionConnections -Scope Global -ErrorAction SilentlyContinue)) {
        $global:SmartM365ShareGateSessionConnections = @{}
    }
    if ($global:SmartM365ShareGateSessionConnections.ContainsKey($key) -and
        -not [object]::ReferenceEquals($global:SmartM365ShareGateSessionConnections[$key], $Site)) {
        throw 'A different ShareGate connection is already registered for this URL and mode.'
    }
    $global:SmartM365ShareGateSessionConnections[$key] = $Site
}

function Clear-SmartM365ShareGateSessionConnections {
    Remove-Variable -Name SmartM365ShareGateSessionConnections -Scope Global -ErrorAction SilentlyContinue
}

function Connect-SmartM365ShareGateEndpointWithExistingCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode,
        [Parameter(Mandatory)]$CredentialSite,
        [switch]$AllowDifferentHost
    )
    $null = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    $address = $CredentialSite.PSObject.Properties['Address']
    if (-not $address -or -not $address.Value) { throw 'The credential source has no site address.' }
    $null = Get-SmartM365ShareGateSessionKey -Url ([string]$address.Value) -Mode $Mode
    $targetUri = [uri]$Url; $sourceUri = [uri]$address.Value
    if ($targetUri.Scheme -ne $sourceUri.Scheme -or $targetUri.Port -ne $sourceUri.Port) {
        throw 'Credential reuse requires the same protocol and port.'
    }
    if ($targetUri.DnsSafeHost -ine $sourceUri.DnsSafeHost -and -not $AllowDifferentHost) {
        throw 'Reusing credentials across different hosts requires explicit -AllowDifferentHost.'
    }
    $site = ShareGate\Connect-Site -Url $Url -UseCredentialsFrom $CredentialSite -ErrorAction Stop
    if (-not $site) { throw 'ShareGate returned no site for the reused credentials.' }
    Register-SmartM365ShareGateSessionConnection -Url $Url -Mode $Mode -Site $site
    return $site
}

function Connect-SmartM365ShareGateEndpointWithSavedConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode,
        [Parameter(Mandatory)]$SavedConnection
    )
    $null = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    $site = ShareGate\Connect-Site -Url $Url -UseConnection $SavedConnection -ErrorAction Stop
    if (-not $site) { throw 'ShareGate returned no site for the saved connection.' }
    Register-SmartM365ShareGateSessionConnection -Url $Url -Mode $Mode -Site $site
    return $site
}

function Connect-SmartM365ShareGateEndpoint {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode
    )
    $key = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    if ((Get-Variable -Name SmartM365ShareGateSessionConnections -Scope Global -ErrorAction SilentlyContinue) -and
        $global:SmartM365ShareGateSessionConnections.ContainsKey($key)) {
        $site = $global:SmartM365ShareGateSessionConnections[$key]
        $address = $site.PSObject.Properties['Address']
        if (-not $address -or -not $address.Value -or (Get-SmartM365ShareGateSessionKey -Url ([string]$address.Value) -Mode $Mode) -ne $key) {
            throw 'Registered ShareGate session connection no longer matches the requested URL.'
        }
        return $site
    }
    $arguments = @{ Url=$Url; ErrorAction='Stop' }
    if ($Mode -eq 'Browser') { $arguments.Browser = $true }
    if ($Mode -eq 'ModernAuth') { $arguments.ModernAuth = $true }
    $site = ShareGate\Connect-Site @arguments
    if (-not $site) { throw "ShareGate Connect-Site returned no site for $Url." }
    return $site
}

function Assert-SmartM365ShareGateDistinctSites {
    param([Parameter(Mandatory = $true)][string]$SourceUrl, [Parameter(Mandatory = $true)][string]$DestinationUrl)
    $source = $null; $destination = $null
    if (-not [uri]::TryCreate($SourceUrl, [UriKind]::Absolute, [ref]$source) -or
        -not [uri]::TryCreate($DestinationUrl, [UriKind]::Absolute, [ref]$destination) -or
        $source.UserInfo -or $source.Query -or $source.Fragment -or
        $destination.UserInfo -or $destination.Query -or $destination.Fragment -or
        $source.Scheme -notin @('http','https') -or
        $destination.Scheme -ne 'https' -or $destination.Port -ne 443 -or $destination.Host -notmatch '(?i)\.sharepoint\.(com|cn|de)$') {
        throw 'Source and destination must be valid site URLs and the destination must be HTTPS SharePoint Online.'
    }
    if ($source.Host -ine $destination.Host) { return }
    if ($source.Scheme -ne 'https' -or $source.Port -ne $destination.Port) {
        throw 'Same-host SharePoint sites must both use the same HTTPS endpoint.'
    }
    $collections = foreach ($site in @($source, $destination)) {
        $segments = @($site.AbsolutePath.Trim('/').Split('/') | Where-Object { $_ })
        if ($segments.Count -lt 2 -or $segments[0] -notin @('sites','teams')) {
            throw 'Same-host copy requires two explicit /sites/ or /teams/ site collections.'
        }
        $name = [uri]::UnescapeDataString($segments[1])
        if ([string]::IsNullOrWhiteSpace($name) -or $name -in @('.','..') -or $name.Contains('%') -or $name -match '[/\\\x00-\x1f]') {
            throw 'Invalid site collection segment in same-host copy.'
        }
        '{0}/{1}' -f $segments[0].ToLowerInvariant(), $name.ToLowerInvariant()
    }
    if ($collections[0] -eq $collections[1]) { throw 'Source and destination must be distinct site collections on the same host.' }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOc2AU96GzAlDE
# 6we9km1NZqX+FBRtOj2cVYbiFj+KV6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILmntCUrM/hbO5d4KNrAMHbwkuj7YQqg69EBkNV0+M1fMA0GCSqG
# SIb3DQEBAQUABIIBgAEr6Y4mhHg4oRAAV4acBgbKEgTfm6EwF7TQTFE5B6GWOHzk
# ct5JgWY0aw/7edQBH4SsZF0R5BMw7rgmacOjWXlc9Ih/1Ul7QTBDvV2yhhkc67y5
# V2hfafnXFlNodHvx2b1NKx18G0Cu5EB2baBBiwY7AbMO+agJRvfz31CZ80uXBPrp
# EgWhI6lkVTBTyFWf9blFM5UKSKkPqlb5Db+OFwx9gzOz+pNLXC1zD/QAS/flxJj0
# PpL+R+sOMjK7Yi37+tNgwnPdjzvuBDXWLWqasLgtZoMAHqwqoOfaWboc4GT4JVXR
# bgz2bibGpERGUVYxcp/0K+XYhm0DJIWCHES9s8fDgcIbbYoOvYfWzakMu2DTtDP4
# vCHGQF3wON2bO8wVFHxwta267WNbXnGviQYkIMCZV5/+C6GTfD5Xl4LTH6yMlm9q
# nXYFSnMDgrUAqiEpDEq5dfoyKOM3vPHZ2jpG9hWF9spRiWk/eAyubCr7dMnZZFmb
# UF4qLiPtToDR/UiPS6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDQyMjM4
# MjRaMC8GCSqGSIb3DQEJBDEiBCBsmQEhAQDRzw8a4/7qLiHKbmk9YhNek+xeJ6m0
# Vv8G/zANBgkqhkiG9w0BAQEFAASCAgA7O3iNJMrbyN4ecPcAZ433dx8ecs2EzPRC
# pT1HYYHpfyNPvKA8tYnKM0UAHvowZvNn4qMmhSwxc480ynZH+t8neeOzkgx6iyI0
# FeYU3I5aJ61X7wHyfkUg57toaV3Tkk2dsufdcCB0SYvGc9MhGkrGispE7pLcgCJP
# 7/bUkQqYqeYhMTgFIZVAx0/ci63OJ5VuuEN0Ottg025hhJvFYAGHj9nhNkO5HnJt
# KDOIa35funJu+XzccWc1N99gPXXR0pkgLU7xIPl6pXGQvyEfhmZsNcP3TEy081/b
# 3ZZzvOZdnWuZMaIQhBx07lfx+QT1ZYJ86O7W/Y08k1r9cr07O5uFmtfvpCKmerGM
# TJRgZI4HHGDlBLoy3zkRNwVyjxw6AOOpKVxv3geW0mZxNjNDrtx5B3gDnHe4RylN
# w5/3aUurhljvBvVZSYZ5eykQ99v5F5EOWe8UjJtjeA6WbGjZREJ+Cjv6K1dQ8Ugy
# dN42+Fx+udb5eKtB1yIcmvdC/NgrWrmM+SZYB5wR0z/NtXsB9EMWFmJOAbMX7tzX
# 30p3JvKOQ/qHAcpewQmMCFM/bGmRMXzeIhCAryQpS9TYgVGmb6gYAJboVLVb+7DG
# kJrkhPEzorz6GONZbnP5drM+cwlxoeuk5Jd08nWXWL3OUaJB0GknpKQRz7J8T/Eu
# jdG+ouXXZA==
# SIG # End signature block
