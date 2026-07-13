<#
.SYNOPSIS
Signs repository PowerShell files with the workplacecloudhub.com Authenticode code-signing certificate.

.DESCRIPTION
Signs PowerShell scripts/modules/manifests tracked by Git, with an option to include
non-ignored untracked files. Intended for manual use and for the repository pre-push hook.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$RepoRoot = $PSScriptRoot,
    [string]$Thumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [string]$ExpectedSignerEmail = 'contact@workplacecloudhub.com',
    [string]$TimestampServer = 'http://timestamp.digicert.com',
    [string[]]$Extensions = @('.ps1', '.psm1', '.psd1', '.ps1xml'),
    [switch]$IncludeUntracked,
    [switch]$Force,
    [switch]$SkipTrustInstall,
    [switch]$SkipLineEndingNormalization
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    Write-Host "[Sign-WorkplaceCloudHub] $Message"
}

function Normalize-Thumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Test-MicrosoftAuthenticodeSignature {
    param([AllowNull()]$Signature)

    if (-not $Signature -or [string]$Signature.Status -ne 'Valid' -or -not $Signature.SignerCertificate) {
        return $false
    }

    $subject = [string]$Signature.SignerCertificate.Subject
    $issuer = [string]$Signature.SignerCertificate.Issuer
    return (($subject -match '(^|,\s*)O=Microsoft Corporation(,|$)' -or $subject -match '(^|,\s*)CN=Microsoft Corporation(,|$)') -and $issuer -match 'Microsoft')
}
function Install-CertificateInCurrentUserStore {
    param(
        [Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $true)][string]$StoreName
    )

    $thumbprint = Normalize-Thumbprint $Certificate.Thumbprint
    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($StoreName, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $existing = @($store.Certificates | Where-Object { (Normalize-Thumbprint $_.Thumbprint) -eq $thumbprint })
        if ($existing.Count -gt 0) {
            return $false
        }

        $publicCert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Certificate.RawData)
        $store.Add($publicCert)
        return $true
    }
    finally {
        if ($store) { $store.Close() }
    }
}

function Get-GitPowerShellFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string[]]$Extensions,
        [switch]$IncludeUntracked
    )

    $patterns = @($Extensions | ForEach-Object { '*{0}' -f $_ })
    $args = @('-C', $Root, 'ls-files')
    if ($IncludeUntracked) { $args += @('--cached', '--others', '--exclude-standard') }
    $args += '--'
    $args += $patterns

    $relativePaths = & git @args
    if ($LASTEXITCODE -ne 0) { throw 'git ls-files failed.' }

    return @($relativePaths | Where-Object { $_ } | Sort-Object -Unique | ForEach-Object {
        [System.IO.Path]::GetFullPath((Join-Path -Path $Root -ChildPath $_))
    })
}

function ConvertTo-WorkplaceCloudHubCrlf {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $preamble = [byte[]]@()
    $encoding = $null
    $offset = 0

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encoding = [System.Text.UTF8Encoding]::new($false, $true)
        $preamble = [byte[]]@(0xEF, 0xBB, 0xBF)
        $offset = 3
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encoding = [System.Text.UnicodeEncoding]::new($false, $false, $true)
        $preamble = [byte[]]@(0xFF, 0xFE)
        $offset = 2
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encoding = [System.Text.UnicodeEncoding]::new($true, $false, $true)
        $preamble = [byte[]]@(0xFE, 0xFF)
        $offset = 2
    }
    else {
        $encoding = [System.Text.UTF8Encoding]::new($false, $true)
    }

    $text = $encoding.GetString($bytes, $offset, $bytes.Length - $offset)
    $crlf = [string][char]13 + [char]10
    $lineEndingPattern = [regex]::Escape([string][char]13 + [char]10) + '|' + [regex]::Escape([string][char]13) + '|' + [regex]::Escape([string][char]10)
    $normalizedText = [regex]::Replace($text, $lineEndingPattern, $crlf)
    if ($normalizedText -ceq $text) { return $false }

    $contentBytes = $encoding.GetBytes($normalizedText)
    $normalizedBytes = if ($preamble.Length -gt 0) { [byte[]]($preamble + $contentBytes) } else { $contentBytes }
    if ($PSCmdlet.ShouldProcess($Path, 'Normalize line endings to CRLF before Authenticode signing')) {
        [System.IO.File]::WriteAllBytes($Path, $normalizedBytes)
    }
    return $true
}
$RepoRoot = [System.IO.Path]::GetFullPath($RepoRoot)
if (-not (Test-Path -LiteralPath (Join-Path -Path $RepoRoot -ChildPath '.git'))) {
    throw "RepoRoot is not a Git repository root: $RepoRoot"
}

$normalizedThumbprint = Normalize-Thumbprint $Thumbprint
$cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { (Normalize-Thumbprint $_.Thumbprint) -eq $normalizedThumbprint } |
    Select-Object -First 1

if (-not $cert) { throw "Code-signing certificate not found in Cert:\CurrentUser\My: $normalizedThumbprint" }
if (-not $cert.HasPrivateKey) { throw "Code-signing certificate has no private key: $normalizedThumbprint" }
$certificateEmail = [string]$cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::EmailName, $false)
if (-not [string]::IsNullOrWhiteSpace($ExpectedSignerEmail) -and $certificateEmail -ine $ExpectedSignerEmail) {
    throw "Code-signing certificate email mismatch. Expected '$ExpectedSignerEmail', got '$certificateEmail'."
}

if (-not $SkipTrustInstall) {
    foreach ($storeName in @('Root', 'TrustedPublisher')) {
        $installed = Install-CertificateInCurrentUserStore -Certificate $cert -StoreName $storeName
        if ($installed) { Write-Info "Installed public certificate in CurrentUser\$storeName ($normalizedThumbprint)." }
    }
}

$files = @(Get-GitPowerShellFiles -Root $RepoRoot -Extensions $Extensions -IncludeUntracked:$IncludeUntracked)
Write-Info ("Files to inspect: {0}" -f $files.Count)

$signed = 0
$skipped = 0
$normalized = 0
$failed = 0
$failures = New-Object System.Collections.Generic.List[string]

foreach ($file in $files) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }

    try {
        $signature = Get-AuthenticodeSignature -FilePath $file -ErrorAction Stop
        $signerThumbprint = ''
        if ($signature.SignerCertificate) { $signerThumbprint = Normalize-Thumbprint $signature.SignerCertificate.Thumbprint }

        if (-not $Force -and [string]$signature.Status -eq 'Valid' -and (Test-MicrosoftAuthenticodeSignature -Signature $signature)) {
            $skipped++
            continue
        }

        $lineEndingsChanged = $false
        if (-not $SkipLineEndingNormalization) {
            $lineEndingsChanged = ConvertTo-WorkplaceCloudHubCrlf -Path $file
            if ($lineEndingsChanged) { $normalized++ }
        }

        if (-not $Force -and -not $lineEndingsChanged -and [string]$signature.Status -eq 'Valid' -and $signerThumbprint -eq $normalizedThumbprint) {
            $skipped++
            continue
        }

        if ($PSCmdlet.ShouldProcess($file, 'Set Authenticode signature')) {
            $signingParams = @{
                FilePath = $file
                Certificate = $cert
                HashAlgorithm = 'SHA256'
                ErrorAction = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($TimestampServer)) {
                $signingParams['TimestampServer'] = $TimestampServer
            }
            $result = Set-AuthenticodeSignature @signingParams
            if ([string]$result.Status -ne 'Valid') {
                throw "Signature status is $($result.Status): $($result.StatusMessage)"
            }
            $verifiedSignature = Get-AuthenticodeSignature -FilePath $file -ErrorAction Stop
            if ([string]$verifiedSignature.Status -ne 'Valid') {
                throw "Signature verification status is $($verifiedSignature.Status): $($verifiedSignature.StatusMessage)"
            }
            if (-not [string]::IsNullOrWhiteSpace($TimestampServer) -and -not $verifiedSignature.TimeStamperCertificate) {
                throw "The signature is valid but no timestamp countersignature was returned by $TimestampServer."
            }
            $signed++
        }
    }
    catch {
        $failed++
        $failures.Add("$file :: $($_.Exception.Message)") | Out-Null
    }
}

Write-Info ("Normalized={0}; Signed={1}; Skipped={2}; Failed={3}" -f $normalized, $signed, $skipped, $failed)
if ($failed -gt 0) {
    $failures | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7N7WFmabYwjFE
# lvk66tFh3Cn+7fiwIpcKCJAoh4eqOaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIEtYC411a+kVTMcsmU3MjjS8GhuqPxtCP107Fq7lJAzaMA0GCSqG
# SIb3DQEBAQUABIIBgC0k1AAoqeSc83i00Z6msURvKqospxI47VbrtMO1HM1mvUYu
# e1NUkrMvQVprABh+igfE9R3gawVoFPPXxUKHjEv2ax/ZslZBVwTg0hKFB8yO2Dlv
# uxpGpTOHGkgrO27fqqFy6kSd4k5KvXIOA7HMjEQEsW7OzdslHX5MGi97MwFqdqCE
# AyN7ITdgg0SXfetA1CBqZ4G61WEaa6O2m11mlAevfoGoGQfbSEtL20QL3Jw2YDqG
# Xol5tU3ynmkqKkd+GBFi2BqOwm2jsGFWx37EVRrDBk3PY+zep+AOLULEelWcWeEl
# ypyO5fKgNKlDgif9zYu2StNT53EoCUxTvq2BKRjoQ0xhW5tpFGnFKekeIa0nbQ83
# K+s9NigqGVA7y1LJrCXNqFYvkxJu2hMJbkVdr+R+xtzX08VnSd0oN7UojIosuZ4Z
# O4ryYsNhCFfAWviPkoBhyYcwfrkB6JSdNjV48khMAm6dsBmldWCasEQOi+1sScQm
# wyyAaqalekBettXsvqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MDZaMC8GCSqGSIb3DQEJBDEiBCBX+79USd9AKXJwPWmFlWO+nQq+Fsm70QgUXpsx
# rZNPijANBgkqhkiG9w0BAQEFAASCAgAl4M1UgO69jA+JF5SIXESXAOdwfK6iZ+dj
# QFsapGHnVR21m7K62kZSaeqZzYrow+WDeu2h0+gaUW0sSG6NI9fE/EHhRjJUIjAY
# UaFQLpAWmJgEVykUyK8dALc8CvrliJg+oh7TARf0kg68ahfXhcoyjrUCkIjccbJe
# +jvXIs+aVkY+1VdySiEHzOXq7CpZMRSg9s3AI6eIpn9sk+RBT1yQRSVDALB5SNib
# cVuc3H2Rn3DyLTm5/wlncLqKamu4YeyojQSn82krFDClih6TJOOT/RO6b/Osin7U
# hL+eAeoGuQD+eHtXv26JS+jcvmg7SxnODq4bgO1jmZDwtnfZuznkgxWLWaWBpqPA
# Jvn4FW6CJ60gPL7amDgJ30ZyeK7+QBaIOB9vRCs3vzTd66Q466dVqJvd18Yiwnzb
# CSWtsAatpzLvaSiJ/DchmwuwmPEXbdSyoYX1kEcGbF3UNOvdJ0PQ7upiPN8n4fFr
# gKfetWbAoM49UKShWpIgRmUaGGBUxgnyyNESSZ9eJlpO5VZUGON32gcNcPtRlS0N
# p7cgDdIxKabA40GoN7N6TC9IlC8UgbtPRDHuU3qLbmTFt9lUyRjxmxNLtfKPWIpN
# 3mAmFKnAgDJfMcWvcWZKvlsHIzwPEek9O5G94krURyk05LKLly4BfmSZDkl5l53V
# ArwqF7wwcA==
# SIG # End signature block
