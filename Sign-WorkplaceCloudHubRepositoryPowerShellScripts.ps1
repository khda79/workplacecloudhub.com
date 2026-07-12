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
    [string]$Thumbprint = 'F01F4A8871B7E349B40564D90F2B2E5BB563720B',
    [string[]]$Extensions = @('.ps1', '.psm1', '.psd1', '.ps1xml'),
    [switch]$IncludeUntracked,
    [switch]$Force,
    [switch]$SkipTrustInstall
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

if (-not $SkipTrustInstall) {
    foreach ($storeName in @('Root', 'TrustedPublisher')) {
        $installed = Install-CertificateInCurrentUserStore -Certificate $cert -StoreName $storeName
        if ($installed) { Write-Info "Installed public certificate in CurrentUser\$storeName ($normalizedThumbprint)." }
    }
}

$files = Get-GitPowerShellFiles -Root $RepoRoot -Extensions $Extensions -IncludeUntracked:$IncludeUntracked
Write-Info ("Files to inspect: {0}" -f $files.Count)

$signed = 0
$skipped = 0
$failed = 0
$failures = New-Object System.Collections.Generic.List[string]

foreach ($file in $files) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }

    try {
        $signature = Get-AuthenticodeSignature -FilePath $file -ErrorAction Stop
        $signerThumbprint = ''
        if ($signature.SignerCertificate) { $signerThumbprint = Normalize-Thumbprint $signature.SignerCertificate.Thumbprint }

        if (-not $Force -and [string]$signature.Status -eq 'Valid') {
            if ($signerThumbprint -eq $normalizedThumbprint -or (Test-MicrosoftAuthenticodeSignature -Signature $signature)) {
                $skipped++
                continue
            }
        }

        if ($PSCmdlet.ShouldProcess($file, 'Set Authenticode signature')) {
            $result = Set-AuthenticodeSignature -FilePath $file -Certificate $cert -HashAlgorithm SHA256 -ErrorAction Stop
            if ([string]$result.Status -ne 'Valid') {
                throw "Signature status is $($result.Status): $($result.StatusMessage)"
            }
            $signed++
        }
    }
    catch {
        $failed++
        $failures.Add("$file :: $($_.Exception.Message)") | Out-Null
    }
}

Write-Info ("Signed={0}; Skipped={1}; Failed={2}" -f $signed, $skipped, $failed)
if ($failed -gt 0) {
    $failures | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBD+W2gOTlLRf7
# m7sSQL5n2Iy11VOv/xdQ/53IFhf6rqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCD6DBmytVxFvXN3d97Abq0VaEwj1t8dnWMzEkd4iDo/+DANBgkqhkiG9w0B
# AQEFAASCAYCTDJmclPS7i/pbXDDK1vpkGJiKlM7Qcb/crnIOZu4cFz73x2rRLzQF
# HHLm+j65ZiFpWxB5nB++DubDTbrGZuo9KB6/VNVk/g0AJszo0FznpkBSA+l9gad/
# /2d8GgVe7sstFsgff+WelhdhjGMGD73XcRQEviAYkuv483jT1glhnIvYrMo6ZUb2
# bYQ6pmd8V3LYu66BN14CF6l+n7/Iv4qL7KaqSVtB0LXPkhnlTgC922mdkIu5HAhy
# KgnLmDvlKv1NjOY9uzXijk8Lmf1Vd5f5U8pbqfJ/OY9T09uoNY1Qi8vjqH+gGBeo
# bWNWuBuWf7/GXyqblpM+JvKdGq/Sx9NcskOdy7uW9q6oKPS7SuO4ERIy9Lx2USXu
# J4NiYd+l94b3G4LRvKqwEgS3QeNlERfmKJ+pqlPxImvuIthM31iI4VukX91ZjMrf
# niyviZ9bWDXLYbSucDiEpiXwFM0eFrS9vscQP4QRsFpRUFiYlDCg5NOEKaBrkxe1
# tBMRntysipM=
# SIG # End signature block
