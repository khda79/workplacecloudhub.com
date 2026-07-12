<#
.SYNOPSIS
Signs repository PowerShell files with the SmartM365 Authenticode code-signing certificate.

.DESCRIPTION
Signs PowerShell scripts/modules/manifests that are tracked by Git, with an option to include
non-ignored untracked files. Intended for manual use and for the repository pre-push hook.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$RepoRoot = (Resolve-Path -LiteralPath (Join-Path -Path $PSScriptRoot -ChildPath '..')).Path,
    [string]$Thumbprint = '526459860BED5BA91ED005483C90182852F26FE0',
    [string[]]$Extensions = @('.ps1', '.psm1', '.psd1', '.ps1xml'),
    [switch]$IncludeUntracked,
    [switch]$Force,
    [switch]$SkipTrustInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    Write-Host "[Sign-SmartM365] $Message"
}

function Normalize-Thumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
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

        if (-not $Force -and [string]$signature.Status -eq 'Valid' -and $signerThumbprint -eq $normalizedThumbprint) {
            $skipped++
            continue
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAaOnCu7t4OOLS/
# AqmnoM77Q9tFhf7VXUr6hSbqvfBjqKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBQx7c604KloziKu7Gl
# dLatNSQw+hzhhk7JSyibtI7pyzANBgkqhkiG9w0BAQEFAASCAYCFvXae2VZTg/hI
# WpJJG6i57vJVeWu2LKThPYVGHitNT2RdT7ZfSq2L2uRDtmgErboRM7CO08dWOLQ1
# 1gIyQCztw3Q0EycoTXRMnVtlKZjror3Plw9SkQjGZRLbtBjoA0cHUsxjGK6aHFnr
# lFAyJkFhNgFAqQDX9pK1J9O0JB4VHqWuQV8lM1Q8PKE9NZ63pyJOXOpKgXUUwveP
# X2FsTB8dczlg7J9i7GNqM7aoZ+MteEiGytgGbT82He1PMu/p18kU/G3lFv0D/I+X
# Qs2Z2aRAZvKbDBTEXFH3Xgs5WBQaa6xZAanmgPvNuRsx71QL73wHieYPKs7xW03K
# 2D9/7NnGVIMTEGMWddHXaULFxQ8UHfoTFr326AEcOpF4WD1+BUdh/fky8bjGXP1B
# k/nkVgW/jrRCQ02KywKBvRwagBB2eRHIAwX8vdlzp2q8vvEZSUgW9O/ziMksN6ak
# B7UbUANuRH+/GIj8TOKsNPXv2/R3fqqYGPGQfFDBkFwZ/VJaskU=
# SIG # End signature block
