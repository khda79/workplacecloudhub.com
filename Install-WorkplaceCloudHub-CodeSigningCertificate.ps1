<#
.SYNOPSIS
Installs the workplacecloudhub.com public code-signing certificate.

.DESCRIPTION
Imports the committed public .cer file into the Windows certificate stores used by
Authenticode validation. The script never installs a private key and refuses a
certificate file whose thumbprint does not match the expected workplacecloudhub.com signer.

Default target stores are CurrentUser\Root and CurrentUser\TrustedPublisher. Run
the script as the same account that runs the signed PowerShell scripts, or use
-StoreLocation LocalMachine from an elevated session to trust the certificate for
all local users.

.PARAMETER CertificatePath
Path to the public certificate .cer file.

.PARAMETER Thumbprint
Expected certificate thumbprint. The certificate is rejected if it does not match.

.PARAMETER StoreLocation
Certificate store location: CurrentUser or LocalMachine.

.PARAMETER StoreNames
Target stores. Defaults to Root and TrustedPublisher.

.PARAMETER Remove
Removes the certificate from the target stores instead of installing it.

.EXAMPLE
.\Install-WorkplaceCloudHub-CodeSigningCertificate.ps1

.EXAMPLE
.\Install-WorkplaceCloudHub-CodeSigningCertificate.ps1 -StoreLocation LocalMachine
#>
#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$CertificatePath = (Join-Path -Path $PSScriptRoot -ChildPath 'Certificates\workplacecloudhub.com-CodeSigning-F01F4A8871B7E349B40564D90F2B2E5BB563720B.cer'),

    [string]$Thumbprint = 'F01F4A8871B7E349B40564D90F2B2E5BB563720B',

    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string]$StoreLocation = 'CurrentUser',

    [ValidateSet('Root', 'TrustedPublisher')]
    [string[]]$StoreNames = @('Root', 'TrustedPublisher'),

    [switch]$Remove
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-InstallMessage {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }

    Write-Host ("[WorkplaceCloudHub-Cert][{0}] {1}" -f $Level, $Message) -ForegroundColor $color
}

function Normalize-Thumbprint {
    param([AllowNull()][string]$Value)

    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Test-CurrentProcessIsElevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Open-CertificateStore {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Location
    )

    $locationEnum = [System.Security.Cryptography.X509Certificates.StoreLocation]::$Location
    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($Name, $locationEnum)
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    return $store
}

$resolvedCertificatePath = [System.IO.Path]::GetFullPath($CertificatePath)
if (-not (Test-Path -LiteralPath $resolvedCertificatePath -PathType Leaf)) {
    throw "Certificate file not found: $resolvedCertificatePath"
}

if ($StoreLocation -eq 'LocalMachine' -and -not (Test-CurrentProcessIsElevated)) {
    throw 'Administrator rights are required to write to LocalMachine certificate stores. Re-run from an elevated PowerShell session or use -StoreLocation CurrentUser.'
}

$expectedThumbprint = Normalize-Thumbprint $Thumbprint
$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($resolvedCertificatePath)
$certificateThumbprint = Normalize-Thumbprint $certificate.Thumbprint

if (-not [string]::IsNullOrWhiteSpace($expectedThumbprint) -and $certificateThumbprint -ne $expectedThumbprint) {
    throw "Certificate thumbprint mismatch. Expected $expectedThumbprint, got $certificateThumbprint from $resolvedCertificatePath"
}

if ($certificate.HasPrivateKey) {
    throw "The certificate file contains a private key and will not be imported by this public trust installer: $resolvedCertificatePath"
}

Write-InstallMessage ("Certificate: subject={0}; thumbprint={1}; source={2}" -f $certificate.Subject, $certificateThumbprint, $resolvedCertificatePath)

$results = New-Object System.Collections.Generic.List[object]
foreach ($storeName in $StoreNames) {
    $store = $null
    try {
        $store = Open-CertificateStore -Name $storeName -Location $StoreLocation
        $existing = @($store.Certificates | Where-Object { (Normalize-Thumbprint $_.Thumbprint) -eq $certificateThumbprint })
        $storePath = "{0}\{1}" -f $StoreLocation, $storeName

        if ($Remove) {
            if ($existing.Count -eq 0) {
                Write-InstallMessage ("Certificate is not present in {0}." -f $storePath) -Level WARN
                $results.Add([pscustomobject]@{ Store = $storePath; Action = 'NotPresent'; Thumbprint = $certificateThumbprint }) | Out-Null
                continue
            }

            foreach ($existingCertificate in $existing) {
                if ($PSCmdlet.ShouldProcess($storePath, "Remove certificate $certificateThumbprint")) {
                    $store.Remove($existingCertificate)
                    Write-InstallMessage ("Removed certificate from {0}." -f $storePath) -Level SUCCESS
                    $results.Add([pscustomobject]@{ Store = $storePath; Action = 'Removed'; Thumbprint = $certificateThumbprint }) | Out-Null
                }
            }
            continue
        }

        if ($existing.Count -gt 0) {
            Write-InstallMessage ("Certificate already present in {0}." -f $storePath) -Level SUCCESS
            $results.Add([pscustomobject]@{ Store = $storePath; Action = 'AlreadyPresent'; Thumbprint = $certificateThumbprint }) | Out-Null
            continue
        }

        if ($PSCmdlet.ShouldProcess($storePath, "Install certificate $certificateThumbprint")) {
            $publicCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certificate.RawData)
            $store.Add($publicCertificate)
            Write-InstallMessage ("Installed certificate in {0}." -f $storePath) -Level SUCCESS
            $results.Add([pscustomobject]@{ Store = $storePath; Action = 'Installed'; Thumbprint = $certificateThumbprint }) | Out-Null
        }
    }
    finally {
        if ($store) { $store.Close() }
    }
}

$results.ToArray()

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAp+h46y9cSmWAb
# 6QBYmCQ0hwyqblXxzcuSpYXSHUhfpqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCA5J/SrGbMX2yg+gyPepu42pwJBRSMTKX849TO+aKa7yTANBgkqhkiG9w0B
# AQEFAASCAYCt87c8MfY/DLwyGxfZczGrYe22fZx1qI4NVHCKIHzCVK3c2A6as2ef
# 4FabbGcgx25JJf+zGUGvxyOfu8Dh9+B++hzHzGIYSR3DWSwpb+JZaMSEbeE8aYJj
# YAl3RaLLCfT3C/U6x1wAx8kU4MbKeviup/tVb3I/49US3eAHPjEtVuNtVPnfMl9P
# y+WPaRmg3Qef+ImCcCZQ2RzkvnF8vl9oXHbZhEYPJ7yLTO3aUBi3qwijX/UezqC5
# Lkz7LYp8eq9byDaDYU3KzgGwJce5WI/BbmgdJTtSjSIE9ExipvKpNosUI/TnMFkH
# hqjRur11tBD7Ot61E41Oxznp/tH7r0Lpb+d+K6pwdTCPG+E2IK5G8KzJUQ6SDHeA
# 7HCkz8wIGQ5ATWXfKmQ/d1PM77JF32SBwA5jXc0wxqGfzTVgs9IPTcPLWboBzFjr
# 4+4PCo9XUqPpsWG5P8KpLdVgi6cho+N5dpixoLy5Ws67b067ZuDeFlTGfH8HNIru
# GLPxIa9e7CU=
# SIG # End signature block
