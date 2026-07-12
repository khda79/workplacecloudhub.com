<#
.SYNOPSIS
Installs the SmartM365 Orchestrator public code-signing certificate.

.DESCRIPTION
Imports the committed public .cer file into the Windows certificate stores used by
Authenticode validation. The script never installs a private key and refuses a
certificate file whose thumbprint does not match the expected SmartM365 signer.

Default target stores are CurrentUser\Root and CurrentUser\TrustedPublisher. Run
the script as the same account that runs the orchestrator scheduled task, or use
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
.\Install-SmartM365-OrchestratorCodeSigningCertificate.ps1

.EXAMPLE
.\Install-SmartM365-OrchestratorCodeSigningCertificate.ps1 -StoreLocation LocalMachine
#>
#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$CertificatePath = (Join-Path -Path $PSScriptRoot -ChildPath 'Certificates\SmartM365-Orchestrator-CodeSigning-526459860BED5BA91ED005483C90182852F26FE0.cer'),

    [string]$Thumbprint = '526459860BED5BA91ED005483C90182852F26FE0',

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

    Write-Host ("[SmartM365-Cert][{0}] {1}" -f $Level, $Message) -ForegroundColor $color
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBvRohd1qq5pgbl
# 4PmJxkLroDpr+6LncvojUctD8dNahaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAvdTTSLDk494wMjfhn
# boudP84Kf2a9OPGWdL6cB2QIsDANBgkqhkiG9w0BAQEFAASCAYCPO2hIR6pCKRKo
# Xqm1AQztHFvuZmqMFGCsFMMQiOxZ+VNytCeV44dfC9VO+z3sMbOdM28fIjq7D+9c
# 1r4o7WENMjlNyV49CxsO+NSWk2w6w7xOA0cVVVSD4gGw1RhFxkYgGPyC+p4Rxx8p
# 2x6sX/LhlATLu4KYhtE6ziaFqwj8L+pTzsossLFn/EJ7RD/X4Ut+H1yuWVW9YGxd
# BZ1JHHudfvQQQscUmmQZtCBm05NhQ8QbxO+wEF0c6qTATWIgUQmSdAQ8Mur4kWCH
# 4WAYGA3ow41pmpQk/siHp/UEK067STgVTAkrUsmM68UqiVhSdKXPO0aJRKf6SsZB
# t6pFQrmXjqFEGG2LTy/NnZ1eIY0BARM6vQbY02a+dGADg/mQUyPMMZi4EdT0lxBF
# Be3uWOsgvrPGe0RrRQG/iLFI3y2bXfdrwWRzEf2inHF+Jsb1Qte5NJu2phz4BFI8
# f6BiaaGxcik4WYtafEP6pbTw+oAB3WO0Do6fvwGlLgZVsO9bsiI=
# SIG # End signature block
