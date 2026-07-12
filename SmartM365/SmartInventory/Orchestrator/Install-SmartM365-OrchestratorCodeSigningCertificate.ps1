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