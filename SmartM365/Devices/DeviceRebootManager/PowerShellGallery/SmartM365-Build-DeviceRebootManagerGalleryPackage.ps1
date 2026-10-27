#Requires -Version 5.1

<#
.SYNOPSIS
Builds or validates the local PowerShell Gallery package for SmartM365 Device Reboot Manager.

.DESCRIPTION
Creates a clean module folder from an explicit allow-list, writes a SHA-256 file
catalog (SmartM365.DeviceRebootManager.Files.cat) over every packaged file, signs
the catalog with the pinned WorkplaceCloudHub code-signing certificate, then
validates PowerShell syntax, manifest/version consistency, the catalog and the
pinned Authenticode signer. It never copies local runtime configuration or
LOCAL_MEMORY.md.

With -ValidatePackagePath, nothing is built: the same validation is run against
an existing package folder. The publish script uses this mode for -PackagePath.

Signature rule (WinVerifyTrust, WINTRUST_ACTION_GENERIC_VERIFY_V2, no revocation
check): a signature is accepted only when the result is 0, CERT_E_UNTRUSTEDROOT
or CERT_E_CHAINING (self-signed root not trusted locally, logged as a warning)
AND the signer thumbprint equals the pinned thumbprint. Any other result, such
as TRUST_E_BAD_DIGEST, TRUST_E_NOSIGNATURE or an expired untimestamped
certificate, is rejected.
.VERSION
1.1.0
#>

[CmdletBinding(DefaultParameterSetName = 'Build')]
param(
    [string]$SourceRoot = (Split-Path -Path $PSScriptRoot -Parent),
    [Parameter(ParameterSetName = 'Build')]
    [string]$OutputRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\PowerShellGallery'),
    [Parameter(Mandatory = $true, ParameterSetName = 'Validate')]
    [string]$ValidatePackagePath,
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [Parameter(ParameterSetName = 'Build')]
    [string]$TimestampServer = 'http://timestamp.digicert.com',
    [Parameter(ParameterSetName = 'Build')]
    [switch]$Force,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ExpectedModuleName = 'SmartM365.DeviceRebootManager'
$script:PackageCatalogFileName = 'SmartM365.DeviceRebootManager.Files.cat'

function Get-NormalizedThumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Assert-SafeChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Child
    )

    $resolvedRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    $resolvedChild = [IO.Path]::GetFullPath($Child).TrimEnd('\', '/')
    $prefix = $resolvedRoot + [IO.Path]::DirectorySeparatorChar
    if ($resolvedChild -eq $resolvedRoot -or -not $resolvedChild.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe build target outside OutputRoot: $resolvedChild"
    }
}

function Assert-PowerShellSyntax {
    param([Parameter(Mandatory = $true)][string]$Path)

    $parseErrors = $null
    [void][Management.Automation.PSParser]::Tokenize(
        (Get-Content -LiteralPath $Path -Raw),
        [ref]$parseErrors
    )
    if ($parseErrors.Count -gt 0) {
        $messages = $parseErrors | ForEach-Object { $_.Message }
        throw ("PowerShell syntax validation failed for {0}: {1}" -f $Path,($messages -join '; '))
    }
}

function Get-WinVerifyTrustResult {
    # Direct WinVerifyTrust (WINTRUST_ACTION_GENERIC_VERIFY_V2, no UI, no
    # revocation check, file choice, VERIFY then CLOSE). Returns the HRESULT.
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not ('SmartM365.DeviceRebootManagerWinTrust' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace SmartM365
{
    public static class DeviceRebootManagerWinTrust
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WINTRUST_FILE_INFO
        {
            public uint cbStruct;
            public string pcwszFilePath;
            public IntPtr hFile;
            public IntPtr pgKnownSubject;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WINTRUST_DATA
        {
            public uint cbStruct;
            public IntPtr pPolicyCallbackData;
            public IntPtr pSIPClientData;
            public uint dwUIChoice;
            public uint fdwRevocationChecks;
            public uint dwUnionChoice;
            public IntPtr pFile;
            public uint dwStateAction;
            public IntPtr hWVTStateData;
            public IntPtr pwszURLReference;
            public uint dwProvFlags;
            public uint dwUIContext;
            public IntPtr pSignatureSettings;
        }

        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int WinVerifyTrust(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, IntPtr pWVTData);

        private const uint WTD_UI_NONE = 2;
        private const uint WTD_REVOKE_NONE = 0;
        private const uint WTD_CHOICE_FILE = 1;
        private const uint WTD_STATEACTION_VERIFY = 1;
        private const uint WTD_STATEACTION_CLOSE = 2;
        private const uint WTD_REVOCATION_CHECK_NONE = 0x10;

        public static int Verify(string path)
        {
            Guid action = new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");
            WINTRUST_FILE_INFO fileInfo = new WINTRUST_FILE_INFO();
            fileInfo.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_FILE_INFO));
            fileInfo.pcwszFilePath = path;
            IntPtr filePointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)));
            IntPtr dataPointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_DATA)));
            try
            {
                Marshal.StructureToPtr(fileInfo, filePointer, false);
                WINTRUST_DATA data = new WINTRUST_DATA();
                data.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_DATA));
                data.dwUIChoice = WTD_UI_NONE;
                data.fdwRevocationChecks = WTD_REVOKE_NONE;
                data.dwUnionChoice = WTD_CHOICE_FILE;
                data.pFile = filePointer;
                data.dwStateAction = WTD_STATEACTION_VERIFY;
                data.dwProvFlags = WTD_REVOCATION_CHECK_NONE;
                Marshal.StructureToPtr(data, dataPointer, false);
                int result = WinVerifyTrust(IntPtr.Zero, action, dataPointer);

                data = (WINTRUST_DATA)Marshal.PtrToStructure(dataPointer, typeof(WINTRUST_DATA));
                data.dwStateAction = WTD_STATEACTION_CLOSE;
                Marshal.StructureToPtr(data, dataPointer, true);
                WinVerifyTrust(IntPtr.Zero, action, dataPointer);
                return result;
            }
            finally
            {
                Marshal.DestroyStructure(filePointer, typeof(WINTRUST_FILE_INFO));
                Marshal.FreeHGlobal(filePointer);
                Marshal.FreeHGlobal(dataPointer);
            }
        }
    }
}
'@
    }
    return [SmartM365.DeviceRebootManagerWinTrust]::Verify($Path)
}

function Get-PinnedSignatureState {
    # Signature rule: WinVerifyTrust must return 0, CERT_E_UNTRUSTEDROOT or
    # CERT_E_CHAINING (self-signed root not trusted locally) and the signer
    # thumbprint must equal the pinned one. Any other result, including
    # TRUST_E_BAD_DIGEST, TRUST_E_NOSIGNATURE or an expired untimestamped
    # certificate, is rejected.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    $actual = if ($signature.SignerCertificate) {
        ([string]$signature.SignerCertificate.Thumbprint).Replace(' ', '').ToUpperInvariant()
    }
    else {
        ''
    }
    $pinnedSignerMatch = ($actual -ne '' -and $actual -eq ([string]$Thumbprint).Replace(' ', '').ToUpperInvariant())
    $trustResult = Get-WinVerifyTrustResult -Path ([IO.Path]::GetFullPath($Path))
    # 0x800B0109 CERT_E_UNTRUSTEDROOT, 0x800B010A CERT_E_CHAINING
    $untrustedChainResults = @([int]-2146762487, [int]-2146762486)
    $chainTrusted = ($trustResult -eq 0)

    [pscustomobject]@{
        Path              = $Path
        Status            = [string]$signature.Status
        TrustResult       = ('0x{0:X8}' -f $trustResult)
        SignerThumbprint  = $actual
        PinnedSignerMatch = $pinnedSignerMatch
        ChainTrusted      = $chainTrusted
        Accepted          = ($pinnedSignerMatch -and ($chainTrusted -or $untrustedChainResults -contains $trustResult))
    }
}

function Assert-PinnedSignature {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $state = Get-PinnedSignatureState -Path $Path -Thumbprint $Thumbprint
    if (-not $state.Accepted) {
        throw ("Authenticode validation failed: file={0}; status={1}; winVerifyTrust={2}; signer={3}" -f
            $Path,$state.Status,$state.TrustResult,$state.SignerThumbprint)
    }
    if (-not $state.ChainTrusted) {
        Write-Warning ("Pinned signer accepted but its certificate chain is not trusted locally: file={0}; winVerifyTrust={1}" -f $Path,$state.TrustResult)
    }
}

function Set-PinnedCatalogSignature {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Thumbprint,
        [string]$Timestamp
    )

    $normalizedThumbprint = Get-NormalizedThumbprint $Thumbprint
    $certificate = Get-ChildItem -LiteralPath "Cert:\CurrentUser\My\$normalizedThumbprint" -ErrorAction SilentlyContinue
    if (-not $certificate -or -not $certificate.HasPrivateKey) {
        throw "Pinned code-signing certificate with private key not found: Cert:\CurrentUser\My\$normalizedThumbprint"
    }

    $signingParameters = @{
        LiteralPath   = $Path
        Certificate   = $certificate
        HashAlgorithm = 'SHA256'
        ErrorAction   = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($Timestamp)) {
        $signingParameters.TimestampServer = $Timestamp
    }
    [void](Set-AuthenticodeSignature @signingParameters)
}

function Get-PackageVersionText {
    param([Parameter(Mandatory = $true)][object]$Manifest)

    $prereleaseLabel = [string]$Manifest.PrivateData.PSData['Prerelease']
    if ([string]::IsNullOrWhiteSpace($prereleaseLabel)) {
        return [string]$Manifest.Version
    }
    return ('{0}-{1}' -f $Manifest.Version,$prereleaseLabel)
}

function Assert-VersionManifest {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedPackageVersion
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Package version manifest not found: $Path"
    }
    $versionManifest = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if ([int]$versionManifest.SchemaVersion -ne 1 -or [string]$versionManifest.PackageVersion -ne $ExpectedPackageVersion) {
        throw ("Package version mismatch: manifest={0}; module={1}; file={2}" -f $versionManifest.PackageVersion,$ExpectedPackageVersion,$Path)
    }
}

function Test-GalleryPackage {
    # Shared validation for a freshly built package and for -ValidatePackagePath.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ProductRoot,
        [Parameter(Mandatory = $true)][string]$Thumbprint,
        [bool]$SkipSignatures
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Package folder not found: $Path"
    }
    $resolvedPath = (Resolve-Path -LiteralPath $Path).ProviderPath

    $builtManifestPath = Join-Path -Path $resolvedPath -ChildPath "$script:ExpectedModuleName.psd1"
    if (-not (Test-Path -LiteralPath $builtManifestPath -PathType Leaf)) {
        throw "Package manifest not found: $builtManifestPath"
    }

    $packagedFiles = @(Get-ChildItem -LiteralPath $resolvedPath -Recurse -File)
    $powerShellFiles = @($packagedFiles | Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
    foreach ($file in $powerShellFiles) {
        Assert-PowerShellSyntax -Path $file.FullName
    }

    $builtManifest = Test-ModuleManifest -Path $builtManifestPath
    if ([string]$builtManifest.Name -ne $script:ExpectedModuleName) {
        throw "Unexpected module name: $($builtManifest.Name)"
    }
    $builtPackageVersion = Get-PackageVersionText -Manifest $builtManifest
    Assert-VersionManifest `
        -Path (Join-Path -Path $resolvedPath -ChildPath 'Runtime\SmartM365-DeviceRebootManager.version.json.txt') `
        -ExpectedPackageVersion $builtPackageVersion
    Assert-VersionManifest `
        -Path (Join-Path -Path $ProductRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt') `
        -ExpectedPackageVersion $builtPackageVersion

    $packagedDetectionPath = Join-Path -Path $resolvedPath -ChildPath 'Runtime\Deploy\SmartM365-DeviceRebootManager-Detection.ps1'
    if (-not (Test-Path -LiteralPath $packagedDetectionPath -PathType Leaf)) {
        throw "Packaged detection script not found: $packagedDetectionPath"
    }
    $detectionVersionMatch = [regex]::Match(
        (Get-Content -LiteralPath $packagedDetectionPath -Raw),
        '\[string\]\$ExpectedVersion\s*=\s*''([^'']+)'''
    )
    if (-not $detectionVersionMatch.Success -or $detectionVersionMatch.Groups[1].Value -ne $builtPackageVersion) {
        throw ("Packaged detection version does not match the package version: package={0}; detection={1}" -f
            $builtPackageVersion,$detectionVersionMatch.Groups[1].Value)
    }

    $catalogFiles = @($packagedFiles | Where-Object { $_.Extension -eq '.cat' })
    $catalogPath = Join-Path -Path $resolvedPath -ChildPath $script:PackageCatalogFileName
    if ($catalogFiles.Count -ne 1 -or $catalogFiles[0].FullName -ne $catalogPath) {
        throw "The package must contain exactly one file catalog: $catalogPath"
    }
    $catalogResult = Test-FileCatalog -Path $resolvedPath -CatalogFilePath $catalogPath -FilesToSkip '*.cat' -Detailed
    if ([string]$catalogResult.Status -ne 'Valid') {
        $differences = New-Object Collections.Generic.List[string]
        foreach ($key in @($catalogResult.CatalogItems.Keys)) {
            if (-not $catalogResult.PathItems.ContainsKey($key)) { $differences.Add("missing=$key") }
            elseif ([string]$catalogResult.PathItems[$key] -ne [string]$catalogResult.CatalogItems[$key]) { $differences.Add("changed=$key") }
        }
        foreach ($key in @($catalogResult.PathItems.Keys)) {
            if (-not $catalogResult.CatalogItems.ContainsKey($key)) { $differences.Add("extra=$key") }
        }
        throw ("Package file catalog validation failed: status={0}; differences={1}" -f
            $catalogResult.Status,(@($differences | Sort-Object) -join ', '))
    }

    if (-not $SkipSignatures) {
        Assert-PinnedSignature -Path $catalogPath -Thumbprint $Thumbprint
        foreach ($file in $powerShellFiles) {
            Assert-PinnedSignature -Path $file.FullName -Thumbprint $Thumbprint
        }
    }

    [pscustomobject]@{
        ModuleName          = [string]$builtManifest.Name
        Version             = [string]$builtManifest.Version
        Prerelease          = [string]$builtManifest.PrivateData.PSData['Prerelease']
        PackageVersion      = $builtPackageVersion
        PackagePath         = $resolvedPath
        CatalogPath         = $catalogPath
        CatalogFileCount    = $catalogResult.CatalogItems.Count
        FileCount           = $packagedFiles.Count
        PowerShellFileCount = $powerShellFiles.Count
    }
}

if ($PSCmdlet.ParameterSetName -eq 'Validate') {
    $validation = Test-GalleryPackage `
        -Path $ValidatePackagePath `
        -ProductRoot $SourceRoot `
        -Thumbprint $ExpectedSignerThumbprint `
        -SkipSignatures ([bool]$SkipAuthenticodeValidation)
    return [pscustomobject]@{
        Mode                     = 'Validate'
        ModuleName               = $validation.ModuleName
        Version                  = $validation.Version
        Prerelease               = $validation.Prerelease
        PackageVersion           = $validation.PackageVersion
        PackagePath              = $validation.PackagePath
        IntuneSourcePath         = Join-Path -Path $validation.PackagePath -ChildPath 'Runtime'
        CatalogPath              = $validation.CatalogPath
        FileCount                = $validation.FileCount
        PowerShellFileCount      = $validation.PowerShellFileCount
        SignaturesValidated      = (-not $SkipAuthenticodeValidation)
        ExpectedSignerThumbprint = Get-NormalizedThumbprint $ExpectedSignerThumbprint
        Ready                    = $true
    }
}

$moduleTemplateRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Module'
$manifestSourcePath = Join-Path -Path $moduleTemplateRoot -ChildPath 'SmartM365.DeviceRebootManager.psd1'
if (-not (Test-Path -LiteralPath $manifestSourcePath -PathType Leaf)) {
    throw "Module manifest not found: $manifestSourcePath"
}

$manifest = Test-ModuleManifest -Path $manifestSourcePath
$moduleName = [string]$manifest.Name
$moduleVersion = [string]$manifest.Version
$prerelease = [string]$manifest.PrivateData.PSData['Prerelease']
$packageVersion = Get-PackageVersionText -Manifest $manifest

Assert-VersionManifest `
    -Path (Join-Path -Path $SourceRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt') `
    -ExpectedPackageVersion $packageVersion

$packagePath = Join-Path -Path (Join-Path -Path $OutputRoot -ChildPath $moduleName) -ChildPath $moduleVersion
Assert-SafeChildPath -Root $OutputRoot -Child $packagePath

if (Test-Path -LiteralPath $packagePath) {
    if (-not $Force) {
        throw "Build target already exists. Use -Force to replace it: $packagePath"
    }
    Remove-Item -LiteralPath $packagePath -Recurse -Force
}

New-Item -ItemType Directory -Path $packagePath -Force | Out-Null

$moduleFiles = @(
    'SmartM365.DeviceRebootManager.psd1'
    'SmartM365.DeviceRebootManager.psm1'
    'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
)

$runtimeFiles = @(
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-Demo.ps1'
    'Start-SmartM365-DeviceRebootManager-Demo.cmd'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
)

$deployFiles = @(
    'SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    'SmartM365-DeviceRebootManager-Detection.ps1'
    'SmartM365-DeviceRebootManager-Install.ps1'
    'SmartM365-DeviceRebootManager-Uninstall.ps1'
)

foreach ($relativePath in $moduleFiles) {
    $sourcePath = Join-Path -Path $moduleTemplateRoot -ChildPath $relativePath
    $destinationPath = Join-Path -Path $packagePath -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required module file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

foreach ($relativePath in $runtimeFiles) {
    $sourcePath = Join-Path -Path $SourceRoot -ChildPath $relativePath
    $destinationPath = Join-Path -Path (Join-Path $packagePath 'Runtime') -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required runtime file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

foreach ($relativePath in $deployFiles) {
    $sourcePath = Join-Path -Path (Join-Path $SourceRoot 'Deploy') -ChildPath $relativePath
    $destinationPath = Join-Path -Path (Join-Path $packagePath 'Runtime\Deploy') -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required deployment file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

# The catalog covers every packaged file, including non-PowerShell runtime files.
$catalogPath = Join-Path -Path $packagePath -ChildPath $script:PackageCatalogFileName
New-FileCatalog -Path $packagePath -CatalogFilePath $catalogPath -CatalogVersion 2.0 | Out-Null
if (-not $SkipAuthenticodeValidation) {
    Set-PinnedCatalogSignature -Path $catalogPath -Thumbprint $ExpectedSignerThumbprint -Timestamp $TimestampServer
}

$validation = Test-GalleryPackage `
    -Path $packagePath `
    -ProductRoot $SourceRoot `
    -Thumbprint $ExpectedSignerThumbprint `
    -SkipSignatures ([bool]$SkipAuthenticodeValidation)

[pscustomobject]@{
    ModuleName               = $moduleName
    Version                  = $moduleVersion
    Prerelease               = $prerelease
    PackageVersion           = $packageVersion
    PackagePath              = $validation.PackagePath
    IntuneSourcePath         = Join-Path -Path $validation.PackagePath -ChildPath 'Runtime'
    CatalogPath              = $validation.CatalogPath
    CatalogSigned            = (-not $SkipAuthenticodeValidation)
    FileCount                = $validation.FileCount
    PowerShellFileCount      = $validation.PowerShellFileCount
    SignaturesValidated      = (-not $SkipAuthenticodeValidation)
    ExpectedSignerThumbprint = Get-NormalizedThumbprint $ExpectedSignerThumbprint
    Ready                    = $true
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAQe+cS5TMczEt5
# vgbTl6it42664OJ1uewZ+fVu3ESVaKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHoHast7CtNFsq6sBiD6QeeCBlFHGLLKaKs8zAAMTYtmMA0GCSqG
# SIb3DQEBAQUABIIBgHL6BcWUo+t5yoVZAFWFZt14U1iS72lNyKM6ZxKKGXSeoiZ8
# 9kuzfvpn/Ui3Y0RNztI8s+TdB6anyGC23vXqVebtKMzaeDF6ErsEBogevVZQ1Iai
# DgSEUQAgXoqWl3eJlbLSlR4uVm1YnX4KvvtX1lUQKnUmaI8OCS5jMr17c1nU5Wlu
# 9eDEKjBsQPTlR/TeY+/2M8xYmq2Zyz+UqWfrVTZcheVNfFduxZ3IQmY8X9itGDQ8
# cKjoMYJddHQypd4XSNpdgjLN7+nhMwIhSBEPyqPSL/JVtqNpAYM4TLip8f6RVtX+
# K9kYHlUbq9bOF9FyPdukvvZJaSgXziLn107K1aDWgtFM3CxXoPS4P7BW8PoO7BMB
# K+i1bLWkFvaS87Wu71uczUhjCBPteNdXjpfHf0gLYNAOmgWkZg8c6ESDMU6txpI4
# GGSAAW49JAsTqaaPfhMfRsKEhv9VK2llFw9fAyKhSc0/P7dHtJgKCQRxcjk4DFN7
# cYrRXqHaIP+w/7DhwqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjQy
# MjJaMC8GCSqGSIb3DQEJBDEiBCB2/9NqVHQZUAwAKBg2HDRjr28JrU9lk4GzI216
# PGGV0TANBgkqhkiG9w0BAQEFAASCAgBswNOVAV10Gc8rfFZR6ZJICMc5MOl83fw0
# rvWxTWZ/k790CEHNzuik7TUGEBBEOJkSvkYKH/Is5WnRSKSE5eHkgHWjq2EstD6m
# C8nc//5HZrzIk5ZWu0amhmUUL2Qcuav/CumU1CixTPs88SctjoGuT8dqyfi7BZjj
# LQrRd7Vjla2yXuZfFAEz3dhjEnqibVnv+NtQ01jXRq145HtSk0tVucK/+oX/5QFB
# 5JLWUhASPFakez2LP38gu1rxq5QrgTbR445ZpnOXqXIIVOfp403c6Rc9/3hj0WNd
# K/7DoJF/UNZkUkwrIz9/7UQd2rbqevJOwJZ/twc8Z/hCyo9I2RUbdm7UcZZVfhCQ
# f93NjTvIIwYXtv/1Q+u0yR3cByKOxiaESWxmFNwc7gqwT4uRpv5oENIzgOnVT/Hh
# 4FFrYWFRoy45sN9A55po2cylUgiA1e3clvf0sALv73Jf4r/gDefePY1fKe/Yw1bX
# p6jPtR4ciwDLepmXMAMG/O22H5zK2LI9VBqcWgOhifEU1U8wq5tPS2jweuY+W1T2
# SOKVVSGAdZpk0f6gyfookoqHRacEbOKG2mnZOm+PlcWweKOs7o1DOCwh6kuvTHxC
# 2JTXKILaV9vWcAtkQz7Czu5Cp6ik+uFIjVrjMwkjxB1wAXQ09kJkowL8ueOVevPm
# Q0v+FONz7g==
# SIG # End signature block
