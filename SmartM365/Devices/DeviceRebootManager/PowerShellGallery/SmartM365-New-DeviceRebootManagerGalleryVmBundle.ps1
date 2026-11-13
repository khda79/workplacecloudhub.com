#Requires -Version 5.1
<# .VERSION
1.0.2
#>

<#
.SYNOPSIS
Builds a signed Device Reboot Manager validation bundle for a clean Windows VM.

.DESCRIPTION
Builds the local Gallery module, copies the guarded VM validation script,
exports only the public signing certificate, writes VM instructions, and
creates a ZIP. It never installs the product or publishes to PowerShell Gallery.
#>

[CmdletBinding()]
param(
    [string]$OutputRoot = (Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath 'SmartM365\DeviceRebootManagerVmBundle'),
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'


function Test-WchDemoAssetPath {
    param([string]$Path)
    $name=$Path.Replace('\','/')
    return $name -match '(?i)(^|/)DemoAssets(/|$)|(^|/)[^/]*demo[-_ ]?assets[^/]*\.(zip|nupkg|7z|rar|tar|tgz|gz|intunewin)$'
}

function Assert-WchNoDemoAssets {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string[]]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $budget=@{Entries=0;Bytes=[long]0}
    function Read-WchReleaseArchive {
        param([IO.Stream]$Stream,[string]$Name,[int]$Depth)
        if($Depth -gt 8){throw "Release archive nesting limit exceeded: $Name"}
        $zip=[IO.Compression.ZipArchive]::new($Stream,[IO.Compression.ZipArchiveMode]::Read,$true)
        try {
            foreach($entry in $zip.Entries){
                $budget.Entries++;$budget.Bytes+=$entry.Length
                if($budget.Entries -gt 100000 -or $budget.Bytes -gt 536870912){throw "Release archive inspection limit exceeded: $Name"}
                if(Test-WchDemoAssetPath $entry.FullName){throw "DemoAssets are local-only: $Name!$($entry.FullName)"}
                if($entry.Name){
                    $content=$entry.Open();$memory=$null
                    try {
                        $head=New-Object byte[] 4;$read=$content.Read($head,0,4)
                        $magic=$read -eq 4 -and $head[0] -eq 80 -and $head[1] -eq 75 -and (($head[2] -eq 3 -and $head[3] -eq 4) -or ($head[2] -eq 5 -and $head[3] -eq 6) -or ($head[2] -eq 7 -and $head[3] -eq 8))
                        if($magic -or $entry.FullName -match '(?i)\.(zip|nupkg)$'){
                            $memory=[IO.MemoryStream]::new();$memory.Write($head,0,$read);$content.CopyTo($memory);$memory.Position=0
                            Read-WchReleaseArchive $memory ($Name+'!'+$entry.FullName) ($Depth+1)
                        }
                    }finally{$content.Dispose();if($memory){$memory.Dispose()}}
                }
            }
        }finally{$zip.Dispose()}
    }
    foreach($inputPath in $Path){
        if(Test-WchDemoAssetPath $inputPath){throw "DemoAssets are local-only: $inputPath"}
        $item=Get-Item -LiteralPath $inputPath -Force -ErrorAction Stop
        $pending=[Collections.Generic.Stack[IO.FileSystemInfo]]::new();$pending.Push($item)
        while($pending.Count){
            $file=$pending.Pop()
            if(Test-WchDemoAssetPath $file.FullName){throw "DemoAssets are local-only: $($file.FullName)"}
            if($file.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "Release input must not be a reparse point: $($file.FullName)"}
            if($file.PSIsContainer){foreach($child in @(Get-ChildItem -LiteralPath $file.FullName -Force -ErrorAction Stop)){$pending.Push($child)}}else{
                $stream=[IO.File]::OpenRead($file.FullName)
                try {
                    $head=New-Object byte[] 4;$read=$stream.Read($head,0,4);$stream.Position=0
                    $magic=$read -eq 4 -and $head[0] -eq 80 -and $head[1] -eq 75 -and (($head[2] -eq 3 -and $head[3] -eq 4) -or ($head[2] -eq 5 -and $head[3] -eq 6) -or ($head[2] -eq 7 -and $head[3] -eq 8))
                    if($magic -or $file.Extension -match '(?i)^\.(zip|nupkg)$'){Read-WchReleaseArchive $stream $file.FullName 0}
                }finally{$stream.Dispose()}
            }
        }
    }
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
        throw "Unsafe bundle target outside OutputRoot: $resolvedChild"
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

function New-ForwardSlashZip {
    # Builds the ZIP with '/' entry names (Compress-Archive on Windows PowerShell
    # 5.1 writes '\', which non-Windows extractors treat as part of the name).
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Assert-WchNoDemoAssets -Path $SourceRoot
    $resolvedSource = [IO.Path]::GetFullPath($SourceRoot).TrimEnd('\', '/')
    $zip = [IO.Compression.ZipFile]::Open($DestinationPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($file in @(Get-ChildItem -LiteralPath $resolvedSource -Recurse -File | Sort-Object FullName)) {
            $entryName = $file.FullName.Substring($resolvedSource.Length).TrimStart('\', '/').Replace('\', '/')
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zip,
                $file.FullName,
                $entryName,
                [IO.Compression.CompressionLevel]::Optimal
            )
        }
    }
    finally {
        $zip.Dispose()
    }
}

$moduleManifestPath = Join-Path -Path $PSScriptRoot -ChildPath 'Module\SmartM365.DeviceRebootManager.psd1'
$moduleManifest = Test-ModuleManifest -Path $moduleManifestPath
$version = [string]$moduleManifest.Version
$prerelease = [string]$moduleManifest.PrivateData.PSData['Prerelease']
$displayVersion = if ([string]::IsNullOrWhiteSpace($prerelease)) { $version } else { "$version-$prerelease" }
$bundleName = "SmartM365-DeviceRebootManager-VMValidation-$displayVersion"
$bundleRoot = Join-Path -Path $OutputRoot -ChildPath $bundleName
$zipPath = Join-Path -Path $OutputRoot -ChildPath "$bundleName.zip"
Assert-SafeChildPath -Root $OutputRoot -Child $bundleRoot
Assert-SafeChildPath -Root $OutputRoot -Child $zipPath

foreach ($targetPath in @($bundleRoot, $zipPath)) {
    if (Test-Path -LiteralPath $targetPath) {
        if (-not $Force) {
            throw "Bundle target already exists. Use -Force to replace it: $targetPath"
        }
        Remove-Item -LiteralPath $targetPath -Recurse -Force
    }
}
New-Item -ItemType Directory -Path $bundleRoot -Force | Out-Null

$buildScript = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1'
$buildResult = & $buildScript `
    -OutputRoot $bundleRoot `
    -ExpectedSignerThumbprint $ExpectedSignerThumbprint `
    -Force

$testsRoot = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'Tests'
$vmValidationSource = Join-Path -Path $testsRoot -ChildPath 'SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'
$vmValidationDestination = Join-Path -Path $bundleRoot -ChildPath 'SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'
if (-not (Test-Path -LiteralPath $vmValidationSource -PathType Leaf)) {
    throw "VM validation script not found: $vmValidationSource"
}
Copy-Item -LiteralPath $vmValidationSource -Destination $vmValidationDestination -Force

$validationSignature = Get-PinnedSignatureState -Path $vmValidationDestination -Thumbprint $ExpectedSignerThumbprint
if ($validationSignature.Accepted -and -not $validationSignature.ChainTrusted) {
    Write-Warning "Pinned signer accepted but its certificate chain is not trusted locally: $vmValidationDestination"
}
if (-not $validationSignature.Accepted) {
    throw 'VM validation script signature is not valid or does not match the pinned signer.'
}

$signingCertificate = Get-ChildItem -LiteralPath "Cert:\CurrentUser\My\$ExpectedSignerThumbprint" -ErrorAction Stop
$certificatePath = Join-Path -Path $bundleRoot -ChildPath 'WorkplaceCloudHub-CodeSigning.cer'
Export-Certificate -Cert $signingCertificate -FilePath $certificatePath -Force | Out-Null

$instructions = @"
SmartM365 Device Reboot Manager VM validation
Package: $($buildResult.ModuleName) $displayVersion

1. Extract the ZIP on an isolated Windows test VM.
2. Open an elevated Windows PowerShell 5.1 session in this folder.
3. Preview only (no installation):
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1
4. Execute the validation. The pinned self-signed signer is accepted without
   trusting it; add -TrustSignerCertificate only to also trust the bundled
   public certificate on this isolated VM:
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1 -Execute
5. Optional full cleanup after the validation:
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1 -Execute -TrustSignerCertificate -CleanupAfterTest

Expected final state without -CleanupAfterTest:
- Device Reboot Manager is installed.
- The main GUI task exists.
- The automatic Gallery update task does not exist.
- The runtime configuration is preserved.
"@
$instructionsPath = Join-Path -Path $bundleRoot -ChildPath 'SmartM365-VM-Validation-README.txt'
Set-Content -LiteralPath $instructionsPath -Value $instructions -Encoding UTF8

New-ForwardSlashZip -SourceRoot $bundleRoot -DestinationPath $zipPath
if (-not (Test-Path -LiteralPath $zipPath -PathType Leaf)) {
    throw "ZIP bundle was not created: $zipPath"
}

[pscustomobject]@{
    Result                     = 'PASS'
    ModuleName                 = $buildResult.ModuleName
    Version                    = $version
    Prerelease                 = $prerelease
    BundleRoot                 = $bundleRoot
    ZipPath                    = $zipPath
    ZipSizeBytes               = (Get-Item -LiteralPath $zipPath).Length
    PackageFileCount           = $buildResult.FileCount
    PackageSignaturesValidated = $buildResult.SignaturesValidated
    SignerCertificateIncluded  = $true
    ProductInstalled           = $false
    PublicationAttempted       = $false
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDA81lYiJxNo3s4
# rZAgzw43pAC1Ke0EtVumVQnQ/oTTR6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJxtKm7dYOlCgRllQVkx4Lcy6m1HCxj1ehUWFUZEuSLHMA0GCSqG
# SIb3DQEBAQUABIIBgJy5E2ecqOai/5jA3AyZkCloBHYoYElkDwmJqPrHe0paiVys
# TbAozrm8L8hjNNvV54EmMmLau18QNLT3bxRIOPxUyy2gXKffSe/jicwCvV1YBE6A
# MCgVeIcarhNWcCJ3mG3ndNilkEWks+dynOSwMhyAvMzUU9qRwXSYkb0Yz9sKDCCI
# /xkoITdss6sBn02j6HqvU/xK/gD4UvZNpX7+Y9WFbL14h1S71O010Ub1QCJdokoc
# QSX7GxHunblvxax6tnq30knUJcfGOdEywNi8yMmf4c9yWm9Dz5bXcAcGCY0eiiQR
# j+bMuL5SJtAzB4K6Fzry0KEUd5K17WVsyJTX4SqvKeRzAO0AWNQQzD92NW6hCebZ
# yfskHfvWo1cMIYNMD0pjj3sY2KbqdvgsKT4Qm18Z5hS3pmKmO7IWhwx9/Lt5Tnv4
# XwdAHqPQVHbfmWGbP9cU34AwYhO1xmO04LsNT3U+LyCUd/Fd3mxH7wmyDoGtalWm
# 9/iyic+G7Bzsr53tXaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMDE1
# NTFaMC8GCSqGSIb3DQEJBDEiBCC9R2aO3tE+5Z1d7Tbd6U1TEqqz/w6/pIQ0lD7l
# p0RZmzANBgkqhkiG9w0BAQEFAASCAgAkFScjwlKt0SGG65AgEx3cCQwaoj6iaT5N
# Gmyb2tVtwy9uW7mdv88gEMutjzA/wcWMP5DZ7ZqPkvbpN8N5tLxu40/BIqfBHeXD
# +RJe8E2eNXQZ5F/Ok0wikssVoItq8mw5Lksvuj1HA4igTVNRx5Vk40JqohSNurEX
# z77S6Sw4ch/BJz1i5pkdSLDqB8Kv9fGfB0SgomxJdrx/irdYsTx3q+YNKDtblUKl
# o1h8rnbqaCiS/zTdY8mcaQyPBCx13ooG1PQxTz2p/5dvd+gfZVEvOZV4l1fdIjhu
# PDfVCZrOTAsXByDgtOyHQ5SJWspgqQFqbFtE2SOhlDLG8U8ut+EEfjSiG+Nc6P9G
# oKrCqrKi+1n2qjam4p96wDbw5af1306m6ro3+ocMRs8LOlyjQZgok1JMiwRPR24x
# jbwlkBvJj9buSuap1zzUZsAGC7q/p87teGwymF3N7Cf8Sq376OdfDw6iDmDV1WKk
# mfAe9DKDACWzsq2E8QvnMS9/7BYn1bHtRoP4oOCx+uJfJctL/YOaXVfX+4P0dtHA
# kvUiR7ItawRDtEFpVKdq0s2E9C9LlBlup09u7ot0793p547eNfzkCqz4tnV9gPm9
# cxUPjJL6+qbh5IleIPQV5/ykpLNBh1vVcqujFXqajhWd90kz8nDAlaRs3ZQFQ6qx
# 1LNOqY1JJw==
# SIG # End signature block
