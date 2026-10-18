#Requires -Version 7.0
<#
.SYNOPSIS
    Packaging engines for Smart IntunePackager.

.DESCRIPTION
    Engines are registered in a name-keyed registry. Each engine provides:
    - Test  : scriptblock(Request) returning a string array of blocking issues (empty = available)
    - Build : scriptblock(Request) returning an object with IntuneWinPath and EngineVersion

    Request is a hashtable with SourceFolder, SetupFile, OutputFolder, Settings (local
    configuration object) and AppRoot. Invoke-IpPackageBuild validates the request,
    runs the selected engine and verifies the produced .intunewin.

    The shipped engine wraps the official Microsoft Win32 Content Prep Tool
    (IntuneWinAppUtil.exe), pinned by version and SHA256, downloaded on demand from the
    official Microsoft GitHub repository and verified for a valid Microsoft Authenticode
    signature before use. A native PowerShell engine can be registered later with the
    same contract.

.VERSION
    1.0.1

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version Latest
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.IntunePackager.Core.psm1') -Scope Local

$script:IpEngines = [ordered]@{}
$script:IpContentPrepToolFileName = 'IntuneWinAppUtil.exe'
$script:IpOfficialToolUrlPattern = '^https://(raw\.githubusercontent\.com/microsoft/Microsoft-Win32-Content-Prep-Tool/|github\.com/microsoft/Microsoft-Win32-Content-Prep-Tool/(raw|releases/download)/)'
$script:IpMicrosoftRootPattern = '^CN=Microsoft Root Certificate Authority( 2010| 2011)?,'

#region Engine registry

function Register-IpPackagingEngine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Description,
        [Parameter(Mandatory = $true)][scriptblock]$Test,
        [Parameter(Mandatory = $true)][scriptblock]$Build
    )
    $script:IpEngines[$Name] = [pscustomobject]@{ Name = $Name; Description = $Description; Test = $Test; Build = $Build }
}

function Unregister-IpPackagingEngine {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    if ($script:IpEngines.Contains($Name)) { $script:IpEngines.Remove($Name) }
}

function Get-IpPackagingEngine {
    [CmdletBinding()]
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return @($script:IpEngines.Values) }
    if (-not $script:IpEngines.Contains($Name)) {
        throw ("Unknown packaging engine '{0}'. Registered engines: {1}" -f $Name, (@($script:IpEngines.Keys) -join ', '))
    }
    return $script:IpEngines[$Name]
}

#endregion

#region Build request

function Test-IpPathIsUnder {
    param([string]$Path, [string]$Parent)
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\') + '\'
    $root = [IO.Path]::GetFullPath($Parent).TrimEnd('\') + '\'
    return $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)
}

function Test-IpPackageBuildRequest {
    <#
    .SYNOPSIS
        Returns blocking issues for a build request. An empty result means valid.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [AllowEmptyString()][string]$SourceFolder,
        [AllowEmptyString()][string]$SetupFile,
        [AllowEmptyString()][string]$OutputFolder
    )

    $errors = [Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($SourceFolder) -or -not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
        $errors.Add("Source folder not found: $SourceFolder")
    }
    if ([string]::IsNullOrWhiteSpace($SetupFile)) { $errors.Add('Setup file is required.') }
    elseif ([IO.Path]::IsPathRooted($SetupFile) -or $SetupFile -match '(^|[\\/])\.\.([\\/]|$)') {
        $errors.Add('Setup file must be a relative path inside the source folder.')
    }
    elseif ($errors.Count -eq 0 -and -not (Test-Path -LiteralPath (Join-Path $SourceFolder $SetupFile) -PathType Leaf)) {
        $errors.Add("Setup file not found in the source folder: $SetupFile")
    }
    if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $errors.Add('Output folder is required.') }
    elseif ($SourceFolder -and (Test-Path -LiteralPath $SourceFolder -PathType Container) -and (Test-IpPathIsUnder -Path $OutputFolder -Parent $SourceFolder)) {
        $errors.Add('Output folder must not be the source folder or one of its subfolders.')
    }
    return [string[]]$errors.ToArray()
}

function Invoke-IpPackageBuild {
    <#
    .SYNOPSIS
        Builds a .intunewin package with the selected engine and returns its metadata.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EngineName,
        [Parameter(Mandatory = $true)][string]$SourceFolder,
        [Parameter(Mandatory = $true)][string]$SetupFile,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $true)][psobject]$Settings,
        [Parameter(Mandatory = $true)][string]$AppRoot
    )

    $issues = @(Test-IpPackageBuildRequest -SourceFolder $SourceFolder -SetupFile $SetupFile -OutputFolder $OutputFolder)
    if ($issues.Count -gt 0) { throw ('Package build request is invalid: ' + ($issues -join ' ')) }

    $engine = Get-IpPackagingEngine -Name $EngineName
    $request = @{
        SourceFolder = [IO.Path]::GetFullPath($SourceFolder)
        SetupFile    = $SetupFile
        OutputFolder = [IO.Path]::GetFullPath($OutputFolder)
        Settings     = $Settings
        AppRoot      = $AppRoot
    }
    $engineIssues = @(& $engine.Test $request)
    if ($engineIssues.Count -gt 0) { throw ("Packaging engine '{0}' is not available: {1}" -f $EngineName, ($engineIssues -join ' ')) }

    [IO.Directory]::CreateDirectory($request.OutputFolder) | Out-Null
    Write-IpLog ("Packaging started. Engine={0}; Source={1}; SetupFile={2}; Output={3}" -f $EngineName, $request.SourceFolder, $SetupFile, $request.OutputFolder)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $engineResult = & $engine.Build $request
    $watch.Stop()

    $intuneWinPath = [string]$engineResult.IntuneWinPath
    if (-not (Test-Path -LiteralPath $intuneWinPath -PathType Leaf)) { throw "Packaging engine did not produce the expected file: $intuneWinPath" }
    $metadata = Read-IpIntuneWinMetadata -Path $intuneWinPath
    $hash = (Get-FileHash -LiteralPath $intuneWinPath -Algorithm SHA256).Hash
    Write-IpLog ("Packaging completed in {0} s. File={1}; SHA256={2}; EncryptedSize={3}" -f $watch.Elapsed.TotalSeconds.ToString("0.0", [Globalization.CultureInfo]::InvariantCulture), $intuneWinPath, $hash, (Format-IpByteSize $metadata.SizeEncrypted))

    return [pscustomobject]@{
        IntuneWinPath = $intuneWinPath
        Sha256        = $hash
        Engine        = $EngineName
        EngineVersion = [string]$engineResult.EngineVersion
        Duration      = $watch.Elapsed
        Metadata      = $metadata
    }
}

#endregion

#region Win32 Content Prep Tool

function Get-IpContentPrepToolPath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][psobject]$ToolSettings,
        [Parameter(Mandatory = $true)][string]$AppRoot
    )

    if ([string]::IsNullOrWhiteSpace($ToolSettings.Version)) { throw 'Configuration error: ContentPrepTool.Version is required.' }
    if ([string]::IsNullOrWhiteSpace($ToolSettings.StoragePath)) { throw 'Configuration error: ContentPrepTool.StoragePath is required.' }
    $storage = Resolve-IpPath -Path $ToolSettings.StoragePath -BaseDirectory $AppRoot
    return (Join-Path -Path (Join-Path -Path $storage -ChildPath ([string]$ToolSettings.Version)) -ChildPath $script:IpContentPrepToolFileName)
}

function Test-IpContentPrepTool {
    <#
    .SYNOPSIS
        Verifies the pinned SHA256 and a valid Microsoft Authenticode signature.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ExpectedSha256
    )

    $reasons = [Collections.Generic.List[string]]::new()
    $hash = ''
    $status = ''
    $signer = ''
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $reasons.Add("File not found: $Path")
    }
    else {
        $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
        if ([string]::IsNullOrWhiteSpace($ExpectedSha256)) { $reasons.Add('Configuration error: ContentPrepTool.Sha256 is empty.') }
        elseif ($hash -ne $ExpectedSha256.Trim().ToUpperInvariant()) { $reasons.Add(("SHA256 mismatch. Expected={0}; Actual={1}" -f $ExpectedSha256, $hash)) }

        $signature = Get-AuthenticodeSignature -LiteralPath $Path
        $status = [string]$signature.Status
        if ($signature.Status -ne [Management.Automation.SignatureStatus]::Valid) { $reasons.Add("Authenticode signature is not valid: $status") }
        $certificate = $signature.SignerCertificate
        if ($null -eq $certificate) { $reasons.Add('No signer certificate.') }
        else {
            $signer = $certificate.Subject
            if ($signer -notmatch '^CN=Microsoft Corporation,' -or $signer -notmatch 'O=Microsoft Corporation') { $reasons.Add("Unexpected signer: $signer") }
            # The signing certificate may be expired while the timestamped signature stays valid,
            # so only the chain structure up to a Microsoft root is checked here.
            $chain = [Security.Cryptography.X509Certificates.X509Chain]::new()
            try {
                $chain.ChainPolicy.RevocationMode = [Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
                $chain.ChainPolicy.VerificationFlags = [Security.Cryptography.X509Certificates.X509VerificationFlags]::IgnoreNotTimeValid
                [void]$chain.Build($certificate)
                $elements = @($chain.ChainElements)
                $root = if ($elements.Count -gt 0) { $elements[-1].Certificate.Subject } else { '' }
                if ($root -notmatch $script:IpMicrosoftRootPattern) { $reasons.Add("Signer chain does not end at a Microsoft root: $root") }
            }
            finally { $chain.Dispose() }
        }
    }

    return [pscustomobject]@{
        Path            = $Path
        IsValid         = ($reasons.Count -eq 0)
        Sha256          = $hash
        SignatureStatus = $status
        Signer          = $signer
        Reasons         = [string[]]$reasons.ToArray()
    }
}

function Install-IpContentPrepTool {
    <#
    .SYNOPSIS
        Ensures the pinned IntuneWinAppUtil.exe is present and verified, downloading it if needed.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][psobject]$ToolSettings,
        [Parameter(Mandatory = $true)][string]$AppRoot,
        [switch]$Force
    )

    $toolPath = Get-IpContentPrepToolPath -ToolSettings $ToolSettings -AppRoot $AppRoot
    if (-not $Force -and (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        $check = Test-IpContentPrepTool -Path $toolPath -ExpectedSha256 ([string]$ToolSettings.Sha256)
        if ($check.IsValid) { return $toolPath }
        Write-IpLog ("Existing Win32 Content Prep Tool failed verification and will be replaced: {0}" -f ($check.Reasons -join ' ')) -Level WARN
    }

    $url = [string]$ToolSettings.DownloadUrl
    if ($url -notmatch $script:IpOfficialToolUrlPattern) {
        throw "Configuration error: ContentPrepTool.DownloadUrl must point to the official Microsoft GitHub repository: $url"
    }
    if (-not $PSCmdlet.ShouldProcess($toolPath, "Download Win32 Content Prep Tool $($ToolSettings.Version) from $url")) { return $toolPath }

    $versionFolder = Split-Path -Path $toolPath -Parent
    $downloadFolder = Join-Path -Path $versionFolder -ChildPath ('.download-' + [guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($downloadFolder) | Out-Null
    $downloadPath = Join-Path -Path $downloadFolder -ChildPath $script:IpContentPrepToolFileName
    try {
        Write-IpLog ("Downloading Win32 Content Prep Tool {0}: {1}" -f $ToolSettings.Version, $url)
        Invoke-WebRequest -Uri $url -OutFile $downloadPath -UseBasicParsing -ErrorAction Stop
        $check = Test-IpContentPrepTool -Path $downloadPath -ExpectedSha256 ([string]$ToolSettings.Sha256)
        if (-not $check.IsValid) { throw ('Downloaded Win32 Content Prep Tool failed verification: ' + ($check.Reasons -join ' ')) }
        Move-Item -LiteralPath $downloadPath -Destination $toolPath -Force
        Write-IpLog ("Win32 Content Prep Tool verified. SHA256={0}; Signer={1}; Path={2}" -f $check.Sha256, $check.Signer, $toolPath)
    }
    finally {
        Remove-Item -LiteralPath $downloadFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    return $toolPath
}

function Get-IpIntuneWinAppUtilArguments {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)][string]$SourceFolder,
        [Parameter(Mandatory = $true)][string]$SetupFile,
        [Parameter(Mandatory = $true)][string]$OutputFolder
    )
    return @('-c', $SourceFolder, '-s', $SetupFile, '-o', $OutputFolder, '-q')
}

function Invoke-IpIntuneWinAppUtil {
    <#
    .SYNOPSIS
        Runs IntuneWinAppUtil.exe, streams its output to the log and returns the produced file path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ToolPath,
        [Parameter(Mandatory = $true)][string]$SourceFolder,
        [Parameter(Mandatory = $true)][string]$SetupFile,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [ValidateRange(1, 1440)][int]$TimeoutMinutes = 120
    )

    $expected = Join-Path -Path $OutputFolder -ChildPath ([IO.Path]::GetFileNameWithoutExtension($SetupFile) + '.intunewin')
    if (Test-Path -LiteralPath $expected) { Remove-Item -LiteralPath $expected -Force }

    $startInfo = [Diagnostics.ProcessStartInfo]::new($ToolPath)
    foreach ($argument in (Get-IpIntuneWinAppUtilArguments -SourceFolder $SourceFolder -SetupFile $SetupFile -OutputFolder $OutputFolder)) {
        $startInfo.ArgumentList.Add($argument)
    }
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    $startInfo.WorkingDirectory = $OutputFolder

    $process = [Diagnostics.Process]::Start($startInfo)
    try {
        $errorTask = $process.StandardError.ReadToEndAsync()
        $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
        $lineTask = $null
        while ($true) {
            if ($null -eq $lineTask) { $lineTask = $process.StandardOutput.ReadLineAsync() }
            if ($lineTask.Wait(500)) {
                $line = $lineTask.Result
                $lineTask = $null
                if ($null -eq $line) { break }
                if (-not [string]::IsNullOrWhiteSpace($line)) { Write-IpLog ('IntuneWinAppUtil: ' + $line.Trim()) }
            }
            if ((Get-Date) -gt $deadline) {
                try { $process.Kill($true) } catch { Write-Verbose $_.Exception.Message }
                throw "IntuneWinAppUtil did not complete within $TimeoutMinutes minutes."
            }
        }
        $process.WaitForExit()
        $standardError = $errorTask.GetAwaiter().GetResult()
        foreach ($line in @([regex]::Split([string]$standardError, '\r?\n') | Where-Object { $_.Trim() })) {
            Write-IpLog ('IntuneWinAppUtil: ' + $line.Trim()) -Level WARN
        }
        if ($process.ExitCode -ne 0) { throw "IntuneWinAppUtil failed with exit code $($process.ExitCode)." }
    }
    finally {
        $process.Dispose()
    }

    if (-not (Test-Path -LiteralPath $expected -PathType Leaf)) { throw "IntuneWinAppUtil completed but the expected output was not found: $expected" }
    return $expected
}

function Invoke-IpIntuneWinAppUtilBuild {
    <#
    .SYNOPSIS
        Build implementation of the IntuneWinAppUtil engine, with optional short-path staging.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Request)

    $settings = $Request.Settings
    $toolSettings = $settings.ContentPrepTool
    $toolPath = Install-IpContentPrepTool -ToolSettings $toolSettings -AppRoot $Request.AppRoot -Confirm:$false
    $timeout = [int]$settings.Packaging.BuildTimeoutMinutes
    if ($timeout -lt 1) { $timeout = 120 }

    if (-not [bool]$settings.Packaging.UseShortPathStaging) {
        $path = Invoke-IpIntuneWinAppUtil -ToolPath $toolPath -SourceFolder $Request.SourceFolder -SetupFile $Request.SetupFile -OutputFolder $Request.OutputFolder -TimeoutMinutes $timeout
        return [pscustomobject]@{ IntuneWinPath = $path; EngineVersion = [string]$toolSettings.Version }
    }

    # Short staging paths avoid MAX_PATH failures in the Content Prep Tool for deep source trees.
    $stagingRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('IP' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $stagingSource = Join-Path $stagingRoot 'S'
    $stagingOutput = Join-Path $stagingRoot 'O'
    try {
        [IO.Directory]::CreateDirectory($stagingSource) | Out-Null
        [IO.Directory]::CreateDirectory($stagingOutput) | Out-Null
        Write-IpLog "Copying source to short staging path: $stagingSource"
        $robocopy = Join-Path -Path $env:SystemRoot -ChildPath 'System32\robocopy.exe'
        & $robocopy $Request.SourceFolder $stagingSource /MIR /R:2 /W:2 /NP /NFL /NDL /NJH /NJS | Out-Null
        if ($LASTEXITCODE -gt 7) { throw "Robocopy staging copy failed with exit code $LASTEXITCODE." }
        $stagedPath = Invoke-IpIntuneWinAppUtil -ToolPath $toolPath -SourceFolder $stagingSource -SetupFile $Request.SetupFile -OutputFolder $stagingOutput -TimeoutMinutes $timeout
        $finalPath = Join-Path -Path $Request.OutputFolder -ChildPath ([IO.Path]::GetFileName($stagedPath))
        Move-Item -LiteralPath $stagedPath -Destination $finalPath -Force
        return [pscustomobject]@{ IntuneWinPath = $finalPath; EngineVersion = [string]$toolSettings.Version }
    }
    finally {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

#endregion

Register-IpPackagingEngine -Name 'IntuneWinAppUtil' `
    -Description 'Official Microsoft Win32 Content Prep Tool (pinned version, SHA256 and Authenticode verified).' `
    -Test {
    param($Request)
    $issues = [Collections.Generic.List[string]]::new()
    $toolSettings = $Request.Settings.ContentPrepTool
    if ($null -eq $toolSettings) { $issues.Add('Configuration error: the ContentPrepTool section is missing.'); return [string[]]$issues.ToArray() }
    foreach ($key in @('Version', 'DownloadUrl', 'Sha256', 'StoragePath')) {
        if ([string]::IsNullOrWhiteSpace([string]$toolSettings.$key)) { $issues.Add("Configuration error: ContentPrepTool.$key is required.") }
    }
    return [string[]]$issues.ToArray()
} `
    -Build { param($Request) Invoke-IpIntuneWinAppUtilBuild -Request $Request }

Export-ModuleMember -Function @(
    'Register-IpPackagingEngine', 'Unregister-IpPackagingEngine', 'Get-IpPackagingEngine',
    'Test-IpPackageBuildRequest', 'Invoke-IpPackageBuild',
    'Get-IpContentPrepToolPath', 'Test-IpContentPrepTool', 'Install-IpContentPrepTool',
    'Get-IpIntuneWinAppUtilArguments', 'Invoke-IpIntuneWinAppUtil', 'Invoke-IpIntuneWinAppUtilBuild'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAib3lv3ISexBsG
# 0YtQ36u3w6c2iZt+lgD5L1+sgvYCqaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKhwkB2CN00moCaSG55MVj0fnETp2bj+8GU44SlesFn8MA0GCSqG
# SIb3DQEBAQUABIIBgBcbpRU74uBnrx130n6TrCwfwlQwoaf2madcGcAmchyZYktO
# IFp2WAPZsU8S6iBJ2jWLR19RPx9JAMrU0R36MWoYSUM+nhJexWYMiepKVrSuxhbY
# FgS1mmu4ck1su86pBab2/OQDvYfOZKGvmN1ntvo1afYi7GHo+sVxgrMMN9cWgu6G
# 0AzwJ752Rm8aF+8uc88iOKDYFRsfxzII8tSvixGtQ4dx+IQgoTSFHCZ6rXSDQhq0
# cvjY2r5+M2rsAoFppA8xeX01h/RCwjyHDjiONknmI7G1KyzADGoTNhUPqopucw4z
# DzrXhK2c5XlF9eP6NVGwDje1f30Q7JJA8W55bM75QVKJ4v7DTIpLCGvZrm9i4AAH
# AlEpzYMJSrBvdpHnEveIWfhWm7GseaFmwr/V6zRammaeIOlLeefhgnX13zrqI1B3
# DZYUYC7Xeqx1wNaoI+2q1yCi3bIoX7XFb6GVQ+ZOIkfRZ+vR4yhO6mXM+xCvWDiF
# KNHWk37hjX+7Zh5KFaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzVaMC8GCSqGSIb3DQEJBDEiBCCr9PtDIjU1vUnpzfC2HLCMhE7nOcsp3lhG//Zx
# BZK7NzANBgkqhkiG9w0BAQEFAASCAgC2dlgW5S2++AOrTckeXrY9tFvOWNkaH/9q
# 7dc0WeKa2LfbpZ5QRVX7QOPN/zHHNAuFh1YHGKjgA585rKT19LamXOn/BzNVRQzn
# 0jQv8G0mHijhlvXhyJsdSPGM4/smU6t8aDXm0jt1pCXR320RxMfCVDDTn+FbNNEm
# 9kZ+ix6JndBctG7GPki5Ukanjbr9l2XqkHoxB+AFdO9WXiWaadhgRhpr1JtoCjrU
# 7OvAaiBPo/K0Ksaeca9ATiG/SJy4ySVNuMJ5V+CnyHM/AV8mf2MScSJmwcTfZr5R
# TtxcPquvhmx/sv62iCB6ib2u0mFJBoSHCaIE+DybBKURE0dyO9nvUiayGv1ZDtuu
# MamoGxqXM4WbSOjGO3mCSuI72LXIkumXOxN+WE6Ngj/0J9TGz2yxXt6JC0tSo//V
# UEG50kNzpWNfvIvqPmP+BOemIR4B9v+cDR91fDJ8RKO2mItPlusUZc9BlkySR69J
# pCiqijjkwrvC3jB88374hfmkTSFbDrrS0K6QcXqFIOpNGQLjNOs1U7xjOBnPnKcI
# CguVAcyecgL9s/Y07SVhaDGhhA4AIGWI7hsmf9b4vQIRmv6tUmXWMAQY4EpL8x6H
# IM/kMHDFouk/i+KxkUl7baWu0IYe6xMU7gNvDldDWyhub9lCfPgUFljH567hF4/0
# pi4i2lMaKw==
# SIG # End signature block
