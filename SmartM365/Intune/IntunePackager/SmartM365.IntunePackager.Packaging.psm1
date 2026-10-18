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
