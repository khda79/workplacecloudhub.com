<#
.SYNOPSIS
    Shared ImportExcel startup health and exact-version missing-file repair.
.VERSION
    1.0.1
#>
# Shared by SmartM365.Core and its Windows PowerShell 5.1 compatibility module.
# Repair only missing files, from a verified copy of the exact installed version.
function Write-SmartM365ImportExcelHealthLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ((Get-Command WriteLog -ErrorAction SilentlyContinue) -and (Get-Variable LogTextFile -Scope Global -ErrorAction SilentlyContinue)) { WriteLog -Message $Message -Level $Level }
    else { Write-Verbose $Message -Verbose }
}

function Get-SmartM365ImportExcelCandidates {
    [CmdletBinding()]
    param([string[]]$ModuleRoots, [switch]$IncludeReferenceRoots)
    if (-not $ModuleRoots) {
        $ModuleRoots = @($env:PSModulePath -split [IO.Path]::PathSeparator)
        $documents = [Environment]::GetFolderPath('MyDocuments')
        foreach ($base in @($documents, $env:ProgramFiles)) {
            if (-not $IncludeReferenceRoots) { continue }
            if ($base) {
                $ModuleRoots += Join-Path $base 'PowerShell/Modules'
                $ModuleRoots += Join-Path $base 'WindowsPowerShell/Modules'
            }
        }
    }
    $seen = @{}
    foreach ($root in @($ModuleRoots | Select-Object -Unique)) {
        if (-not $root) { continue }
        $moduleRoot = Join-Path $root 'ImportExcel'
        if (-not (Test-Path -LiteralPath $moduleRoot -PathType Container)) { continue }
        $bases = @($moduleRoot) + @(Get-ChildItem -LiteralPath $moduleRoot -Directory -ErrorAction Stop | ForEach-Object FullName)
        foreach ($base in $bases) {
            $manifest = Join-Path $base 'ImportExcel.psd1'
            if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
                $payload = @(Get-ChildItem -LiteralPath $base -File -Force -ErrorAction Stop)
                if ($payload.Count -gt 0) { throw "Incomplete ImportExcel installation has no manifest: $base. Manual version qualification required; no reinstall attempted." }
                continue
            }
            $manifest = [IO.Path]::GetFullPath($manifest)
            if ($seen.ContainsKey($manifest)) { continue }
            $seen[$manifest] = $true
            # Data-only parsing also discovers installations with missing root modules.
            $data = Import-PowerShellDataFile -LiteralPath $manifest -ErrorAction Stop
            if (-not $data.ModuleVersion -or -not $data.GUID) { throw "Invalid ImportExcel manifest: $manifest" }
            [pscustomobject]@{ Path = $manifest; ModuleBase = [IO.Path]::GetDirectoryName($manifest); Version = [version]$data.ModuleVersion; Guid = [guid]$data.GUID }
        }
    }
}

function Sort-SmartM365ImportExcelCandidates {
    param([object[]]$Candidates)
    # PS5 Sort-Object is not stable for equal versions. Keep module-path priority.
    $ordered = @(for ($index = 0; $index -lt $Candidates.Count; $index++) {
        [pscustomobject]@{ Candidate=$Candidates[$index]; Priority=$index }
    })
    $ordered | Sort-Object @{Expression={$_.Candidate.Version};Descending=$true}, Priority | ForEach-Object Candidate
}

function Get-SmartM365ImportExcelExecutionPolicy {
    [CmdletBinding()]
    param()
    # Load the built-in module before any Exchange snap-in can affect type data.
    # Never force a reload, remove type data, or change the effective policy.
    try {
        $securityManifest = Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Security/Microsoft.PowerShell.Security.psd1'
        Import-Module -Name $securityManifest -Global -ErrorAction Stop
        $policy = [string](Microsoft.PowerShell.Security\Get-ExecutionPolicy -ErrorAction Stop)
        if ($policy -notin @('Restricted','AllSigned','RemoteSigned','Unrestricted','Bypass')) {
            throw 'The effective execution policy could not be determined unambiguously.'
        }
        return $policy
    } catch {
        $message = 'ImportExcel host environment check failed: Microsoft.PowerShell.Security or Get-ExecutionPolicy is unavailable. Initialize this dependency before loading the Exchange snap-in. No ImportExcel repair, installation or policy change was attempted. ' + $_.Exception.Message
        $exception = New-Object System.InvalidOperationException($message, $_.Exception)
        $record = New-Object Management.Automation.ErrorRecord($exception, 'SmartM365.ImportExcel.HostEnvironment', [Management.Automation.ErrorCategory]::ResourceUnavailable, $PSHOME)
        $PSCmdlet.ThrowTerminatingError($record)
    }
}

function Test-SmartM365ImportExcelLoad {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ManifestPath, [ValidateRange(1,60)][int]$TimeoutSeconds = 45)
    # An isolated, same-edition process prevents a failed EPPlus load contaminating the collector.
    # No policy bypass: inherit the process policy and preserve the effective parent policy.
    $policy = Get-SmartM365ImportExcelExecutionPolicy
    $engineName = if ($PSVersionTable.PSVersion.Major -ge 6) { 'pwsh.exe' } else { 'powershell.exe' }
    $engine = Join-Path $PSHOME $engineName
    $encodedPath = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($ManifestPath)))
    $body = @'
$ErrorActionPreference = 'Stop'
$manifestPath = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__PATH__'))
try {
    $module = Import-Module -Name $manifestPath -PassThru -ErrorAction Stop -WarningAction SilentlyContinue
    $expectedBase = [IO.Path]::GetDirectoryName($manifestPath)
    $loaded = @($module | Where-Object { $_.Name -eq 'ImportExcel' -and $_.ModuleBase -ieq $expectedBase })
    if ($loaded.Count -ne 1) { throw 'ImportExcel loaded from an unexpected location.' }
    foreach ($name in @('Import-Excel','Export-Excel','Open-ExcelPackage','Close-ExcelPackage','Get-ExcelSheetInfo')) {
        $command = Get-Command -Name $name -Module ImportExcel -ErrorAction Stop
        if ($command.Module.ModuleBase -ine $expectedBase) { throw "Unexpected command provenance: $name" }
    }
    $result = @{ Success=$true; Error=''; ErrorId=''; Version=[string]$loaded[0].Version }
} catch {
    $result = @{ Success=$false; Error=$_.Exception.Message; ErrorId=$_.FullyQualifiedErrorId; Version='' }
}
[Console]::WriteLine('SMARTIMPORTEXCEL_RESULT:' + ($result | ConvertTo-Json -Compress))
if (-not $result.Success) { exit 1 }
'@
    $body = $body.Replace('__PATH__', $encodedPath)
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $engine
    $start.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy ' + $policy + ' -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($body))
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        if (-not $process.Start()) { throw 'ImportExcel load probe could not start.' }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $process.Kill()
            throw 'ImportExcel load probe timed out; no automatic repair attempted.'
        }
        $output = $stdout.GetAwaiter().GetResult()
        $errors = $stderr.GetAwaiter().GetResult()
        $lines = @($output -split '\r?\n' | Where-Object { $_.StartsWith('SMARTIMPORTEXCEL_RESULT:') })
        if ($lines.Count -ne 1) { throw "ImportExcel probe returned no unambiguous result: $errors" }
        $result = $lines[0].Substring('SMARTIMPORTEXCEL_RESULT:'.Length) | ConvertFrom-Json
        if ($result.Success -and $process.ExitCode -ne 0) { throw 'ImportExcel probe exit code conflicts with its result.' }
        return $result
    } finally { $process.Dispose() }
}

function Get-SmartM365ImportExcelTree {
    param([Parameter(Mandatory)][string]$ModuleBase)
    $base = [IO.Path]::GetFullPath($ModuleBase).TrimEnd('\','/')
    # Cloud files can carry reparse flags without being links. Refuse actual links
    # and unavailable files, without rejecting a healthy hydrated OneDrive copy.
    foreach ($item in @((Get-Item -LiteralPath $base -Force -ErrorAction Stop)) + @(Get-ChildItem -LiteralPath $base -Recurse -Force -ErrorAction Stop)) {
        if (($item.PSObject.Properties['LinkType'] -and $item.LinkType) -or (($item.Attributes -band [IO.FileAttributes]::Offline) -ne 0)) { throw "ImportExcel repair refuses links or unavailable files: $($item.FullName)" }
        if ($item.PSIsContainer -or $item.Name -eq 'PSGetModuleInfo.xml') { continue }
        $relative = $item.FullName.Substring($base.Length + 1)
        [pscustomobject]@{ RelativePath=$relative; Path=$item.FullName; Hash=(Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256 -ErrorAction Stop).Hash }
    }
}

function Repair-SmartM365ImportExcelMissingFiles {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Target, [Parameter(Mandatory)]$Reference, [Parameter(Mandatory)][string]$RepairRoot)
    if ($Target.Version -ne $Reference.Version -or $Target.Guid -ne $Reference.Guid) { throw 'ImportExcel repair requires the exact same version and module GUID.' }
    foreach ($candidate in @($Target,$Reference)) {
        if ([IO.Path]::GetFullPath($candidate.Path) -ine (Join-Path ([IO.Path]::GetFullPath($candidate.ModuleBase)) 'ImportExcel.psd1')) { throw 'ImportExcel repair descriptor has an inconsistent manifest path.' }
        $base = [IO.Path]::GetFullPath($candidate.ModuleBase).TrimEnd('\','/')
        $repair = [IO.Path]::GetFullPath($RepairRoot).TrimEnd('\','/')
        if ($repair -ieq $base -or $repair.StartsWith($base + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'ImportExcel backups/staging must be outside module installations.' }
    }
    $referenceFiles = @(Get-SmartM365ImportExcelTree -ModuleBase $Reference.ModuleBase)
    $targetFiles = if (Test-Path -LiteralPath $Target.ModuleBase) { @(Get-SmartM365ImportExcelTree -ModuleBase $Target.ModuleBase) } else { @() }
    $targetMap = @{}
    foreach ($file in $targetFiles) { $targetMap[$file.RelativePath] = $file }
    $missing = @()
    foreach ($file in $referenceFiles) {
        if ($targetMap.ContainsKey($file.RelativePath)) {
            if ($targetMap[$file.RelativePath].Hash -ne $file.Hash) { throw "ImportExcel repair refused: existing file differs from reference: $($file.RelativePath). No files overwritten." }
        } else { $missing += $file }
    }
    if ($missing.Count -eq 0) { throw 'ImportExcel load failure is not explained by missing files; automatic repair refused.' }
    # Verify access before the backup or any repair, without overwriting an existing file.
    $null = New-Item -ItemType Directory -Path $Target.ModuleBase -Force -ErrorAction Stop
    $accessPath = Join-Path $Target.ModuleBase ('.smartm365-access-' + [guid]::NewGuid().ToString('N'))
    $access = [IO.File]::Open($accessPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $access.Dispose()
    Remove-Item -LiteralPath $accessPath -ErrorAction Stop
    $backup = Join-Path $RepairRoot ('repair-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $backup -ErrorAction Stop
    foreach ($file in $targetFiles) {
        $saved = Join-Path (Join-Path $backup 'before') $file.RelativePath
        $null = New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($saved)) -Force -ErrorAction Stop
        [IO.File]::Copy($file.Path, $saved, $false)
        if ((Get-FileHash -LiteralPath $saved -Algorithm SHA256).Hash -ne $file.Hash) { throw 'ImportExcel backup verification failed.' }
    }
    $receipt = [ordered]@{ Version=[string]$Target.Version; Target=$Target.Path; Reference=$Reference.Path; StartedUtc=[DateTime]::UtcNow.ToString('o'); Existing=$targetFiles; Missing=$missing; State='BackedUp' }
    $receiptPath = Join-Path $backup 'repair.json.txt'
    $receipt | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $receiptPath -Encoding UTF8 -ErrorAction Stop
    $created = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($file in $missing) {
            $destination = Join-Path $Target.ModuleBase $file.RelativePath
            $null = New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($destination)) -Force -ErrorAction Stop
            # Exclusive create: a concurrent writer's file is never replaced.
            $inputStream = [IO.File]::OpenRead($file.Path)
            $outputStream = $null
            try {
                $outputStream = [IO.File]::Open($destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
                $created.Add([pscustomobject]@{ Path=$destination; Hash=$file.Hash })
                $inputStream.CopyTo($outputStream)
            } finally {
                if ($outputStream) { $outputStream.Dispose() }
                $inputStream.Dispose()
            }
            if ((Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash -ne $file.Hash) { throw 'ImportExcel restored file hash mismatch.' }
        }
        # Verify the entire reference again; a writer/source change during copying is fatal.
        foreach ($file in $referenceFiles) {
            if ((Get-FileHash -LiteralPath (Join-Path $Target.ModuleBase $file.RelativePath) -Algorithm SHA256).Hash -ne $file.Hash) { throw 'ImportExcel installation changed during repair.' }
        }
        $probe = Test-SmartM365ImportExcelLoad -ManifestPath $Target.Path
        if (-not $probe.Success) { throw "Repaired ImportExcel still cannot load: $($probe.Error)" }
        $receipt.State = 'Verified'
        $receipt | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $receiptPath -Encoding UTF8 -ErrorAction Stop
        Write-SmartM365ImportExcelHealthLog "ImportExcel missing-file repair verified; version=$($Target.Version); restored=$($missing.Count); backup=$backup."
        return $backup
    } catch {
        $repairError = $_
        foreach ($file in $created) {
            # Only owned, unchanged files are rolled back. No recursive removal or prior-file deletion.
            if ((Test-Path -LiteralPath $file.Path -PathType Leaf) -and (Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -eq $file.Hash) {
                Remove-Item -LiteralPath $file.Path -ErrorAction Stop
            }
        }
        $receipt.State = 'Failed'
        $receipt | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $receiptPath -Encoding UTF8 -ErrorAction Stop
        throw "ImportExcel repair failed; existing files preserved; backup=$backup. $($repairError.Exception.Message)"
    }
}

function Initialize-SmartM365ImportExcel {
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser','AllUsers')][string]$Scope = 'CurrentUser',
        [string]$Repository = 'PSGallery',
        [switch]$InstallIfMissing,
        [switch]$SkipRepositoryRepair,
        [switch]$CheckOnly,
        [string]$RequiredModuleRoot,
        [string]$ManifestPath,
        [string]$RepairRoot = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'SmartM365/ModuleRepairs')
    )
    # Fail before candidate discovery, installation, promotion or repair.
    $null = Get-SmartM365ImportExcelExecutionPolicy
    if ($Scope -eq 'AllUsers' -and -not $RequiredModuleRoot) {
        $flavor = if ($PSVersionTable.PSVersion.Major -ge 6) { 'PowerShell' } else { 'WindowsPowerShell' }
        $RequiredModuleRoot = Join-Path $env:ProgramFiles ($flavor + '/Modules/ImportExcel')
    }
    $candidates = @(Sort-SmartM365ImportExcelCandidates -Candidates @(Get-SmartM365ImportExcelCandidates))
    $references = @(Sort-SmartM365ImportExcelCandidates -Candidates @(Get-SmartM365ImportExcelCandidates -IncludeReferenceRoots))
    $selected = if ($ManifestPath) { @($references | Where-Object Path -IEQ ([IO.Path]::GetFullPath($ManifestPath))) | Select-Object -First 1 }
        elseif ($RequiredModuleRoot) { @($references | Where-Object { $_.ModuleBase.StartsWith([IO.Path]::GetFullPath($RequiredModuleRoot).TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) }) | Select-Object -First 1 }
        else { $candidates | Select-Object -First 1 }
    if ($ManifestPath -and -not $selected) { throw 'The requested ImportExcel manifest was not discovered; no alternative installation selected.' }
    if ($selected -and $RequiredModuleRoot -and -not $selected.ModuleBase.StartsWith([IO.Path]::GetFullPath($RequiredModuleRoot).TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
        throw 'The requested ImportExcel manifest is outside the required installation scope; no CurrentUser fallback permitted.'
    }
    if ($CheckOnly -and -not $selected) { throw 'ImportExcel check-only: requested installation is absent; no installation or promotion attempted.' }
    if (-not $selected -and $RequiredModuleRoot -and $references.Count -gt 0) {
        # Preserve ProxyAddresses' AllUsers promotion, but only from a verified copy.
        $reference = $null
        foreach ($candidate in $references) {
            if ($candidate.Version -ne $references[0].Version) { continue }
            if ((Test-SmartM365ImportExcelLoad $candidate.Path).Success) { $reference = $candidate; break }
        }
        if (-not $reference) { throw 'No healthy highest-version ImportExcel copy available for AllUsers promotion.' }
        $base = Join-Path $RequiredModuleRoot ([string]$reference.Version)
        $selected = [pscustomobject]@{ Path=(Join-Path $base 'ImportExcel.psd1'); ModuleBase=$base; Version=$reference.Version; Guid=$reference.Guid }
    }
    if (-not $selected) {
        if (-not $InstallIfMissing) { throw 'ImportExcel is not installed; no collector or workbook processing started.' }
        if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) { throw 'ImportExcel is missing and Install-Module is unavailable.' }
        Write-SmartM365ImportExcelHealthLog "ImportExcel missing; initial installation from $Repository with Scope=$Scope." 'WARNING'
        Install-Module -Name ImportExcel -Scope $Scope -Repository $Repository -Force -AllowClobber -ErrorAction Stop
        # Do not install twice; enforce the requested root on the read-back.
        return Initialize-SmartM365ImportExcel -Scope $Scope -Repository $Repository -RequiredModuleRoot $RequiredModuleRoot -SkipRepositoryRepair:$SkipRepositoryRepair -RepairRoot $RepairRoot
    }
    $probe = Test-SmartM365ImportExcelLoad $selected.Path
    $backup = ''
    if (-not $probe.Success) {
        if ($CheckOnly) { throw "ImportExcel check-only load failed; no repair attempted: $($probe.Error)" }
        if ($probe.ErrorId -match 'PSSecurityException|UnauthorizedAccess|UnauthorizedAccessException' -or $probe.Error -match 'not digitally signed|running scripts is disabled|execution policy|Access.*denied') {
            throw "ImportExcel security/access failure; policy and permissions were not changed: $($probe.Error)"
        }
        $hash = [Security.Cryptography.SHA256]::Create()
        try { $key = ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($selected.Path.ToLowerInvariant())))).Replace('-','') } finally { $hash.Dispose() }
        $mutex = New-Object Threading.Mutex($false, ('Global\SmartM365-ImportExcel-' + $key))
        $locked = $false
        try {
            try { $locked = $mutex.WaitOne(45000) } catch [Threading.AbandonedMutexException] { $locked = $true }
            if (-not $locked) { throw 'ImportExcel repair is locked by another process; retry after that run.' }
            $probe = Test-SmartM365ImportExcelLoad $selected.Path
            if (-not $probe.Success) {
                $reference = $null
                foreach ($candidate in $references) {
                    if ($candidate.Path -ieq $selected.Path -or $candidate.Version -ne $selected.Version -or $candidate.Guid -ne $selected.Guid) { continue }
                    if ((Test-SmartM365ImportExcelLoad $candidate.Path).Success) { $reference = $candidate; break }
                }
                if (-not $reference) {
                    if ($SkipRepositoryRepair) { throw 'No healthy same-version local ImportExcel copy. Repository repair is disabled for this mode.' }
                    # Exact version only. No repository trust, provider or execution-policy changes.
                    $repositoryInfo = Get-PSRepository -Name $Repository -ErrorAction Stop
                    if ($repositoryInfo.InstallationPolicy -ne 'Trusted') { throw "Repository '$Repository' is not trusted; automatic download refused." }
                    if (-not (Get-Command Save-Module -ErrorAction SilentlyContinue)) { throw 'Save-Module is unavailable; install package-management prerequisites separately.' }
                    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) { throw 'NuGet provider is unavailable; no automatic provider installation attempted.' }
                    $staging = Join-Path $RepairRoot ('staging-' + [guid]::NewGuid().ToString('N'))
                    $null = New-Item -ItemType Directory -Path $staging -ErrorAction Stop
                    Save-Module -Name ImportExcel -RequiredVersion ([string]$selected.Version) -Repository $Repository -Path $staging -ErrorAction Stop
                    $reference = @(Get-SmartM365ImportExcelCandidates -ModuleRoots @($staging) | Where-Object { $_.Version -eq $selected.Version -and $_.Guid -eq $selected.Guid }) | Select-Object -First 1
                    if (-not $reference -or -not (Test-SmartM365ImportExcelLoad $reference.Path).Success) { throw 'Downloaded exact-version ImportExcel reference failed validation.' }
                }
                $backup = Repair-SmartM365ImportExcelMissingFiles -Target $selected -Reference $reference -RepairRoot $RepairRoot
            }
        } finally {
            if ($locked) { $mutex.ReleaseMutex() }
            $mutex.Dispose()
        }
    }
    # Pin the checked path. Never let name-based autoload select the broken user copy again.
    Import-Module -Name $selected.Path -Global -ErrorAction Stop
    $loaded = @(Get-Module ImportExcel | Where-Object { $_.ModuleBase -ieq $selected.ModuleBase -and $_.Version -eq $selected.Version })
    if ($loaded.Count -ne 1) { throw 'ImportExcel parent-process import provenance verification failed.' }
    foreach ($name in @('Import-Excel','Export-Excel','Open-ExcelPackage','Close-ExcelPackage','Get-ExcelSheetInfo')) {
        $command = Get-Command -Name $name -Module ImportExcel -ErrorAction Stop
        if ($command.Module.ModuleBase -ine $selected.ModuleBase) { throw "ImportExcel parent command provenance verification failed: $name" }
    }
    Write-SmartM365ImportExcelHealthLog "ImportExcel startup load verified; version=$($selected.Version); manifest=$($selected.Path)."
    [pscustomobject]@{ Path=$selected.Path; Version=$selected.Version; Repaired=([bool]$backup); BackupPath=$backup }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBLtxSohymt7Eb
# ZXTP8M+7r1FJDRjd7SDGSTKoSxdgTqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKdOZvtboAKqkdfSFuWPA7ezrA9JuXHGCJ5tA3o/kcKGMA0GCSqG
# SIb3DQEBAQUABIIBgA66w/lcfsCPZ9437ryv97zHLUoxpkdlAV+cg9ilG3zREmMd
# YEYYrFelnx4vFt0eNzaeqApxB0gKBAHsz06HDRq5LW+VF+7a86ivmK+5rj4FSlAB
# Iad5k5eaE1IbYQ6aSzTU+OL02CgSqIBe4QuuPkoWQVUhHWYqPO5tAAakOg1JCQDe
# 8xk7J6Gy0oGzCDuMnwN0XUnoGzySfSBwa8+jmeqezbE9oKyXoRgTavP+4aNT7DJ9
# tMxp2sSokVGOwfAA6p26bJhekj4wntJSbOQzKg5MskQpinnhuWhyxdB/7BQ3lgs2
# 9v57eSvB6BQwqySsmyN0hKbnTQRUfHLqDONH1JcGNjI3ZLNO9Yo5DNQjKB7Ib4Yb
# o62OuMiRtegYBYzEbg/5J3fN2uImd4vZK7pO4G+HWZKWzX5uLP9SbTHWkkH6pvre
# 3Psw2pBo2rHHNbWnOUIXD4IYSv3/uDuZ/aV/i3uWzKWvQ5yEeqr1vVUJuUtI8REn
# vJAPXKqPNHbIfDno76GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMDA4
# NDlaMC8GCSqGSIb3DQEJBDEiBCCKwJiZqmU8i71sys7fIQnd1xxHl+00owBzeITw
# n6vjgTANBgkqhkiG9w0BAQEFAASCAgB3cjnPqJYWn8/HaaEidZrLn2mWsVN66AG7
# GnnpqMx1NBr6tvnOO9ugJrgKHc75rzG6b38sv4QXyzZ9WVDnim6Xh0daWdkOIhSe
# fls0XPYUxCLz9L52ZgQfVAVBiGVJPtLUeS/NGRcUO4AOdZL2jgoPZJE6xvpw1mZJ
# TYt0n7ghaK7vRXlKrJTS0r2Qtc7mkyYSLW25KstpDLMUPeMGz9EqvqvgGDKEXrLr
# REitWStAGvJA0WCcuYykR0o3AC4wNxkRzHK2DcKfUyz5hwyJYcQK/fc/zvHLd0bY
# MZbYVeF+ytjawktJwDlXG8ypvPTI62m72V9RMI6Qor/+Pmn7Q4mi2O0PVwb7WJHR
# jpwt0B4W1WOBtjFGgZTKVbAgfYbesU8bmLMh1EWAaT0ruoUKcRPcemn2pmaBVD1L
# SSSIA/u25qKTE4K6AiAL5mM8zV59mdErF0xHNS/NDF045Ij4gpbfoSIJ3cWZbYLc
# r9WKcWNmu/1MblIsrGUJgrdnwK1ykl/cgrYuNyiQHEdhWsYt7+OwfeC4Nk/TSsBr
# A+TsJKNF2SK9bapXeYR9iJGh5VfR3zLhDvpa3w0kJ2IjJEgGCDcZe5YJTfzh117L
# rg/jHbRY7PD50+GXw89eKrziVRn+E9AVIuix5Ug+mqZZQtV3ljMdwLD0rBqLKifD
# rjPKwr2DQw==
# SIG # End signature block
