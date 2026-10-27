# JSON transport names are independent from the JSON payload. Import has no side effects.
Set-StrictMode -Version 2.0
# PS7 can launch Windows PowerShell with a module path that excludes its native modules.
# Load this engine's trusted built-in module, never a caller-provided module search result.
if ($PSVersionTable.PSEdition -eq 'Desktop') {
    Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
}

function Get-SmartM365JsonNames {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.EndsWith('.json.txt', [StringComparison]::OrdinalIgnoreCase)) {
        $legacy = $full.Substring(0, $full.Length - 4)
    } elseif ($full.EndsWith('.json', [StringComparison]::OrdinalIgnoreCase)) {
        $legacy = $full
    } else { throw "Not an owned JSON transport filename: $Path" }
    [pscustomobject]@{ Legacy = $legacy; Preferred = $legacy + '.txt'; Lock = $legacy + '.transport.lock'; Journal = $legacy + '.migration.log' }
}

function ConvertTo-SmartM365JsonQualifiedRoot {
    param([Parameter(Mandatory)][string]$Path)
    if ($Path -notmatch '^[A-Za-z]:\\[^\\]' -or $Path -match '[*?{}]' -or
        $Path.Substring(2).Contains(':') -or $Path.Contains('/') -or
        @($Path.Split('\') | Where-Object { $_ -in @('.', '..') -or ($_ -and $_ -match '[. ]$') }).Count) {
        throw 'Qualified local data roots must be exact absolute local directory paths.'
    }
    [IO.Path]::GetFullPath($Path).TrimEnd('\')
}

function Get-SmartM365JsonPathNode {
    param([Parameter(Mandatory)][string]$Path)
    if (-not ('SmartM365.JsonTransport.PathTagsV1' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
namespace SmartM365.JsonTransport {
    public static class PathTagsV1 {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct FindData {
            public uint Attributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME Creation, Access, Write;
            public uint SizeHigh, SizeLow, Tag, Reserved;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string Name;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string Alternate;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr FindFirstFileW(string path, out FindData data);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FindClose(IntPtr handle);
        public static FindData? Read(string path) {
            FindData data;
            IntPtr handle = FindFirstFileW(path, out data);
            if (handle == new IntPtr(-1)) {
                int error = Marshal.GetLastWin32Error();
                if (error == 2 || error == 3) return null;
                throw new Win32Exception(error, "Cannot inspect JSON path: " + path);
            }
            try { return data; } finally { FindClose(handle); }
        }
    }
}
'@ -ErrorAction Stop
    }
    [SmartM365.JsonTransport.PathTagsV1]::Read($Path)
}

function Test-SmartM365JsonCloudTag {
    param([uint32]$Tag)
    # Only IO_REPARSE_TAG_CLOUD and CLOUD_1..F; never a name-surrogate tag.
    ($Tag -band [uint32]4294905855) -eq [uint32]2415919130
}

function Assert-SmartM365JsonQualifiedChain {
    param([string]$Path, [string]$CloudRoot, [string]$JunctionRoot = '', [string]$JunctionTarget = '')
    $cursor = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    $volumeRoot = [IO.Path]::GetPathRoot($cursor).TrimEnd('\')
    while ($cursor -and $cursor -ne $volumeRoot) {
        $node = Get-SmartM365JsonPathNode $cursor
        if ($null -ne $node -and ($node.Attributes -band [uint32]1024)) {
            if ($cursor.Equals($JunctionRoot, [StringComparison]::OrdinalIgnoreCase)) {
                if ($node.Tag -ne [uint32]2684354563) { throw "Approved data alias is not a junction: $cursor" }
                $targets = @((Get-Item -LiteralPath $cursor -Force -ErrorAction Stop).Target)
                if ($targets.Count -ne 1 -or
                    -not (ConvertTo-SmartM365JsonQualifiedRoot ([string]$targets[0])).Equals($JunctionTarget, [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Approved data junction target changed: $cursor"
                }
            } else {
                $inCloudChain = $CloudRoot -and ($cursor.Equals($CloudRoot, [StringComparison]::OrdinalIgnoreCase) -or
                    $cursor.StartsWith($CloudRoot + '\', [StringComparison]::OrdinalIgnoreCase) -or
                    $CloudRoot.StartsWith($cursor + '\', [StringComparison]::OrdinalIgnoreCase))
                if (-not $inCloudChain -or -not (Test-SmartM365JsonCloudTag $node.Tag)) {
                    throw "Unqualified reparse point in JSON path: $cursor"
                }
            }
        }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
}

function Assert-SmartM365JsonQualifiedDataPath {
    param([string]$Path, [object[]]$Qualifications)
    $full = [IO.Path]::GetFullPath($Path)
    $matches = @($Qualifications | Where-Object {
        $full.StartsWith($_.Root + '\', [StringComparison]::OrdinalIgnoreCase) -or
        $full.StartsWith($_.Target + '\', [StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -gt 1) { throw 'Overlapping qualified local data roots are not permitted.' }
    if ($matches.Count -eq 0) {
        Assert-SmartM365JsonQualifiedChain -Path $full -CloudRoot ''
        return
    }
    $qualification = $matches[0]
    $root = $qualification.Root; $target = $qualification.Target
    $targetNode = Get-SmartM365JsonPathNode $target
    if ($null -eq $targetNode -or -not ($targetNode.Attributes -band [uint32]16)) { throw 'Qualified data target directory is missing.' }
    Assert-SmartM365JsonQualifiedChain -Path $target -CloudRoot $target
    if ($root.Equals($target, [StringComparison]::OrdinalIgnoreCase)) {
        Assert-SmartM365JsonQualifiedChain -Path $full -CloudRoot $target
        return
    }
    if (-not $root.Equals($target, [StringComparison]::OrdinalIgnoreCase)) {
        $rootNode = Get-SmartM365JsonPathNode $root
        if ($null -eq $rootNode -or $rootNode.Tag -ne [uint32]2684354563) { throw 'Qualified data alias junction is missing or changed.' }
        Assert-SmartM365JsonQualifiedChain -Path $root -CloudRoot '' -JunctionRoot $root -JunctionTarget $target
    }
    if ($full.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) {
        $physical = $target + $full.Substring($root.Length)
        Assert-SmartM365JsonQualifiedChain -Path $physical -CloudRoot $target
        Assert-SmartM365JsonQualifiedChain -Path $full -CloudRoot $root -JunctionRoot $root -JunctionTarget $target
    } else {
        Assert-SmartM365JsonQualifiedChain -Path $full -CloudRoot $target
    }
}

function Assert-SmartM365JsonUnlinkedPath {
    param([Parameter(Mandatory)][string]$Path, [switch]$AllowQualifiedDataRoot)
    if ($AllowQualifiedDataRoot) {
        $qualifications = @((Get-SmartM365JsonTransportPolicy).QualifiedLocalDataRoots)
        if ($qualifications.Count) {
            Assert-SmartM365JsonQualifiedDataPath -Path $Path -Qualifications $qualifications
            return
        }
    }
    $cursor = [IO.Path]::GetFullPath($Path)
    $volumeRoot = [IO.Path]::GetPathRoot($cursor).TrimEnd('\', '/')
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked path is not eligible for JSON migration: $cursor" }
        }
        if ($cursor.TrimEnd('\', '/') -eq $volumeRoot) { break }
        $parent = [IO.Path]::GetDirectoryName($cursor)
        if ($parent -eq $cursor) { break }
        $cursor = $parent
    }
}

function Get-SmartM365JsonReadPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$Optional)
    $names = Get-SmartM365JsonNames $Path
    # Test existence, not just Leaf: a directory at the preferred name must not enable fallback.
    foreach ($candidate in @($names.Preferred, $names.Legacy)) {
        if (Test-Path -LiteralPath $candidate -ErrorAction Stop) {
            if (-not (Test-Path -LiteralPath $candidate -PathType Leaf -ErrorAction Stop)) { throw "JSON path is not a file: $candidate" }
            return $candidate
        }
    }
    if (-not $Optional) { throw "JSON file missing: $($names.Preferred) (legacy: $($names.Legacy))" }
}

function Get-SmartM365JsonHash {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '') }
    finally { $sha.Dispose() }
}

function ConvertFrom-SmartM365JsonBytes {
    param([Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes, [scriptblock]$Validate)
    if ($Bytes.Length -eq 0) { throw 'Empty JSON payload.' }
    $memory = [IO.MemoryStream]::new($Bytes, $false)
    $reader = [IO.StreamReader]::new($memory, [Text.UTF8Encoding]::new($false, $true), $true)
    try { $content = $reader.ReadToEnd() } finally { $reader.Dispose(); $memory.Dispose() }
    if ([string]::IsNullOrWhiteSpace($content)) { throw 'Blank JSON payload.' }
    try { $document = ConvertFrom-Json -InputObject $content -ErrorAction Stop }
    catch {
        $_.Exception.Data['SmartM365JsonInvalidSyntax'] = $true
        throw
    }
    if ($null -eq $document -and $content.TrimStart().StartsWith('[')) { $document = @() }
    if ($null -eq $document) { throw 'Null JSON payload cannot represent a valid owned state.' }
    if ($Validate) { & $Validate $document | Out-Null }
    # Preserve array shape instead of PowerShell pipeline unrolling.
    return ,$document
}

function Read-SmartM365JsonDocument {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [scriptblock]$Validate)
    $selected = Get-SmartM365JsonReadPath $Path
    $bytes = [IO.File]::ReadAllBytes($selected)
    # No catch/fallback: an invalid preferred file must surface to the owner.
    $document = ConvertFrom-SmartM365JsonBytes -Bytes $bytes -Validate $Validate
    $names = Get-SmartM365JsonNames $Path
    $coexistence = 'Single'
    if ($selected -eq $names.Preferred -and (Test-Path -LiteralPath $names.Legacy)) {
        $oldBytes = [IO.File]::ReadAllBytes($names.Legacy)
        $coexistence = if ((Get-SmartM365JsonHash $oldBytes) -eq (Get-SmartM365JsonHash $bytes)) { 'Identical' } else { 'Conflict' }
        if ($coexistence -eq 'Conflict') {
            # A regenerated replacement may have committed just before interruption.
            # Only the durable, exact old/new hash pair can authorize that coexistence.
            $receipt = Get-SmartM365LastJsonMigrationEvent $names.Journal
            if ($receipt -and $receipt.Phase -in @('ReplacementPrepared','ReplacementPublished') -and $receipt.SHA256 -eq (Get-SmartM365JsonHash $bytes) -and $receipt.Detail -eq (Get-SmartM365JsonHash $oldBytes)) {
                $coexistence = 'PublishedReplacement'
            } else { throw "Conflicting JSON transport names: $($names.Legacy) and $selected" }
        }
    }
    [pscustomobject]@{ Path = $selected; Document = $document; SHA256 = (Get-SmartM365JsonHash $bytes); Coexistence = $coexistence }
}

function Enter-SmartM365JsonTransportLock {
    param([Parameter(Mandatory)][string]$Path, [int]$TimeoutSeconds = 10)
    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        try { return [IO.File]::Open($Path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch [IO.IOException] {
            if ([datetime]::UtcNow -ge $deadline) { throw "JSON transport lock unavailable: $Path" }
            Start-Sleep -Milliseconds 100
        }
    } while ($true)
}

function Write-SmartM365JsonMigrationEvent {
    param([string]$Journal, [string]$Owner, [string]$Phase, [string]$SHA256, [string]$Detail)
    $line = ([ordered]@{ Utc = [datetime]::UtcNow.ToString('o'); Owner = $Owner; Phase = $Phase; SHA256 = $SHA256; Detail = $Detail } | ConvertTo-Json -Compress) + [Environment]::NewLine
    if([IO.File]::Exists($Journal)){
        $check=[IO.File]::Open($Journal,'Open','Read','ReadWrite')
        try{if($check.Length -gt 0){$null=$check.Seek(-1,'End');if($check.ReadByte() -ne 10){$line=[Environment]::NewLine+$line}}}finally{$check.Dispose()}
    }
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($line)
    $stream = [IO.File]::Open($Journal, 'Append', 'Write', 'Read')
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
}

function Get-SmartM365LastJsonMigrationEvent {
    param([Parameter(Mandatory)][string]$Path)
    if(-not [IO.File]::Exists($Path)){return $null}
    $lines=@(Get-Content -LiteralPath $Path -Tail 32 -ErrorAction Stop)
    for($i=$lines.Count-1;$i -ge 0;$i--){
        if([string]::IsNullOrWhiteSpace($lines[$i])){continue}
        try{
            $event=ConvertFrom-Json -InputObject $lines[$i] -ErrorAction Stop
            if(-not $event.PSObject.Properties['Phase'] -or -not $event.PSObject.Properties['Owner']){throw 'Incomplete journal event.'}
            return $event
        }catch{Write-Warning "Incomplete migration journal tail preserved; checking preceding durable event: $Path"}
    }
    throw "Migration journal has no recoverable event: $Path"
}

function Move-SmartM365OwnedJsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][scriptblock]$Validate,
        [string[]]$CompatibleJournalOwners = @(),
        [switch]$QualifiedUnc,
        [switch]$RemoveIdenticalLegacy,
        [int]$LockTimeoutSeconds = 10
    )
    if ([IO.Path]::IsPathRooted($RelativePath)) { throw 'Migration requires an owner-relative filename.' }
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $names = Get-SmartM365JsonNames (Join-Path $Root $RelativePath)
    if (-not $names.Legacy.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { throw 'Migration path escapes the owner root.' }
    if ($names.Legacy.StartsWith('\\') -and -not $QualifiedUnc) { throw 'UNC migration requires qualification on the target share before activation.' }
    foreach ($path in @($names.Legacy, $names.Preferred, $names.Lock, $names.Journal)) { Assert-SmartM365JsonUnlinkedPath $path }
    $selected = Get-SmartM365JsonReadPath $names.Legacy -Optional
    if (-not $selected) { return [pscustomobject]@{ Status = 'Absent'; Path = $names.Preferred; SHA256 = '' } }
    $lock = Enter-SmartM365JsonTransportLock $names.Lock $LockTimeoutSeconds
    $journalOwnerRejected = $false
    try {
        $current = Read-SmartM365JsonDocument $names.Legacy -Validate $Validate
        $hash = $current.SHA256
        if ($current.Path -eq $names.Preferred) {
            if($current.Coexistence -eq 'PublishedReplacement'){throw 'Regenerated replacement must be completed by its producer before persistent migration.'}
            if ($current.Coexistence -eq 'Identical' -and $RemoveIdenticalLegacy) {
                Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'DeduplicatePrepared' $hash ''
                if ((Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash -ne $hash) { throw 'Legacy file changed during migration.' }
                [IO.File]::Delete($names.Legacy)
                Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $hash 'Identical legacy removed'
            } elseif ($current.Coexistence -eq 'Identical') {
                Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'PendingLegacy' $hash 'Publication/compatibility confirmation required'
                return [pscustomobject]@{ Status = 'PendingLegacy'; Path = $names.Preferred; SHA256 = $hash }
            } else {
                $last=Get-SmartM365LastJsonMigrationEvent $names.Journal
                if($last -and $last.Phase -ne 'Completed'){
                    if($last.Owner -ne $Owner -and $last.Owner -notin $CompatibleJournalOwners){$journalOwnerRejected=$true;throw 'Migration journal belongs to another owner.'}
                    Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $hash 'Resumed after rename'
                }
            }
            return [pscustomobject]@{ Status = 'Completed'; Path = $names.Preferred; SHA256 = $hash }
        }
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Prepared' $hash ''
        if ((Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash -ne $hash) { throw 'Legacy file changed before rename.' }
        # Same directory; never emulate failed rename with copy/delete, especially on UNC.
        [IO.File]::Move($names.Legacy, $names.Preferred)
        if ((Get-FileHash -LiteralPath $names.Preferred -Algorithm SHA256).Hash -ne $hash) { throw 'JSON integrity failure after rename.' }
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $hash ''
        [pscustomobject]@{ Status = 'Completed'; Path = $names.Preferred; SHA256 = $hash }
    } catch {
        $failure = $_
        if (-not $journalOwnerRejected) {
            try { Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Failed' '' $failure.Exception.Message } catch { Write-Warning 'Could not persist JSON migration failure journal.' }
        }
        throw $failure
    } finally { $lock.Dispose() }
}

function Resolve-SmartM365JsonQualifiedUncToken {
    param([Parameter(Mandatory)][string]$Value)
    # Only the deployment roots can be referenced: the SmartM365 folder holding this module and the
    # WorkspaceRootPath of the global file. Tenant overlays and script settings never qualify a share.
    foreach ($match in [regex]::Matches($Value, '\{\{([^}]*)\}\}')) {
        if ($match.Groups[1].Value -notin @('WorkspaceRootPath', 'SmartM365RootPath')) { throw "Unsupported qualified UNC token: $($match.Value)" }
    }
    $tokens = @{ SmartM365RootPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\') }
    if ($Value -match '\{\{WorkspaceRootPath\}\}') {
        $workspace = ''
        # The global file is read directly: resolving a qualification must not depend on the policy being resolved.
        $globalPath = Join-Path $PSScriptRoot '..\..\Config\SmartM365.global.local.json'
        if (Get-SmartM365JsonReadPath -Path $globalPath -Optional) {
            $globalConfig = (Read-SmartM365JsonDocument -Path $globalPath).Document
            if ($globalConfig.PSObject.Properties['WorkspaceRootPath']) { $workspace = [string]$globalConfig.WorkspaceRootPath }
        }
        if ([string]::IsNullOrWhiteSpace($workspace)) { $workspace = '{{SmartM365RootPath}}' }
        $tokens['WorkspaceRootPath'] = $workspace
    }
    $resolved = $Value
    for ($i = 0; $i -lt 5 -and $resolved -match '\{\{'; $i++) {
        foreach ($match in [regex]::Matches($resolved, '\{\{([A-Za-z0-9_.-]+)\}\}')) {
            if (-not $tokens.ContainsKey($match.Groups[1].Value)) { return $null }
            $resolved = $resolved.Replace($match.Value, [string]$tokens[$match.Groups[1].Value])
        }
    }
    if ($resolved -match '\{\{') { return $null }
    return $resolved
}

function Get-SmartM365JsonTransportPolicy {
    [CmdletBinding()]
    param()
    $path = Join-Path $PSScriptRoot '../../Config/SmartM365-JsonTransport.policy.psd1'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'JSON transport deployment policy missing. Incomplete deployment.' }
    $policy = Import-PowerShellDataFile -LiteralPath $path -ErrorAction Stop
    if ($policy.Mode -notin @('Readers', 'JsonText')) { throw 'Unsupported JSON transport deployment mode.' }
    $policy.QualifiedLocalDataRoots = @()
    # Deployment-specific qualifications are private data, never executable policy.
    $localPath = Join-Path $PSScriptRoot '../../Config/SmartM365-JsonTransport.policy.local.json'
    if (Get-SmartM365JsonReadPath -Path $localPath -Optional) {
        $local = (Read-SmartM365JsonDocument -Path $localPath).Document
        foreach ($property in $local.PSObject.Properties) {
            if ($property.Name -notin @('QualifiedUncRoots','QualifiedSharePointDrives','QualifiedLocalDataRoots')) { throw 'Unsupported local JSON transport policy setting.' }
        }
        foreach ($name in @('QualifiedUncRoots','QualifiedSharePointDrives')) {
            if (-not $local.PSObject.Properties[$name]) { continue }
            $values = @()
            foreach ($value in @($local.$name)) {
                if ($value -isnot [string] -or [string]::IsNullOrWhiteSpace($value) -or $value -match '[*?]') { throw 'Invalid qualified destination.' }
                if ($name -eq 'QualifiedUncRoots' -and $value -match '\{\{') {
                    # A tokenized root follows the deployment; on a deployment where it is not a UNC path it qualifies nothing.
                    $value = Resolve-SmartM365JsonQualifiedUncToken -Value $value
                    if ([string]::IsNullOrWhiteSpace($value) -or -not $value.StartsWith('\\')) { continue }
                }
                if ($name -eq 'QualifiedUncRoots' -and $value -notmatch '^\\\\[^\\/]+\\[^\\/]+(?:\\[^\\/]+)*\\?$') { throw 'Qualified UNC destination must be an exact absolute UNC path.' }
                if ($name -eq 'QualifiedUncRoots' -and @($value.Split('\') | Where-Object { $_ -in @('.','..') }).Count) { throw 'Qualified UNC destination cannot contain traversal segments.' }
                $values += $value
            }
            $policy[$name] = @($values | Select-Object -Unique)
        }
        if ($local.PSObject.Properties['QualifiedLocalDataRoots']) {
            $roots = @()
            foreach ($entry in @($local.QualifiedLocalDataRoots)) {
                if ($null -eq $entry -or $entry -isnot [pscustomobject] -or
                    @($entry.PSObject.Properties).Count -ne 2 -or
                    -not $entry.PSObject.Properties['Root'] -or -not $entry.PSObject.Properties['Target'] -or
                    $entry.Root -isnot [string] -or $entry.Target -isnot [string]) { throw 'Invalid qualified local data root entry.' }
                $roots += [pscustomobject]@{
                    Root = ConvertTo-SmartM365JsonQualifiedRoot $entry.Root
                    Target = ConvertTo-SmartM365JsonQualifiedRoot $entry.Target
                }
            }
            $policy.QualifiedLocalDataRoots = $roots
        }
    }
    return $policy
}

function Get-SmartM365JsonLockErrorCode {
    param([Parameter(Mandatory)][Exception]$Exception)
    for ($errorCause = $Exception; $null -ne $errorCause; $errorCause = $errorCause.InnerException) {
        $code = $errorCause.HResult -band 65535
        if ($code -in @(32, 33)) { return $code }
    }
    return 0
}

<#
.SYNOPSIS
Temporary file for an atomic replace: same folder (same volume), short name. A name built from the target
name plus a GUID adds about 40 characters and fails under MAX_PATH (260) with Windows PowerShell 5.1 and
the .NET Framework when the target already sits deep (launcher cache copies of the repository).
#>
function Get-SmartM365SiblingTemporaryPath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidatePattern('^\.[A-Za-z0-9]{1,10}$')][string]$Extension = '.tmp'
    )
    $parent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Path))
    [IO.Path]::Combine($parent, ('~sm{0}{1}' -f [guid]::NewGuid().ToString('N').Substring(0, 12), $Extension))
}

function Write-SmartM365JsonBytesAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes,
        [Parameter(Mandatory)][scriptblock]$Validate,
        [string]$ExpectedSHA256 = '',
        [switch]$RetireLegacyAfterPublication,
        [string]$Owner = '',
        [int]$LockTimeoutSeconds = 10,
        # Configuration enrichment can wait longer without changing other callers' budgets.
        [ValidateRange(0, 3600)][int]$ReplacementTimeoutSeconds = $LockTimeoutSeconds
    )
    $names = Get-SmartM365JsonNames $Path
    # Caller chooses the deployment-phase path. Serialization is never performed here.
    $target = [IO.Path]::GetFullPath($Path)
    foreach ($candidate in @($target, $names.Lock)) { Assert-SmartM365JsonUnlinkedPath $candidate -AllowQualifiedDataRoot }
    $null = ConvertFrom-SmartM365JsonBytes -Bytes $Bytes -Validate $Validate
    $lock = Enter-SmartM365JsonTransportLock $names.Lock $LockTimeoutSeconds
    $temporary = Get-SmartM365SiblingTemporaryPath -Path $target -Extension '.pending'
    $backup = Get-SmartM365SiblingTemporaryPath -Path $target -Extension '.previous'
    $restoreAttributes = $null
    $published = $false
    try {
        foreach ($candidate in @($target, $names.Lock, $temporary, $backup)) { Assert-SmartM365JsonUnlinkedPath $candidate -AllowQualifiedDataRoot }
        $legacyHash = ''
        if ($RetireLegacyAfterPublication) {
            if ($target -ne $names.Preferred -or [string]::IsNullOrWhiteSpace($Owner)) { throw 'Replacement retirement requires a preferred target and an explicit owner.' }
            foreach ($candidate in @($names.Legacy, $names.Journal)) { Assert-SmartM365JsonUnlinkedPath $candidate }
            if (Get-SmartM365JsonReadPath $Path -Optional) {
                $prior = Read-SmartM365JsonDocument -Path $Path -Validate $Validate
                if ($prior.Coexistence -eq 'PublishedReplacement') {
                    # Finish the already verified publication before preparing another one.
                    $oldHash = (Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash
                    $receipt = Get-SmartM365LastJsonMigrationEvent $names.Journal
                    if ($receipt.Owner -ne $Owner -or $receipt.Detail -ne $oldHash -or $receipt.SHA256 -ne $prior.SHA256) { throw 'Replacement recovery receipt does not match this owner and file pair.' }
                    [IO.File]::Delete($names.Legacy)
                    Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $prior.SHA256 'Recovered previously published replacement'
                }
            }
            if ([IO.File]::Exists($names.Legacy)) { $legacyHash = (Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash }
        }
        $stream = [IO.File]::Open($temporary, 'CreateNew', 'Write', 'None')
        try { $stream.Write($Bytes, 0, $Bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
        $expected = Get-SmartM365JsonHash $Bytes
        if ((Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash -ne $expected) { throw 'Staged JSON integrity check failed.' }
        if ($legacyHash) { Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'ReplacementPrepared' $expected $legacyHash }
        $publishWatch = [Diagnostics.Stopwatch]::StartNew()
        $retryDelayMilliseconds = 250
        $attempt = 0
        $lastLockError = $null
        $lockCode = 0
        if ([IO.File]::Exists($target)) {
            $attributes = [IO.File]::GetAttributes($target)
            if ($attributes -band [IO.FileAttributes]::ReadOnly) {
                $restoreAttributes = $attributes
                [IO.File]::SetAttributes($target, ($attributes -band (-bnot [IO.FileAttributes]::ReadOnly)))
            }
        }
        while ($true) {
            if ($null -ne $lastLockError -and $publishWatch.Elapsed.TotalSeconds -ge $ReplacementTimeoutSeconds) {
                $failure = [IO.IOException]::new(("Atomic replacement timed out for '{0}' after {1:N1}s and {2} attempt(s): destination remains locked (Win32 {3})." -f $target, $publishWatch.Elapsed.TotalSeconds, $attempt, $lockCode), $lastLockError)
                $failure.Data['SmartM365JsonFailureKind'] = 'LockTimeout'
                throw $failure
            }
            $attempt++
            try {
                foreach ($candidate in @($target, $temporary, $backup)) { Assert-SmartM365JsonUnlinkedPath $candidate -AllowQualifiedDataRoot }
                # An external writer may have changed the destination during a retry delay.
                if ($ExpectedSHA256 -eq 'ABSENT') {
                    if (Test-Path -LiteralPath $target) { throw 'JSON target was created by another writer.' }
                } elseif ($ExpectedSHA256) {
                    # Native IO preserves Win32 lock codes that PS5 Get-FileHash can discard.
                    if (-not [IO.File]::Exists($target) -or (Get-SmartM365JsonHash -Bytes ([IO.File]::ReadAllBytes($target))) -ne $ExpectedSHA256) {
                        throw 'JSON target changed since it was read; refusing a lost update.'
                    }
                }
                if ([IO.File]::Exists($target)) {
                    # No non-atomic emulation when the filesystem cannot replace atomically.
                    [IO.File]::Replace($temporary, $target, $backup)
                } else { [IO.File]::Move($temporary, $target) }
                $published = $true
                break
            } catch {
                $lockCode = Get-SmartM365JsonLockErrorCode -Exception $_.Exception
                if (-not $lockCode -or -not [IO.File]::Exists($temporary)) {
                    throw [IO.IOException]::new("Atomic replacement failed for '$target': $($_.Exception.Message)", $_.Exception)
                }
                $lastLockError = $_.Exception
                $remainingMilliseconds = [Math]::Ceiling(($ReplacementTimeoutSeconds - $publishWatch.Elapsed.TotalSeconds) * 1000)
                if ($remainingMilliseconds -gt 0) {
                    $delay = [int][Math]::Min($retryDelayMilliseconds, $remainingMilliseconds)
                    Write-Information ("[{0}] JSON destination locked (Win32 {1}): {2}; attempt={3}; retry in {4:N2}s; timeout={5}s." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $lockCode, $target, $attempt, ($delay / 1000.0), $ReplacementTimeoutSeconds) -InformationAction Continue
                    Start-Sleep -Milliseconds $delay
                    $retryDelayMilliseconds = [Math]::Min(5000, $retryDelayMilliseconds * 2)
                }
            }
        }
        if ($attempt -gt 1) {
            Write-Information ("[{0}] JSON replacement succeeded after {1} attempt(s), {2:N1}s: {3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $attempt, $publishWatch.Elapsed.TotalSeconds, $target) -InformationAction Continue
        }
        if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $expected) { throw "Published JSON integrity failure; previous version retained at $backup" }
        if ($legacyHash) {
            Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'ReplacementPublished' $expected $legacyHash
            if ((Get-FileHash -LiteralPath $names.Legacy -Algorithm SHA256).Hash -ne $legacyHash) { throw 'Legacy JSON changed during replacement; both files preserved.' }
            [IO.File]::Delete($names.Legacy)
            Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Completed' $expected 'Regenerated replacement verified before legacy retirement'
        }
        if ([IO.File]::Exists($backup)) { [IO.File]::Delete($backup) }
        [pscustomobject]@{ Path = $target; SHA256 = $expected; Bytes = $Bytes.Length }
    } finally {
        # Keep an unconsumed staging file on failure for recovery; never touch another owner's file.
        try {
            if (-not $published -and $null -ne $restoreAttributes -and [IO.File]::Exists($target)) { [IO.File]::SetAttributes($target, $restoreAttributes) }
        } finally { $lock.Dispose() }
    }
}

function Complete-SmartM365JsonConsumption {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string]$Owner,[Parameter(Mandatory)][string]$ExpectedSHA256)
    $names = Get-SmartM365JsonNames $Path
    foreach ($candidate in @($names.Legacy,$names.Preferred,$names.Lock,$names.Journal)) { Assert-SmartM365JsonUnlinkedPath $candidate }
    $lock = Enter-SmartM365JsonTransportLock $names.Lock 10
    try {
        $current = Read-SmartM365JsonDocument $Path
        if ($current.SHA256 -ne $ExpectedSHA256) { throw 'JSON request changed before consumption; newer request preserved.' }
        if ($current.Coexistence -notin @('Single','Identical')) { throw 'Unexpected request coexistence; files preserved.' }
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'ConsumptionPrepared' $ExpectedSHA256 ''
        if ($current.Coexistence -eq 'Identical') { [IO.File]::Delete($names.Legacy) }
        [IO.File]::Delete($current.Path)
        Write-SmartM365JsonMigrationEvent $names.Journal $Owner 'Consumed' $ExpectedSHA256 ''
    } finally { $lock.Dispose() }
}

function Get-SmartM365JsonTemplateName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    # Committed templates stay next to the script, even when the private file lives in the SmartInventory configuration root.
    return (Get-SmartM365JsonNames (ConvertFrom-SmartM365InventoryConfigLocation -Path $Path)).Legacy + '.template'
}

function Get-SmartM365InventoryCodeRoot {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    # The SmartInventory folder of the deployment holding this module (a launcher cache maps to its own copy).
    return [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\SmartInventory')).TrimEnd('\')
}

<#
.SYNOPSIS
Optional root for the private SmartInventory files: SmartInventoryConfigRootPath of the global file.
Empty or absent keeps every file next to its script. Deployment setting: tenant overlays never move it.
Only {{WorkspaceRootPath}} and {{SmartM365RootPath}} are resolved, so a root below WorkspaceRootPath stays qualified.
#>
function Get-SmartM365InventoryConfigRoot {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $globalPath = Join-Path $PSScriptRoot '..\..\Config\SmartM365.global.local.json'
    if (-not (Get-SmartM365JsonReadPath -Path $globalPath -Optional)) { return '' }
    $globalConfig = (Read-SmartM365JsonDocument -Path $globalPath).Document
    if ($globalConfig -isnot [pscustomobject]) { return '' }
    $property = $globalConfig.PSObject.Properties['SmartInventoryConfigRootPath']
    if ($null -eq $property) { return '' }
    $value = [string]$property.Value
    if ([string]::IsNullOrWhiteSpace($value) -or $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return '' }
    $resolved = Resolve-SmartM365JsonQualifiedUncToken -Value $value
    if ([string]::IsNullOrWhiteSpace($resolved) -or $resolved -match '[*?]' -or -not [IO.Path]::IsPathRooted($resolved) -or
        @($resolved.Split('\') | Where-Object { $_ -in @('.', '..') }).Count) {
        throw "SmartInventoryConfigRootPath must resolve to an absolute folder: $value"
    }
    $root = [IO.Path]::GetFullPath($resolved).TrimEnd('\')
    $codeRoot = Get-SmartM365InventoryCodeRoot
    if (($root + '\').StartsWith($codeRoot + '\', [StringComparison]::OrdinalIgnoreCase) -or
        ($codeRoot + '\').StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "SmartInventoryConfigRootPath must be outside the SmartInventory code folder: $root"
    }
    return $root
}

<#
.SYNOPSIS
Maps a private JSON file of the SmartInventory code folder to the configuration root, same relative path.
Other paths, and an empty SmartInventoryConfigRootPath, are returned unchanged. Fails closed when the file
still exists only next to the script: the files must be moved first, never read from a second source.
#>
function ConvertTo-SmartM365InventoryConfigLocation {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    $codeRoot = Get-SmartM365InventoryCodeRoot
    if (-not $full.StartsWith($codeRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { return $Path }
    $configRoot = Get-SmartM365InventoryConfigRoot
    if (-not $configRoot) { return $Path }
    $target = Join-Path $configRoot $full.Substring($codeRoot.Length + 1)
    if (-not (Get-SmartM365JsonReadPath -Path $target -Optional) -and (Get-SmartM365JsonReadPath -Path $full -Optional)) {
        throw ("SmartInventory private file still in the code folder: {0}. SmartInventoryConfigRootPath is {1}; move the private files there before running." -f $full, $configRoot)
    }
    $parent = Split-Path -Path $target -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop | Out-Null }
    return $target
}

function ConvertFrom-SmartM365InventoryConfigLocation {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Path)
    # Inverse mapping, used to find the committed template of a relocated file.
    $configRoot = Get-SmartM365InventoryConfigRoot
    if (-not $configRoot) { return $Path }
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($configRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { return $Path }
    return (Join-Path (Get-SmartM365InventoryCodeRoot) $full.Substring($configRoot.Length + 1))
}

function Publish-SmartM365RegeneratedJsonBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][byte[]]$Bytes,
        [Parameter(Mandatory)][scriptblock]$Validate
    )
    $names = Get-SmartM365JsonNames $Path
    $policy = Get-SmartM365JsonTransportPolicy
    $target = Get-SmartM365JsonReadPath $Path -Optional
    if ($policy.Mode -eq 'JsonText') {
        if ($names.Preferred.StartsWith('\\')) {
            $qualified = @($policy.QualifiedUncRoots | Where-Object { $names.Preferred.StartsWith(([IO.Path]::GetFullPath($_).TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
            if (-not $qualified) { throw 'UNC replacement requires qualification on the target share before activation.' }
        }
        $target = $names.Preferred
    } elseif (-not $target) { $target = $names.Legacy }
    $retire = $policy.Mode -eq 'JsonText'
    Write-SmartM365JsonBytesAtomically -Path $target -Bytes $Bytes -Validate $Validate -RetireLegacyAfterPublication:$retire -Owner $Owner
}

function Resolve-SmartM365OwnedJsonPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][scriptblock]$Validate,
        [string[]]$CompatibleJournalOwners = @()
    )
    $names = Get-SmartM365JsonNames $Path
    $policy = Get-SmartM365JsonTransportPolicy
    if ($policy.Mode -eq 'JsonText' -and $names.Legacy.StartsWith('\\')) {
        $qualifiedRoot = @($policy.QualifiedUncRoots | Where-Object { $names.Legacy.StartsWith(([IO.Path]::GetFullPath($_).TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
        if (-not $qualifiedRoot) { throw 'UNC JSON activation requires qualification on the target share.' }
    }
    $existing = Get-SmartM365JsonReadPath $Path -Optional
    if ($existing) {
        $null = Read-SmartM365JsonDocument $Path -Validate $Validate
        if ($policy.Mode -eq 'JsonText') {
            $qualified = $false
            foreach ($root in @($policy.QualifiedUncRoots)) {
                if ($names.Legacy.StartsWith(([IO.Path]::GetFullPath($root).TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase)) { $qualified = $true }
            }
            $result = Move-SmartM365OwnedJsonFile -Root ([IO.Path]::GetDirectoryName($names.Legacy)) -RelativePath ([IO.Path]::GetFileName($names.Legacy)) -Owner $Owner -Validate $Validate -CompatibleJournalOwners $CompatibleJournalOwners -QualifiedUnc:$qualified -RemoveIdenticalLegacy
            return $result.Path
        }
        return $existing
    }
    if ($policy.Mode -eq 'JsonText') { return $names.Preferred }
    return $names.Legacy
}

function Resolve-SmartM365WeeklyManifestPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HistoryRootPath,
        [Parameter(Mandatory)][string]$HistoryLabel,
        [scriptblock]$ValidateOwner
    )
    # The owner supplies its exact history root. Never recurse into arbitrary JSON.
    $rootFull = [IO.Path]::GetFullPath($HistoryRootPath).TrimEnd('\', '/')
    # Automatic CSV publication uses the generic label; explicit saves use a family label.
    # Permit both directions only for the documented collector-relative history root.
    $genericLabel = 'SmartM365 inventory'
    $compatibleLabels = @($HistoryLabel, $genericLabel)
    $familyLabelsByRoot = @{
        'Exchange\OnPrem\ServersAndStorage\WeeklyHistory' = @('Exchange on-prem infrastructure and readiness')
        'Exchange\OnPrem\Mailboxes\WeeklyHistory' = @('Exchange on-prem mailboxes', 'Exchange on-prem remote mailboxes', 'Exchange on-prem mailbox daily stats')
        'Exchange\OnPrem\CalendarPermissions\WeeklyHistory' = @('Exchange on-premises calendar permissions')
        'M365\Licensing\Licenses\WeeklyHistory' = @('M365 licenses inventory')
        'M365\Backup\PolicyScope\WeeklyHistory' = @('M365 Backup policy scope inventory')
        'M365\Teams\Inventory\WeeklyHistory' = @('Microsoft Teams inventory')
        'M365\SharePoint\Inventory\WeeklyHistory' = @('SmartM365 SharePoint Online inventory')
    }
    $relativeMatch = [regex]::Match($rootFull, '(?i)[\\/]DATA-ALL[\\/](?<relative>.+[\\/]WeeklyHistory)$')
    if ($relativeMatch.Success) {
        $familyLabels = $familyLabelsByRoot[$relativeMatch.Groups['relative'].Value.Replace('/', '\')]
        if ($familyLabels -and ($HistoryLabel -eq $genericLabel -or $HistoryLabel -in $familyLabels)) {
            $compatibleLabels += $familyLabels
        }
    }
    $journalLabels = if ($ValidateOwner) { @($HistoryLabel) } else { $compatibleLabels }
    $compatibleJournalOwners = @($journalLabels | Sort-Object -Unique | ForEach-Object { 'WeeklyHistory:' + $_ })
    $plans = @()
    foreach ($folder in @(Get-ChildItem -LiteralPath $HistoryRootPath -Directory -ErrorAction Stop | Where-Object { $_.Name -match '^\d{4}-W(?:0[1-9]|[1-4]\d|5[0-3])$' })) {
        $legacy = Join-Path $folder.FullName 'manifest.json'
        if (-not (Get-SmartM365JsonReadPath $legacy -Optional)) { continue }
        $week = $folder.Name
        $validator = {
            param($document)
            if ($document.Week -ne $week) { throw 'Weekly manifest week mismatch.' }
            if ($ValidateOwner) { & $ValidateOwner $document | Out-Null }
            else {
                if ($document.HistoryLabel -notin $compatibleLabels) { throw "Weekly manifest owner mismatch: $legacy (expected '$HistoryLabel', found '$($document.HistoryLabel)')." }
                $labelChanged = $document.HistoryLabel -ne $HistoryLabel
                $recordedRoot = [IO.Path]::GetFullPath([string]$document.HistoryRootPath).TrimEnd('\', '/')
                if ($recordedRoot -ne $rootFull) {
                    # A relocated workspace keeps the tenant and collector-relative history identity.
                    # Do not dereference the historical path or rewrite the historical payload.
                    $identityPattern = '(?i)[\\/]Tenants[\\/](?<identity>[^\\/]+[\\/]DATA-ALL[\\/].+[\\/]WeeklyHistory)$'
                    $recordedIdentity = [regex]::Match($recordedRoot, $identityPattern)
                    $currentIdentity = [regex]::Match($rootFull, $identityPattern)
                    if (-not $recordedIdentity.Success -or -not $currentIdentity.Success -or
                        $recordedIdentity.Groups['identity'].Value.Replace('/', '\') -ine $currentIdentity.Groups['identity'].Value.Replace('/', '\')) {
                        throw 'Weekly manifest history root mismatch (tenant/collector identity differs).'
                    }
                }
                # Aliases never bypass tenant/root identity or completeness validation.
                if ($recordedRoot -ne $rootFull -or $labelChanged) {
                    foreach ($csv in @($document.Files)) {
                        if ([string]::IsNullOrWhiteSpace([string]$csv) -or [IO.Path]::GetFileName([string]$csv) -ne $csv -or -not [IO.File]::Exists((Join-Path $folder.FullName $csv))) {
                            throw 'Relocated or relabelled weekly manifest snapshot is incomplete.'
                        }
                    }
                }
            }
            if (-not $document.PSObject.Properties['Files']) { throw 'Weekly manifest Files is missing.' }
            foreach ($name in @($document.Files)) {
                if ([string]::IsNullOrWhiteSpace([string]$name) -or [IO.Path]::GetFileName([string]$name) -ne $name -or $name -notmatch '(?i)\.csv$') { throw 'Weekly manifest contains an invalid CSV filename.' }
            }
        }.GetNewClosure()
        $before = Read-SmartM365JsonDocument $legacy -Validate $validator
        $plans += [pscustomobject]@{ Legacy=$legacy; Before=$before.Path; Validate=$validator }
    }
    # Preflight all existing weeks before mutating any; divergent pairs fail closed.
    foreach ($plan in $plans) {
        $resolved = Resolve-SmartM365OwnedJsonPath -Path $plan.Legacy -Owner ('WeeklyHistory:' + $HistoryLabel) -CompatibleJournalOwners $compatibleJournalOwners -Validate $plan.Validate
        [pscustomobject]@{ Path=$resolved; Renamed=($resolved -ne $plan.Before) }
    }
}

function Resolve-SmartM365JsonConfigurationPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $Path = ConvertTo-SmartM365InventoryConfigLocation -Path $Path
    $validate = { param($document) if ($document -isnot [System.Management.Automation.PSCustomObject]) { throw 'Configuration must be a JSON object.' } }
    Resolve-SmartM365OwnedJsonPath -Path $Path -Owner ('Configuration:' + [IO.Path]::GetFileName($Path)) -Validate $validate
}

Export-ModuleMember -Function Get-SmartM365JsonNames, Get-SmartM365JsonReadPath, Read-SmartM365JsonDocument, Move-SmartM365OwnedJsonFile, Resolve-SmartM365JsonConfigurationPath, Get-SmartM365JsonTransportPolicy, Get-SmartM365JsonTemplateName, Resolve-SmartM365OwnedJsonPath, Write-SmartM365JsonBytesAtomically, Resolve-SmartM365WeeklyManifestPaths, Publish-SmartM365RegeneratedJsonBytes, Complete-SmartM365JsonConsumption, Get-SmartM365SiblingTemporaryPath, Get-SmartM365InventoryConfigRoot, ConvertTo-SmartM365InventoryConfigLocation, ConvertFrom-SmartM365InventoryConfigLocation

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCm9cMRmXn+Zyi
# X5eMPABXXsPNuFdtam3TFsuy/JXm/aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOMuEnjI5D4/VE+a3pwm8iBj3zUM44GT9xXXllnPU5M8MA0GCSqG
# SIb3DQEBAQUABIIBgHVntkAms7puVD1msGQXKsUcDJ5jArf4pDGp/Pzam53LyRK1
# gDAy4uWZtrCgZ0jXImLn6SfegIw/19Og9cQLIigEOxBgcjMjORBr+L4u0/KKuz8D
# e8lJ+Lt1tncJHXXWAVB0juuz/QZyPjdMGNY5q9MhvUE7MjGWbAOjvoL5UfSXCGia
# eCl+nz010K03f8xgYIh9QxpGVKxm/pC5QSvIkSgnXJ6HubGrZRYe2IlJqoWQiIIx
# QSeXxxkKMt5RDrZvqf58FUXPQwkX3gzv1fzXxWUP9EZaxWA2QGpfmjLedCdUOxEL
# tC0Wb9ViLpuvfbVusF3pADr1SmoriPYVw2sMRAFUj2H5GpQIbwqoJUC5KoYTbDRw
# Dwbzz6WQ43yJFzmNIyejSp4ZN6sxCAjxkyHXUwH+mtTKhA40G0/+GxrRoENtyGfq
# 1uHiYhQG1NXMr8Q8j3vW4q4xahkA+oPtXbqUHYZr4ocmoTnoxjs64/A9a0ykAydj
# tOtOYoVjpW7Qe1NOQaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcyMDUy
# NTJaMC8GCSqGSIb3DQEJBDEiBCAg7L7V/7giSTdy/zA7kjjiBQ0qxue9AakcSURZ
# syb+FDANBgkqhkiG9w0BAQEFAASCAgB+AaU6cRN8YF8QQEOXb7uxOb0eOZGMePlz
# oL2/007cWsgAyqYNMTYQLnDUMxRHPhm24Up2wDIc7nzPjuPGASXZlWBX6xxy0rD4
# 4W9GDRul05ziIFH6gU7LfdzB3g0iRFODIk6voVSApkEmBXZYXBV0AxeSOYgd5Cx1
# d5nvhRueon3PeiWCvUvMtpJ0OArbDsxcWax/S7E5VW+MXXpSUZ3QUILOttCDQYIy
# Z6HHQTeq09HVEhdYwkOLsY17Gis58CmNuD71jvTBa1BDapt4shezvs4Qn8FaVaaB
# GWE6PQfmGejh+DEO+AgcA4U61g1n2wrq1OFWZcMMyzFQMCGFHNqKdByV0Npv3OG4
# U41FCZGzYfvVF/TAS7bSxAvOwSrGUK9/fawDeWk4XaO/29FzT19Yo64j/VVjpSGS
# zjJRBJarcIe9HJbvpTYQzkx9XDUQkB5Eik/adNrGWUYJAHRZd4WOeljUvvAclFg8
# YpM9jExwvL3RKUWk82XMiIa1b9QdouDEExyZnHch0I/qfPel+q+vppiDQVZhm+rf
# 6UG6zY6ihcogwKaQsPL59SiU7e0fsjxBSsREPcjqXAuG76K/HE37tRQB4QmSC4Y0
# TpMWej9bz18lBNg33yWy2yOH0trmUqHF7bbRs75245pvSwNJBq9eKEtgAJP3zrL0
# 8zRBETF7uQ==
# SIG # End signature block
