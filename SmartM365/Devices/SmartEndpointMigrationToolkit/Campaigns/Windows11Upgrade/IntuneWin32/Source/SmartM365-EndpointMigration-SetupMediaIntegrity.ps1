<#
.SYNOPSIS
    Validates a Windows setup media folder against its SmartM365 SHA256 manifest.
.DESCRIPTION
    Dot-sourced helper of the Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune package. Validates every RelativePath, Length and SHA256 row of SmartM365-SetupMediaManifest.sha256.csv using long-path aware hashing.
    Also carries the shared, read-only setup media resolution (region identical to the endpoint script) used by the installer and the runner to select new, legacy Intune-owned or adoptable media before any copy.
    Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/IntuneWin32/Source/SmartM365-SetupMediaIntegrity.ps1 1.0.0 (see Docs\PROVENANCE.json.txt).
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

function ConvertTo-SmartM365LongLiteralPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith('\\?\')) { return $fullPath }
    if ($fullPath.StartsWith('\\')) { return '\\?\UNC\' + $fullPath.Substring(2) }
    return '\\?\' + $fullPath
}

function Get-SmartM365SetupMediaFileHash {
    param([Parameter(Mandatory = $true)][string]$Path)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead((ConvertTo-SmartM365LongLiteralPath -Path $Path))
        $hashBytes = $sha.ComputeHash($stream)
        return (($hashBytes | ForEach-Object { $_.ToString('X2') }) -join '')
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $sha) { $sha.Dispose() }
    }
}

function Test-SmartM365SetupMediaIntegrity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$MediaRoot,
        [string]$ManifestFileName = 'SmartM365-SetupMediaManifest.sha256.csv'
    )

    if (-not (Test-Path -LiteralPath $MediaRoot -PathType Container)) {
        throw "Setup media root not found: $MediaRoot"
    }

    $manifestPath = Join-Path $MediaRoot $ManifestFileName
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Setup media integrity manifest not found: $manifestPath"
    }

    try {
        $rootFull = [System.IO.Path]::GetFullPath($MediaRoot).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
        $rows = @(Import-Csv -LiteralPath $manifestPath -ErrorAction Stop)
    }
    catch {
        throw ("Setup media integrity manifest cannot be read. Manifest={0}; Error={1}" -f $manifestPath,$_.Exception.Message)
    }

    if ($rows.Count -eq 0) { throw "Setup media integrity manifest is empty: $manifestPath" }
    foreach ($column in @('RelativePath','Length','SHA256')) {
        if (-not $rows[0].PSObject.Properties[$column]) {
            throw "Setup media integrity manifest is missing required column '$column': $manifestPath"
        }
    }

    $checkedFiles = 0
    $checkedBytes = 0L
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $relativePath = ([string]$row.RelativePath).Trim().TrimStart('\', '/')
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
            throw ("Setup media integrity manifest contains an empty RelativePath. Manifest={0}; Row={1}" -f $manifestPath,$rowNumber)
        }
        if ([System.IO.Path]::IsPathRooted($relativePath) -or $relativePath -match '(^|[\\/])\.\.([\\/]|$)') {
            throw ("Setup media integrity manifest contains an unsafe RelativePath. Manifest={0}; Row={1}; RelativePath={2}" -f $manifestPath,$rowNumber,$relativePath)
        }
        if ($relativePath -match '[<>:"|?*\x00-\x1F]') {
            throw ("Setup media integrity manifest contains an invalid RelativePath character. Manifest={0}; Row={1}; RelativePath={2}" -f $manifestPath,$rowNumber,$relativePath)
        }

        try {
            $fileFull = [System.IO.Path]::GetFullPath((Join-Path $MediaRoot $relativePath))
        }
        catch {
            throw ("Setup media integrity manifest path cannot be resolved. Manifest={0}; Row={1}; RelativePath={2}; Error={3}" -f $manifestPath,$rowNumber,$relativePath,$_.Exception.Message)
        }
        if (-not $fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw ("Setup media integrity manifest path escapes media root. Manifest={0}; Row={1}; RelativePath={2}; File={3}" -f $manifestPath,$rowNumber,$relativePath,$fileFull)
        }
        if (-not (Test-Path -LiteralPath $fileFull -PathType Leaf)) {
            throw ("Setup media integrity check failed. File missing. Manifest={0}; Row={1}; RelativePath={2}; File={3}" -f $manifestPath,$rowNumber,$relativePath,$fileFull)
        }

        $expectedLength = 0L
        if (-not [int64]::TryParse([string]$row.Length, [ref]$expectedLength)) {
            throw ("Setup media integrity manifest contains an invalid Length. Manifest={0}; Row={1}; RelativePath={2}; Length={3}" -f $manifestPath,$rowNumber,$relativePath,$row.Length)
        }
        $item = Get-Item -LiteralPath $fileFull -ErrorAction Stop
        if ([int64]$item.Length -ne $expectedLength) {
            throw ("Setup media integrity check failed. Length mismatch. Manifest={0}; Row={1}; RelativePath={2}; File={3}; ExpectedLength={4}; ActualLength={5}" -f $manifestPath,$rowNumber,$relativePath,$fileFull,$expectedLength,$item.Length)
        }

        $expectedHash = ([string]$row.SHA256).Trim().ToUpperInvariant()
        if ($expectedHash -notmatch '^[A-F0-9]{64}$') {
            throw ("Setup media integrity manifest contains an invalid SHA256. Manifest={0}; Row={1}; RelativePath={2}; SHA256={3}" -f $manifestPath,$rowNumber,$relativePath,$row.SHA256)
        }
        try { $actualHash = (Get-SmartM365SetupMediaFileHash -Path $fileFull).ToUpperInvariant() }
        catch {
            throw ("Setup media integrity check failed. Cannot hash file. Manifest={0}; Row={1}; RelativePath={2}; File={3}; Error={4}" -f $manifestPath,$rowNumber,$relativePath,$fileFull,$_.Exception.Message)
        }
        if ($actualHash -ne $expectedHash) {
            throw ("Setup media integrity check failed. SHA256 mismatch. Manifest={0}; Row={1}; RelativePath={2}; File={3}; ExpectedSHA256={4}; ActualSHA256={5}" -f $manifestPath,$rowNumber,$relativePath,$fileFull,$expectedHash,$actualHash)
        }

        $checkedFiles++
        $checkedBytes += [int64]$item.Length
    }

    return [pscustomobject]@{
        MediaRoot = [System.IO.Path]::GetFullPath($MediaRoot)
        ManifestPath = $manifestPath
        Files = $checkedFiles
        Bytes = $checkedBytes
    }
}

#region SEMT setup media resolution v1
# Shared, read-only resolution of the Windows setup media used by the Windows 11 Campaign.
# The same region is copied unchanged into the endpoint script and into the Intune package
# helper (a test checks that both copies are identical). It never moves, deletes, writes or
# changes ACLs: it only decides which media folder may be used and why.
#
# Decisions:
#   NewCacheValid       valid media in the cache of this product
#   LegacyIntuneOwned   valid media of the previous Windows 11 toolkit still owned by its Intune
#                       package: used in place, read-only, external to this product
#   LegacyAdoptable     valid, idle, unmanaged legacy media that can be moved to the new cache
#   InUse               the selected media is used by a running process or a live lock
#   Indeterminate       the state could not be read (registry, processes, locks, permissions,
#                       folder tree): nothing is done, the caller stops
#   Untrusted           legacy media is present but a non-administrator owns or can modify it, or
#                       it contains a reparse point: never used (setup.exe runs as SYSTEM), its
#                       permissions are never changed, and a second copy needs authorization
#   Incompatible        media present but not the requested media (language, version)
#   Corrupt             media present but invalid
#   Absent              no media of this product and no legacy media found
# LegacyValidRemains is true when valid legacy media exists but is not used. DuplicationRequired
# is true when legacy media is present and not used (valid or untrusted): a second copy then
# requires an explicit duplication authorization.

$script:SemtLegacySetupMediaRegistryKey = 'SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages'

function Test-SemtPathHasReparsePoint {
    param([Parameter(Mandatory = $true)][string]$Path)
    $current = [System.IO.Path]::GetFullPath($Path)
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path -LiteralPath $current) {
            try {
                $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
                if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
            }
            catch { return $true }
        }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    return $false
}

function Test-SemtTreeHasReparsePoint {
    # Stops at the first reparse point and never descends into one (no junction is followed).
    param([Parameter(Mandatory = $true)][string]$Path)
    if (Test-SemtPathHasReparsePoint -Path $Path) { return $true }
    $pending = New-Object System.Collections.Generic.Queue[string]
    $pending.Enqueue([System.IO.Path]::GetFullPath($Path))
    try {
        while ($pending.Count -gt 0) {
            $folder = $pending.Dequeue()
            foreach ($item in @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop)) {
                if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
                if ($item.PSIsContainer) { $pending.Enqueue($item.FullName) }
            }
        }
    }
    catch { return $true }
    return $false
}

function Get-SemtTreeReparseState {
    # Found: a reparse point exists on the path or in the tree. Unreadable: the tree could not be
    # enumerated, so the absence of a reparse point is not proven. Never descends into one.
    param([Parameter(Mandatory = $true)][string]$Path)
    if (Test-SemtPathHasReparsePoint -Path $Path) { return [pscustomobject]@{ State = 'Found'; Detail = ('Reparse point on the path of {0}.' -f $Path) } }
    $pending = New-Object System.Collections.Generic.Queue[string]
    $pending.Enqueue([System.IO.Path]::GetFullPath($Path))
    try {
        while ($pending.Count -gt 0) {
            $folder = $pending.Dequeue()
            foreach ($item in @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop)) {
                if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ State = 'Found'; Detail = ('Reparse point: {0}' -f $item.FullName) } }
                if ($item.PSIsContainer) { $pending.Enqueue($item.FullName) }
            }
        }
    }
    catch { return [pscustomobject]@{ State = 'Unreadable'; Detail = ('Media folder tree could not be read: {0}' -f $_.Exception.Message) } }
    return [pscustomobject]@{ State = 'None'; Detail = '' }
}

function Get-SemtLegacySetupMediaOwners {
    # Intune packages of the previous toolkit whose SetupCacheFolder is this media folder.
    param(
        [Parameter(Mandatory = $true)][string]$Leaf,
        [string]$RegistryKeyPath = $script:SemtLegacySetupMediaRegistryKey
    )
    $owners = New-Object System.Collections.Generic.List[string]
    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        try {
            $packages = $baseKey.OpenSubKey($RegistryKeyPath, $false)
            if ($packages) {
                try {
                    foreach ($name in @($packages.GetSubKeyNames())) {
                        $package = $packages.OpenSubKey($name, $false)
                        if (-not $package) { continue }
                        try {
                            $cacheFolder = [string]$package.GetValue('SetupCacheFolder', '')
                            $installState = [string]$package.GetValue('InstallState', '')
                            if ($cacheFolder -ieq $Leaf -and $installState -notin @('','Uninstalled','Removed')) { $owners.Add(('{0}:{1}' -f $name,$installState)) }
                        }
                        finally { $package.Dispose() }
                    }
                }
                finally { $packages.Dispose() }
            }
        }
        finally { $baseKey.Dispose() }
    }
    catch {
        return [pscustomobject]@{ State = 'Unreadable'; Owners = @(); Detail = ('Legacy Intune package registry could not be read: {0}' -f $_.Exception.Message) }
    }
    return [pscustomobject]@{ State = 'Read'; Owners = @($owners.ToArray()); Detail = '' }
}

function Get-SemtSetupMediaUsage {
    # A media folder is in use when a live process runs from it or copies into it, or when a
    # media lock of its folder is held by a live process. Unreadable state is Indeterminate.
    param(
        [Parameter(Mandatory = $true)][string]$MediaPath,
        [string[]]$LockRoots = @()
    )
    $full = [System.IO.Path]::GetFullPath($MediaPath).TrimEnd('\')
    $leaf = Split-Path -Path $full -Leaf
    $signals = New-Object System.Collections.Generic.List[string]
    foreach ($lockRoot in @($LockRoots | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if (-not (Test-Path -LiteralPath $lockRoot -PathType Container)) { continue }
        foreach ($lockDir in @(Get-ChildItem -LiteralPath $lockRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*{0}*' -f $leaf) })) {
            $payloadPath = Join-Path $lockDir.FullName 'lock.json'
            if (-not (Test-Path -LiteralPath $payloadPath -PathType Leaf)) {
                if ($lockDir.LastWriteTimeUtc -gt (Get-Date).ToUniversalTime().AddMinutes(-5)) { $signals.Add(('Recent lock without payload: {0}' -f $lockDir.FullName)) }
                continue
            }
            $payload = $null
            try { $payload = Get-Content -LiteralPath $payloadPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
            catch { return [pscustomobject]@{ State = 'Indeterminate'; Detail = ('Media lock is unreadable: {0}' -f $payloadPath) } }
            $lockPid = 0
            if ($payload -and $payload.PSObject.Properties['PID']) { [void][int]::TryParse([string]$payload.PID, [ref]$lockPid) }
            if ($lockPid -gt 0 -and (Get-Process -Id $lockPid -ErrorAction SilentlyContinue)) { $signals.Add(('Lock {0} held by PID {1}' -f $lockDir.Name,$lockPid)) }
        }
    }
    try {
        $escaped = [regex]::Escape($full)
        foreach ($process in @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)) {
            $text = '{0} {1}' -f [string]$process.ExecutablePath,[string]$process.CommandLine
            if ($text -match ('(?i){0}(\\|"|\s|$)' -f $escaped)) { $signals.Add(('Process {0}:{1}' -f $process.Name,$process.ProcessId)) }
        }
    }
    catch { return [pscustomobject]@{ State = 'Indeterminate'; Detail = ('Running processes could not be listed: {0}' -f $_.Exception.Message) } }
    if ($signals.Count -gt 0) { return [pscustomobject]@{ State = 'InUse'; Detail = ($signals -join '; ') } }
    return [pscustomobject]@{ State = 'Free'; Detail = '' }
}

function Test-SemtMediaWriteRestricted {
    # Legacy media is run as SYSTEM, so it is trusted only when no principal other than SYSTEM,
    # Administrators, TrustedInstaller or CREATOR OWNER owns or can modify any of its items.
    # State: Trusted, Untrusted (insufficient protection, proven), or Unreadable (the permissions
    # or the folder tree could not be read: the protection is not proven either way).
    # Read-only: the permissions are never changed.
    param([Parameter(Mandatory = $true)][string]$Path)
    $trustedSids = @('S-1-5-18','S-1-5-32-544','S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464','S-1-3-0')
    $writeRights = [System.Security.AccessControl.FileSystemRights]'WriteData, AppendData, WriteExtendedAttributes, DeleteSubdirectoriesAndFiles, Delete, ChangePermissions, TakeOwnership'
    try { $items = @(Get-Item -LiteralPath $Path -Force -ErrorAction Stop) + @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Stop) }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Media folder tree could not be read: {0}' -f $_.Exception.Message) } }
    foreach ($item in $items) {
        try {
            $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
            $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
            $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
        }
        catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Permissions of {0} could not be read: {1}' -f $item.FullName,$_.Exception.Message) } }
        if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
        if ($trustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM or Administrators.' -f $owner.Value,$item.FullName) } }
        foreach ($rule in $rules) {
            if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
            if (($rule.FileSystemRights -band $writeRights) -eq 0) { continue }
            if ($trustedSids -notcontains $rule.IdentityReference.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can modify {1}.' -f $rule.IdentityReference.Value,$item.FullName) } }
        }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Get-SemtMediaTrustState {
    # Normalizes a trust provider result. A missing or unknown result is Unreadable, never Untrusted
    # or Trusted, so an unverifiable state always stops the caller.
    param([AllowNull()]$Trust)
    if ($null -eq $Trust) { return 'Unreadable' }
    if ($Trust.PSObject.Properties['State']) {
        $state = [string]$Trust.State
        if ($state -in @('Trusted','Untrusted','Unreadable')) { return $state }
        return 'Unreadable'
    }
    if ($Trust.PSObject.Properties['Trusted']) { if ([bool]$Trust.Trusted) { return 'Trusted' } else { return 'Untrusted' } }
    return 'Unreadable'
}

function Invoke-SemtSetupMediaValidator {
    param([Parameter(Mandatory = $true)][scriptblock]$ValidateMedia, [Parameter(Mandatory = $true)][string]$Path)
    try {
        $result = & $ValidateMedia $Path
        if ($null -eq $result) { return [pscustomobject]@{ Valid = $false; Category = 'Corrupt'; Detail = 'Validator returned nothing.' } }
        return $result
    }
    catch { return [pscustomobject]@{ Valid = $false; Category = 'Corrupt'; Detail = $_.Exception.Message } }
}

function Resolve-SemtSetupMedia {
    param(
        [Parameter(Mandatory = $true)][string]$NewCachePath,
        [string[]]$LegacyCandidates = @(),
        [Parameter(Mandatory = $true)][scriptblock]$ValidateMedia,
        [scriptblock]$OwnerProvider = { param($Path) Get-SemtLegacySetupMediaOwners -Leaf (Split-Path -Path $Path -Leaf) },
        [scriptblock]$UsageProvider = { param($Path) Get-SemtSetupMediaUsage -MediaPath $Path },
        [scriptblock]$TrustProvider = { param($Path) Test-SemtMediaWriteRestricted -Path $Path },
        [switch]$AllowAdoption
    )

    $newFull = [System.IO.Path]::GetFullPath($NewCachePath).TrimEnd('\')
    $result = [ordered]@{
        Decision = 'Absent'; Path = ''; Source = 'None'; External = $false; Owner = ''
        NewCachePath = $newFull; NewCacheState = 'Absent'; NewCacheDetail = ''
        LegacyPath = ''; LegacyState = 'Absent'; LegacyDetail = ''; LegacyValidRemains = $false
        UntrustedPath = ''; UntrustedDetail = ''; DuplicationRequired = $false
        Detail = ''
    }

    if (Test-Path -LiteralPath (Join-Path $newFull 'setup.exe') -PathType Leaf) {
        $check = Invoke-SemtSetupMediaValidator -ValidateMedia $ValidateMedia -Path $newFull
        if ($check.Valid) {
            $result.NewCacheState = 'Valid'
            $result.Decision = 'NewCacheValid'; $result.Path = $newFull; $result.Source = 'NewCache'
            $result.Detail = 'Valid media in the cache of this product.'
            return [pscustomobject]$result
        }
        $result.NewCacheState = [string]$check.Category; $result.NewCacheDetail = [string]$check.Detail
    }
    elseif (Test-Path -LiteralPath $newFull) { $result.NewCacheState = 'Corrupt'; $result.NewCacheDetail = 'Cache folder exists without setup.exe.' }

    foreach ($candidate in @($LegacyCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [System.IO.Path]::GetFullPath($_).TrimEnd('\') } | Select-Object -Unique)) {
        if ($candidate -ieq $newFull) { continue }
        if (-not (Test-Path -LiteralPath (Join-Path $candidate 'setup.exe') -PathType Leaf)) { continue }
        $result.LegacyPath = $candidate
        $owners = & $OwnerProvider $candidate
        if ($null -eq $owners -or [string]$owners.State -ne 'Read') {
            $result.LegacyState = 'Indeterminate'; $result.LegacyDetail = if ($owners) { [string]$owners.Detail } else { 'Owner provider returned nothing.' }
            break
        }
        $usage = & $UsageProvider $candidate
        if ($null -eq $usage -or [string]$usage.State -eq 'Indeterminate') {
            $result.LegacyState = 'Indeterminate'; $result.LegacyDetail = if ($usage) { [string]$usage.Detail } else { 'Usage provider returned nothing.' }
            break
        }
        if ([string]$usage.State -eq 'InUse') { $result.LegacyState = 'InUse'; $result.LegacyDetail = [string]$usage.Detail; break }
        # Present legacy media is never reported as Absent: an unreadable tree or unreadable
        # permissions stop the caller (Indeterminate); proven insufficient protection is Untrusted.
        $reparse = Get-SemtTreeReparseState -Path $candidate
        if ([string]$reparse.State -eq 'Unreadable') { $result.LegacyState = 'Indeterminate'; $result.LegacyDetail = [string]$reparse.Detail; break }
        if ([string]$reparse.State -eq 'Found') {
            $result.LegacyState = 'Untrusted'; $result.LegacyDetail = ('Legacy media contains a reparse point; it is not used. {0}' -f $reparse.Detail)
            if ([string]::IsNullOrWhiteSpace($result.UntrustedPath)) { $result.UntrustedPath = $candidate; $result.UntrustedDetail = $result.LegacyDetail }
            continue
        }
        $trust = $null
        try { $trust = & $TrustProvider $candidate } catch { $trust = [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = $_.Exception.Message } }
        $trustState = Get-SemtMediaTrustState -Trust $trust
        $trustDetail = if ($trust -and $trust.PSObject.Properties['Detail']) { [string]$trust.Detail } else { 'no result' }
        if ($trustState -eq 'Unreadable') { $result.LegacyState = 'Indeterminate'; $result.LegacyDetail = ('Legacy media permissions could not be verified: {0}' -f $trustDetail); break }
        if ($trustState -eq 'Untrusted') {
            $result.LegacyState = 'Untrusted'; $result.LegacyDetail = ('Legacy media is not write-restricted; it is not used and its permissions are not changed: {0}' -f $trustDetail)
            if ([string]::IsNullOrWhiteSpace($result.UntrustedPath)) { $result.UntrustedPath = $candidate; $result.UntrustedDetail = $result.LegacyDetail }
            continue
        }
        $check = Invoke-SemtSetupMediaValidator -ValidateMedia $ValidateMedia -Path $candidate
        if (-not $check.Valid) { $result.LegacyState = [string]$check.Category; $result.LegacyDetail = [string]$check.Detail; continue }
        if (@($owners.Owners).Count -gt 0) {
            $result.LegacyState = 'IntuneOwned'; $result.Owner = (@($owners.Owners) -join ', ')
            $result.LegacyDetail = ('Owned by the legacy Intune package {0}.' -f $result.Owner)
            break
        }
        $reasons = New-Object System.Collections.Generic.List[string]
        if (-not $AllowAdoption) { $reasons.Add('adoption disabled') }
        if (Test-Path -LiteralPath $newFull) { $reasons.Add('the new cache folder already exists') }
        if ([System.IO.Path]::GetPathRoot($candidate) -ine [System.IO.Path]::GetPathRoot($newFull)) { $reasons.Add('other volume') }
        if ((Test-SemtTreeHasReparsePoint -Path $candidate) -or (Test-SemtPathHasReparsePoint -Path (Split-Path -Path $newFull -Parent))) { $reasons.Add('reparse point') }
        if ($reasons.Count -eq 0) { $result.LegacyState = 'Adoptable'; $result.LegacyDetail = 'Idle, unmanaged and valid.' }
        else { $result.LegacyState = 'NotAdoptable'; $result.LegacyDetail = ('Valid but not adoptable: {0}.' -f ($reasons -join ', ')) }
        break
    }

    switch ($result.LegacyState) {
        'IntuneOwned' { $result.Decision = 'LegacyIntuneOwned'; $result.Path = $result.LegacyPath; $result.Source = 'LegacyExternal'; $result.External = $true }
        'Adoptable' { $result.Decision = 'LegacyAdoptable'; $result.Path = $result.LegacyPath; $result.Source = 'LegacyAdoptable' }
        'InUse' { $result.Decision = 'InUse'; $result.Path = $result.LegacyPath }
        'Indeterminate' { $result.Decision = 'Indeterminate'; $result.Path = $result.LegacyPath }
        default {
            if ($result.LegacyState -eq 'NotAdoptable') { $result.LegacyValidRemains = $true }
            if (-not [string]::IsNullOrWhiteSpace($result.UntrustedPath) -and -not $result.LegacyValidRemains) {
                # Untrusted legacy media is reported as such, even when another candidate was
                # incompatible or corrupt afterwards.
                $result.Decision = 'Untrusted'
                $result.LegacyState = 'Untrusted'; $result.LegacyPath = $result.UntrustedPath; $result.LegacyDetail = $result.UntrustedDetail
            }
            elseif ($result.NewCacheState -in @('Incompatible','Corrupt')) { $result.Decision = $result.NewCacheState }
            elseif ($result.LegacyState -in @('Incompatible','Corrupt')) { $result.Decision = $result.LegacyState }
            else { $result.Decision = 'Absent' }
            if ($result.LegacyValidRemains -or -not [string]::IsNullOrWhiteSpace($result.UntrustedPath)) { $result.DuplicationRequired = $true }
        }
    }
    $result.Detail = ('NewCache={0}; Legacy={1} {2}' -f $result.NewCacheState,$result.LegacyState,$result.LegacyDetail).Trim()
    return [pscustomobject]$result
}
#endregion SEMT setup media resolution v1

#region SEMT Intune setup media selection
# Intune use of the shared resolution: the installer, the runner and the requirement/detection
# logic all decide from the same rules which media folder is used, before any copy.

$script:SemtIntuneLegacyWindows11Root = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'

function Test-SemtIntuneSetupMediaCandidate {
    # Read-only. Incompatible: not the requested language or not the packaged media build.
    # Corrupt: missing files or failed SHA256 manifest validation.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$ExpectedLanguage,
        [string]$ReferenceManifestPath
    )
    if (-not (Test-Path -LiteralPath (Join-Path $Path 'setup.exe') -PathType Leaf)) { return [pscustomobject]@{ Valid = $false; Category = 'Corrupt'; Detail = ('setup.exe missing: {0}' -f $Path) } }
    if (-not (Test-Path -LiteralPath (Join-Path $Path 'sources\install.wim') -PathType Leaf) -and -not (Test-Path -LiteralPath (Join-Path $Path 'sources\install.esd') -PathType Leaf)) { return [pscustomobject]@{ Valid = $false; Category = 'Corrupt'; Detail = ('sources\install.wim or install.esd missing: {0}' -f $Path) } }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedLanguage) -and $ExpectedLanguage -notin @('Any','MatchSystem')) {
        $langIni = Join-Path $Path 'sources\lang.ini'
        if (-not (Test-Path -LiteralPath $langIni -PathType Leaf)) { return [pscustomobject]@{ Valid = $false; Category = 'Incompatible'; Detail = ('Setup media language could not be detected (sources\lang.ini missing): {0}' -f $Path) } }
        $langText = Get-Content -LiteralPath $langIni -Raw -ErrorAction Stop
        if ($langText -notmatch ('(?im)^\s*{0}\s*=' -f [regex]::Escape($ExpectedLanguage))) { return [pscustomobject]@{ Valid = $false; Category = 'Incompatible'; Detail = ('Setup media language mismatch. Expected={0}; MediaRoot={1}' -f $ExpectedLanguage,$Path) } }
    }
    $candidateManifest = Join-Path $Path 'SmartM365-SetupMediaManifest.sha256.csv'
    if (-not (Test-Path -LiteralPath $candidateManifest -PathType Leaf)) { return [pscustomobject]@{ Valid = $false; Category = 'Corrupt'; Detail = ('Setup media integrity manifest not found: {0}' -f $candidateManifest) } }
    # A required reference that is missing refuses the media: the comparison is never skipped, and
    # the manifest that comes with the media is not, on its own, a proof of authenticity.
    if (-not [string]::IsNullOrWhiteSpace($ReferenceManifestPath)) {
        if (-not (Test-Path -LiteralPath $ReferenceManifestPath -PathType Leaf)) { return [pscustomobject]@{ Valid = $false; Category = 'Corrupt'; Detail = ('Reference setup media manifest of the package not found: {0}' -f $ReferenceManifestPath) } }
        $reference = (Get-SmartM365SetupMediaFileHash -Path $ReferenceManifestPath).ToUpperInvariant()
        $actual = (Get-SmartM365SetupMediaFileHash -Path $candidateManifest).ToUpperInvariant()
        if ($reference -ne $actual) { return [pscustomobject]@{ Valid = $false; Category = 'Incompatible'; Detail = ('Setup media differs from the packaged media (manifest SHA256 {0} instead of {1}): {2}' -f $actual,$reference,$Path) } }
    }
    try {
        $integrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $Path
        return [pscustomobject]@{ Valid = $true; Category = 'Valid'; Detail = ('Files={0}; Bytes={1}' -f $integrity.Files,$integrity.Bytes) }
    }
    catch { return [pscustomobject]@{ Valid = $false; Category = 'Corrupt'; Detail = $_.Exception.Message } }
}

function Get-SemtIntuneLegacySetupMediaCandidates {
    param([Parameter(Mandatory = $true)][string]$SetupCacheFolder, [string]$LegacyRoot = $script:SemtIntuneLegacyWindows11Root)
    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add((Join-Path (Join-Path $LegacyRoot 'SetupMedia') $SetupCacheFolder))
    try {
        $lastRunPath = Join-Path $LegacyRoot 'LastRun.json'
        if (Test-Path -LiteralPath $lastRunPath -PathType Leaf) {
            $lastRun = Get-Content -LiteralPath $lastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($lastRun -and $lastRun.PSObject.Properties['SetupCachePath']) {
                $recorded = [string]$lastRun.SetupCachePath
                # Normalized first, then anchored to <LegacyRoot>\SetupMedia\ (no '..' or sibling-prefix escape).
                if (-not [string]::IsNullOrWhiteSpace($recorded)) {
                    $recordedFull = [System.IO.Path]::GetFullPath($recorded).TrimEnd('\')
                    $legacyMediaRoot = [System.IO.Path]::GetFullPath((Join-Path $LegacyRoot 'SetupMedia')).TrimEnd('\') + '\'
                    if ($recordedFull.StartsWith($legacyMediaRoot, [System.StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Path $recordedFull -Leaf) -ieq $SetupCacheFolder) { $candidates.Add($recordedFull) }
                }
            }
        }
    }
    catch { }
    return @($candidates | ForEach-Object { [System.IO.Path]::GetFullPath($_) } | Select-Object -Unique)
}

function Resolve-SemtIntuneSetupMedia {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$SetupCacheFolder,
        [string]$Language,
        [string]$ReferenceManifestPath,
        [string]$LegacyRoot = $script:SemtIntuneLegacyWindows11Root,
        [scriptblock]$OwnerProvider = { param($Path) Get-SemtLegacySetupMediaOwners -Leaf (Split-Path -Path $Path -Leaf) },
        [scriptblock]$UsageProvider
    )
    $script:SemtIntuneMediaValidation = [pscustomobject]@{ Language = $Language; Reference = $ReferenceManifestPath }
    if ($null -eq $UsageProvider) {
        $script:SemtIntuneMediaLockRoots = @((Join-Path $LegacyRoot 'Locks'), (Join-Path $DataRoot 'Locks'))
        $UsageProvider = { param($Path) Get-SemtSetupMediaUsage -MediaPath $Path -LockRoots $script:SemtIntuneMediaLockRoots }
    }
    $newCache = Join-Path (Join-Path $DataRoot 'SetupMedia') $SetupCacheFolder
    return Resolve-SemtSetupMedia -NewCachePath $newCache -LegacyCandidates @(Get-SemtIntuneLegacySetupMediaCandidates -SetupCacheFolder $SetupCacheFolder -LegacyRoot $LegacyRoot) -ValidateMedia { param($Path) Test-SemtIntuneSetupMediaCandidate -Path $Path -ExpectedLanguage $script:SemtIntuneMediaValidation.Language -ReferenceManifestPath $script:SemtIntuneMediaValidation.Reference } -OwnerProvider $OwnerProvider -UsageProvider $UsageProvider -AllowAdoption
}

function Test-SemtIntuneMediaDecisionUsable {
    param([Parameter(Mandatory = $true)][string]$Decision)
    return ($Decision -in @('NewCacheValid','LegacyIntuneOwned','LegacyAdoptable'))
}
#endregion SEMT Intune setup media selection
