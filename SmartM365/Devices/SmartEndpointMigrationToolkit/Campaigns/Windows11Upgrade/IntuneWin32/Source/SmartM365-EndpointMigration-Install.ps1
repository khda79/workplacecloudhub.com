<#
.SYNOPSIS
    Installs the Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune endpoint package.
.DESCRIPTION
    Installs the endpoint, runner and helper scripts under %ProgramFiles% (writable by administrators only, owner and permissions verified), optionally copies packaged setup media or validates an existing cache under ProgramData (data only), registers package detection state, and starts a SYSTEM scheduled task that runs the runner from %ProgramFiles%.
    Uses the Smart Endpoint Migration Toolkit identities (data root, registry root, scheduled task) so it never collides with the previous Windows11UpgradeToolkit Intune package.
    Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/IntuneWin32/Source/Install.ps1 1.0.9 (see Docs\PROVENANCE.json.txt).
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [string]$DataRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade',
    # Scripts run as SYSTEM are installed here; empty means %ProgramFiles%\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade.
    [string]$ProgramRoot = '',
    [string]$ProgramFilesPath = $env:ProgramFiles,
    [string]$ProgramDataPath = $env:ProgramData,
    [string]$TaskName = 'SmartM365 SEMT Windows 11 Upgrade - Intune',
    [switch]$Uninstall
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:DataFolderTrusted = $false

$endpointScriptFileName = 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$runnerScriptFileName = 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1'
$integrityHelperFileName = 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
$packageRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$integrityHelperPackagePath = Join-Path $packageRoot $integrityHelperFileName
$manifestPath = Join-Path $packageRoot 'PackageManifest.json.txt'
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "PackageManifest.json.txt not found: $manifestPath" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($ProgramRoot)) { $ProgramRoot = Join-Path $ProgramFilesPath 'WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade' }
$programIntuneRoot = Join-Path $ProgramRoot 'Intune'
# Scripts of earlier qualification builds were copied here; they are removed and never run.
$legacyIntuneDataRoot = Join-Path $DataRoot 'Intune'
$logRoot = Join-Path $DataRoot 'Logs\Intune'
$setupCacheRoot = Join-Path $DataRoot 'SetupMedia'
$registrySubKeyRoot = 'SOFTWARE\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\IntunePackages'
$registrySubKey = "$registrySubKeyRoot\$([string]$manifest.PackageId)"
$packageMode = 'WithMedia'
if ($manifest.PSObject.Properties['PackageMode'] -and -not [string]::IsNullOrWhiteSpace([string]$manifest.PackageMode)) { $packageMode = [string]$manifest.PackageMode }
$requiresExistingSetupCache = $false
if ($manifest.PSObject.Properties['RequiresExistingSetupCache']) { $requiresExistingSetupCache = [bool]$manifest.RequiresExistingSetupCache }
$packageMediaRoot = Join-Path $packageRoot ("SetupMedia\{0}" -f $manifest.SetupCacheFolder)
$targetMediaRoot = Join-Path $setupCacheRoot ([string]$manifest.SetupCacheFolder)

#region SEMT trusted program folder v1
# Scripts that run as SYSTEM (scheduled tasks, Intune runner, retry after reboot) are installed
# under %ProgramFiles%, which only administrators can write. ProgramData keeps only data (logs,
# state, results, setup media). The same region is copied unchanged into both endpoint scripts,
# the Intune installer and the Intune runner (a test checks that all copies are identical).
#
# A folder is trusted when it and every item below it are owned by SYSTEM, Administrators or
# TrustedInstaller and no other principal can write, append, delete, change permissions or take
# ownership. Permissions that cannot be read are never trusted. An existing folder that is not
# trusted is never repaired automatically: the operation stops (SEMT_PROGRAM_FOLDER_UNTRUSTED).

$script:SemtProgramTrustedSids = @('S-1-5-18','S-1-5-32-544','S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')
$script:SemtProgramFolderSddl = 'O:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;0x1200a9;;;BU)'

function Get-SemtProgramRoot {
    param([Parameter(Mandatory = $true)][string]$Campaign, [string]$ProgramFilesPath = $env:ProgramFiles)
    if ([string]::IsNullOrWhiteSpace($ProgramFilesPath)) { throw 'SEMT_PROGRAM_FOLDER_UNTRUSTED: the Program Files folder is unknown.' }
    return (Join-Path $ProgramFilesPath ('WCH\SmartM365\SmartEndpointMigrationToolkit\{0}' -f $Campaign))
}

function Test-SemtTrustedProgramPath {
    # Read-only. State: Trusted, Untrusted (proven) or Unreadable (not proven either way).
    param([Parameter(Mandatory = $true)][string]$Path, [switch]$Recurse)
    # Write data/append/extended attributes, delete child, delete, WRITE_DAC, WRITE_OWNER and the
    # generic write/all bits that inherit-only entries may carry.
    $writeMask = 0x00000002 -bor 0x00000004 -bor 0x00000010 -bor 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x40000000 -bor 0x10000000
    try {
        $items = @(Get-Item -LiteralPath $Path -Force -ErrorAction Stop)
        if ($Recurse -and $items[0].PSIsContainer) { $items += @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Stop) }
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Program folder could not be read: {0}' -f $_.Exception.Message) } }
    foreach ($item in $items) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Reparse point in the program folder: {0}' -f $item.FullName) } }
        try {
            $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
            $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
            $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
        }
        catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Permissions of {0} could not be read: {1}' -f $item.FullName,$_.Exception.Message) } }
        if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
        if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
        foreach ($rule in $rules) {
            if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
            if (([int64]$rule.FileSystemRights -band $writeMask) -eq 0) { continue }
            $sid = $rule.IdentityReference.Value
            if ($script:SemtProgramTrustedSids -contains $sid) { continue }
            # CREATOR OWNER only applies to items created later, by principals that can already write.
            if ($sid -eq 'S-1-3-0' -and ($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
            return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can modify {1}.' -f $sid,$item.FullName) }
        }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Set-SemtProgramFolderAcl {
    # Owner Administrators; protected DACL: SYSTEM and Administrators full control, Users read and
    # execute. Every other entry is removed. Requires an elevated or SYSTEM context.
    param([Parameter(Mandatory = $true)][string]$Path)
    $security = New-Object System.Security.AccessControl.DirectorySecurity
    $security.SetSecurityDescriptorSddlForm($script:SemtProgramFolderSddl)
    Set-Acl -LiteralPath $Path -AclObject $security -ErrorAction Stop
}

function Initialize-SemtTrustedProgramFolder {
    # Creates the folder chain below Program Files with a protected ACL, or verifies an existing
    # chain. Throws SEMT_PROGRAM_FOLDER_UNTRUSTED when any component is not trusted afterwards.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramFiles)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $baseTrust = Test-SemtTrustedProgramPath -Path $base
    if (-not $baseTrust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $base,$baseTrust.State,$baseTrust.Detail) }
    $current = $base
    foreach ($segment in @($target.Substring($base.Length + 1).Split('\'))) {
        $current = Join-Path $current $segment
        if (-not (Test-Path -LiteralPath $current)) {
            New-Item -ItemType Directory -Path $current -ErrorAction Stop | Out-Null
            try { Set-SemtProgramFolderAcl -Path $current }
            catch {
                # A folder created here that cannot be protected is removed again.
                Remove-Item -LiteralPath $current -Force -ErrorAction SilentlyContinue
                throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $current,$_.Exception.Message)
            }
        }
        elseif (-not (Test-Path -LiteralPath $current -PathType Container)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} exists and is not a folder.' -f $current) }
        $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $target)
        if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $current,$trust.State,$trust.Detail) }
    }
    return $target
}

function Install-SemtTrustedProgramFile {
    # Copies through a staging name inside the trusted folder, verifies the SHA256 against the
    # source, renames, then verifies the hash and the permissions of the installed file.
    param([Parameter(Mandatory = $true)][string]$Source, [Parameter(Mandatory = $true)][string]$Destination)
    $expected = (Get-FileHash -LiteralPath $Source -Algorithm SHA256 -ErrorAction Stop).Hash
    $folder = Split-Path -Path $Destination -Parent
    $staging = Join-Path $folder ('.{0}.{1}.tmp' -f (Split-Path -Path $Destination -Leaf),([guid]::NewGuid().ToString('N')))
    try {
        Copy-Item -LiteralPath $Source -Destination $staging -ErrorAction Stop
        if ((Get-FileHash -LiteralPath $staging -Algorithm SHA256).Hash -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: staged copy of {0} differs from the source.' -f $Source) }
        Move-Item -LiteralPath $staging -Destination $Destination -Force -ErrorAction Stop
    }
    finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Force -ErrorAction SilentlyContinue } }
    $actual = (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
    if ($actual -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: {0} Expected={1}; Actual={2}' -f $Destination,$expected,$actual) }
    $trust = Test-SemtTrustedProgramPath -Path $Destination
    if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $Destination,$trust.State,$trust.Detail) }
    return [pscustomobject]@{ Path = $Destination; Sha256 = $actual }
}

# ProgramData data folders (logs, state, results, setup media whose setup.exe runs as SYSTEM).
# Scope: only the SEMT components (SmartEndpointMigrationToolkit and the folders below it) are
# ever changed. The shared parents above it (WCH, WCH\SmartM365), used by other products and by
# the legacy caches reused in place, are verified read-only and never changed; when they are not
# protected, the operation stops. A read-only pass checks the whole scope before the first
# change: shared parents, SEMT components, the provenance of existing content and, when a
# permission replacement is planned, every reparse point and hard link below it.
# A permission replacement never makes existing content trusted: an item that another principal
# could modify, or any content of a SEMT component that was not protected, has no verified
# provenance; it is refused and left unchanged (SEMT_DATA_FOLDER_UNTRUSTED), never adopted.
$script:SemtDataProductFolderName = 'SmartEndpointMigrationToolkit'

function Get-SemtHardLinkCount {
    # Number of names of a file; -1 when it cannot be read.
    param([Parameter(Mandatory = $true)][string]$Path)
    $fsutil = Join-Path $env:SystemRoot 'System32\fsutil.exe'
    $output = @(& $fsutil hardlink list $Path 2>$null)
    if ($LASTEXITCODE -ne 0) { return -1 }
    return @($output | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
}

function Assert-SemtNoLinkedItem {
    # Read-only. Before a DACL replacement propagates inheritance below Path: no reparse point
    # (never followed) and only files with a single name.
    param([Parameter(Mandatory = $true)][string]$Path)
    $pending = New-Object System.Collections.Generic.Queue[string]
    $pending.Enqueue([System.IO.Path]::GetFullPath($Path))
    while ($pending.Count -gt 0) {
        $folder = $pending.Dequeue()
        try { $children = @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop) }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $folder,$_.Exception.Message) }
        foreach ($child in $children) {
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: reparse point {0}; nothing was changed.' -f $child.FullName) }
            if ($child.PSIsContainer) { $pending.Enqueue($child.FullName); continue }
            $links = Get-SemtHardLinkCount -Path $child.FullName
            if ($links -ne 1) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} has {1} hard link name(s); nothing was changed.' -f $child.FullName,$links) }
        }
    }
}

function Test-SemtCanonicalFolderAcl {
    # True when the folder already has exactly the protected ACL that this region sets.
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        if ($acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value -ne 'S-1-5-32-544') { return $false }
        if (-not $acl.AreAccessRulesProtected) { return $false }
        $rules = @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]) | ForEach-Object { '{0}:{1}:{2}:{3}' -f $_.IdentityReference.Value,[int64]$_.FileSystemRights,[int]$_.InheritanceFlags,$_.AccessControlType })
        $expected = @('S-1-5-18:2032127:3:Allow','S-1-5-32-544:2032127:3:Allow','S-1-5-32-545:1179817:3:Allow')
        return (@($rules | Sort-Object) -join '|') -eq (@($expected | Sort-Object) -join '|')
    }
    catch { return $false }
}

function Test-SemtTrustedSharedParent {
    # Read-only; a shared parent is never changed. Trusted when its owner is SYSTEM,
    # Administrators or TrustedInstaller and no other principal can delete it, delete its
    # children, change its permissions or take ownership. Creating new items (other products) is
    # allowed: it cannot alter the existing SEMT folder below.
    param([Parameter(Mandatory = $true)][string]$Path)
    # Delete child, delete, WRITE_DAC, WRITE_OWNER and generic all.
    $mask = 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x10000000
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
        $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
        $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Shared folder {0} could not be read: {1}' -f $Path,$_.Exception.Message) } }
    if (-not $item.PSIsContainer) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is not a folder.' -f $item.FullName) } }
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is a reparse point.' -f $item.FullName) } }
    if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
    if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
        # Inherit-only entries apply to new children only, not to this folder.
        if (($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
        if (([int64]$rule.FileSystemRights -band $mask) -eq 0) { continue }
        if ($script:SemtProgramTrustedSids -contains $rule.IdentityReference.Value) { continue }
        return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can delete, rename or change the permissions of {1}.' -f $rule.IdentityReference.Value,$item.FullName) }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Assert-SemtDataFolderAnchor {
    # The data folder must be below the base inside a SmartEndpointMigrationToolkit folder, and
    # the base and its parents must not be reparse points.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$BasePath)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $segments = @($target.Substring($base.Length + 1).Split('\'))
    $productIndex = -1
    for ($i = 0; $i -lt $segments.Count; $i++) { if ($segments[$i] -ieq $script:SemtDataProductFolderName) { $productIndex = $i; break } }
    if ($productIndex -lt 0 -or $productIndex -ge ($segments.Count - 1)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not a campaign folder below {1}.' -f $target,$script:SemtDataProductFolderName) }
    $current = $base
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}' -f $current,$_.Exception.Message) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point.' -f $current) }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    return [pscustomobject]@{ Base = $base; Target = $target; Segments = $segments; ProductIndex = $productIndex }
}

function Get-SemtDataFolderPlan {
    # Read-only pass over the whole scope. Returns one step per component (Keep, Create or
    # Replace); throws SEMT_DATA_FOLDER_UNTRUSTED, with nothing changed, when the operation
    # cannot be done within the SEMT scope or when existing content has no verified provenance.
    param([Parameter(Mandatory = $true)]$Anchor)
    $plan = New-Object System.Collections.Generic.List[object]
    $current = $Anchor.Base
    $missing = $false
    for ($index = 0; $index -lt $Anchor.Segments.Count; $index++) {
        $current = Join-Path $current $Anchor.Segments[$index]
        $shared = $index -lt $Anchor.ProductIndex
        if ($missing -or -not (Test-Path -LiteralPath $current)) {
            $missing = $true
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $shared; Action = 'Create'; WasUntrusted = $false })
            continue
        }
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
        if (-not $item.PSIsContainer) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} exists and is not a folder; nothing was changed.' -f $current) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point; nothing was changed.' -f $current) }
        if ($shared) {
            $trust = Test-SemtTrustedSharedParent -Path $current
            if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} ({1}) {2} It is used by other products and is never changed; nothing was changed.' -f $current,$trust.State,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $true; Action = 'Keep'; WasUntrusted = $false })
            continue
        }
        $trust = Test-SemtTrustedProgramPath -Path $current
        if ($trust.State -eq 'Unreadable') { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} (Unreadable) {1}; nothing was changed.' -f $current,$trust.Detail) }
        if (-not $trust.Trusted) {
            # Content of a SEMT folder that was not protected has no verified provenance.
            try { $children = @(Get-ChildItem -LiteralPath $current -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
            if ($children.Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} was not protected ({1}) and is not empty; its content has no verified provenance and is not used; nothing was changed.' -f $current,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = 'Replace'; WasUntrusted = $true })
            continue
        }
        $action = if (Test-SemtCanonicalFolderAcl -Path $current) { 'Keep' } else { 'Replace' }
        $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = $action; WasUntrusted = $false })
    }
    # Existing content must already be trusted, item by item: below the data root and, when a
    # permission replacement is planned, everywhere it propagates to (other SEMT campaigns
    # included), so that the replacement never makes untrusted content look trusted.
    $firstReplace = @($plan | Where-Object { $_.Action -eq 'Replace' } | Select-Object -First 1)
    $existing = @($plan | Where-Object { $_.Action -ne 'Create' -and -not $_.Shared })
    $checkFrom = if ($firstReplace.Count -gt 0) { $firstReplace[0].Path } elseif (-not $missing) { $Anchor.Target } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($checkFrom)) {
        $chain = @($existing | ForEach-Object { $_.Path })
        $start = [array]::IndexOf([string[]]$chain, [string]$checkFrom)
        for ($i = $start; $i -lt $chain.Count; $i++) {
            try { $children = @(Get-ChildItem -LiteralPath $chain[$i] -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $chain[$i],$_.Exception.Message) }
            foreach ($child in $children) {
                if ($i -lt ($chain.Count - 1) -and $child.FullName -ieq $chain[$i + 1]) { continue }
                $trust = Test-SemtTrustedProgramPath -Path $child.FullName -Recurse
                if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance and is not used ({0}) {1}; nothing was changed.' -f $trust.State,$trust.Detail) }
            }
        }
    }
    # A DACL replacement propagates inheritance below the folder: the whole subtree of the first
    # replaced component is checked for reparse points and hard links before any change.
    if ($firstReplace.Count -gt 0) { Assert-SemtNoLinkedItem -Path $firstReplace[0].Path }
    return $plan
}

function Initialize-SemtTrustedDataFolder {
    # Protects the SEMT data folder chain below ProgramData and verifies it. Throws
    # SEMT_DATA_FOLDER_UNTRUSTED when it cannot be made trusted within the SEMT scope.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath
    $plan = Get-SemtDataFolderPlan -Anchor $anchor
    foreach ($step in $plan) {
        if ($step.Action -eq 'Create') {
            New-Item -ItemType Directory -Path $step.Path -ErrorAction Stop | Out-Null
            if ($step.Shared) {
                # Created with inherited permissions only: a shared folder is never changed.
                $trust = Test-SemtTrustedSharedParent -Path $step.Path
                if (-not $trust.Trusted) {
                    Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                    throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} created here is not protected ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail)
                }
                continue
            }
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch {
                Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                throw ('SEMT_DATA_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $step.Path,$_.Exception.Message)
            }
        }
        elseif ($step.Action -eq 'Replace') {
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: owner and permissions of {0} could not be replaced: {1}' -f $step.Path,$_.Exception.Message) }
        }
        else { continue }
        # A folder created here, or one that was not protected (empty in the read-only pass), must
        # still be empty once protected: nothing written meanwhile is adopted.
        if ($step.Action -eq 'Create' -or $step.WasUntrusted) {
            if (@(Get-ChildItem -LiteralPath $step.Path -Force -ErrorAction Stop).Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: content appeared in {0} before it was protected; it is not used.' -f $step.Path) }
        }
        $trust = Test-SemtTrustedProgramPath -Path $step.Path
        if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail) }
    }
    $final = Test-SemtTrustedDataFolder -Path $anchor.Target -BasePath $anchor.Base
    if (-not $final.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $anchor.Target,$final.State,$final.Detail) }
    return $anchor.Target
}

function Test-SemtTrustedDataFolder {
    # Read-only verification used before anything is run: shared parents protected, every SEMT
    # component trusted, and the data root trusted with everything below it.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    try { $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = $_.Exception.Message } }
    $current = $anchor.Base
    for ($index = 0; $index -lt $anchor.Segments.Count; $index++) {
        $current = Join-Path $current $anchor.Segments[$index]
        if ($index -lt $anchor.ProductIndex) { $trust = Test-SemtTrustedSharedParent -Path $current }
        else { $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $anchor.Target) }
        if (-not $trust.Trusted) { return $trust }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}
#endregion SEMT trusted program folder v1

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Protect-ToolkitPathAcl {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$Directory
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }
    $icacls = Join-Path $env:SystemRoot 'System32\icacls.exe'
    if (-not (Test-Path -LiteralPath $icacls -PathType Leaf)) {
        Write-InstallLog ("ACL hardening skipped because icacls.exe was not found. Path={0}" -f $Path) 'WARN'
        return
    }

    $grants = if ($Directory) {
        @('*S-1-5-18:(OI)(CI)F','*S-1-5-32-544:(OI)(CI)F','*S-1-5-32-545:(OI)(CI)RX')
    }
    else {
        @('*S-1-5-18:F','*S-1-5-32-544:F','*S-1-5-32-545:RX')
    }

    $output = & $icacls $Path '/inheritance:r' '/grant:r' $grants 2>&1
    $exitCode = [int]$LASTEXITCODE
    if ($exitCode -ne 0) {
        Write-InstallLog ("ACL hardening failed. Path={0}; ExitCode={1}; Output={2}" -f $Path,$exitCode,(($output | Out-String).Trim())) 'WARN'
        return
    }

    Write-InstallLog ("ACL hardening applied. Path={0}; Directory={1}" -f $Path,[bool]$Directory)
}

function Open-Registry64LocalMachine {
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
}

function Set-Registry64String {
    param(
        [string]$SubKey,
        [string]$Name,
        [string]$Value
    )

    $baseKey = Open-Registry64LocalMachine
    try {
        $key = $baseKey.CreateSubKey($SubKey)
        try { $key.SetValue($Name, [string]$Value, [Microsoft.Win32.RegistryValueKind]::String) }
        finally { if ($key) { $key.Dispose() } }
    }
    finally { $baseKey.Dispose() }
}

function Remove-Registry64SubKeyTree {
    param([string]$SubKey)

    $baseKey = Open-Registry64LocalMachine
    try { $baseKey.DeleteSubKeyTree($SubKey, $false) }
    finally { $baseKey.Dispose() }
}

function Write-InstallLog {
    param([string]$Message, [string]$Level = 'INFO')
    # Nothing is written under the data root before its folder chain is trusted.
    if (-not $script:DataFolderTrusted) { Write-Output ('{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message); return }
    New-Directory -Path $logRoot
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -LiteralPath (Join-Path $logRoot 'SEMT-Intune-Install.log') -Value $line -Encoding UTF8
}

function Get-OsFamily {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $build = [int]$os.BuildNumber
        if ($build -ge 22000 -or ([string]$os.Caption) -match 'Windows 11') { return 'Windows11' }
        if ([string]$os.Caption -match 'Windows 10') { return 'Windows10' }
        return 'Other'
    }
    catch { return 'Unknown' }
}

function Get-DirectorySizeBytes {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return [int64]0 }
    $sum = [int64]0
    foreach ($file in @(Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue)) { $sum += [int64]$file.Length }
    return $sum
}

function Get-SetupCacheLockName {
    param([Parameter(Mandatory = $true)][string]$CachePath)

    $leaf = Split-Path -Path $CachePath -Leaf
    if ([string]::IsNullOrWhiteSpace($leaf)) { $leaf = 'Default' }
    $safe = $leaf -replace '[^A-Za-z0-9._-]', '_'
    return "SetupCache-$safe.lock"
}

function Test-SetupCacheLockProcessAlive {
    param([AllowNull()][object]$ProcessId)

    $pidValue = 0
    if ($null -eq $ProcessId -or -not [int]::TryParse([string]$ProcessId, [ref]$pidValue) -or $pidValue -le 0) { return $false }
    try { [void](Get-Process -Id $pidValue -ErrorAction Stop); return $true } catch { return $false }
}

function New-SetupCacheLockPayload {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][string]$LockPath,
        [Parameter(Mandatory = $true)][string]$OwnerToken,
        [Parameter(Mandatory = $true)][string]$Purpose,
        [Parameter(Mandatory = $true)][int]$LeaseMinutes
    )

    $nowUtc = (Get-Date).ToUniversalTime()
    $processName = ''
    try { $processName = (Get-Process -Id $PID -ErrorAction Stop).ProcessName } catch { }
    return [pscustomobject]@{
        OwnerToken = $OwnerToken
        Owner = 'SEMT_Windows11Upgrade_IntuneInstall'
        ComputerName = $env:COMPUTERNAME
        PID = $PID
        ProcessName = $processName
        Purpose = $Purpose
        CachePath = $CachePath
        LockPath = $LockPath
        StartedUtc = $nowUtc.ToString('o')
        LastSeenUtc = $nowUtc.ToString('o')
        ExpiresUtc = $nowUtc.AddMinutes($LeaseMinutes).ToString('o')
    }
}

function Read-SetupCacheLockPayload {
    param([Parameter(Mandatory = $true)][string]$LockPath)

    $payloadPath = Join-Path $LockPath 'lock.json.txt'
    if (-not (Test-Path -LiteralPath $payloadPath -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $payloadPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop) } catch { return $null }
}

function Save-SetupCacheLockPayload {
    param(
        [Parameter(Mandatory = $true)][string]$LockPath,
        [Parameter(Mandatory = $true)][object]$Payload
    )

    $payloadPath = Join-Path $LockPath 'lock.json.txt'
    $Payload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $payloadPath -Encoding UTF8
}

function Format-SetupCacheLockPayload {
    param([AllowNull()][object]$Payload)

    if ($null -eq $Payload) { return 'LockPayload=<missing>' }
    return ("Owner={0}; PID={1}; Process={2}; Purpose={3}; StartedUtc={4}; LastSeenUtc={5}; ExpiresUtc={6}" -f $Payload.Owner,$Payload.PID,$Payload.ProcessName,$Payload.Purpose,$Payload.StartedUtc,$Payload.LastSeenUtc,$Payload.ExpiresUtc)
}

function Acquire-SetupCacheLock {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][string]$Purpose,
        [int]$LeaseMinutes = 90
    )

    $lockRoot = Join-Path $DataRoot 'Locks'
    New-Directory -Path $lockRoot
    $lockPath = Join-Path $lockRoot (Get-SetupCacheLockName -CachePath $CachePath)
    $ownerToken = [guid]::NewGuid().ToString('N')

    while ($true) {
        try {
            New-Item -ItemType Directory -Path $lockPath -ErrorAction Stop | Out-Null
            $payload = New-SetupCacheLockPayload -CachePath $CachePath -LockPath $lockPath -OwnerToken $ownerToken -Purpose $Purpose -LeaseMinutes $LeaseMinutes
            Save-SetupCacheLockPayload -LockPath $lockPath -Payload $payload
            Write-InstallLog ("Acquired setup cache lock: CachePath={0}; LeaseMinutes={1}; Lock={2}; {3}" -f $CachePath,$LeaseMinutes,$lockPath,(Format-SetupCacheLockPayload -Payload $payload))
            return [pscustomobject]@{ LockPath = $lockPath; OwnerToken = $ownerToken; CachePath = $CachePath }
        }
        catch {
            $payload = Read-SetupCacheLockPayload -LockPath $lockPath
            $nowUtc = (Get-Date).ToUniversalTime()
            if ($null -eq $payload) {
                $lockItem = $null
                try { $lockItem = Get-Item -LiteralPath $lockPath -ErrorAction Stop } catch { }
                if ($null -ne $lockItem -and $lockItem.LastWriteTimeUtc -lt $nowUtc.AddMinutes(-5)) {
                    Write-InstallLog ("Reclaiming empty stale setup cache lock: CachePath={0}; Lock={1}; LastWriteTimeUtc={2}" -f $CachePath,$lockPath,$lockItem.LastWriteTimeUtc.ToString('o')) 'WARN'
                    Remove-Item -LiteralPath $lockPath -Recurse -Force -ErrorAction SilentlyContinue
                    continue
                }
                throw ("Setup cache is locked by another SmartM365 process. CachePath={0}; Lock={1}; {2}" -f $CachePath,$lockPath,(Format-SetupCacheLockPayload -Payload $payload))
            }
            $expiresUtc = [datetime]::MinValue
            $expired = $true
            if ($null -ne $payload -and [datetime]::TryParse([string]$payload.ExpiresUtc, [ref]$expiresUtc)) { $expired = ($expiresUtc.ToUniversalTime() -lt $nowUtc) }
            $processAlive = if ($null -ne $payload) { Test-SetupCacheLockProcessAlive -ProcessId $payload.PID } else { $false }

            if ($expired -and -not $processAlive) {
                Write-InstallLog ("Reclaiming stale setup cache lock: CachePath={0}; Lock={1}; {2}" -f $CachePath,$lockPath,(Format-SetupCacheLockPayload -Payload $payload)) 'WARN'
                Remove-Item -LiteralPath $lockPath -Recurse -Force -ErrorAction SilentlyContinue
                continue
            }

            throw ("Setup cache is locked by another SmartM365 process. CachePath={0}; Lock={1}; {2}" -f $CachePath,$lockPath,(Format-SetupCacheLockPayload -Payload $payload))
        }
    }
}

function Release-SetupCacheLock {
    param([AllowNull()][object]$Lock)

    if ($null -eq $Lock) { return }
    $payload = Read-SetupCacheLockPayload -LockPath $Lock.LockPath
    if ($null -ne $payload -and [string]$payload.OwnerToken -ne [string]$Lock.OwnerToken) {
        Write-InstallLog ("Skipping setup cache lock release because owner changed. Lock={0}; {1}" -f $Lock.LockPath,(Format-SetupCacheLockPayload -Payload $payload)) 'WARN'
        return
    }
    Remove-Item -LiteralPath $Lock.LockPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-InstallLog ("Released setup cache lock: CachePath={0}; Lock={1}" -f $Lock.CachePath,$Lock.LockPath)
}
function Set-PackageDetectionState {
    param([string]$InstallState = 'Installed')

    Write-InstallLog "Writing Intune detection registry state to HKLM:\$registrySubKey (64-bit registry view). InstallState=$InstallState"
    Set-Registry64String -SubKey $registrySubKey -Name PackageId -Value ([string]$manifest.PackageId)
    Set-Registry64String -SubKey $registrySubKey -Name PackageVersion -Value ([string]$manifest.PackageVersion)
    Set-Registry64String -SubKey $registrySubKey -Name Language -Value ([string]$manifest.Language)
    Set-Registry64String -SubKey $registrySubKey -Name MediaId -Value ([string]$manifest.MediaId)
    Set-Registry64String -SubKey $registrySubKey -Name SetupCacheFolder -Value ([string]$manifest.SetupCacheFolder)
    Set-Registry64String -SubKey $registrySubKey -Name PackageMode -Value $packageMode
    Set-Registry64String -SubKey $registrySubKey -Name InstallState -Value $InstallState
    Set-Registry64String -SubKey $registrySubKey -Name InstalledUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))
    $media = Get-Variable -Name SelectedMedia -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($media) {
        Set-Registry64String -SubKey $registrySubKey -Name SetupMediaPath -Value ([string]$media.Path)
        Set-Registry64String -SubKey $registrySubKey -Name SetupMediaSource -Value ([string]$media.Source)
        Set-Registry64String -SubKey $registrySubKey -Name SetupMediaExternal -Value ([string]([bool]$media.External))
    }
    if ($InstallState -eq 'Installed') {
        Set-Registry64String -SubKey $registrySubKey -Name RepairReason -Value ''
        Set-Registry64String -SubKey $registrySubKey -Name RepairRequiredUtc -Value ''
    }
}

function Test-SetupCacheReady {
    param([Parameter(Mandatory = $true)][string]$Path)

    $setupExe = Join-Path $Path 'setup.exe'
    $installWim = Join-Path $Path 'sources\install.wim'
    if (-not (Test-Path -LiteralPath $setupExe -PathType Leaf)) { throw "Local setup cache is missing setup.exe: $Path" }
    if (-not (Test-Path -LiteralPath $installWim -PathType Leaf)) { throw "Local setup cache is missing sources\install.wim: $Path" }
    $integrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $Path
    Write-InstallLog ("Setup cache integrity validated. Files={0}; Bytes={1}; Root={2}" -f $integrity.Files,$integrity.Bytes,$integrity.MediaRoot)
    return $true
}

function Remove-InstalledProgramFiles {
    # Removes only the files this package installed under Program Files, then the empty folders.
    foreach ($path in @((Join-Path $programIntuneRoot $runnerScriptFileName),(Join-Path $programIntuneRoot $integrityHelperFileName),(Join-Path $programIntuneRoot 'PackageManifest.json.txt'),(Join-Path $ProgramRoot $endpointScriptFileName))) {
        if (Test-Path -LiteralPath $path -PathType Leaf) { Remove-Item -LiteralPath $path -Force -ErrorAction Stop }
    }
    foreach ($folder in @($programIntuneRoot,$ProgramRoot)) {
        if ((Test-Path -LiteralPath $folder -PathType Container) -and @(Get-ChildItem -LiteralPath $folder -Force).Count -eq 0) { Remove-Item -LiteralPath $folder -Force -ErrorAction Stop }
    }
}

function Remove-LegacyDataScripts {
    # Earlier qualification builds copied the runner, helper and manifest under ProgramData.
    foreach ($name in @($runnerScriptFileName,$integrityHelperFileName,'PackageManifest.json.txt')) {
        $path = Join-Path $legacyIntuneDataRoot $name
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop; Write-InstallLog ("Removed a script copy from ProgramData (scripts now run from Program Files only): {0}" -f $path) 'WARN' }
            catch { Write-InstallLog ("Could not remove a script copy from ProgramData; it is no longer referenced by any task. Path={0}; Error={1}" -f $path,$_.Exception.Message) 'WARN' }
        }
    }
}

# SEMT data folder boundary: owner Administrators and a protected DACL on the SEMT components of
# the data folder chain after a read-only check of the whole scope (shared parents verified, never
# changed; content without verified provenance refused), then verification; failure blocks the install.
$script:DataFolderTrusted = $false
$dataFolderError = ''
try {
    [void](Initialize-SemtTrustedDataFolder -Path $DataRoot -BasePath $ProgramDataPath)
    $script:DataFolderTrusted = $true
}
catch { $dataFolderError = $_.Exception.Message }

if ($Uninstall) {
    Write-InstallLog "Uninstall requested for $($manifest.PackageId)."
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    Remove-InstalledProgramFiles
    Remove-LegacyDataScripts
    Remove-Registry64SubKeyTree -SubKey $registrySubKey
    exit 0
}

if (-not $script:DataFolderTrusted) {
    Write-InstallLog ("SEMT_DATA_FOLDER_UNTRUSTED: the data folder chain could not be made trusted; nothing was installed, copied or registered. {0}" -f $dataFolderError) 'ERROR'
    exit 1
}
Write-InstallLog "Installing $($manifest.PackageId) version $($manifest.PackageVersion)."
Write-InstallLog ("Data folder chain trusted (owner Administrators, protected DACL, verified): {0}" -f $DataRoot)
# ProgramData keeps only data (logs, state, results, setup media); no script there is ever run.
# The subfolders inherit the protected DACL of the data root.
New-Directory -Path $setupCacheRoot
New-Directory -Path $logRoot

if ((Get-OsFamily) -eq 'Windows11') {
    Write-InstallLog 'Device is already Windows 11 during package install. Writing detection state, removing scheduled task if present, and exiting success.'
    Set-PackageDetectionState -InstallState 'Installed'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    exit 0
}

if (-not (Test-Path -LiteralPath $integrityHelperPackagePath -PathType Leaf)) { throw "Setup media integrity helper not found: $integrityHelperPackagePath" }
. $integrityHelperPackagePath

# Scripts run as SYSTEM go to Program Files: the folder chain is created with a protected ACL
# (owner Administrators) or verified; any untrusted component stops the install before any media
# copy or task registration. Each copy is verified against the package hash.
try {
    [void](Initialize-SemtTrustedProgramFolder -Path $programIntuneRoot -BasePath $ProgramFilesPath)
    $installedFiles = @(
        (Install-SemtTrustedProgramFile -Source (Join-Path $packageRoot $endpointScriptFileName) -Destination (Join-Path $ProgramRoot $endpointScriptFileName)),
        (Install-SemtTrustedProgramFile -Source (Join-Path $packageRoot $runnerScriptFileName) -Destination (Join-Path $programIntuneRoot $runnerScriptFileName)),
        (Install-SemtTrustedProgramFile -Source $integrityHelperPackagePath -Destination (Join-Path $programIntuneRoot $integrityHelperFileName)),
        (Install-SemtTrustedProgramFile -Source $manifestPath -Destination (Join-Path $programIntuneRoot 'PackageManifest.json.txt'))
    )
    $programTrust = Test-SemtTrustedProgramPath -Path $ProgramRoot -Recurse
    if (-not $programTrust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $ProgramRoot,$programTrust.State,$programTrust.Detail) }
}
catch {
    Write-InstallLog ("SEMT_PROGRAM_FOLDER_UNTRUSTED: scripts were not installed and no task was registered. {0}" -f $_.Exception.Message) 'ERROR'
    exit 1
}
foreach ($installed in $installedFiles) { Write-InstallLog ("Installed program file: {0}; SHA256={1}" -f $installed.Path,$installed.Sha256) }
Remove-LegacyDataScripts

# Media selection before any copy (shared rules with the runner and the endpoint):
# - NewCacheValid: the cache of this product is used; nothing is copied.
# - LegacyIntuneOwned: media of the previous toolkit still owned by its Intune package is used in
#   place, read-only and external to this product (no move, no ACL or marker change, no cleanup).
# - LegacyAdoptable: idle unmanaged legacy media; nothing is copied, the endpoint moves it with
#   rollback protection on its first run.
# - InUse / Indeterminate: nothing is changed; Intune retries (1618) or the install fails explicitly
#   (unreadable permissions or folder tree of legacy media are Indeterminate).
# - Untrusted: legacy media present but not protected; never used, its permissions are not changed;
#   SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED unless the package allows duplication (WithMedia only).
# - Absent / Corrupt / Incompatible: WithMedia copies its packaged media (unless legacy media that
#   remains would be duplicated without authorization); WithCacheOnly fails because it carries no media.
$referenceManifest = if ($requiresExistingSetupCache) { '' } else { Join-Path $packageMediaRoot 'SmartM365-SetupMediaManifest.sha256.csv' }
$allowDuplication = ($manifest.PSObject.Properties['AllowSetupMediaDuplication'] -and [bool]$manifest.AllowSetupMediaDuplication)
$mediaResolution = Resolve-SemtIntuneSetupMedia -DataRoot $DataRoot -SetupCacheFolder ([string]$manifest.SetupCacheFolder) -Language ([string]$manifest.Language) -ReferenceManifestPath $referenceManifest
Write-InstallLog ("Setup media resolution: Decision={0}; Path={1}; External={2}; Owner={3}; {4}" -f $mediaResolution.Decision,$mediaResolution.Path,$mediaResolution.External,$mediaResolution.Owner,$mediaResolution.Detail)
$script:SelectedMedia = [pscustomobject]@{ Path = $mediaResolution.Path; Source = $mediaResolution.Source; External = [bool]$mediaResolution.External; Decision = $mediaResolution.Decision }

switch ($mediaResolution.Decision) {
    'InUse' {
        Write-InstallLog ("Setup media is in use; nothing was changed. Intune will retry. {0}" -f $mediaResolution.LegacyDetail) 'WARN'
        exit 1618
    }
    'Indeterminate' {
        Write-InstallLog ("SEMT_SETUP_MEDIA_STATE_INDETERMINATE: setup media state could not be read; nothing was changed. {0}" -f $mediaResolution.LegacyDetail) 'ERROR'
        exit 1
    }
    'Untrusted' {
        # Present legacy media without sufficient protection: never used, permissions never changed.
        if (-not $allowDuplication) {
            Write-InstallLog ("SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED: legacy media at {0} is present but not protected against modification by non-administrators; it is not used, its permissions were not changed and nothing was copied ({1}). Secure or remove the legacy media, or build the package with -AllowSetupMediaDuplication." -f $mediaResolution.LegacyPath,$mediaResolution.LegacyDetail) 'ERROR'
            exit 1
        }
        if ($requiresExistingSetupCache) {
            Write-InstallLog ("SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED: legacy media at {0} is not protected and the cache-only package carries no media. Use the WithMedia package. {1}" -f $mediaResolution.LegacyPath,$mediaResolution.LegacyDetail) 'ERROR'
            exit 1
        }
    }
}

if (Test-SemtIntuneMediaDecisionUsable -Decision $mediaResolution.Decision) {
    if ($mediaResolution.Decision -eq 'NewCacheValid') { Protect-ToolkitPathAcl -Path $targetMediaRoot -Directory }
    else { Write-InstallLog ("Using media outside the cache of this product without copying it ({0}). Its permissions, markers and content are not changed. Path={1}" -f $mediaResolution.Decision,$mediaResolution.Path) 'WARN' }
}
elseif ($requiresExistingSetupCache) {
    Write-InstallLog ("SEMT_SETUP_MEDIA_NOT_REUSABLE: cache-only package and no reusable setup media (Decision={0}). Use the WithMedia package for this device. {1}" -f $mediaResolution.Decision,$mediaResolution.Detail) 'ERROR'
    exit 1
}
elseif (($mediaResolution.LegacyValidRemains -or $mediaResolution.DuplicationRequired) -and -not $allowDuplication) {
    Write-InstallLog ("SEMT_SETUP_MEDIA_DUPLICATION_BLOCKED: valid legacy media remains at {0} and is not reusable ({1}). Nothing was copied. Unassign the legacy Intune app, free the legacy media, or build the package with -AllowSetupMediaDuplication." -f $mediaResolution.LegacyPath,$mediaResolution.LegacyDetail) 'ERROR'
    exit 1
}
else {
    if ($mediaResolution.LegacyValidRemains -or $mediaResolution.DuplicationRequired) { Write-InstallLog ("SEMT_SETUP_MEDIA_DUPLICATION_ALLOWED: second setup media copy allowed by the package (-AllowSetupMediaDuplication) although legacy media remains at {0} (Decision={1}); the legacy media is not changed." -f $mediaResolution.LegacyPath,$mediaResolution.Decision) 'WARN' }
    $packageIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $packageMediaRoot
    Write-InstallLog ("Packaged setup media integrity validated before installation. Files={0}; Bytes={1}; Root={2}" -f $packageIntegrity.Files,$packageIntegrity.Bytes,$packageIntegrity.MediaRoot)
    if (-not (Test-Path -LiteralPath (Join-Path $packageMediaRoot 'setup.exe') -PathType Leaf)) { throw "Packaged setup.exe not found: $packageMediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $packageMediaRoot 'sources\install.wim') -PathType Leaf)) { throw "Packaged sources\install.wim not found: $packageMediaRoot" }

    $cacheLock = $null
    try {
        $cacheLock = Acquire-SetupCacheLock -CachePath $targetMediaRoot -Purpose ("IntunePackageInstall:{0}" -f $manifest.PackageId) -LeaseMinutes 90
        if (Test-Path -LiteralPath $targetMediaRoot -PathType Container) {
            Write-InstallLog ("Removing the invalid cache of this product before recopy ({0}): {1}" -f $mediaResolution.NewCacheState,$targetMediaRoot)
            Remove-Item -LiteralPath $targetMediaRoot -Recurse -Force -ErrorAction Stop
        }
        Write-InstallLog ("Copying packaged setup media to local cache (Decision={0}): {1} -> {2}" -f $mediaResolution.Decision,$packageMediaRoot,$targetMediaRoot)
        New-Directory -Path $targetMediaRoot
        $robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
        $installRobocopyLog = Join-Path $logRoot 'SEMT-Intune-Install-Robocopy.log'
        & $robocopy $packageMediaRoot $targetMediaRoot /MIR /R:2 /W:5 /NP /NFL /NDL "/LOG+:$installRobocopyLog" | Out-Null
        $copyExit = [int]$LASTEXITCODE
        Write-InstallLog "Robocopy setup media copy exit code: $copyExit; Log=$installRobocopyLog"
        if ($copyExit -gt 7) { throw "Robocopy install media copy failed with exit code $copyExit. Log=$installRobocopyLog" }
        [void](Test-SetupCacheReady -Path $targetMediaRoot)
        Protect-ToolkitPathAcl -Path $targetMediaRoot -Directory
        $script:SelectedMedia = [pscustomobject]@{ Path = $targetMediaRoot; Source = 'PackageCopy'; External = $false; Decision = 'NewCacheValid' }
    }
    catch {
        if ($_.Exception.Message -like 'Setup cache is locked by another SmartM365 process.*') {
            Write-InstallLog ("Setup cache is locked by another SmartM365 process. Intune will retry. Detail={0}" -f $_.Exception.Message) 'WARN'
            exit 1618
        }
        throw
    }
    finally {
        Release-SetupCacheLock -Lock $cacheLock
    }
}
Set-PackageDetectionState -InstallState 'Installed'

# Final verification before the SYSTEM task is registered: data chain (setup media included) and
# program folder.
$dataTrust = Test-SemtTrustedDataFolder -Path $DataRoot -BasePath $ProgramDataPath
$programTrust = Test-SemtTrustedProgramPath -Path $ProgramRoot -Recurse
if (-not $dataTrust.Trusted -or -not $programTrust.Trusted) {
    Write-InstallLog ("SEMT_DATA_FOLDER_UNTRUSTED: verification before task registration failed; no task was registered. Data={0} {1}; Program={2} {3}" -f $dataTrust.State,$dataTrust.Detail,$programTrust.State,$programTrust.Detail) 'ERROR'
    exit 1
}
$runner = Join-Path $programIntuneRoot $runnerScriptFileName
$runnerLog = Join-Path $logRoot 'SEMT-Intune-Runner.log'
$taskArgument = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -DataRoot "{1}" -TaskName "{2}"' -f $runner,$DataRoot,$TaskName
$powerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$action = New-ScheduledTaskAction -Execute $powerShellExe -Argument $taskArgument
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddHours(2) -RepetitionInterval (New-TimeSpan -Hours 2) -RepetitionDuration (New-TimeSpan -Days 30)
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-InstallLog ("Scheduled task action: {0} {1}" -f $powerShellExe,$taskArgument)
Write-InstallLog "Runner log expected at: $runnerLog"
Start-ScheduledTask -TaskName $TaskName

Write-InstallLog "Install completed and scheduled task started: $TaskName; next scheduled retry is every 2 hours."
exit 0

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCXqotM9LS9u2gb
# dXBkpbq3MpMzZZ06YOFag6HIyYl4xKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB9zrLrdocnPEKrQPXoLKaJ
# BOmwjnyTHJ3Q6rL3UvNMaTANBgkqhkiG9w0BAQEFAASCAYAQuL5kuZacHSuI6Ft7
# +OTaGsLMerdoxK+iO5mkUwS9BfRTiZMyTClDXsAMQYan6hPcBsR9VQVL5wkZwkea
# pQLhTZ44r54+H4J/u3I/JUMtRaCDpEL95smMeEM7xdAzNkeL4HTTh2jHlWxmnNdj
# LtiiYJ1rMuLZTi1q3unNVTvkgNK6OJy5axE/R2nMUi29QucA5h4oEmvNhuJmlJmB
# H1Q4AD3d8M2xYi2w22ZopILtE6T2KzPb0Y3h/aJGWS/JSjT57FPVrlJ7z/1Qp7/w
# V1YVuoqTh50WmIUdUHrv/m0NmmA9kL/Z6rxWJRwNRMUv9HobOyqogWbX6VJMavy7
# jiXKLUlCvXSeIaC6mECHql6ekrcJ0sYTXVPPGS/xUHNvOizB14L1H2zmIL5bwuH+
# FJY6q+rC6g/mgshTVIxReodxslcr9JZeC1hdxLXX5rBepKeVA2oLVvDHxNj3ETrn
# pV+jqb26z2WfTXlP7WeMDjZfqnTFlszDCPb+BW070WEDCTI=
# SIG # End signature block
