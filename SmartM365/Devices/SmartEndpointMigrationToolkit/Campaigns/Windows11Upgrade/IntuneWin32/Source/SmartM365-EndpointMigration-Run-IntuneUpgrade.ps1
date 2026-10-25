<#
.SYNOPSIS
    Runs the Smart Endpoint Migration Toolkit Windows 11 Upgrade endpoint from an Intune scheduled task.
.DESCRIPTION
    Executes the endpoint script installed under %ProgramFiles% against the packaged setup media cache and removes the scheduled task once the device is already Windows 11.
    The runner, its helper, the package manifest and the endpoint are loaded from its own Program Files folder only after that folder is verified (owner and permissions); ProgramData holds only data.
    Each run passes a new CorrelationId to the endpoint and reads the attempt outcome from <DataRoot>\State\Results\<CorrelationId>.json, falling back to LastRun.json.
    Legacy Windows11UpgradeToolkit setup media is reused by the endpoint itself (-ReuseLegacySetupMedia); this runner never moves or cleans it.
    Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/IntuneWin32/Source/Run-IntuneUpgrade.ps1 1.0.11 (see Docs\PROVENANCE.json.txt).
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [string]$DataRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade',
    [string]$TaskName = 'SmartM365 SEMT Windows 11 Upgrade - Intune',
    [string]$ProgramDataPath = $env:ProgramData,
    [int]$RunGuardHours = 2,
    [ValidateRange(1, 365)][int]$LogRetentionDays = 7,
    [ValidateRange(10, 5000)][int]$MaxEndpointLogFiles = 200
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:DataFolderTrusted = $false

# <ProgramRoot>\Intune\<runner>, <ProgramRoot>\Intune\<helper and manifest>, <ProgramRoot>\<endpoint>.
$programIntuneRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$programRoot = Split-Path -Parent $programIntuneRoot
$logRoot = Join-Path $DataRoot 'Logs\Intune'
$manifestPath = Join-Path $programIntuneRoot 'PackageManifest.json'
$integrityHelperPath = Join-Path $programIntuneRoot 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
$endpointScript = Join-Path $programRoot 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$setupCacheRoot = Join-Path $DataRoot 'SetupMedia'
$resultsRoot = Join-Path $DataRoot 'State\Results'
$lastRunPath = Join-Path $DataRoot 'LastRun.json'
$registrySubKeyRoot = 'SOFTWARE\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\IntunePackages'

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

function Invoke-IntuneLogMaintenance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 365)][int]$RetentionDays = 7,
        [ValidateRange(1, 5000)][int]$MaximumEndpointFiles = 200
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
    $removedByAge = 0
    $removedByCount = 0
    $failed = 0
    $files = @(Get-ChildItem -LiteralPath $Path -Filter 'Endpoint_*.log' -File -ErrorAction Stop)

    foreach ($file in @($files | Where-Object { $_.LastWriteTime -lt $cutoff })) {
        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $removedByAge++
        }
        catch { $failed++ }
    }

    $remaining = @(Get-ChildItem -LiteralPath $Path -Filter 'Endpoint_*.log' -File -ErrorAction Stop |
        Sort-Object LastWriteTime,Name -Descending)
    foreach ($file in @($remaining | Select-Object -Skip $MaximumEndpointFiles)) {
        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $removedByCount++
        }
        catch { $failed++ }
    }

    $remainingCount = @(Get-ChildItem -LiteralPath $Path -Filter 'Endpoint_*.log' -File -ErrorAction Stop).Count
    return [pscustomobject]@{
        Scanned        = $files.Count
        RemovedByAge   = $removedByAge
        RemovedByCount = $removedByCount
        Failed         = $failed
        Remaining      = $remainingCount
    }
}

function Write-RunnerLog {
    param([string]$Message, [string]$Level = 'INFO')
    # Nothing is written under the data root before its folder chain is verified.
    if (-not $script:DataFolderTrusted) { Write-Output ('{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message); return }
    New-Directory -Path $logRoot
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -LiteralPath (Join-Path $logRoot 'SEMT-Intune-Runner.log') -Value $line -Encoding UTF8
}

function Get-EndpointAttemptOutcome {
    param(
        [Parameter(Mandatory = $true)][string]$ResultsRoot,
        [Parameter(Mandatory = $true)][string]$LastRunPath,
        [Parameter(Mandatory = $true)][string]$CorrelationId
    )

    $resultPath = Join-Path $ResultsRoot ('{0}.json' -f $CorrelationId)
    if (Test-Path -LiteralPath $resultPath -PathType Leaf) {
        try {
            $result = Get-Content -LiteralPath $resultPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($result -and [string]$result.CorrelationId -eq $CorrelationId -and [string]$result.State -eq 'Completed') {
                return [pscustomobject]@{ Source = 'AttemptResult'; Path = $resultPath; Completed = $true; Status = [string]$result.Status; ExitCode = [string]$result.ExitCode; NextAction = [string]$result.NextAction }
            }
            $state = if ($result) { [string]$result.State } else { '' }
            return [pscustomobject]@{ Source = 'AttemptResult'; Path = $resultPath; Completed = $false; Status = ('ATTEMPT_NOT_COMPLETED:{0}' -f $state); ExitCode = ''; NextAction = '' }
        }
        catch { }
    }

    if (Test-Path -LiteralPath $LastRunPath -PathType Leaf) {
        try {
            $lastRun = Get-Content -LiteralPath $LastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $status = ''
            $exitCode = ''
            $nextAction = ''
            if ($lastRun.PSObject.Properties['Status']) { $status = [string]$lastRun.Status }
            if ($lastRun.PSObject.Properties['ExitCode']) { $exitCode = [string]$lastRun.ExitCode }
            if ($lastRun.PSObject.Properties['NextAction']) { $nextAction = [string]$lastRun.NextAction }
            $sameAttempt = ($lastRun.PSObject.Properties['CorrelationId'] -and [string]$lastRun.CorrelationId -eq $CorrelationId)
            return [pscustomobject]@{ Source = 'LastRun'; Path = $LastRunPath; Completed = [bool]$sameAttempt; Status = $status; ExitCode = $exitCode; NextAction = $nextAction }
        }
        catch { }
    }

    return [pscustomobject]@{ Source = 'None'; Path = ''; Completed = $false; Status = ''; ExitCode = ''; NextAction = '' }
}

function Invoke-RunnerLogMaintenance {
    param([string]$Phase)

    try {
        $result = Invoke-IntuneLogMaintenance -Path $logRoot -RetentionDays $LogRetentionDays -MaximumEndpointFiles $MaxEndpointLogFiles
        $level = if ($result.Failed -gt 0) { 'WARN' } else { 'INFO' }
        Write-RunnerLog ("Intune log maintenance completed. Phase={0}; RetentionDays={1}; MaximumEndpointFiles={2}; Scanned={3}; RemovedByAge={4}; RemovedByCount={5}; Failed={6}; Remaining={7}" -f $Phase,$LogRetentionDays,$MaxEndpointLogFiles,$result.Scanned,$result.RemovedByAge,$result.RemovedByCount,$result.Failed,$result.Remaining) $level
    }
    catch {
        Write-RunnerLog ("Intune log maintenance failed without blocking the upgrade. Phase={0}; Error={1}" -f $Phase,$_.Exception.Message) 'WARN'
    }
}

function Open-Registry64LocalMachine {
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
}

function Set-Registry64String {
    param([string]$SubKey, [string]$Name, [string]$Value)

    $baseKey = Open-Registry64LocalMachine
    try {
        $key = $baseKey.CreateSubKey($SubKey)
        try { $key.SetValue($Name, [string]$Value, [Microsoft.Win32.RegistryValueKind]::String) }
        finally { if ($key) { $key.Dispose() } }
    }
    finally { $baseKey.Dispose() }
}

function Set-PackageRepairRequired {
    param(
        [Parameter(Mandatory = $true)][object]$Manifest,
        [Parameter(Mandatory = $true)][string]$Reason
    )

    $packageId = [string]$Manifest.PackageId
    if ([string]::IsNullOrWhiteSpace($packageId)) { throw 'Cannot mark RepairRequired because PackageId is empty.' }
    $subKey = "$registrySubKeyRoot\$packageId"
    $safeReason = if ($Reason.Length -gt 2048) { $Reason.Substring(0, 2048) } else { $Reason }
    Set-Registry64String -SubKey $subKey -Name PackageId -Value $packageId
    Set-Registry64String -SubKey $subKey -Name PackageVersion -Value ([string]$Manifest.PackageVersion)
    Set-Registry64String -SubKey $subKey -Name InstallState -Value 'RepairRequired'
    Set-Registry64String -SubKey $subKey -Name RepairReason -Value $safeReason
    Set-Registry64String -SubKey $subKey -Name RepairRequiredUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))
    Write-RunnerLog ("Marked Intune package RepairRequired so detection triggers reinstall. PackageId={0}; Reason={1}" -f $packageId,$safeReason) 'WARN'
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-RunnerLog ("Removed scheduled task while package repair is required. TaskName={0}" -f $TaskName) 'WARN'
    }
    catch {
        Write-RunnerLog ("Could not remove scheduled task after marking RepairRequired. TaskName={0}; Error={1}" -f $TaskName,$_.Exception.Message) 'WARN'
    }
}

function Set-PackageMediaState {
    param([Parameter(Mandatory = $true)][object]$Manifest, [Parameter(Mandatory = $true)][object]$Resolution)
    $packageId = [string]$Manifest.PackageId
    if ([string]::IsNullOrWhiteSpace($packageId)) { return }
    $subKey = "$registrySubKeyRoot\$packageId"
    try {
        Set-Registry64String -SubKey $subKey -Name SetupMediaPath -Value ([string]$Resolution.Path)
        Set-Registry64String -SubKey $subKey -Name SetupMediaSource -Value ([string]$Resolution.Source)
        Set-Registry64String -SubKey $subKey -Name SetupMediaExternal -Value ([string]([bool]$Resolution.External))
    }
    catch { Write-RunnerLog ("Could not record the selected setup media in the package state. Error={0}" -f $_.Exception.Message) 'WARN' }
}
function Set-PackageInstalledForWindows11 {
    param([Parameter(Mandatory = $true)][object]$Manifest)

    $packageId = [string]$Manifest.PackageId
    if ([string]::IsNullOrWhiteSpace($packageId)) { throw 'Cannot mark Installed because PackageId is empty.' }
    $subKey = "$registrySubKeyRoot\$packageId"
    Set-Registry64String -SubKey $subKey -Name PackageId -Value $packageId
    Set-Registry64String -SubKey $subKey -Name PackageVersion -Value ([string]$Manifest.PackageVersion)
    Set-Registry64String -SubKey $subKey -Name InstallState -Value 'Installed'
    Set-Registry64String -SubKey $subKey -Name CompletionReason -Value 'AlreadyWindows11'
    Set-Registry64String -SubKey $subKey -Name RepairReason -Value ''
    Set-Registry64String -SubKey $subKey -Name RepairRequiredUtc -Value ''
    Set-Registry64String -SubKey $subKey -Name InstalledUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))
    Write-RunnerLog ("Device is Windows 11. Marked Intune package Installed. PackageId={0}; PackageVersion={1}" -f $packageId,$Manifest.PackageVersion)
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

trap {
    try { Write-RunnerLog ("Runner failed: {0}" -f $_.Exception.Message) 'ERROR' } catch { }
    exit 1
}

# SEMT data folder boundary: nothing is loaded, written or started unless the program folder and
# the data folder chain (setup media included) are verified as trusted.
$script:DataFolderTrusted = $false
$programTrust = Test-SemtTrustedProgramPath -Path $programRoot -Recurse
if (-not $programTrust.Trusted) {
    Write-RunnerLog ("SEMT_PROGRAM_FOLDER_UNTRUSTED: the program folder is not protected; nothing is loaded or started. Path={0}; State={1}; {2}" -f $programRoot,$programTrust.State,$programTrust.Detail) 'ERROR'
    exit 1
}
$dataTrust = Test-SemtTrustedDataFolder -Path $DataRoot -BasePath $ProgramDataPath
if (-not $dataTrust.Trusted) {
    $reason = ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $DataRoot,$dataTrust.State,$dataTrust.Detail)
    # Registry only: the package is marked for repair so that Intune reinstalls it and the
    # installer hardens the chain again; this task is removed and nothing is started.
    try {
        $repairManifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $repairSubKey = "$registrySubKeyRoot\$([string]$repairManifest.PackageId)"
        Set-Registry64String -SubKey $repairSubKey -Name InstallState -Value 'RepairRequired'
        Set-Registry64String -SubKey $repairSubKey -Name RepairReason -Value $(if ($reason.Length -gt 2048) { $reason.Substring(0, 2048) } else { $reason })
        Set-Registry64String -SubKey $repairSubKey -Name RepairRequiredUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))
    }
    catch { Write-RunnerLog ("Package repair state could not be written: {0}" -f $_.Exception.Message) 'ERROR' }
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop } catch { }
    Write-RunnerLog ("{0} Nothing was loaded, written under the data root or started." -f $reason) 'ERROR'
    exit 1
}
$script:DataFolderTrusted = $true
Invoke-RunnerLogMaintenance -Phase 'Startup'
Write-RunnerLog ("Runner started. DataRoot={0}; TaskName={1}; RunGuardHours={2}; LogRetentionDays={3}; MaxEndpointLogFiles={4}; User={5}; Computer={6}; PID={7}" -f $DataRoot,$TaskName,$RunGuardHours,$LogRetentionDays,$MaxEndpointLogFiles,([Security.Principal.WindowsIdentity]::GetCurrent().Name),$env:COMPUTERNAME,$PID)

$osFamily = Get-OsFamily
if ($osFamily -eq 'Windows11') {
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
        try {
            $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            Set-PackageInstalledForWindows11 -Manifest $manifest
        }
        catch { Write-RunnerLog ("Windows 11 was detected, but package state could not be refreshed. Intune OS detection remains authoritative. Error={0}" -f $_.Exception.Message) 'WARN' }
    }
    else {
        Write-RunnerLog ("Windows 11 was detected before package manifest validation. Intune OS detection remains authoritative. MissingManifest={0}" -f $manifestPath) 'WARN'
    }
    Write-RunnerLog 'Device is already Windows 11. Removing scheduled task.'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    exit 0
}

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Package manifest not found: $manifestPath" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if (-not (Test-Path -LiteralPath $endpointScript -PathType Leaf)) { throw "Endpoint script not found: $endpointScript" }
$endpointParseErrors = $null
[void][System.Management.Automation.PSParser]::Tokenize((Get-Content -LiteralPath $endpointScript -Raw), [ref]$endpointParseErrors)
if ($endpointParseErrors) {
    $parseSummary = (($endpointParseErrors | Select-Object -First 5 | ForEach-Object { "Line={0}; Column={1}; Message={2}" -f $_.Token.StartLine,$_.Token.StartColumn,$_.Message }) -join ' | ')
    Write-RunnerLog ("Endpoint script parse validation failed: {0}" -f $parseSummary) 'ERROR'
    exit 1
}
Write-RunnerLog ("Endpoint script parse validation succeeded: {0}" -f $endpointScript)
$setupProcesses = @(Get-Process -Name setup,setuphost,setupprep -ErrorAction SilentlyContinue)
if ($setupProcesses.Count -gt 0) {
    Write-RunnerLog ("Setup process already running; skipping this cycle. Processes={0}" -f (($setupProcesses | Select-Object -ExpandProperty Id) -join ','))
    exit 0
}

# The media folder is resolved with the same rules as the installer and the endpoint. Only the
# absence of any usable media marks the package for repair; media in use or in an unreadable
# state skips this cycle. A valid legacy cache owned by the previous Intune package is a normal,
# external media source and never triggers a repair. Untrusted legacy media is never used: the
# repair reason is SEMT_MEDIA_UNTRUSTED and the reinstall copies only when duplication is allowed.
if (-not (Test-Path -LiteralPath $integrityHelperPath -PathType Leaf)) { throw "Setup media integrity helper not found: $integrityHelperPath" }
. $integrityHelperPath
$cacheFolder = [string]$manifest.SetupCacheFolder
$mediaResolution = Resolve-SemtIntuneSetupMedia -DataRoot $DataRoot -SetupCacheFolder $cacheFolder -Language ([string]$manifest.Language)
Write-RunnerLog ("Setup media resolution before endpoint launch: Decision={0}; Path={1}; External={2}; Owner={3}; {4}" -f $mediaResolution.Decision,$mediaResolution.Path,$mediaResolution.External,$mediaResolution.Owner,$mediaResolution.Detail)
if ($mediaResolution.Decision -in @('InUse','Indeterminate')) {
    Write-RunnerLog ("Setup media is {0}; skipping this cycle without changing anything. {1}" -f $mediaResolution.Decision,$mediaResolution.LegacyDetail) 'WARN'
    exit 0
}
if (-not (Test-SemtIntuneMediaDecisionUsable -Decision $mediaResolution.Decision)) {
    $repairReason = ('SEMT_MEDIA_{0}: no usable setup media. {1}' -f $mediaResolution.Decision.ToUpperInvariant(),$mediaResolution.Detail)
    Set-PackageRepairRequired -Manifest $manifest -Reason $repairReason
    throw $repairReason
}
$cachePath = $mediaResolution.Path
Set-PackageMediaState -Manifest $manifest -Resolution $mediaResolution
$correlationId = [guid]::NewGuid().ToString('N')
$endpointArgs = @(
    '-RunGuardHours', [string]$RunGuardHours,
    '-DirectSetupUpgrade',
    '-SkipVirtualMachines',
    '-AllowReboot',
    '-AllowSetupCompletionRebootWhenNoUser',
    '-AllowSetupProfileRepair',
    '-AllowDiskCleanup',
    '-SetupExecutionMode', 'LocalCache',
    '-SetupMediaId', [string]$manifest.MediaId,
    '-SetupLanguage', [string]$manifest.Language,
    '-SetupDynamicUpdate', 'Disable',
    '-SetupCacheRoot', $setupCacheRoot,
    '-SetupProcessHeartbeatSeconds', '300',
    '-SetupProcessTimeoutMinutes', '0',
    '-SetupMediaCopyTimeoutMinutes', '180',
    '-ForceRequiredRebootWhenUptimeOverDays', '7',
    '-DataRoot', $DataRoot,
    '-CorrelationId', $correlationId,
    '-ReuseLegacySetupMedia',
    '-SetupMediaPath', $cachePath
)
if ($manifest.PSObject.Properties['AllowSetupMediaDuplication'] -and [bool]$manifest.AllowSetupMediaDuplication) { $endpointArgs += '-AllowSetupMediaDuplication' }

Write-RunnerLog ("Starting endpoint script for package {0}; Language={1}; SetupMediaPath={2}; Source={3}; CorrelationId={4}" -f $manifest.PackageId,$manifest.Language,$cachePath,$mediaResolution.Source,$correlationId)
$powerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$endpointProcessArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $endpointScript) + $endpointArgs
$endpointStdOut = Join-Path $logRoot ("Endpoint_{0}_stdout.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$endpointStdErr = Join-Path $logRoot ("Endpoint_{0}_stderr.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$process = Start-Process -FilePath $powerShellExe -ArgumentList $endpointProcessArgs -Wait -PassThru -NoNewWindow -RedirectStandardOutput $endpointStdOut -RedirectStandardError $endpointStdErr
$exit = [int]$process.ExitCode
Invoke-RunnerLogMaintenance -Phase 'PostRun'
if ($exit -ne 0) {
    try {
        $postRun = Resolve-SemtIntuneSetupMedia -DataRoot $DataRoot -SetupCacheFolder $cacheFolder -Language ([string]$manifest.Language)
        Write-RunnerLog ("Setup media resolution after endpoint exit {0}: Decision={1}; Path={2}" -f $exit,$postRun.Decision,$postRun.Path)
        if (Test-SemtIntuneMediaDecisionUsable -Decision $postRun.Decision) { Set-PackageMediaState -Manifest $manifest -Resolution $postRun }
        elseif ($postRun.Decision -notin @('InUse','Indeterminate')) { Set-PackageRepairRequired -Manifest $manifest -Reason ('SEMT_MEDIA_{0}: no usable setup media after endpoint run. {1}' -f $postRun.Decision.ToUpperInvariant(),$postRun.Detail) }
    }
    catch { Write-RunnerLog ("Post-run setup media check failed; package state unchanged. Error={0}" -f $_.Exception.Message) 'ERROR' }
}
$attemptOutcome = Get-EndpointAttemptOutcome -ResultsRoot $resultsRoot -LastRunPath $lastRunPath -CorrelationId $correlationId
$outcomeLevel = if ($attemptOutcome.Completed) { 'INFO' } else { 'WARN' }
Write-RunnerLog ("Endpoint attempt outcome. CorrelationId={0}; Source={1}; Completed={2}; Status={3}; ExitCode={4}; NextAction={5}; Path={6}" -f $correlationId,$attemptOutcome.Source,$attemptOutcome.Completed,$attemptOutcome.Status,$attemptOutcome.ExitCode,$attemptOutcome.NextAction,$attemptOutcome.Path) $outcomeLevel
$latestEndpointLog = Get-ChildItem -LiteralPath (Join-Path $DataRoot 'Logs') -Filter 'SEMT_Windows11Upgrade_*.log' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestEndpointLog) {
    Write-RunnerLog ("Endpoint script exited with code {0}. LatestEndpointLog={1}; LastWriteTime={2}; StdOut={3}; StdErr={4}" -f $exit,$latestEndpointLog.FullName,$latestEndpointLog.LastWriteTime,$endpointStdOut,$endpointStdErr)
}
else {
    Write-RunnerLog ("Endpoint script exited with code {0}. LatestEndpointLog=<not found>; StdOut={1}; StdErr={2}" -f $exit,$endpointStdOut,$endpointStdErr) 'WARN'
}

if ((Get-OsFamily) -eq 'Windows11') {
    Set-PackageInstalledForWindows11 -Manifest $manifest
    Write-RunnerLog 'Device is Windows 11 after endpoint run. Removing scheduled task.'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}

exit 0
