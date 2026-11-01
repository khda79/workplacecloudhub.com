<#
.SYNOPSIS
    Creates on the target the scheduled tasks used by the HybridJoin and Windows11Upgrade campaigns.

.DESCRIPTION
    Script name: SmartM365-Invoke-EndpointMigration-Prerequisites.ps1
    Campaign: Prerequisites of the Smart Endpoint Migration Toolkit (autonomous single-file endpoint
    script). Runs locally as SYSTEM (PsExec or a remote SYSTEM task) in Windows PowerShell 5.1.
    Writes under -DataRoot, by default C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Prerequisites.

    The HybridJoin and Windows11Upgrade endpoint scripts never create a scheduled task: they only
    start, enable or disable the tasks created here.
    -HybridJoinTasks, in \WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\:
      SEMT-HJ-UserDsregStatus, SEMT-HJ-UserRefreshPrt, SEMT-HJ-UserMdmAutoEnroll (on demand, group
      INTERACTIVE), SEMT-HJ-RunUserAutoEnrollAtLogon (SYSTEM, at logon, created disabled), and
      \SmartM365-SEMT-HybridJoin-RetryAfterReboot (SYSTEM, at startup, created disabled).
    -Windows11UpgradeTasks: \SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot (SYSTEM, at startup,
      created disabled).
    A task that already has the expected definition is left unchanged. A task with another
    definition is updated and keeps its enabled state: a campaign may have enabled it on purpose.
    -AuditOnly only reports the tasks that are present, missing or different.
    Each attempt writes State\Results\<CorrelationId>.json.txt for the technician worker.

.PARAMETER HybridJoinTasks
    Creates or checks the tasks of the HybridJoin campaign.

.PARAMETER Windows11UpgradeTasks
    Creates or checks the task of the Windows11Upgrade campaign.

.PARAMETER AuditOnly
    Reports present, missing or different tasks. Nothing is created or changed.

.PARAMETER SkipVirtualMachines
    Skips detected virtual machines before any task operation.

.PARAMETER RetryAfterRebootDelaySeconds
    Common endpoint parameter of the toolkit. Not used: the retry-after-reboot delay is written by
    the campaign that schedules a retry, in its RetryAfterReboot.json.txt state.

.PARAMETER RetryAfterRebootMaxAttempts
    Common endpoint parameter of the toolkit. Not used (see RetryAfterRebootDelaySeconds).

.PARAMETER IgnoreRunGuard
    Common endpoint parameter of the toolkit. This campaign has no local run guard; the shared
    target guard (one toolkit operation at a time) always applies.

.PARAMETER DataRoot
    Campaign runtime folder. Defaults to C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Prerequisites.

.PARAMETER CorrelationId
    Attempt identifier supplied by the technician worker. A local run generates one.

.VERSION
    1.0.0

.EXITCODES
    0 = Every requested task is present with its expected definition (or was created/updated)
    3 = Skipped / audit with missing or different tasks / nothing selected / target busy
    1 = Error (task creation failed, untrusted data folder)

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$HybridJoinTasks,
    [switch]$Windows11UpgradeTasks,
    [switch]$AuditOnly,
    [switch]$SkipVirtualMachines,
    [ValidateRange(0, 3600)][int]$RetryAfterRebootDelaySeconds = 300,
    [ValidateRange(1, 30)][int]$RetryAfterRebootMaxAttempts = 3,
    [switch]$IgnoreRunGuard,
    [string]$DataRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Prerequisites',
    [string]$CorrelationId
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Endpoint scripts run in Windows PowerShell 5.1 (the worker starts them with powershell.exe). Any
# other engine is refused here, before any state is written. A 32-bit host would resolve the
# Program Files paths of the tasks to Program Files (x86): it is refused too.
if ([string]$PSVersionTable['PSEdition'] -eq 'Core' -or $PSVersionTable['PSVersion'] -lt [version]'5.1') {
    Write-Output ('FINAL_STATUS=UNSUPPORTED_POWERSHELL_HOST; EXIT_CODE=3; NEXT_ACTION=RUN_WITH_WINDOWS_POWERSHELL_5_1; CORRELATION_ID={0}; DETAIL={1} {2} is not supported; run this script with powershell.exe (Windows PowerShell 5.1).' -f [string]$CorrelationId,[string]$PSVersionTable['PSEdition'],[string]$PSVersionTable['PSVersion'])
    exit 3
}
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Write-Output ('FINAL_STATUS=UNSUPPORTED_POWERSHELL_HOST; EXIT_CODE=3; NEXT_ACTION=RUN_WITH_64BIT_WINDOWS_POWERSHELL; CORRELATION_ID={0}; DETAIL=32-bit Windows PowerShell is not supported on a 64-bit system.' -f [string]$CorrelationId)
    exit 3
}

$script:ScriptVersion = '1.0.0'
$script:SemtCampaign = 'Prerequisites'
$script:RunId = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:ComputerName = $env:COMPUTERNAME
$DataRoot = [System.IO.Path]::GetFullPath($DataRoot)
$script:LogDir = Join-Path $DataRoot 'Logs'
$script:OutputDir = Join-Path $DataRoot 'Output'
$script:StateDir = Join-Path $DataRoot 'State'
$script:LogPath = Join-Path $script:LogDir ('SEMT_Prerequisites_{0}_{1}.log' -f $script:ComputerName,$script:RunId)
$script:CsvPath = Join-Path $script:OutputDir ('SEMT_Prerequisites_{0}_{1}.csv' -f $script:ComputerName,$script:RunId)
$script:AttemptStartUtc = (Get-Date).ToUniversalTime()
$script:SemtTargetGuardStatus = ''
$script:LogRetentionDays = 30
if ([string]::IsNullOrWhiteSpace($CorrelationId)) { $CorrelationId = [guid]::NewGuid().ToString('N') }
$script:Counters = [ordered]@{ TasksExpected = 0; TasksPresent = 0; TasksCreated = 0; TasksUpdated = 0; TasksMissing = 0; TasksFailed = 0 }
$script:TaskRows = New-Object System.Collections.Generic.List[object]

function Write-PrerequisitesLog {
    param([Parameter(Mandatory = $true)][string]$Message,[ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO')
    try {
        if (-not (Test-Path -LiteralPath $script:LogDir -PathType Container)) { return }
        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -LiteralPath $script:LogPath -Value @(($Message -split "`r?`n") | ForEach-Object { '{0} [{1}] {2}' -f $stamp,$Level,$_ }) -Encoding UTF8
    }
    catch { }
}

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

#region SEMT shared target guards v1 (identical in every endpoint script; checked by Tests\Core)
# Target-side protections shared by every campaign of the Smart Endpoint Migration Toolkit.
# - One operation at a time on a target, whatever the campaign: named Global mutex held by the
#   running endpoint plus a diagnostic lease file. The mutex is released by Windows if the
#   process dies; the lease file only explains who holds it.
# - A campaign can leave a pending-operation marker that outlives its process (for example a
#   Windows Setup completion reboot). Other campaigns stop until the marker expires, the
#   computer has restarted, or the owning campaign clears it.
# - Previous toolkits (IntuneHybridJoinToolkit, Windows11UpgradeToolkit) are inspected
#   READ-ONLY. Installed files or old history never block; a real active operation blocks;
#   an unreadable state blocks with an explicit status.

$script:SemtTargetMutexName = 'Global\SmartM365_SEMT_TargetOperation'
$script:SemtTargetMutex = $null
$script:SemtTargetMutexAcquired = $false
$script:SemtLegacyWindows11Root = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'

function Get-SemtTargetProductRoot {
    param([Parameter(Mandatory = $true)][string]$CampaignDataRoot)
    return (Split-Path -Path $CampaignDataRoot -Parent)
}

function Get-SemtTargetLastBootUtc {
    try { return ((Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime).ToUniversalTime() }
    catch { return $null }
}

function Write-SemtTargetJson {
    param([Parameter(Mandatory = $true)][string]$Path,[Parameter(Mandatory = $true)]$Data)
    $parent = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $temp = Join-Path -Path $parent -ChildPath ('.{0}.{1}.tmp' -f (Split-Path -Path $Path -Leaf),[guid]::NewGuid().ToString('N'))
    try {
        $Data | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $temp -Encoding UTF8 -Force -ErrorAction Stop
        Move-Item -LiteralPath $temp -Destination $Path -Force -ErrorAction Stop
    }
    finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
}

function ConvertTo-SemtTargetJsonDateText {
    # Same dates on every engine: DateTime values created by PowerShell 7 JSON parsing become
    # round-trip text again (UTC/offset as UTC 'o', no zone kept without zone).
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Unspecified) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss.FFFFFFF', [Globalization.CultureInfo]::InvariantCulture) }
        return $Value.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in @($Value.PSObject.Properties)) { $property.Value = ConvertTo-SemtTargetJsonDateText -Value $property.Value }
        return $Value
    }
    if ($Value -is [System.Collections.IList]) {
        for ($index = 0; $index -lt $Value.Count; $index++) { $Value[$index] = ConvertTo-SemtTargetJsonDateText -Value $Value[$index] }
        return ,$Value
    }
    return $Value
}
function Read-SemtTargetJson {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    # Keep ISO dates as text on PowerShell 7 too, so UTC round-trip comparisons stay exact.
    if ((Get-Command -Name ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { return ($raw | ConvertFrom-Json -DateKind String -ErrorAction Stop) }
    return (ConvertTo-SemtTargetJsonDateText -Value ($raw | ConvertFrom-Json -ErrorAction Stop))
}

function Enter-SemtTargetOperationGuard {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot,
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][string]$CorrelationId
    )

    $productRoot = Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot
    $leasePath = Join-Path -Path $productRoot -ChildPath 'State\TargetOperation.json.txt'
    $pendingPath = Join-Path -Path $productRoot -ChildPath 'State\TargetPendingOperation.json.txt'
    try {
        $script:SemtTargetMutex = New-Object System.Threading.Mutex($false, $script:SemtTargetMutexName)
        try { $script:SemtTargetMutexAcquired = $script:SemtTargetMutex.WaitOne(0) }
        catch [System.Threading.AbandonedMutexException] { $script:SemtTargetMutexAcquired = $true }
    }
    catch { $script:SemtTargetMutexAcquired = $false }
    if (-not $script:SemtTargetMutexAcquired) {
        $holder = ''
        try {
            $lease = Read-SemtTargetJson -Path $leasePath
            if ($lease) { $holder = ('Campaign={0}; RunId={1}; ProcessId={2}; StartedUtc={3}' -f $lease.Campaign,$lease.RunId,$lease.ProcessId,$lease.StartedUtc) }
        }
        catch { $holder = 'Lease unreadable.' }
        if ($script:SemtTargetMutex) { try { $script:SemtTargetMutex.Dispose() } catch { } ; $script:SemtTargetMutex = $null }
        return [pscustomobject]@{ Acquired = $false; Status = 'TARGET_OPERATION_ACTIVE'; NextAction = 'WAIT_ACTIVE_TARGET_OPERATION'; Detail = ("Another Smart Endpoint Migration Toolkit operation is active on this computer. {0}" -f $holder) }
    }

    try {
        $pending = Read-SemtTargetJson -Path $pendingPath
        if ($pending -and [string]$pending.Campaign -ne $Campaign) {
            $expires = [datetime]::MinValue
            $expiresOk = [datetime]::TryParse([string]$pending.ExpiresUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$expires)
            $bootAtWrite = [datetime]::MinValue
            $bootOk = [datetime]::TryParse([string]$pending.LastBootUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$bootAtWrite)
            $currentBoot = Get-SemtTargetLastBootUtc
            $rebootedSince = ($bootOk -and $null -ne $currentBoot -and [math]::Abs(($currentBoot - $bootAtWrite.ToUniversalTime()).TotalMinutes) -gt 2 -and [bool]$pending.ClearedByReboot)
            if ($expiresOk -and (Get-Date).ToUniversalTime() -lt $expires.ToUniversalTime() -and -not $rebootedSince) {
                Exit-SemtTargetOperationGuard -CampaignDataRoot $CampaignDataRoot
                return [pscustomobject]@{ Acquired = $false; Status = 'TARGET_PENDING_OPERATION'; NextAction = 'WAIT_OTHER_CAMPAIGN_PENDING_OPERATION'; Detail = ("Campaign {0} left a pending operation on this computer: {1}. Expires={2}." -f $pending.Campaign,$pending.Reason,$pending.ExpiresUtc) }
            }
        }
    }
    catch {
        Exit-SemtTargetOperationGuard -CampaignDataRoot $CampaignDataRoot
        return [pscustomobject]@{ Acquired = $false; Status = 'TARGET_STATE_UNREADABLE'; NextAction = 'VERIFY_TARGET_STATE_FILES'; Detail = ("Pending-operation marker could not be read: {0}" -f $_.Exception.Message) }
    }

    $processStart = ''
    try { $processStart = (Get-Process -Id $PID -ErrorAction Stop).StartTime.ToUniversalTime().ToString('o') } catch { $processStart = '' }
    try {
        Write-SemtTargetJson -Path $leasePath -Data ([pscustomobject]@{ SchemaVersion = 1; Campaign = $Campaign; RunId = $RunId; CorrelationId = $CorrelationId; ProcessId = $PID; ProcessStartUtc = $processStart; StartedUtc = (Get-Date).ToUniversalTime().ToString('o') })
    }
    catch { }
    return [pscustomobject]@{ Acquired = $true; Status = 'TARGET_OPERATION_ACQUIRED'; NextAction = ''; Detail = 'Target operation guard acquired.' }
}

function Exit-SemtTargetOperationGuard {
    param([Parameter(Mandatory = $true)][string]$CampaignDataRoot)

    if ($script:SemtTargetMutexAcquired) {
        $leasePath = Join-Path -Path (Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot) -ChildPath 'State\TargetOperation.json.txt'
        try {
            $lease = Read-SemtTargetJson -Path $leasePath
            if ($lease -and [int]$lease.ProcessId -eq $PID) { Remove-Item -LiteralPath $leasePath -Force -ErrorAction SilentlyContinue }
        }
        catch { }
        if ($script:SemtTargetMutex) { try { $script:SemtTargetMutex.ReleaseMutex() } catch { } }
    }
    if ($script:SemtTargetMutex) { try { $script:SemtTargetMutex.Dispose() } catch { } }
    $script:SemtTargetMutex = $null
    $script:SemtTargetMutexAcquired = $false
}

function Set-SemtTargetPendingOperation {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot,
        [Parameter(Mandatory = $true)][string]$Reason,
        [ValidateRange(1, 168)][int]$Hours = 48,
        [bool]$ClearedByReboot = $true
    )

    $path = Join-Path -Path (Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot) -ChildPath 'State\TargetPendingOperation.json.txt'
    $boot = Get-SemtTargetLastBootUtc
    Write-SemtTargetJson -Path $path -Data ([pscustomobject]@{ SchemaVersion = 1; Campaign = $Campaign; Reason = $Reason; SinceUtc = (Get-Date).ToUniversalTime().ToString('o'); ExpiresUtc = (Get-Date).ToUniversalTime().AddHours($Hours).ToString('o'); LastBootUtc = $(if ($boot) { $boot.ToString('o') } else { '' }); ClearedByReboot = $ClearedByReboot })
}

function Clear-SemtTargetPendingOperation {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot
    )

    $path = Join-Path -Path (Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot) -ChildPath 'State\TargetPendingOperation.json.txt'
    try {
        $pending = Read-SemtTargetJson -Path $path
        if ($pending -and [string]$pending.Campaign -eq $Campaign) { Remove-Item -LiteralPath $path -Force -ErrorAction Stop; return $true }
    }
    catch { }
    return $false
}

function Test-SemtWindowsSetupActive {
    $signals = New-Object System.Collections.Generic.List[string]
    try {
        $setupKey = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\Setup' -ErrorAction Stop
        foreach ($name in @('SystemSetupInProgress','UpgradeInProgress')) {
            $property = $setupKey.PSObject.Properties[$name]
            if ($property -and [string]$property.Value -ne '0') { $signals.Add(('{0}={1}' -f $name,$property.Value)) }
        }
    }
    catch { }
    foreach ($processName in @('SetupHost','SetupPrep','setup')) {
        foreach ($process in @(Get-Process -Name $processName -ErrorAction SilentlyContinue)) {
            $path = ''
            try { $path = [string]$process.Path } catch { $path = '' }
            if ($processName -ne 'setup' -or $path -match '(?i)\\(SetupMedia|\$WINDOWS\.~BT|sources)\\' -or [string]::IsNullOrWhiteSpace($path)) { $signals.Add(('{0}:{1}' -f $process.ProcessName,$process.Id)) }
        }
    }
    return [pscustomobject]@{ Active = ($signals.Count -gt 0); Detail = ($signals -join '; ') }
}

function Write-SemtAttemptResult {
    <#
    Writes <DataRoot>\State\Results\<CorrelationId>.json.txt. The technician worker only trusts a
    result whose CorrelationId matches its own attempt. Keeps the newest 50 result files.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Data
    )

    $resultsRoot = Join-Path -Path $CampaignDataRoot -ChildPath 'State\Results'
    $path = Join-Path -Path $resultsRoot -ChildPath ('{0}.json.txt' -f [string]$Data['CorrelationId'])
    Write-SemtTargetJson -Path $path -Data ([pscustomobject]$Data)
    try {
        @(Get-ChildItem -LiteralPath $resultsRoot -Filter '*.json.txt' -File -ErrorAction Stop | Sort-Object LastWriteTimeUtc -Descending | Select-Object -Skip 50) | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    }
    catch { }
    return $path
}
#endregion SEMT shared target guards v1

#region SEMT prerequisite tasks v1 (identical in every endpoint script; checked by Tests\Core)
# Scheduled tasks are created only by the Prerequisites Campaign (disabled by default). The
# HybridJoin and Windows11Upgrade endpoints start, enable or disable these tasks; they never create,
# change or delete a task definition, and a missing task is reported, not created.
# The retry-after-reboot task has a fixed command: the runner under Program Files reads the script
# path and arguments from RetryAfterReboot.json.txt next to it. The task is therefore created
# disabled in advance and enabled only while a campaign waits for a restart.
# Principals are compared by SID: Task Scheduler shows localized account names.

$script:SemtHybridJoinTaskPath = '\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\'

function Get-SemtRetryAfterRebootTaskSpec {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('HybridJoin','Windows11Upgrade')][string]$Campaign,
        [string]$ProgramFilesPath = $env:ProgramFiles,
        [string]$WindowsPath = $env:WINDIR
    )
    $runner = Join-Path $ProgramFilesPath ('WCH\SmartM365\SmartEndpointMigrationToolkit\{0}\RetryAfterReboot\RetryAfterRebootRunner.ps1' -f $Campaign)
    return [pscustomobject]@{
        Campaign = $Campaign
        TaskPath = '\'
        TaskName = ('SmartM365-SEMT-{0}-RetryAfterReboot' -f $Campaign)
        Execute = (Join-Path $WindowsPath 'System32\WindowsPowerShell\v1.0\powershell.exe')
        Arguments = ('-NoProfile -ExecutionPolicy RemoteSigned -File "{0}"' -f $runner)
        RunnerPath = $runner
    }
}

function ConvertTo-SemtTaskPrincipalSid {
    param([AllowNull()][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    if ($Name -match '^S-1-[0-9-]+$') { return $Name }
    try { return (New-Object System.Security.Principal.NTAccount($Name)).Translate([System.Security.Principal.SecurityIdentifier]).Value }
    catch { return '' }
}

function Get-SemtScheduledTaskState {
    # Read-only. Exists is $null when the Task Scheduler could not be read.
    param([Parameter(Mandatory = $true)][string]$TaskPath,[Parameter(Mandatory = $true)][string]$TaskName)

    $state = [pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Exists = $false; Enabled = $false; Actions = @(); TriggerCount = 0; UserSid = ''; GroupSid = ''; Detail = 'Task not found.' }
    $readErrors = $null
    $task = @(Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue -ErrorVariable readErrors) | Select-Object -First 1
    if ($null -eq $task) {
        $failure = @($readErrors | Where-Object { [string]$_.CategoryInfo.Category -ne 'ObjectNotFound' }) | Select-Object -First 1
        if ($failure) { $state.Exists = $null; $state.Detail = ('Task Scheduler read failed: {0}' -f $failure.Exception.Message) }
        return $state
    }
    $state.Exists = $true
    $state.Enabled = [bool]$task.Settings.Enabled
    $state.Actions = @($task.Actions | ForEach-Object {
        $execute = if ($_.PSObject.Properties['Execute']) { [Environment]::ExpandEnvironmentVariables(([string]$_.Execute).Trim()) } else { '' }
        $arguments = if ($_.PSObject.Properties['Arguments']) { ([string]$_.Arguments).Trim() } else { '' }
        '{0}|{1}' -f $execute,$arguments
    })
    $state.TriggerCount = @($task.Triggers).Count
    $state.UserSid = ConvertTo-SemtTaskPrincipalSid -Name ([string]$task.Principal.UserId)
    $state.GroupSid = ConvertTo-SemtTaskPrincipalSid -Name ([string]$task.Principal.GroupId)
    $state.Detail = ('Enabled={0}; State={1}' -f $state.Enabled,[string]$task.State)
    return $state
}

function Test-SemtRetryAfterRebootTaskState {
    # True when the task exists, runs as SYSTEM and starts exactly the fixed runner command.
    param([Parameter(Mandatory = $true)]$State,[Parameter(Mandatory = $true)]$Spec)
    if ($State.Exists -ne $true -or $State.UserSid -ne 'S-1-5-18' -or @($State.Actions).Count -ne 1) { return $false }
    return ([string]@($State.Actions)[0] -ieq ('{0}|{1}' -f $Spec.Execute,$Spec.Arguments))
}

function Set-SemtScheduledTaskEnabled {
    # Enables or disables an existing task; the definition is not changed.
    param([Parameter(Mandatory = $true)][string]$TaskPath,[Parameter(Mandatory = $true)][string]$TaskName,[Parameter(Mandatory = $true)][bool]$Enabled)
    if ($Enabled) { [void](Enable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop) }
    else { [void](Disable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop) }
}
#endregion SEMT prerequisite tasks v1

function Write-PrerequisitesAttemptResult {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Running','Completed')][string]$State,
        [string]$Status = '',
        [int]$ExitCode = 0,
        [string]$NextAction = '',
        [string]$Detail = ''
    )

    $data = [ordered]@{
        SchemaVersion = 2
        Product = 'SmartEndpointMigrationToolkit'
        Campaign = $script:SemtCampaign
        ScriptVersion = $script:ScriptVersion
        RunId = $script:RunId
        CorrelationId = $CorrelationId
        ComputerName = $script:ComputerName
        State = $State
        StartTimeUtc = $script:AttemptStartUtc.ToString('o')
        EndTimeUtc = $(if ($State -eq 'Completed') { (Get-Date).ToUniversalTime().ToString('o') } else { '' })
        Status = $Status
        ExitCode = $(if ($State -eq 'Completed') { [string]$ExitCode } else { '' })
        NextAction = $NextAction
        Detail = $Detail
        LogPath = $script:LogPath
        CsvPath = $(if (Test-Path -LiteralPath $script:CsvPath -PathType Leaf) { $script:CsvPath } else { '' })
        TranscriptPath = ''
        TargetGuardStatus = $script:SemtTargetGuardStatus
        AuditOnly = [bool]$AuditOnly
    }
    foreach ($name in @($script:Counters.Keys)) { $data[$name] = [string]$script:Counters[$name] }
    $data['TasksDetail'] = (@($script:TaskRows | ForEach-Object { '{0}={1}' -f $_.TaskName,$_.Result }) -join '; ')
    return (Write-SemtAttemptResult -CampaignDataRoot $DataRoot -Data $data)
}

function Complete-PrerequisitesRun {
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][int]$ExitCode,
        [Parameter(Mandatory = $true)][string]$NextAction,
        [Parameter(Mandatory = $true)][string]$Detail
    )

    Write-PrerequisitesLog ('Final Status={0}; ExitCode={1}; NextAction={2}; Detail={3}' -f $Status,$ExitCode,$NextAction,$Detail)
    try { [void](Write-PrerequisitesAttemptResult -State Completed -Status $Status -ExitCode $ExitCode -NextAction $NextAction -Detail $Detail) }
    catch { Write-PrerequisitesLog ('Attempt result write failed: {0}' -f $_.Exception.Message) 'WARN' }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $DataRoot
    Write-Output ('FINAL_STATUS={0}; EXIT_CODE={1}; NEXT_ACTION={2}; CORRELATION_ID={3}; DETAIL={4}' -f $Status,$ExitCode,$NextAction,$CorrelationId,($Detail -replace '[\r\n]+',' '))
    exit $ExitCode
}

function Test-PrerequisitesVirtualMachine {
    try {
        $system = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $signature = '{0} {1}' -f [string]$system.Manufacturer,[string]$system.Model
        foreach ($pattern in @('Virtual Machine','VMware','VirtualBox','KVM','QEMU','Xen','HVM domU','Parallels','BHYVE','OpenStack','Google Compute Engine','Amazon EC2')) {
            if ($signature -match [regex]::Escape($pattern)) { return [pscustomobject]@{ IsVirtual = $true; Detail = ('{0}; Pattern={1}' -f $signature,$pattern) } }
        }
        return [pscustomobject]@{ IsVirtual = $false; Detail = $signature }
    }
    catch { return [pscustomobject]@{ IsVirtual = $false; Detail = ('Computer system query failed: {0}' -f $_.Exception.Message) } }
}

function Get-PrerequisiteTaskDefinitions {
    <#
    Expected definitions. Except for SEMT-HJ-UserRefreshPrt and SEMT-HJ-UserMdmAutoEnroll (now
    started without cmd.exe), they are the definitions that the HybridJoin and Windows11Upgrade
    endpoints created before this campaign existed, so tasks already present are kept as they are.
    Older copies of those two tasks are updated.
    #>
    $definitions = New-Object System.Collections.Generic.List[object]
    $schtasks = Join-Path $env:WINDIR 'System32\schtasks.exe'
    if ($HybridJoinTasks) {
        # Only the dsregcmd /status output is read back (through cmd.exe redirection). refreshprt and
        # AutoEnrollMDM start the executable directly: their result is checked by the HybridJoin
        # endpoint through the status, the MDM registry and the DeviceManagement logs.
        $userTasks = @(
            @{ Name = 'SEMT-HJ-UserDsregStatus'; Description = 'capture dsregcmd /status in the active interactive user context.'; Execute = 'cmd.exe'; Arguments = '/d /c %windir%\System32\dsregcmd.exe /status > %windir%\Temp\SEMT-HJ-UserDsregStatus_%USERNAME%.txt 2>&1' },
            @{ Name = 'SEMT-HJ-UserRefreshPrt'; Description = 'run dsregcmd /refreshprt in the active interactive user context.'; Execute = (Join-Path $env:WINDIR 'System32\dsregcmd.exe'); Arguments = '/refreshprt' },
            @{ Name = 'SEMT-HJ-UserMdmAutoEnroll'; Description = 'run deviceenroller /AutoEnrollMDM in the active interactive user context.'; Execute = (Join-Path $env:WINDIR 'System32\deviceenroller.exe'); Arguments = '/c /AutoEnrollMDM' }
        )
        foreach ($userTask in $userTasks) {
            $definitions.Add([pscustomobject]@{
                Campaign = 'HybridJoin'; TaskPath = $script:SemtHybridJoinTaskPath; TaskName = $userTask.Name
                Description = 'Smart Endpoint Migration Toolkit - Hybrid Join - ' + $userTask.Description
                Principal = 'Interactive'; Trigger = 'None'; EnabledWhenCreated = $true; ExecutionTimeLimit = 'PT10M'
                Actions = @([pscustomobject]@{ Execute = $userTask.Execute; Arguments = $userTask.Arguments })
            })
        }
        $definitions.Add([pscustomobject]@{
            Campaign = 'HybridJoin'; TaskPath = $script:SemtHybridJoinTaskPath; TaskName = 'SEMT-HJ-RunUserAutoEnrollAtLogon'
            Description = 'Smart Endpoint Migration Toolkit - Hybrid Join - trigger user-context auto-enrollment helper tasks at next user logon. Enabled by the HybridJoin endpoint only when needed.'
            Principal = 'System'; Trigger = 'Logon'; EnabledWhenCreated = $false; ExecutionTimeLimit = 'PT10M'
            Actions = @(
                [pscustomobject]@{ Execute = $schtasks; Arguments = ('/Run /TN "{0}SEMT-HJ-UserDsregStatus"' -f $script:SemtHybridJoinTaskPath) },
                [pscustomobject]@{ Execute = $schtasks; Arguments = ('/Run /TN "{0}SEMT-HJ-UserMdmAutoEnroll"' -f $script:SemtHybridJoinTaskPath) }
            )
        })
    }
    $retryCampaigns = @()
    if ($HybridJoinTasks) { $retryCampaigns += 'HybridJoin' }
    if ($Windows11UpgradeTasks) { $retryCampaigns += 'Windows11Upgrade' }
    foreach ($campaign in $retryCampaigns) {
        $spec = Get-SemtRetryAfterRebootTaskSpec -Campaign $campaign
        $definitions.Add([pscustomobject]@{
            Campaign = $campaign; TaskPath = $spec.TaskPath; TaskName = $spec.TaskName
            Description = ('Smart Endpoint Migration Toolkit - {0} - resume the repair after a restart. Enabled by the {0} endpoint only while a retry is scheduled.' -f $campaign)
            Principal = 'System'; Trigger = 'Boot'; EnabledWhenCreated = $false; ExecutionTimeLimit = 'PT72H'
            Actions = @([pscustomobject]@{ Execute = $spec.Execute; Arguments = $spec.Arguments })
            Spec = $spec
        })
    }
    return @($definitions.ToArray())
}

function Test-PrerequisiteTaskMatch {
    param([Parameter(Mandatory = $true)]$Definition,[Parameter(Mandatory = $true)]$State)

    if ($State.Exists -ne $true) { return $false }
    if ($Definition.PSObject.Properties['Spec']) { return ((Test-SemtRetryAfterRebootTaskState -State $State -Spec $Definition.Spec) -and $State.TriggerCount -ge 1) }
    $expected = @($Definition.Actions | ForEach-Object { '{0}|{1}' -f [Environment]::ExpandEnvironmentVariables([string]$_.Execute),[string]$_.Arguments })
    $actual = @($State.Actions)
    if ($expected.Count -ne $actual.Count) { return $false }
    for ($index = 0; $index -lt $expected.Count; $index++) { if ([string]$actual[$index] -ine [string]$expected[$index]) { return $false } }
    if ($Definition.Principal -eq 'Interactive') { return ($State.GroupSid -eq 'S-1-5-4') }
    if ($State.UserSid -ne 'S-1-5-18') { return $false }
    return ($Definition.Trigger -eq 'None' -or $State.TriggerCount -ge 1)
}

function Get-PrerequisiteTaskFolder {
    param([Parameter(Mandatory = $true)]$Service,[Parameter(Mandatory = $true)][string]$TaskPath)

    $path = '\' + $TaskPath.Trim('\')
    if ($path -eq '\') { return $Service.GetFolder('\') }
    try { return $Service.GetFolder($path) } catch { }
    $current = $Service.GetFolder('\')
    $currentPath = ''
    foreach ($part in @($path.Trim('\').Split('\') | Where-Object { $_ })) {
        $currentPath = '{0}\{1}' -f $currentPath,$part
        try { $current = $Service.GetFolder($currentPath) }
        catch {
            try { [void]$current.CreateFolder($part, $null) }
            catch { if ($_.Exception.HResult -ne -2147024713) { throw } }
            $current = $Service.GetFolder($currentPath)
        }
    }
    return $current
}

function Register-PrerequisiteTask {
    param([Parameter(Mandatory = $true)]$Definition,[Parameter(Mandatory = $true)][bool]$Enabled)

    $service = New-Object -ComObject Schedule.Service
    $service.Connect()
    $folder = Get-PrerequisiteTaskFolder -Service $service -TaskPath $Definition.TaskPath
    $task = $service.NewTask(0)
    $task.RegistrationInfo.Author = 'Smart Endpoint Migration Toolkit - Prerequisites'
    $task.RegistrationInfo.Description = $Definition.Description
    $task.Settings.Enabled = $Enabled
    $task.Settings.Hidden = $false
    $task.Settings.AllowDemandStart = $true
    $task.Settings.StartWhenAvailable = $true
    $task.Settings.DisallowStartIfOnBatteries = $false
    $task.Settings.StopIfGoingOnBatteries = $false
    $task.Settings.ExecutionTimeLimit = $Definition.ExecutionTimeLimit
    $task.Settings.MultipleInstances = 2
    switch ($Definition.Trigger) {
        'Logon' { $trigger = $task.Triggers.Create(9); $trigger.Enabled = $true }
        'Boot' { $trigger = $task.Triggers.Create(8); $trigger.Enabled = $true }
    }
    foreach ($actionDefinition in @($Definition.Actions)) {
        $action = $task.Actions.Create(0)
        $action.Path = $actionDefinition.Execute
        $action.Arguments = $actionDefinition.Arguments
    }
    if ($Definition.Principal -eq 'Interactive') {
        $task.Principal.GroupId = 'S-1-5-4'
        $task.Principal.LogonType = 4
        $task.Principal.RunLevel = 0
        [void]$folder.RegisterTaskDefinition($Definition.TaskName, $task, 6, $null, $null, 4, $null)
    }
    else {
        $task.Principal.UserId = 'SYSTEM'
        $task.Principal.LogonType = 5
        $task.Principal.RunLevel = 1
        [void]$folder.RegisterTaskDefinition($Definition.TaskName, $task, 6, 'SYSTEM', $null, 5, $null)
    }
}

function Invoke-PrerequisiteTask {
    param([Parameter(Mandatory = $true)]$Definition)

    $fullName = '{0}{1}' -f $Definition.TaskPath,$Definition.TaskName
    $script:Counters.TasksExpected++
    $before = Get-SemtScheduledTaskState -TaskPath $Definition.TaskPath -TaskName $Definition.TaskName
    $beforeText = if ($null -eq $before.Exists) { 'Unreadable' } elseif (-not $before.Exists) { 'Missing' } elseif (Test-PrerequisiteTaskMatch -Definition $Definition -State $before) { 'Present' } else { 'Different' }
    $result = $beforeText
    $detail = $before.Detail
    $enabledAfter = $before.Enabled
    if ($beforeText -eq 'Present') { $script:Counters.TasksPresent++ }
    elseif ($AuditOnly -or $beforeText -eq 'Unreadable') {
        if ($beforeText -eq 'Unreadable') { $script:Counters.TasksFailed++ } else { $script:Counters.TasksMissing++ }
    }
    else {
        # An existing task keeps its enabled state: a campaign may be waiting on it.
        $enabled = if ($beforeText -eq 'Missing') { [bool]$Definition.EnabledWhenCreated } else { [bool]$before.Enabled }
        try {
            Register-PrerequisiteTask -Definition $Definition -Enabled $enabled
            $after = Get-SemtScheduledTaskState -TaskPath $Definition.TaskPath -TaskName $Definition.TaskName
            if (Test-PrerequisiteTaskMatch -Definition $Definition -State $after) {
                $result = if ($beforeText -eq 'Missing') { 'Created' } else { 'Updated' }
                if ($result -eq 'Created') { $script:Counters.TasksCreated++ } else { $script:Counters.TasksUpdated++ }
                $detail = $after.Detail
                $enabledAfter = $after.Enabled
            }
            else {
                $result = 'Failed'
                $script:Counters.TasksFailed++
                $detail = ('Registered, but the definition read back differs. {0}' -f $after.Detail)
            }
        }
        catch {
            $result = 'Failed'
            $script:Counters.TasksFailed++
            $detail = $_.Exception.Message -replace '[\r\n]+',' '
        }
    }
    Write-PrerequisitesLog ('Task {0}: Campaign={1}; Before={2}; Result={3}; Enabled={4}; Detail={5}' -f $fullName,$Definition.Campaign,$beforeText,$result,$enabledAfter,$detail) $(if ($result -in @('Failed','Unreadable')) { 'ERROR' } elseif ($result -in @('Missing','Different')) { 'WARN' } else { 'INFO' })
    $script:TaskRows.Add([pscustomobject]@{ ComputerName = $script:ComputerName; Campaign = $Definition.Campaign; TaskName = $fullName; Before = $beforeText; Result = $result; Enabled = [string]$enabledAfter; EnabledWhenCreated = [string]$Definition.EnabledWhenCreated; AuditOnly = [string][bool]$AuditOnly; Detail = $detail })
}

function Remove-PrerequisitesOldEvidence {
    $cutoff = (Get-Date).AddDays(-$script:LogRetentionDays)
    foreach ($folder in @($script:LogDir,$script:OutputDir)) {
        @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff }) | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    }
}

# SEMT data folder boundary: the SEMT components of the data folder chain get owner Administrators
# and a protected DACL after a read-only check; failure stops the endpoint without writing there.
try { [void](Initialize-SemtTrustedDataFolder -Path $DataRoot) }
catch {
    Write-Output ('FINAL_STATUS=DATA_FOLDER_UNTRUSTED; EXIT_CODE=1; NEXT_ACTION=VERIFY_DATA_FOLDER_PERMISSIONS; CORRELATION_ID={0}; DETAIL={1}' -f [string]$CorrelationId,($_.Exception.Message -replace '[\r\n]+',' '))
    exit 1
}
foreach ($folder in @($script:LogDir,$script:OutputDir,$script:StateDir)) {
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
}
try { [void](Write-PrerequisitesAttemptResult -State Running -Status 'RUNNING') } catch { Write-PrerequisitesLog ('Attempt start marker write failed: {0}' -f $_.Exception.Message) 'WARN' }
Write-PrerequisitesLog ('Script start. Version={0}; RunId={1}; CorrelationId={2}; DataRoot={3}; HybridJoinTasks={4}; Windows11UpgradeTasks={5}; AuditOnly={6}; SkipVirtualMachines={7}; IgnoreRunGuard={8} (no local run guard in this campaign).' -f $script:ScriptVersion,$script:RunId,$CorrelationId,$DataRoot,[bool]$HybridJoinTasks,[bool]$Windows11UpgradeTasks,[bool]$AuditOnly,[bool]$SkipVirtualMachines,[bool]$IgnoreRunGuard)

$guard = Enter-SemtTargetOperationGuard -Campaign $script:SemtCampaign -CampaignDataRoot $DataRoot -RunId $script:RunId -CorrelationId $CorrelationId
$script:SemtTargetGuardStatus = $guard.Status
if (-not $guard.Acquired) { Complete-PrerequisitesRun -Status $guard.Status -ExitCode 3 -NextAction $guard.NextAction -Detail $guard.Detail }
$setupActivity = Test-SemtWindowsSetupActive
if ($setupActivity.Active) { Complete-PrerequisitesRun -Status 'TARGET_WINDOWS_SETUP_ACTIVE' -ExitCode 3 -NextAction 'WAIT_WINDOWS_SETUP_COMPLETION' -Detail ('Windows Setup is active on this computer; the Prerequisites campaign is postponed. Signals={0}' -f $setupActivity.Detail) }
Remove-PrerequisitesOldEvidence

if ($SkipVirtualMachines) {
    $virtual = Test-PrerequisitesVirtualMachine
    if ($virtual.IsVirtual) { Complete-PrerequisitesRun -Status 'SKIPPED_VIRTUAL_MACHINE' -ExitCode 3 -NextAction 'NO_ACTION_VIRTUAL_MACHINE' -Detail ('Virtual machine skipped by -SkipVirtualMachines. {0}' -f $virtual.Detail) }
}
if (-not $HybridJoinTasks -and -not $Windows11UpgradeTasks) {
    Complete-PrerequisitesRun -Status 'PREREQUISITES_NOTHING_SELECTED' -ExitCode 3 -NextAction 'SELECT_PREREQUISITES_TASKS' -Detail 'Neither -HybridJoinTasks nor -Windows11UpgradeTasks was given; no task was checked.'
}

try {
    foreach ($definition in @(Get-PrerequisiteTaskDefinitions)) { Invoke-PrerequisiteTask -Definition $definition }
    try { $script:TaskRows | Export-Csv -LiteralPath $script:CsvPath -NoTypeInformation -Encoding UTF8 }
    catch { Write-PrerequisitesLog ('Result CSV write failed: {0}' -f $_.Exception.Message) 'WARN' }
}
catch {
    Complete-PrerequisitesRun -Status 'ERROR' -ExitCode 1 -NextAction 'CHECK_TARGET_ENDPOINT_LOG' -Detail ('Unexpected error: {0}' -f ($_.Exception.Message -replace '[\r\n]+',' '))
}

$summary = ('Expected={0}; Present={1}; Created={2}; Updated={3}; Missing={4}; Failed={5}. {6}' -f $script:Counters.TasksExpected,$script:Counters.TasksPresent,$script:Counters.TasksCreated,$script:Counters.TasksUpdated,$script:Counters.TasksMissing,$script:Counters.TasksFailed,(@($script:TaskRows | ForEach-Object { '{0}={1}' -f $_.TaskName,$_.Result }) -join '; '))
if ($script:Counters.TasksFailed -gt 0) { Complete-PrerequisitesRun -Status 'PREREQUISITES_TASK_FAILED' -ExitCode 1 -NextAction 'CHECK_EDR_OR_TASK_SCHEDULER' -Detail $summary }
if ($AuditOnly) {
    if ($script:Counters.TasksMissing -gt 0) { Complete-PrerequisitesRun -Status 'AUDIT_PREREQUISITES_MISSING' -ExitCode 3 -NextAction 'RUN_PREREQUISITES_WITHOUT_AUDIT' -Detail ('AuditOnly: nothing was created or changed. {0}' -f $summary) }
    Complete-PrerequisitesRun -Status 'AUDIT_PREREQUISITES_PRESENT' -ExitCode 0 -NextAction 'NO_ACTION_PREREQUISITES_READY' -Detail ('AuditOnly: nothing was created or changed. {0}' -f $summary)
}
Complete-PrerequisitesRun -Status 'PREREQUISITES_READY' -ExitCode 0 -NextAction 'NO_ACTION_PREREQUISITES_READY' -Detail $summary

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7MdcHlanTLurM
# Zd1TirA9lbu22fxCiH+m5q/wAPC2KqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEfJtse7xr1kgMT4/JgjdY3Fwk2vUTIJ+t5LYp7U9FjnMA0GCSqG
# SIb3DQEBAQUABIIBgBzyh+t6MZ5lEg7z79RM980zrpYDqglJZhqHoDxjeS51gMfX
# c535rrriDdseLpHj0oGZEOiCPYuspZS0T+u+CyqRwb67ProBL10j+WUAh3xKELTA
# SAi4I4yo3wKmKOMT8/THYCSAVB1ddqtmTNh1owvd/K/62hArbSyVfm5cDAI0saVj
# OgzDyuynjaJlmdoiiQRuQCqeHCUIKX8/xrtTxfAm7QsVFusNkDgxYhypg1MtpoVe
# v9JGmF2vvRbd+VFAgjlo7EzD2bSRHgYOALyAhzMPxGh5nhqfCUzRx4n+0uJdXo+G
# +V8j0qNPkDkb/f6NdKnyDCqh5uCoCMaZXIlFKPxu6S4GLAJ073qH/7wGrayD4oid
# 5LnUGHs4xbdql8o3hbvwNdRqFdI5cEl5ko3/wDdPnj0uFYznq8c5+gdjHllhuq4Q
# VKq9jr+afHYpctCTuBta16J5vp/zNnKDi6/O/YvpUrUDwYgC9iTVMtvh3WQ7b6Qg
# Mcs0pbHrG3QfhJsWf6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDExMDM4
# MTVaMC8GCSqGSIb3DQEJBDEiBCCUll1ip9fmkEXw4aifLfVnwOLmotjG5nxtxP/u
# vuFXuDANBgkqhkiG9w0BAQEFAASCAgCirtQjmFT7wtVMhWbzF9gYh8UQmwfR9XNg
# DrMKeaXwF8RfXpiInVDx1da94iS4L3j/HDyyZYfpW7ABeDm8UzV5qZVsLa3XpFN7
# ZhIBj7qGXbeq/h4sqQR6FQuhZ1QyiykW977bFrmaP/u14XRYrkbQMZb6/Co5/lXz
# BXCw+mqynELkyN+g/LirV01ScgdckwV6q0aGMOsuAdy+7EUEC8gighUhSlIqcbsr
# F6gcHwh4M6ji8dk4izEXSB9W5NVm0WXSM92Wzk6ykHccCp+qE0QpX17tkn2Xl76/
# gcPw3/tScmQQwIqvmQne2rwqyIRqjbEDLNZugV3H7Bu4KP0SJAUDxdHgb3W5ZcYf
# LIC7V0G5PvQFN1Qs7qwyXG6DFd3xELu8jJzH3rQUTVOiu2rMlF6G5bWSQ+o+JDjD
# lFOtUlQEue8igMP3nKflTlPvZUAOr9rrwvVbQkI41HjXas1IfTmz8ltZrnJ7kI2z
# bn5v6mb4WBZWeOG110DTIoKOX+m2EZMQpN+bpmVFVAOKI3WxgMz8Eh0nUqT7VC+f
# AptKGbpyI3oAyJGaYLxl2/KhoUuytDF2r0lL5CcIEMFh/zxUN9Y1xvgANnpo6FiF
# NvME7L+1fM/T9mSCUG8IJ3+tCSxWiZxmESXFJ+L0oREwKxS+YcM0RIeVFJb8DPQj
# YW5+/ORV5w==
# SIG # End signature block
