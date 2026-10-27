#Requires -Version 5.1
<#
.SYNOPSIS
Temporary switch tool: moves the private SmartInventory files to SmartInventoryConfigRootPath, then activates it.

.DESCRIPTION
Run once, on one collection server, while the orchestrator task is disabled and no SmartInventory run is in
progress on any server. Without -Apply the script only previews: it lists every file it would move, archive or
just report, and the checks that -Apply enforces.

-Apply:
1. writes a ZIP backup of every listed file to the archive folder and verifies each entry (SHA256);
2. copies the private configuration files and their transport lock/journal files to the configuration root,
   same relative paths, and verifies SHA256 (an existing identical target is kept, a different one stops);
3. renames each original to <name>.moved-<timestamp> (reversible with -Rollback);
4. removes obsolete configurations (no owner script and no template next to them) and manual backups,
   which then stay only in the ZIP;
5. sets SmartInventoryConfigRootPath in the global configuration and checks the mapping of every moved file.
Interrupted .pending/.previous replacement files are only reported: they stay in place and in the ZIP.

-Rollback renames the *.moved-* files back (when the original name is free) and clears
SmartInventoryConfigRootPath. Files under the configuration root are kept.

Remove this script once every deployment has switched.

.PARAMETER Tenant
Tenant profile loaded for the startup context. Default test.

.PARAMETER ConfigRootPath
SmartInventoryConfigRootPath value to activate. Default {{WorkspaceRootPath}}\SmartInventory\Config.
Only the {{WorkspaceRootPath}} and {{SmartM365RootPath}} tokens are supported.

.PARAMETER ArchiveDirectory
Folder of the ZIP backup and of the CSV journal. Default {{WorkspaceRootPath}}\SmartInventory\Archive.

.PARAMETER Apply
Performs the switch. Without it, nothing is written.

.PARAMETER Rollback
Restores the *.moved-* originals and clears SmartInventoryConfigRootPath.

.PARAMETER OtherServersStopped
Required with -Apply or -Rollback: confirms that the orchestrator task is disabled and that no SmartInventory
run is in progress on the other servers. This computer is checked by the script.

.EXAMPLE
.\SmartM365-Move-InventoryConfigToRoot.ps1 -Tenant prod

.EXAMPLE
.\SmartM365-Move-InventoryConfigToRoot.ps1 -Tenant prod -Apply -OtherServersStopped

.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$ConfigRootPath = '{{WorkspaceRootPath}}\SmartInventory\Config',
    [string]$ArchiveDirectory = '{{WorkspaceRootPath}}\SmartInventory\Archive',
    [switch]$Apply,
    [switch]$Rollback,
    [switch]$OtherServersStopped
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$ScriptVersion = '1.0.0'
$global:SmartM365ScriptName = 'SmartM365-Move-InventoryConfigToRoot'
$global:SmartM365ExecutionStartTime = Get-Date

function Write-Host {
    param([Parameter(Position = 0)][object]$Object = '', [ConsoleColor]$ForegroundColor = [ConsoleColor]::Gray)
    foreach ($line in ([string]$Object -split "\r?\n")) {
        Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line) -ForegroundColor $ForegroundColor
    }
}

$smartRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
$codeRoot = Join-Path $smartRoot 'SmartInventory'
. (Join-Path $smartRoot 'Config\SmartM365-TenantContext.ps1')
$null = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $smartRoot

$jsonTransport = Get-Module SmartM365.JsonTransport
if ($null -eq $jsonTransport -or $jsonTransport.Version -lt [version]'1.1.0') {
    throw 'SmartM365.JsonTransport 1.1.0 or later is required: update the deployment before the switch.'
}

function Resolve-RootSetting([string]$Value, [string]$Name) {
    # Same token resolution as the module, so the checked folder is the one the scripts will use.
    $resolved = & $jsonTransport { param($v) Resolve-SmartM365JsonQualifiedUncToken -Value $v } $Value
    if ([string]::IsNullOrWhiteSpace($resolved) -or -not [IO.Path]::IsPathRooted($resolved)) { throw "$Name must resolve to an absolute folder: $Value" }
    return [IO.Path]::GetFullPath($resolved).TrimEnd('\')
}

function Get-FileSha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }

function Get-StreamSha256([IO.Stream]$Stream) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($Stream))).Replace('-', '') } finally { $sha.Dispose() }
}

function Get-InventoryFileAction([IO.FileInfo]$File) {
    $name = $File.Name
    if ($name -match '\.moved-\d{8}-\d{6}$') { return '' }
    if ($name -match '\.json(\.txt)?\..+\.(pending|previous)$') { return 'Report' }
    if ($name -match '(?i)\.(bak|backup)([-_.]|$)|\.before-') { return 'Archive' }
    if ($name -match '^Orchestrator-Jobs\.json(\.txt)?(\.transport\.lock|\.migration\.log)?$') { return 'Move' }
    if ($name -match '\.consumers\.json(\.txt)?(\.transport\.lock|\.migration\.log)?$') { return 'Move' }
    $match = [regex]::Match($name, '^(?<base>.+)\.local\.json(\.txt)?(\.transport\.lock|\.migration\.log)?$')
    if (-not $match.Success) { return '' }
    $base = $match.Groups['base'].Value
    # Owner script: same base name, or the base of a script-scoped registry such as <script>.<Registry>.local.json.
    $owner = @(Get-ChildItem -LiteralPath $File.DirectoryName -Filter '*.ps1' -File | Where-Object {
            $base -eq $_.BaseName -or $base.StartsWith($_.BaseName + '.', [StringComparison]::OrdinalIgnoreCase) })
    $templates = @(Get-ChildItem -LiteralPath $File.DirectoryName -File | Where-Object {
            $_.Name -like ($base + '.local.json*.template') })
    if ($owner.Count -or $templates.Count) { return 'Move' }
    return 'Archive'
}

function Assert-NoLocalRun {
    # Without elevation the command lines of the service account processes are not visible.
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host 'Not elevated: the running processes of other accounts cannot be checked.' -ForegroundColor Yellow
        return $false
    }
    $running = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop | Where-Object {
            $_.ProcessId -ne $PID -and $_.CommandLine -and
            ($_.CommandLine.IndexOf($smartRoot, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or $_.CommandLine -match 'SmartM365LauncherCache|SmartM365\\LauncherCache') })
    foreach ($process in $running) { Write-Host ('Running: PID {0} {1}' -f $process.ProcessId, $process.CommandLine) -ForegroundColor Yellow }
    $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object {
            $_.State -ne 'Disabled' -and (@($_.Actions | ForEach-Object { '{0} {1}' -f $_.Execute, $_.Arguments }) -join ' ').IndexOf($smartRoot, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
    foreach ($task in $tasks) { Write-Host ('Scheduled task not disabled: {0}{1} ({2})' -f $task.TaskPath, $task.TaskName, $task.State) -ForegroundColor Yellow }
    return ($running.Count -eq 0 -and $tasks.Count -eq 0)
}

function Set-GlobalInventoryConfigRoot([string]$Value) {
    $globalPath = Join-Path $smartRoot 'Config\SmartM365.global.local.json'
    $read = Read-SmartM365JsonDocument -Path $globalPath
    $document = $read.Document
    if ($document.PSObject.Properties['SmartInventoryConfigRootPath']) { $document.SmartInventoryConfigRootPath = $Value }
    else { $document | Add-Member -NotePropertyName 'SmartInventoryConfigRootPath' -NotePropertyValue $Value }
    $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes(($document | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $read.SHA256 -Validate {
        param($candidate)
        if ($candidate -isnot [pscustomobject]) { throw 'Global configuration must remain an object.' }
    } | Out-Null
    Write-Host ("SmartInventoryConfigRootPath = '{0}' in {1}" -f $Value, $read.Path) -ForegroundColor Green
}

$status = 'Success'
try {
    $currentSetting = ''
    $globalRead = Read-SmartM365JsonDocument -Path (Join-Path $smartRoot 'Config\SmartM365.global.local.json')
    if ($globalRead.Document.PSObject.Properties['SmartInventoryConfigRootPath']) { $currentSetting = [string]$globalRead.Document.SmartInventoryConfigRootPath }
    Write-Host ("Script {0} | SmartInventory code folder: {1}" -f $ScriptVersion, $codeRoot)
    Write-Host ("Current SmartInventoryConfigRootPath: '{0}'" -f $currentSetting)

    if ($Rollback) {
        if (-not $OtherServersStopped) { throw '-Rollback requires -OtherServersStopped once the other servers are stopped.' }
        if (-not (Assert-NoLocalRun)) { throw 'A SmartInventory run or an enabled SmartInventory scheduled task exists on this computer.' }
        $configRoot = ''
        if (-not [string]::IsNullOrWhiteSpace($currentSetting)) { $configRoot = Resolve-RootSetting $currentSetting 'SmartInventoryConfigRootPath' }
        $restored = 0
        foreach ($moved in @(Get-ChildItem -LiteralPath $codeRoot -Recurse -File -Force | Where-Object { $_.Name -match '\.moved-\d{8}-\d{6}$' })) {
            $original = $moved.FullName -replace '\.moved-\d{8}-\d{6}$', ''
            if (Test-Path -LiteralPath $original) { Write-Host ('Kept (original name in use): {0}' -f $moved.FullName) -ForegroundColor Yellow; continue }
            if ($configRoot) {
                $target = Join-Path $configRoot $original.Substring($codeRoot.Length + 1)
                if ((Test-Path -LiteralPath $target -PathType Leaf) -and (Get-FileSha256 $target) -ne (Get-FileSha256 $moved.FullName)) {
                    Write-Host ('Changed since the switch, the configuration root copy is newer: {0}' -f $target) -ForegroundColor Yellow
                }
            }
            [IO.File]::Move($moved.FullName, $original)
            $restored++
        }
        Set-GlobalInventoryConfigRoot ''
        Write-Host ('Rollback completed: {0} file(s) restored next to the scripts.' -f $restored) -ForegroundColor Green
        return
    }

    $configRoot = Resolve-RootSetting $ConfigRootPath 'ConfigRootPath'
    $archiveRoot = Resolve-RootSetting $ArchiveDirectory 'ArchiveDirectory'
    foreach ($candidate in @($configRoot, $archiveRoot)) {
        if (($candidate + '\').StartsWith($codeRoot + '\', [StringComparison]::OrdinalIgnoreCase) -or ($codeRoot + '\').StartsWith($candidate + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Folder must be outside the SmartInventory code folder: $candidate"
        }
    }
    if (($archiveRoot + '\').StartsWith($configRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'The archive folder must be outside the configuration root.' }
    if (-not [string]::IsNullOrWhiteSpace($currentSetting) -and $currentSetting -ne $ConfigRootPath) {
        throw ("SmartInventoryConfigRootPath is already set to '{0}'; run -Rollback first or pass the same value." -f $currentSetting)
    }
    Write-Host ('Configuration root: {0}' -f $configRoot)
    Write-Host ('Archive folder: {0}' -f $archiveRoot)

    $plan = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-ChildItem -LiteralPath $codeRoot -Recurse -File -Force)) {
        $action = Get-InventoryFileAction $file
        if (-not $action) { continue }
        $relative = $file.FullName.Substring($codeRoot.Length + 1)
        $target = Join-Path $configRoot $relative
        $targetState = ''
        $hash = Get-FileSha256 $file.FullName
        if ($action -eq 'Move' -and (Test-Path -LiteralPath $target)) {
            if ((Test-Path -LiteralPath $target -PathType Leaf) -and (Get-FileSha256 $target) -eq $hash) { $targetState = 'Identical' } else { $targetState = 'CONFLICT' }
        }
        $plan.Add([pscustomobject]@{ Action = $action; RelativePath = $relative; Length = $file.Length; SHA256 = $hash; TargetState = $targetState; Result = '' })
    }
    foreach ($group in @($plan | Group-Object Action | Sort-Object Name)) {
        Write-Host ('--- {0}: {1} file(s)' -f $group.Name, $group.Count) -ForegroundColor Cyan
        foreach ($item in @($group.Group | Sort-Object RelativePath)) {
            $suffix = ''
            if ($item.TargetState) { $suffix = '  [target {0}]' -f $item.TargetState }
            Write-Host ('{0,10:N0}  {1}{2}' -f $item.Length, $item.RelativePath, $suffix)
        }
    }
    $conflicts = @($plan | Where-Object TargetState -eq 'CONFLICT')
    $localClear = Assert-NoLocalRun

    if (-not $Apply) {
        if ($conflicts.Count) { Write-Host ('{0} target file(s) already exist with different content: -Apply will stop.' -f $conflicts.Count) -ForegroundColor Yellow }
        if (-not $localClear) { Write-Host 'Before -Apply: disable the orchestrator task and wait for the runs to end on every server.' -ForegroundColor Yellow }
        Write-Host 'Preview only. Run again with -Apply -OtherServersStopped to switch.' -ForegroundColor Green
        return
    }
    if (-not $OtherServersStopped) { throw '-Apply requires -OtherServersStopped once the orchestrator is stopped on the other servers.' }
    if (-not $localClear) { throw 'A SmartInventory run or an enabled SmartInventory scheduled task exists on this computer.' }
    if ($conflicts.Count) { throw ('Target files already exist with different content: {0}' -f (($conflicts | ForEach-Object RelativePath) -join ', ')) }
    if (-not $plan.Count) { throw 'No SmartInventory private file found: nothing to switch.' }

    # 1. Verified ZIP backup of every listed file.
    New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null
    $baseName = 'SmartInventory-ConfigMove-{0}-{1}' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss')
    $zipPath = Join-Path $archiveRoot ($baseName + '.zip')
    $journalPath = Join-Path $archiveRoot ($baseName + '.csv')
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($item in $plan) {
            $null = [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $codeRoot $item.RelativePath), $item.RelativePath.Replace('\', '/'))
        }
    } finally { $zip.Dispose() }
    $zip = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        foreach ($item in $plan) {
            $entry = $zip.GetEntry($item.RelativePath.Replace('\', '/'))
            if ($null -eq $entry) { throw "ZIP entry missing: $($item.RelativePath)" }
            $stream = $entry.Open()
            try { $entryHash = Get-StreamSha256 $stream } finally { $stream.Dispose() }
            if ($entryHash -ne $item.SHA256) { throw "ZIP entry hash mismatch: $($item.RelativePath)" }
        }
    } finally { $zip.Dispose() }
    Write-Host ('ZIP backup verified: {0} ({1} file(s))' -f $zipPath, $plan.Count) -ForegroundColor Green

    $suffix = '.moved-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    try {
        # 2-3. Copy, verify, then rename the original.
        foreach ($item in @($plan | Where-Object Action -eq 'Move')) {
            $source = Join-Path $codeRoot $item.RelativePath
            $target = Join-Path $configRoot $item.RelativePath
            if ($item.TargetState -ne 'Identical') {
                New-Item -ItemType Directory -Path (Split-Path -Path $target -Parent) -Force | Out-Null
                Copy-Item -LiteralPath $source -Destination $target -ErrorAction Stop
            }
            if ((Get-FileSha256 $target) -ne $item.SHA256) { throw "Copy verification failed: $target" }
            [IO.File]::Move($source, $source + $suffix)
            $item.Result = 'Moved'
        }
        # 4. Obsolete configurations and manual backups stay only in the ZIP.
        foreach ($item in @($plan | Where-Object Action -eq 'Archive')) {
            $source = Join-Path $codeRoot $item.RelativePath
            if ((Get-FileSha256 $source) -ne $item.SHA256) { throw "File changed after the ZIP backup: $source" }
            [IO.File]::Delete($source)
            $item.Result = 'Archived'
        }
        foreach ($item in @($plan | Where-Object Action -eq 'Report')) {
            $item.Result = 'KeptInPlace'
            Write-Host ('Interrupted replacement kept in place, compare it with the active file: {0}' -f $item.RelativePath) -ForegroundColor Yellow
        }
        # 5. Activation and mapping check.
        Set-GlobalInventoryConfigRoot $ConfigRootPath
        foreach ($item in @($plan | Where-Object { $_.Action -eq 'Move' -and $_.RelativePath -match '\.json(\.txt)?$' })) {
            $mapped = ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path $codeRoot $item.RelativePath)
            if ($mapped -ne (Join-Path $configRoot $item.RelativePath)) { throw "Mapping check failed: $($item.RelativePath) -> $mapped" }
        }
        Write-Host ('Switch completed. Moved: {0}; archived: {1}; kept in place: {2}.' -f @($plan | Where-Object Result -eq 'Moved').Count, @($plan | Where-Object Result -eq 'Archived').Count, @($plan | Where-Object Result -eq 'KeptInPlace').Count) -ForegroundColor Green
        Write-Host 'Next: enable the orchestrator task on every server and check the first runs.' -ForegroundColor Green
    } catch {
        # Never leave renamed originals without an active configuration root: put them back.
        $activated = (Read-SmartM365JsonDocument -Path (Join-Path $smartRoot 'Config\SmartM365.global.local.json')).Document.PSObject.Properties['SmartInventoryConfigRootPath']
        if ($null -eq $activated -or [string]::IsNullOrWhiteSpace([string]$activated.Value)) {
            foreach ($item in @($plan | Where-Object Result -eq 'Moved')) {
                $source = Join-Path $codeRoot $item.RelativePath
                if (-not (Test-Path -LiteralPath $source) -and (Test-Path -LiteralPath ($source + $suffix))) { [IO.File]::Move($source + $suffix, $source); $item.Result = 'RestoredAfterFailure' }
            }
            Write-Host 'Switch failed before activation: moved originals were restored; archived files remain in the ZIP.' -ForegroundColor Yellow
        }
        throw
    } finally {
        $plan | Export-Csv -LiteralPath $journalPath -NoTypeInformation -Encoding UTF8
        Write-Host ('Journal: {0}' -f $journalPath)
    }
}
catch {
    $status = 'Failed'
    Write-Host ('ERROR: {0}' -f $_.Exception.Message) -ForegroundColor Red
    throw
}
finally {
    Write-SmartM365CompletionBanner -Status $status -ScriptName $global:SmartM365ScriptName -StartedAt $global:SmartM365ExecutionStartTime
}
