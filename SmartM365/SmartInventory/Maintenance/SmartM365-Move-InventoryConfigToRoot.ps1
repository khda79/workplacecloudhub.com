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
1.0.1
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
$ScriptVersion = '1.0.1'
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
    # COM handler and other non-exec actions have no Execute property.
    $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object {
            $_.State -ne 'Disabled' -and (@($_.Actions | Where-Object { $_.PSObject.Properties['Execute'] } | ForEach-Object { '{0} {1}' -f $_.Execute, $_.Arguments }) -join ' ').IndexOf($smartRoot, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCTYLbn6Plv61Vp
# d+jq+YR0Pcrb5Sl9MMzW8UnR4pYjy6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIMnmZE7pkHJpAuRZ1kFOkSw3Ui7qED7C5aOuIRjBg/aMA0GCSqG
# SIb3DQEBAQUABIIBgG/Q5hgzS0i2UGpoEp+VH20yBMTO8ZMZFi5zN/YYjKW2NGmx
# Kn0u4NCXYa4gMSfuqmPZweueSM5ybReX3ry6KH90XMcvMy2Mg/BoU0MYsxX4cZM+
# aLpqZyDyFTDVLXPjVV3tNyMG4eNNo6awc4dnaTX5CN21vX2Y2139/+sBoAjn7L/w
# hfwNku0VfpaBOyvGtsFiCC05wCoX0Q8XgDXPA+Ux9jvmTRkpFOzYF7tGj/jGCgGq
# 90PiUARUwhsWH0pkrXZTNeUTGf4e2PzJb62E0S4aPM7teGXk2BSn+AitLuYUqNag
# qN9pWzpMvElSHqOgkHQxb7yt13XmS+fRgu8j6bVPPknHY9eiVqskKevhuUNwmT8P
# MzzbOKX/t8BFRb44m5r1d40lS9ye29CytCEsFWBAfn6aQVgGWckSQ2u0+8cI+4z4
# vtQ2E9CqPHDlUAissUFW7odL7jI4Zp1+LRhaOgV60XAMSzaN0fayG/vTovYVZ61f
# iiFT47DAGGosXOqzM6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxMzM1
# NDFaMC8GCSqGSIb3DQEJBDEiBCAAeEijQtC3rcZEG7wVPmzrU9folCm+kOqh2dYx
# jQFWmzANBgkqhkiG9w0BAQEFAASCAgAZ8Glg2/gU8ttuQ+JtcuTtNbF0nmkZND1b
# 5CLCIwC9iEmTfdPkrMw/3qfgWnJvaj0R0uBH3gwjQNzS1A5889JEHBihUFgDEBSk
# ISlZm1Kvnhv6gjNVPT16YAqgEVDxN8iwve7p6vEkVTzWXTqwKlYvOi1Ul5dMZa4m
# 4TZYusPlqtc2/e+ayjI8+CctXhdzpSSJ9BjG3LTO/FwA4A31xYcEVnNIeVfLGS0O
# IvLbhS5L7O5xmyztx4ooGKRMJZkJRS9DLRbvY52rLfYO5CH1hNqcXMeWepZuP7C1
# Hoqz13QVTgp9lhPcJl4pjCwjtC7eWtaD01vHXElFjE0uRXmfN6X54FxPU+kwKUOL
# qVkHHKDf1/UoRO6ZcpKgwl+IzjMAYELmhGBRVuxRHn2vH2mCiCOqIGeqpgLkIqlv
# tClaj95Qa0xRGhZzaHbsysbeVEDwLPZaNtfP7i6zdMsSTHthuNPP+qzgbOofCJVN
# ybw8tzkv9WmZPwoe2TZ8De0IYlswJPg2FdE91hdxKBFfYIHwIaNtlR0NE+CJfsh+
# RtJRHIH+wbDqZa1MENxAG20SsMa2IL2Ko/TFI+9Rda0fy8B0oML9bfycv12DoVSe
# B8yjtaZjDTRseqF2K7mWZUn7e7fQQKEtqbZrinPTz1yjUOiRW4jBFGkHOJf1IQHb
# SUsAgyaaeA==
# SIG # End signature block
