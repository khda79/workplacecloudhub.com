#Requires -Version 5.1
<# Executes only the helper-copy blocks extracted from the real launchers, in temporary fixtures. #>
[CmdletBinding()]
param([string]$LauncherDirectory)
$ErrorActionPreference = 'Stop'
$checks = 0
function Check([bool]$Ok, [string]$Message) { if (-not $Ok) { throw $Message }; $script:checks++ }
$root = Join-Path ([IO.Path]::GetTempPath()) ('mailbox-cache-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root -Force
$smartRoot = Split-Path $PSScriptRoot -Parent
$launcherRoot = Join-Path $smartRoot 'SmartInventory/Launchers/OnPremises'
if ($LauncherDirectory) { $launcherRoot = $LauncherDirectory }
$sourceHelper = Join-Path $smartRoot 'SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1'
$ps5 = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
$names = @('Start-SmartM365-Exchange-Local-Mailboxes-Inventory.cmd','Start-SmartM365-Exchange-Local-Mailboxes-InventoryWithADPermissions.cmd','Start-SmartM365-Exchange-Local-Mailboxes-Inventory-OnlyADPermissions.cmd')
try {
    foreach ($name in $names) {
        $text = [IO.File]::ReadAllText((Join-Path $launcherRoot $name))
        $match = [regex]::Match($text, '(?s)rem BEGIN Mailboxes identity helper cache.*?rem END Mailboxes identity helper cache')
        Check $match.Success ($name + ': dependency-copy block missing')
        Check ($text.IndexOf($match.Value) -lt $text.IndexOf('-File "%LOCAL_SCRIPT_DIR%')) ($name + ': cache must precede collector')
        Check ($text -match "Get-ChildItem.*'%LOCAL_COMMON_DIR%'") ($name + ': helper included in unblock')
        foreach ($scenario in @('Present','MissingSource','CopyFailure','MissingAfterCopy')) {
            $folder = Join-Path $root ($names.IndexOf($name).ToString() + '-' + $scenario)
            $source = Join-Path $folder 'source'; $cache = Join-Path $folder 'cache'
            $common = Join-Path $source 'SmartInventory/Common'
            $null = New-Item -ItemType Directory -Path $common -Force
            if ($scenario -ne 'MissingSource') { Copy-Item -LiteralPath $sourceHelper -Destination $common }
            $stub = Join-Path $folder 'stub'
            $null = New-Item -ItemType Directory -Path $stub -Force
            # Force a copy error or a misleading success-without-file, without touching real caches.
            if ($scenario -in @('CopyFailure','MissingAfterCopy')) {
                $exit = if ($scenario -eq 'CopyFailure') { 8 } else { 0 }
                "@exit /b $exit" | Set-Content -LiteralPath (Join-Path $stub 'robocopy.cmd') -Encoding ASCII
            }
            $harness = Join-Path $folder 'copy-only.cmd'
            $copyBlock = $match.Value
            # CALL returns from a batch stub; the production executable needs no CALL.
            if ($scenario -in @('CopyFailure','MissingAfterCopy')) { $copyBlock = $copyBlock -replace '(?m)^robocopy ', 'call robocopy ' }
            $lines = @('@echo off','setlocal EnableExtensions',('set "SMARTM365_ROOT={0}"' -f $source),('set "LOCAL_SMARTM365_ROOT={0}"' -f $cache),('set "CACHE_LOG={0}"' -f (Join-Path $folder 'copy.log')),('set "POWERSHELL5={0}"' -f $ps5),('set "PATH={0};%PATH%"' -f $stub),$copyBlock,'echo COPY_GATE_PASSED','exit /b 0')
            [IO.File]::WriteAllText($harness, ($lines -join "`r`n"), [Text.ASCIIEncoding]::new())
            $output = (& $env:ComSpec /d /c ('"' + $harness + '"') 2>&1) -join "`n"
            $rc = $LASTEXITCODE
            if ($scenario -eq 'Present') {
                Check ($rc -eq 0 -and $output -match 'COPY_GATE_PASSED') ($name + ': copy success: ' + $output)
                $cachedHelper = Join-Path $cache 'SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1'
                Check ((Get-FileHash $sourceHelper).Hash -eq (Get-FileHash $cachedHelper).Hash) ($name + ': copied bytes differ')
                $loaded = & $ps5 -NoProfile -Command ". '$cachedHelper'; if (-not (Get-Command Read-IdentityDate -ErrorAction SilentlyContinue)) { exit 7 }; 'HELPER_LOADED'"
                Check ($LASTEXITCODE -eq 0 -and $loaded -contains 'HELPER_LOADED') ($name + ': helper must load under PS5.1')
            }
            else {
                Check ($rc -eq 1 -and $output -notmatch 'COPY_GATE_PASSED') ($name + ': must stop before collector: ' + $scenario + ': ' + $output)
                Check ($output -match '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[ERROR\].*cache log:') ($name + ': actionable timestamped error: ' + $scenario)
                Check ($output.Contains($source) -and $output.Contains($cache)) ($name + ': source/destination diagnostic')
            }
            Write-Output ('PASS: ' + $name + ' / ' + $scenario)
        }
    }
}
finally {
    if (-not ([IO.Path]::GetFullPath($root)).StartsWith(([IO.Path]::GetFullPath([IO.Path]::GetTempPath())).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe fixture cleanup path' }
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
Write-Output ("PASS: {0} Mailboxes cache checks; {1}; no launcher entry point or collector executed." -f $checks, $PSVersionTable.PSVersion)
