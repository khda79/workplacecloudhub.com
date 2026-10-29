<#
.SYNOPSIS
Offline test of the exact family match of RemoveOldFiles (NamePattern).

.VERSION
1.0.0

.NOTES
Loads RemoveOldFiles from the PowerShell 7 Core and the Windows PowerShell 5.1 compatibility module.
Checks that a count retention on Name_* keeps the newest Name_<timestamp> copies and never counts
or removes a longer family such as Name_Archive_<timestamp>, that MaxItems copies belong to their
family, that files written in this run are never removed, and that the legacy call without
NamePattern is unchanged. Synthetic files under TEMP only.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Test-SmartM365InterruptedRun { $false }
function Test-FileLocked { param($Path) $false }
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RemoveOldFiles-' + [guid]::NewGuid().ToString('N'))

function New-Family {
    param([string]$Folder)
    $null = New-Item -ItemType Directory -Path $Folder -Force
    $base = Get-Date '2026-10-01 04:00'
    foreach ($index in 0..5) {
        $stamp = $base.AddDays($index)
        foreach ($name in @(
                ('Exchange_EXO_Mailboxes_AllDomains_{0:yyyyMMdd_HHmmss}.csv' -f $stamp),
                ('Exchange_EXO_Mailboxes_AllDomains_Archive_{0:yyyyMMdd_HHmmss}.csv' -f $stamp))) {
            $path = Join-Path $Folder $name
            [IO.File]::WriteAllText($path, 'x')
            [IO.File]::SetLastWriteTime($path, $stamp)
        }
    }
    $maxItems = Join-Path $Folder ('Exchange_EXO_Mailboxes_AllDomains_MAXITEMS-10_{0:yyyyMMdd_HHmmss}.csv' -f $base.AddDays(-1))
    [IO.File]::WriteAllText($maxItems, 'x')
    [IO.File]::SetLastWriteTime($maxItems, $base.AddDays(-1))
}

try {
    $global:csvGeneratedPaths = [Collections.Generic.List[string]]::new()
    $global:LogTextFile = ''
    foreach ($relative in @('../Modules/SmartM365.Core/SmartM365.Core.psm1', '../Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $relative), [ref]$tokens, [ref]$errors)
        if ($errors.Count) { throw ($errors | Out-String) }
        $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'RemoveOldFiles' }, $true)
        Check ($null -ne $definition) "RemoveOldFiles is missing from $relative."
        . ([scriptblock]::Create($definition.Extent.Text))

        $base = 'Exchange_EXO_Mailboxes_AllDomains'
        $pattern = '^' + [regex]::Escape($base) + '(?:_MAXITEMS-\d+)?_\d{8}_\d{6}\.csv$'
        $folder = Join-Path $root ([guid]::NewGuid().ToString('N'))
        New-Family $folder
        $protected = Join-Path $folder 'Exchange_EXO_Mailboxes_AllDomains_20261001_040000.csv'
        $global:csvGeneratedPaths.Clear(); $global:csvGeneratedPaths.Add($protected)
        RemoveOldFiles -FolderPath $folder -FilePattern "$base`_*.csv" -NamePattern $pattern -MaxFiles 3
        $names = @(Get-ChildItem -LiteralPath $folder -File | ForEach-Object Name)
        $own = @($names | Where-Object { $_ -match $pattern })
        $archive = @($names | Where-Object { $_ -like '*_Archive_*' })
        Check ($archive.Count -eq 6) "A longer family was counted or removed in ${relative}: $($archive.Count) archive copies left."
        Check ($own.Count -eq 4) "The family kept $($own.Count) copies instead of three plus the current-run file in $relative."
        Check ($own -contains 'Exchange_EXO_Mailboxes_AllDomains_20261006_040000.csv' -and $own -contains 'Exchange_EXO_Mailboxes_AllDomains_20261004_040000.csv') "The newest copies were not kept in $relative."
        Check ($own -contains 'Exchange_EXO_Mailboxes_AllDomains_20261001_040000.csv') "A file written in this run was removed in $relative."
        Check (-not ($own -like '*MAXITEMS*')) "The oldest MaxItems copy of the family was not removed in $relative."

        # Legacy call without NamePattern: the wildcard still spans both families.
        $legacy = Join-Path $root ([guid]::NewGuid().ToString('N'))
        New-Family $legacy
        $global:csvGeneratedPaths.Clear()
        RemoveOldFiles -FolderPath $legacy -FilePattern "$base`_*.csv" -MaxFiles 3
        Check (@(Get-ChildItem -LiteralPath $legacy -File).Count -eq 3) "The call without NamePattern changed behaviour in $relative."
    }
}
finally {
    Remove-Variable -Name csvGeneratedPaths, LogTextFile -Scope Global -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
"Test-SmartM365RemoveOldFilesNamePatternOffline: $script:passed checks passed."
