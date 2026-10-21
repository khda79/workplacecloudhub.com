#Requires -Version 7.0
<#
.SYNOPSIS
Offline guard: every SmartM365.Core function that a script or module outside Core calls is exported by
the Core module, so the call resolves in production (tests that dot-source single Core functions do not
prove that). A script that defines the function itself is not affected.
.VERSION
1.0.0
.NOTES
Parses repository scripts only; imports the Core module without running any collector. No tenant,
network or synchronized data access.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$coreFolder = Join-Path $root 'Modules/SmartM365.Core'
# Exports of every module of the Core folder (JSON transport functions come from their own module).
$exported = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$modules = foreach ($manifest in Get-ChildItem -LiteralPath $coreFolder -File -Filter '*.psd1') { Import-Module $manifest.FullName -Force -PassThru }
foreach ($module in $modules) { foreach ($name in $module.ExportedFunctions.Keys) { [void]$exported.Add($name) } }

# Functions defined in the Core module file itself. Helper files it dot-sources (inventory lifecycle,
# receipts...) are also dot-sourced by the scripts that use them, so they are not checked here.
$defined = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$coreAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $coreFolder 'SmartM365.Core.psm1'), [ref]$null, [ref]$null)
foreach ($function in $coreAst.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) { [void]$defined.Add($function.Name) }
$hidden = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($name in $defined) { if (-not $exported.Contains($name)) { [void]$hidden.Add($name) } }

$violations = [Collections.Generic.List[string]]::new()
$scanned = 0
$sources = Get-ChildItem -LiteralPath $root -Recurse -File -Include '*.ps1', '*.psm1' | Where-Object {
    $_.FullName -notlike "$coreFolder*" -and $_.FullName -notmatch '[\\/]Tests?[\\/]' -and $_.Name -notlike 'Test-*'
}
foreach ($source in $sources) {
    $ast = [Management.Automation.Language.Parser]::ParseFile($source.FullName, [ref]$null, [ref]$null)
    $local = [Collections.Generic.HashSet[string]]::new([string[]]@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name), [StringComparer]::OrdinalIgnoreCase)
    foreach ($call in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] }, $true)) {
        $name = $call.GetCommandName()
        if (-not $name -or -not $hidden.Contains($name) -or $local.Contains($name)) { continue }
        # A script block run inside the module (& $coreModule { ... }) sees internal functions.
        $insideModule = $false
        for ($node = $call.Parent; $node -and -not $insideModule; $node = $node.Parent) {
            if ($node -is [Management.Automation.Language.CommandAst] -and $node.InvocationOperator -eq 'Ampersand' -and
                $node.CommandElements[0] -is [Management.Automation.Language.VariableExpressionAst] -and
                @($node.CommandElements | Where-Object { $_ -is [Management.Automation.Language.ScriptBlockExpressionAst] }).Count -gt 0) { $insideModule = $true }
        }
        if (-not $insideModule) {
            $violations.Add(('{0}:{1} {2}' -f $source.FullName.Substring($root.Length + 1), $call.Extent.StartLineNumber, $name))
        }
    }
    $scanned++
}
foreach ($module in $modules) { Remove-Module $module -Force -ErrorAction SilentlyContinue }
if ($violations.Count -gt 0) { throw ("Core functions called outside Core but not exported:`n" + ($violations -join "`n")) }
"PASS: $scanned scripts call only exported SmartM365.Core functions ($($hidden.Count) internal, $($exported.Count) exported); PowerShell $($PSVersionTable.PSVersion)."
