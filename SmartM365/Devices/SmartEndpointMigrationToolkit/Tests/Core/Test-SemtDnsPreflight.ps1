#Requires -Version 5.1
<#
.SYNOPSIS
    Tests adaptive DNS preflight with deterministic, offline lookups.
.VERSION
    1.0.0
#>
$ErrorActionPreference = 'Stop'
$network = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'Modules\SmartEndpointMigration.Core\Functions\Semt.Network.ps1'
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($network, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Network helper failed to parse.' }
$fn = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-SemtDnsSample' }, $true)
if (-not $fn) { throw 'DNS preflight function missing.' }
. ([scriptblock]::Create($fn.Extent.Text))
function Get-Random { param([Parameter(ValueFromPipeline=$true)]$InputObject, [int]$Count) begin { $items = @() } process { $items += $InputObject } end { $items | Select-Object -First $Count } }
function Resolve-SemtConnectionTarget {
    param([string]$ComputerName,[string]$DomainSuffix)
    $script:lookups += $ComputerName
    [pscustomobject]@{ Computer=$ComputerName; ConnectionTarget=$ComputerName; Resolved=($ComputerName -eq $script:resolveName) }
}
function Assert-Case {
    param([int]$Size,[int]$SampleSize,[string]$Success,[int]$Expected,[bool]$AllFailed)
    $script:lookups = @(); $script:resolveName = $Success
    $names = @(for ($i = 1; $i -le $Size; $i++) { 'PC-{0:d3}' -f $i })
    $r = Test-SemtDnsSample -ComputerNames $names -SampleSize $SampleSize
    if ($r.Tested -ne $Expected -or $r.Total -ne $Size -or $r.AllFailed -ne $AllFailed -or $script:lookups.Count -ne $Expected -or @($r.Samples).Count -ne $Expected) { throw "Unexpected DNS result for size=$Size sample=$SampleSize success=$Success`: tested=$($r.Tested) total=$($r.Total) allFailed=$($r.AllFailed)" }
    if (@($script:lookups | Select-Object -Unique).Count -ne $Expected) { throw 'A computer was resolved twice.' }
}
Assert-Case 0 5 '' 0 $false
Assert-Case 1 5 '' 1 $true
Assert-Case 11 5 'PC-003' 5 $false
Assert-Case 11 5 'PC-008' 8 $false
Assert-Case 11 5 '' 11 $true
Assert-Case 100 5 '' 50 $true
Assert-Case 100 5 'PC-050' 50 $false
Assert-Case 100 50 '' 50 $true
Assert-Case 3 5 'PC-002' 3 $false
Write-Output 'Adaptive DNS preflight: 9 offline cases passed.'
