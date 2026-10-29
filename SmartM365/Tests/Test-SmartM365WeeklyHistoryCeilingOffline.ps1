<#
.SYNOPSIS
Offline test of the global WeeklyHistoryMaxRetentionWeeks ceiling on weekly history retention.

.VERSION
1.0.0

.NOTES
Loads Save-SmartM365WeeklyInventoryHistory and Get-SmartM365WeeklyHistoryMaxRetentionWeeks from the
PowerShell 7 Core and the Windows PowerShell 5.1 compatibility module, with the global configuration
stubbed. Checks the default ceiling of 12 weeks, a configured ceiling, a smaller script value, the
0 (keep all) script value, the 0 ceiling (no ceiling) and an unreadable value. Synthetic folders only;
SharePoint upload disabled.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$transport = Get-Module SmartM365.JsonTransport
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-WeeklyCeiling-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory $root
$script:passed = 0
$script:warnings = 0
$script:globalConfig = [pscustomobject]@{}
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function WriteLog { param($Message, $Level) if ($Level -eq 'WARNING') { $script:warnings++ } }
function Get-SmartM365IsoWeekName { '2026-W39' }
function Get-SmartM365WeeklyHistoryFileName { param($Path) [IO.Path]::GetFileName($Path) }
function Copy-SmartM365FileAtomically { param($SourcePath, $DestinationPath) Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath }
function Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) return 'receipt' }
function Get-SmartM365EffectiveModuleGlobalConfig { $script:globalConfig }

function Invoke-Case {
    param([object]$Ceiling, [int]$ScriptWeeks)
    $script:globalConfig = if ($null -eq $Ceiling) { [pscustomobject]@{} } else { [pscustomobject]@{ WeeklyHistoryMaxRetentionWeeks = $Ceiling } }
    $case = Join-Path $root ([guid]::NewGuid().ToString('N'))
    $history = Join-Path $case 'WeeklyHistory'
    # Fifteen older weeks plus the current one.
    foreach ($number in 1..15) { $null = New-Item -ItemType Directory (Join-Path $history ('2026-W{0:00}' -f $number)) -Force }
    $source = Join-Path $case 'Inventory.csv'
    [IO.File]::WriteAllText($source, "Id`r`n42`r`n")
    Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks $ScriptWeeks
    $weeks = @(Get-ChildItem -LiteralPath $history -Directory | Sort-Object Name -Descending | ForEach-Object Name)
    Check ($weeks[0] -eq '2026-W39') 'The current week was not kept first.'
    return $weeks
}

try {
    $global:EnableSharePointUpload = $false
    & $transport { function script:Get-SmartM365JsonTransportPolicy { @{Mode='Readers';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
    foreach ($relative in @('../Modules/SmartM365.Core/SmartM365.Core.psm1', '../Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $relative), [ref]$tokens, [ref]$errors)
        if ($errors.Count) { throw ($errors | Out-String) }
        foreach ($name in 'Save-SmartM365WeeklyInventoryHistory', 'Get-SmartM365WeeklyHistoryMaxRetentionWeeks') {
            $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
            Check ($null -ne $definition) "$name is missing from $relative."
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        $weeks = Invoke-Case -Ceiling $null -ScriptWeeks 52
        Check ($weeks.Count -eq 12) "Default ceiling kept $($weeks.Count) weeks in $relative."
        Check ($weeks[-1] -eq '2026-W05') "Default ceiling kept the wrong weeks in ${relative}: $($weeks -join ',')"
        $weeks = Invoke-Case -Ceiling 3 -ScriptWeeks 52
        Check ($weeks.Count -eq 3) "Ceiling 3 kept $($weeks.Count) weeks in $relative."
        $weeks = Invoke-Case -Ceiling '3' -ScriptWeeks 52
        Check ($weeks.Count -eq 3) "String ceiling kept $($weeks.Count) weeks in $relative."
        $weeks = Invoke-Case -Ceiling 3 -ScriptWeeks 2
        Check ($weeks.Count -eq 2) "A smaller script value did not apply in $relative."
        $weeks = Invoke-Case -Ceiling 3 -ScriptWeeks 0
        Check ($weeks.Count -eq 3) "Script value 0 escaped the ceiling in $relative."
        $weeks = Invoke-Case -Ceiling 0 -ScriptWeeks 0
        Check ($weeks.Count -eq 16) "Ceiling 0 with script value 0 did not keep every week in $relative."
        $weeks = Invoke-Case -Ceiling 0 -ScriptWeeks 52
        Check ($weeks.Count -eq 16) "Ceiling 0 did not remove the ceiling in $relative."
        $before = $script:warnings
        $weeks = Invoke-Case -Ceiling 'twelve' -ScriptWeeks 52
        Check ($weeks.Count -eq 12) "An unreadable ceiling did not fall back to 12 in $relative."
        Check ($script:warnings -gt $before) "An unreadable ceiling was not logged in $relative."
    }
}
finally {
    Remove-Variable -Name EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
[pscustomobject]@{ Passed = $script:passed; Scope = 'Synthetic weekly folders; global configuration stubbed; no upload' }
