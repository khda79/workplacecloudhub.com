<#
.SYNOPSIS
Offline test of the duration estimates of the orchestrator GUI: estimated time left in Running
jobs, estimated duration in Pending jobs, both frames stacked in Operations > Live.
.VERSION
1.0.0
.NOTES
Runs the GUI helpers extracted from the script on synthetic rows. No WPF window, tenant
initialization, shared folder or orchestrator.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$guiPath = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator-GUI.ps1'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$guiText = [IO.File]::ReadAllText($guiPath)
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$null, [ref]$null)
foreach ($name in @('Get-OrchestratorGuiPropertyValue', 'Get-GuiJobDurationEstimate', 'Add-GuiJobDurationEstimates')) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Missing GUI function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}

# Average of the recent successful runs first, then the catalog estimate, else nothing.
$script:HealthByName = @{ 'Avg-Job' = [pscustomobject]@{ AverageDurationMinutes = 30.4 }; 'Zero-Job' = [pscustomobject]@{ AverageDurationMinutes = 0 } }
$script:DraftJobs = [pscustomobject]@{ Jobs = @([pscustomobject]@{ Name = 'Avg-Job'; EstimatedDurationMinutes = 90 }, [pscustomobject]@{ Name = 'Zero-Job'; EstimatedDurationMinutes = 12 }, [pscustomobject]@{ Name = 'Bare-Job' }) }
Check ((Get-GuiJobDurationEstimate -JobName 'Avg-Job').Minutes -eq 30) 'The average of the recent runs must win over the catalog estimate.'
Check ((Get-GuiJobDurationEstimate -JobName 'Zero-Job').Minutes -eq 12 -and (Get-GuiJobDurationEstimate -JobName 'Zero-Job').Source -like 'Catalog*') 'Without a usable average the catalog estimate must be used.'
Check ($null -eq (Get-GuiJobDurationEstimate -JobName 'Bare-Job') -and $null -eq (Get-GuiJobDurationEstimate -JobName 'Unknown-Job')) 'A job without history or catalog estimate has no estimate.'

$operations = [pscustomobject]@{
    Running = @(
        [pscustomobject]@{ Job = 'Avg-Job'; DurationMinutes = 10 }
        [pscustomobject]@{ Job = 'Zero-Job'; DurationMinutes = 20 }
        [pscustomobject]@{ Job = 'Bare-Job'; DurationMinutes = 5 }
    )
    Pending = @([pscustomobject]@{ Job = 'Avg-Job' }, [pscustomobject]@{ Job = 'Bare-Job' })
}
Add-GuiJobDurationEstimates -Operations $operations
Check ($operations.Running[0].EstimatedRemaining -eq '~20 min' -and $operations.Running[0].EstimateSource -like 'Average*30 min') "A running job must show the estimated time left ($($operations.Running[0].EstimatedRemaining))."
Check ($operations.Running[1].EstimatedRemaining -eq 'over by 8 min') "A run beyond its estimate must say by how much ($($operations.Running[1].EstimatedRemaining))."
Check ($operations.Running[2].EstimatedRemaining -eq '') 'A job without estimate must leave the cell empty.'
Check ($operations.Pending[0].EstimatedMinutes -eq 30 -and $null -eq $operations.Pending[1].EstimatedMinutes) 'Pending jobs must show the estimated duration.'
Add-GuiJobDurationEstimates -Operations $null
Check ($true) 'A failed operations read must be ignored.'

# GUI wiring: both refresh paths add the estimates; columns and stacked frames in Live.
Check (([regex]::Matches($guiText, '(?i)Add-GuiJobDurationEstimates -Operations \$operations')).Count -eq 2) 'Both operations refresh paths must add the estimates.'
Check ($guiText -match 'Header="Est\. remaining" Binding="\{Binding EstimatedRemaining\}"' -and $guiText -match 'Header="Est\. duration \(min\)" Binding="\{Binding EstimatedMinutes\}"') 'The estimate columns are missing.'
$live = $guiText.Substring($guiText.IndexOf('<TabItem Header="Live">'))
$live = $live.Substring(0, $live.IndexOf('</TabItem>'))
Check ($live -match '<GroupBox Header="Running jobs" Margin="0,0,0,6">' -and $live -match '<GroupBox Grid\.Row="1" Header="Pending jobs \(why they wait\)"' -and $live -notmatch 'Grid\.Column="1" Header="Pending jobs') 'Running and pending jobs must be stacked.'
"PASS: GUI duration estimates, $($script:passed) offline checks."
