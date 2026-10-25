<#
.SYNOPSIS
    Offline tests for the Smart Endpoint Migration Toolkit WPF GUI and its demonstration.
.DESCRIPTION
    Static checks of the window markup and launch helpers, then the GUI -ValidateOnly and the
    demo (-Demo -ValidateOnly with simulated button clicks) in child Windows PowerShell 5.1
    STA processes. Nothing is launched: no CMD, Run, inventory, PsExec or endpoint action.
    The technician state is redirected with SEMT_RUNTIME_ROOT_OVERRIDE and the package is
    copied to a temp folder so no GUI state file is written into the source tree.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$sourceRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-GuiTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$packageRoot = Join-Path $tempRoot 'SmartEndpointMigrationToolkit'
$previousOverride = $env:SEMT_RUNTIME_ROOT_OVERRIDE
$previousUpdate = $env:SEMT_GUI_UPDATE_CHECK
try {
    New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
    foreach ($item in @('Assets','Campaigns','Config','Modules','Scripts')) { Copy-Item -LiteralPath (Join-Path $sourceRoot $item) -Destination $packageRoot -Recurse -Force }
    foreach ($item in @('Runs','Waves\HybridJoin','Waves\Windows11Upgrade')) { New-Item -ItemType Directory -Path (Join-Path $packageRoot $item) -Force | Out-Null }
    Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter '*.local.json' -File | Remove-Item -Force
    $env:SEMT_RUNTIME_ROOT_OVERRIDE = Join-Path $tempRoot 'Runtime'
    $env:SEMT_GUI_UPDATE_CHECK = '0'
    $guiPath = Join-Path $packageRoot 'Scripts\SmartM365-EndpointMigration-GUI.ps1'
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)

    #region Static
    Test-Step 'GUI parses without error' { @($errors).Count -eq 0 }
    $assignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml' }, $true)
    $xaml = [string]$assignment.Right.Expression.Value
    Test-Step 'window markup is a literal here-string (demo framework contract)' { $assignment.Right.Expression -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $xaml -notmatch '\$\(|\$script:' }
    foreach ($name in @('CampaignCombo','WaveCombo','WaveModeCombo','LaunchWaveButton','LaunchAllButton','StopAllButton','SingleComputerText','LaunchSingleButton','NewWaveNameText','CreateWaveButton','AutomaticPreviewButton','AutomaticCreateButton','AutomaticRefreshInventoryButton','RunsGrid','OpenRunReportButton','StopSelectedRunButton','ForceStopSelectedRunButton','OptionsHost','ResetOptionsButton','HeaderLogoImage')) {
        Write-TestResult -Condition ($xaml -match ('x:Name="{0}"' -f $name)) -Message ("control present: {0}" -f $name)
    }
    Test-Step 'no Batch/LOT wording in the interface' { $xaml -notmatch '\b(LOT|LOTs|Lot|Batch|batch)\b' }
    Test-Step 'Preview and Create are separate controls; Create states it never launches' { $xaml -match 'AutomaticPreviewButton' -and $xaml -match 'never launches the Wave' }
    Test-Step 'Launch All is labelled as limited to the selected Campaign in its confirmation' { [IO.File]::ReadAllText($guiPath) -match 'Launch All Waves of this Campaign only' }
    Test-Step 'stop states remote endpoints are reported separately' { [IO.File]::ReadAllText($guiPath) -match 'Remote endpoints may continue and must be verified' }
    $guiText = [IO.File]::ReadAllText($guiPath)
    foreach ($functionName in @('ConvertTo-PowerShellLiteral','Get-AutomaticWaveCommandText','Get-AutomaticFilterSignature','Split-GuiList','Get-GuiRunCommand','Assert-CmdLiteralValue','ConvertTo-CmdArgument','Get-GuiStatePath','New-GuiSelectionFile')) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    $script:PackageRoot = $packageRoot
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
    $filters = @{ WaveName = 'WAVE-AUTO-TEST'; Prefixes = @('CTSO-'); Contains = @("O'Brien"); ExcludeStaleAd = $true; LastLogonDays = 30; Source = 'Both'; ExcludeIntunePresent = $true }
    foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
        $engine = Join-Path $packageRoot ("Campaigns\{0}\Scripts\SmartM365-EndpointMigration-{0}-New-AutomaticWave.ps1" -f $campaign)
        $engineAst = [System.Management.Automation.Language.Parser]::ParseFile($engine, [ref]$null, [ref]$null)
        $engineParameters = @($engineAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        $preview = Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence'
        $create = Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence' -Create
        $used = @([regex]::Matches($create, '(?<=\s)-(?<Name>[A-Za-z]+)') | ForEach-Object { $_.Groups['Name'].Value } | Select-Object -Unique)
        $unknown = @($used | Where-Object { $engineParameters -notcontains $_ })
        Write-TestResult -Condition ($unknown.Count -eq 0) -Message ("{0}: Automatic Wave command uses only engine parameters" -f $campaign) -Detail ($unknown -join ', ')
        Write-TestResult -Condition ($preview -notmatch '-Create\b' -and $create -match '-Create$') -Message ("{0}: Preview never passes -Create; Create does" -f $campaign)
        Write-TestResult -Condition ($create -match "'O''Brien'") -Message ("{0}: filter values are quoted as PowerShell literals" -f $campaign)
    }
    Test-Step 'filter signature changes when a filter changes (Create needs a matching Preview)' {
        $a = Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $filters
        $changed = $filters.Clone(); $changed.LastLogonDays = 31
        $a -ne (Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $changed)
    }
    $run = Get-GuiRunCommand -Campaign Windows11Upgrade -WavePath 'C:\Waves\W 1' -RunMode Once -IgnoreRunGuard $true -SelectionPath 'C:\Sel\s.json'
    Test-Step 'Run command uses Windows PowerShell 5.1, the engine, the GUI selection and the run mode' { $run -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $run -match 'SmartM365-Invoke-EndpointMigrationWithPsExec\.ps1' -and $run -match '-GuiSelectionPath' -and $run -match '-RunMode Once' -and $run -match '-IgnoreRunGuard' -and $run -match '"C:\\Waves\\W 1"' }
    Test-Step 'Single PC command passes -ComputerName and no Wave' { $fictionalComputer = 'PC-1'; $c = Get-GuiRunCommand -Campaign HybridJoin -ComputerName $fictionalComputer -RunMode Once -IgnoreRunGuard $false -SelectionPath 'C:\s.json'; $c -match '-ComputerName PC-1' -and $c -notmatch '-WavePath' -and $c -notmatch '-IgnoreRunGuard' -and $c -match '-GuiSelectionPath' }
    $catalog = Import-SemtOptionCatalog -PackageRoot $packageRoot
    $values = @{ ThrottleLimit = 4; 'Windows11Upgrade.SetupSourcePath' = ''; DryRun = $true }
    $selection = New-GuiSelectionFile -Catalog $catalog -Campaign Windows11Upgrade -Values $values
    $captured = Read-SemtJsonFile -Path $selection
    Test-Step 'GUI selection is stored per Campaign under Runs\<Campaign>\_GuiSelections' { $selection -like (Join-Path $packageRoot 'Runs\Windows11Upgrade\_GuiSelections\*.json') -and $captured.Campaign -eq 'Windows11Upgrade' }
    Test-Step 'GUI selection omits empty text options so Wave/global values still apply' { [int]$captured.Options.ThrottleLimit -eq 4 -and -not $captured.Options.PSObject.Properties['Windows11Upgrade.SetupSourcePath'] }
    Test-Step 'engine accepts the GUI selection only for its own Campaign' {
        $ok = $false
        try { [void](Read-SemtGuiSelection -Path $selection -Campaign HybridJoin) } catch { $ok = $true }
        $layer = Read-SemtGuiSelection -Path $selection -Campaign Windows11Upgrade
        $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign Windows11Upgrade -WaveLayer @{ ThrottleLimit = 2 } -GuiLayer $layer
        $ok -and $resolved['ThrottleLimit'].Value -eq 4 -and $resolved['ThrottleLimit'].Source -eq 'Gui'
    }
    #endregion

    #region Windows PowerShell 5.1 STA
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $validate = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $guiPath -ValidateOnly 2>&1 | ForEach-Object { [string]$_ })
    Test-Step 'GUI -ValidateOnly loads under Windows PowerShell 5.1 STA for both Campaigns' { $LASTEXITCODE -eq 0 -and @($validate | Where-Object { $_ -match '^ValidateOnly: GUI 1\.0\.0 loaded; campaign=(HybridJoin|Windows11Upgrade); options=\d+' }).Count -eq 2 }
    Test-Step 'GUI -ValidateOnly writes no GUI state file' { @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter 'GuiState.*' -File).Count -eq 0 }
    $demoScript = Join-Path $tempRoot 'demo-check.ps1'
    @'
param($Gui)
$ErrorActionPreference = 'Stop'
$state = . $Gui -Demo -ValidateOnly
'VIEWS=' + ($state.ViewNames -join ',')
'OPTIONS1=' + $state.Data.Editors.Count
$state.Controls.CampaignCombo.SelectedIndex = 1
'CAMPAIGN=' + $state.Data.Campaign + ';OPTIONS2=' + $state.Data.Editors.Count
foreach ($name in 'LaunchWaveButton','LaunchAllButton','LaunchSingleButton','CreateWaveButton','AutomaticPreviewButton','AutomaticCreateButton','StopAllButton','ForceStopSelectedRunButton','OpenRunReportButton','ResetOptionsButton','RefreshButton') {
    $state.Controls[$name].RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
}
'FAILED=' + @($state.Activity | Where-Object { $_ -match 'Demo action failed' }).Count
'OUTPUT=' + $state.LastOutput
'NOLAUNCH=' + (@($state.Activity | Where-Object { $_ -match 'No launcher, PsExec, worker' }).Count -gt 0)
$state.Window.Close()
'@ | Set-Content -LiteralPath $demoScript -Encoding ASCII
    $demo = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $demoScript -Gui $guiPath 2>&1 | ForEach-Object { [string]$_ })
    $demoText = $demo -join "`n"
    Test-Step 'demo loads every view' { $demoText -match 'VIEWS=Waves,Single PC,New Wave,Automatic Wave,Runs,Options' }
    Test-Step 'demo shows the real catalog options for both Campaigns' { $demoText -match 'OPTIONS1=\d{2}' -and $demoText -match 'CAMPAIGN=Windows11Upgrade;OPTIONS2=\d{2}' }
    Test-Step 'every demo action completes without error' { $demoText -match 'FAILED=0' }
    Test-Step 'demo report is generated in the demo temp folder only' { $demoText -match 'OUTPUT=.*SmartM365-Demo.*Run-Report-DEMO\.html' }
    Test-Step 'demo states that nothing was launched' { $demoText -match 'NOLAUNCH=True' }
    Test-Step 'demo wrote no Wave, Run or configuration into the package' { @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Waves') -Directory -Recurse | Where-Object { $_.Parent.Name -in @('HybridJoin','Windows11Upgrade') }).Count -eq 0 -and @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter '*.local.json' -File).Count -eq 0 }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    $env:SEMT_RUNTIME_ROOT_OVERRIDE = $previousOverride
    $env:SEMT_GUI_UPDATE_CHECK = $previousUpdate
    try { if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force } } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0
