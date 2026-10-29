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
    Test-Step 'main window uses the full height of the screen work area and fits on small screens' { $sizeText = [IO.File]::ReadAllText($guiPath); $sizeText -match '\$window\.Height = \$workArea\.Height' -and $sizeText -match '\$window\.MinHeight = \[Math\]::Min\(\$window\.MinHeight, \$workArea\.Height\)' -and $sizeText -match '\$window\.Top = \$workArea\.Top' }
    Test-Step 'no Batch/LOT wording in the interface' { $xaml -notmatch '\b(LOT|LOTs|Lot|Batch|batch)\b' }
    Test-Step 'Preview and Create are separate controls; Create states it never launches' { $xaml -match 'AutomaticPreviewButton' -and $xaml -match 'never launches the Wave' }
    Test-Step 'Launch All is labelled as limited to the selected Campaign in its confirmation' { [IO.File]::ReadAllText($guiPath) -match 'Launch All Waves of this Campaign only' }
    Test-Step 'stop states remote endpoints are reported separately' { [IO.File]::ReadAllText($guiPath) -match 'Remote endpoints may continue and must be verified' }
    $guiText = [IO.File]::ReadAllText($guiPath)
    foreach ($functionName in @('ConvertTo-PowerShellLiteral','Get-AutomaticWaveCommandText','Get-AutomaticFilterSignature','Split-GuiList','Get-GuiRunCommand','Assert-CmdLiteralValue','ConvertTo-CmdArgument','Get-GuiStatePath','New-GuiSelectionFile','Get-InventoryRefreshCommandText','Get-AdInventoryDomains','Get-InventoryRefreshPlan','Get-AutomaticWaveProcessText','Get-AutomaticWaveLastStage','Get-AutomaticWaveResultText')) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    $script:PackageRoot = $packageRoot
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
    # Wave options tab (replaces the "Wave options" button): Options first, then Activity; editors read and
    # write Wave.local.json with inheritance. A fictional Wave in a temporary folder; nothing is launched.
    Test-Step 'Waves view: Options tab first (selected), then Activity; no Wave options button' {
        $xaml -notmatch 'OpenWaveConfigButton' -and $xaml -match '(?s)x:Name="WaveSideTabs".*?x:Name="WaveOptionsTab" Header="Options" IsSelected="True".*?x:Name="ActivityTab" Header="Activity".*?x:Name="ActivityText"' -and $xaml -match 'x:Name="WaveOptionsHost"' -and $xaml -match 'x:Name="SaveWaveOptionsButton"'
    }
    foreach ($functionName in @('Register-NumericStepper','Add-SemtGuiOptionControls','Get-SemtGuiOptionValues','Update-WaveOptionsView','Save-WaveOptions')) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $script:CategoryOrder = @('Mode','Actions','Setup','Safety','Schedule','Parallelism','Network','Inventory','Evidence')
    $script:InheritChoiceLabel = '(inherit)'
    $script:Catalog = Import-SemtOptionCatalog -PackageRoot $packageRoot
    $script:CurrentCampaign = 'HybridJoin'
    $window = New-Object System.Windows.Window
    $savedStatus = [System.Collections.Generic.List[string]]::new()
    function Add-Status { param([string]$Message,[string]$Title = 'Ready') $savedStatus.Add($Title) }
    $controls = @{ WaveOptionsHost = (New-Object System.Windows.Controls.StackPanel); WaveOptionsTitleText = (New-Object System.Windows.Controls.TextBlock); SaveWaveOptionsButton = (New-Object System.Windows.Controls.Button); InheritWaveOptionsButton = (New-Object System.Windows.Controls.Button) }
    $optionsWave = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-WaveOptions-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $optionsWave -Force | Out-Null
    try {
        $waveFile = Join-Path $optionsWave 'Wave.local.json'
        Copy-Item -LiteralPath (Join-Path $packageRoot 'Campaigns\HybridJoin\Templates\Wave\Wave.local.json.template') -Destination $waveFile
        $template = Get-Content -LiteralPath $waveFile -Raw | ConvertFrom-Json
        $template.Description = 'Fictional Wave'
        $template.Options.ThrottleLimit = 5
        $template | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $waveFile -Encoding UTF8
        $script:SelectedWave = [pscustomobject]@{ Name = 'WAVE-OPTIONS-TEST'; Path = $optionsWave }
        Update-WaveOptionsView
        $editors = $script:WaveEditors
        Test-Step 'Wave options: one editor per visible option; the Wave value and inheritance are shown' {
            $editors.Count -gt 10 -and [string]$editors['ThrottleLimit'].Control.Text -eq '5' -and $null -eq $editors['DryRun'].Control.IsChecked -and $editors['DryRun'].Control.IsThreeState -and [string]$editors['DelayBetweenCyclesMinutes'].Control.Text -eq '' -and $controls.SaveWaveOptionsButton.IsEnabled
        }
        $choice = @($editors.Values | Where-Object { [string]$_.Definition.Type -eq 'Choice' } | Select-Object -First 1)
        Test-Step 'Wave options: a choice offers (inherit), selected while the Wave inherits' { $choice.Count -eq 0 -or ([string]$choice[0].Control.Items[0] -eq '(inherit)' -and [string]$choice[0].Control.SelectedItem -eq '(inherit)') }
        $editors['ThrottleLimit'].Control.Text = '7'
        $editors['DryRun'].Control.IsChecked = $true
        Save-WaveOptions
        $saved = Get-Content -LiteralPath $waveFile -Raw | ConvertFrom-Json
        Test-Step 'Save Wave options writes the set values and keeps inheritance and the other properties' {
            [int]$saved.Options.ThrottleLimit -eq 7 -and $saved.Options.DryRun -eq $true -and [string]$saved.Options.DelayBetweenCyclesMinutes -eq '__INHERIT__' -and [string]$saved.Description -eq 'Fictional Wave' -and [string]$saved.Campaign -eq 'HybridJoin' -and $savedStatus -contains 'Wave options saved'
        }
        $effective = Resolve-SemtEffectiveOptions -Catalog $script:Catalog -Campaign HybridJoin -WaveLayer $saved.Options
        Test-Step 'saved Wave options are used by the option resolution (Wave layer)' { $effective['ThrottleLimit'].Value -eq 7 -and $effective['ThrottleLimit'].Source -eq 'Wave' -and $effective['DelayBetweenCyclesMinutes'].Source -ne 'Wave' }
        $script:WaveEditors['ThrottleLimit'].Control.Text = '9999'
        Test-Step 'an invalid Wave option is refused and the file is unchanged' { $failed = $false; try { Save-WaveOptions } catch { $failed = $_.Exception.Message -match 'SEMT_OPTION_INVALID' }; $failed -and [int]((Get-Content -LiteralPath $waveFile -Raw | ConvertFrom-Json).Options.ThrottleLimit) -eq 7 }
        Save-WaveOptions -InheritAll
        $inherited = Get-Content -LiteralPath $waveFile -Raw | ConvertFrom-Json
        Test-Step 'Inherit all sets every visible option to __INHERIT__' { [string]$inherited.Options.ThrottleLimit -eq '__INHERIT__' -and [string]$inherited.Options.DryRun -eq '__INHERIT__' }
        $script:SelectedWave = $null
        Update-WaveOptionsView
        Test-Step 'without a selected Wave the options are empty and Save is disabled' { $controls.WaveOptionsHost.Children.Count -eq 0 -and -not $controls.SaveWaveOptionsButton.IsEnabled }
    }
    finally {
        Remove-Item -LiteralPath $optionsWave -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path Function:\Add-Status -ErrorAction SilentlyContinue
    }
    $filters = @{ WaveName = 'WAVE-AUTO-TEST'; Prefixes = @('CTSO-'); Contains = @("O'Brien"); ExcludeStaleAd = $true; LastLogonDays = 30; Source = 'Both'; ExcludeIntunePresent = $true; AdDomain = 'fr.example' }
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
        $reuseCreate = Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence' -Create -PreviewEvidencePath 'C:\Evidence\Preview'
        $reuseUsed = @([regex]::Matches($reuseCreate, '(?<=\s)-(?<Name>[A-Za-z]+)') | ForEach-Object { $_.Groups['Name'].Value } | Where-Object { $engineParameters -notcontains $_ })
        Write-TestResult -Condition ($reuseCreate -match "-Create -PreviewEvidencePath 'C:\\Evidence\\Preview'$" -and $reuseUsed.Count -eq 0 -and (Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence' -PreviewEvidencePath 'C:\Evidence\Preview') -notmatch 'PreviewEvidencePath') -Message ("{0}: Create passes the previewed evidence to the engine (Preview never does)" -f $campaign)
        $allDomains = $filters.Clone(); $allDomains.AdDomain = ''
        Write-TestResult -Condition ($preview -match "-AdDomainFilter @\('fr\.example'\)" -and (Get-AutomaticWaveCommandText -Campaign $campaign -Filters $allDomains -EvidenceRoot 'C:\Evidence') -notmatch 'AdDomainFilter') -Message ("{0}: AD domain is passed only when a domain is chosen" -f $campaign)
    }
    Test-Step 'filter signature changes when a filter changes (Create needs a matching Preview)' {
        $a = Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $filters
        $changed = $filters.Clone(); $changed.LastLogonDays = 31
        $domain = $filters.Clone(); $domain.AdDomain = 'be.example'
        $a -ne (Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $changed) -and $a -ne (Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $domain)
    }
    $run = Get-GuiRunCommand -Campaign Windows11Upgrade -WavePath 'C:\Waves\W 1' -RunMode Once -IgnoreRunGuard $true -SelectionPath 'C:\Sel\s.json'
    Test-Step 'Run command uses Windows PowerShell 5.1, the engine, the GUI selection and the run mode' { $run -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $run -match 'SmartM365-Invoke-EndpointMigrationWithPsExec\.ps1' -and $run -match '-GuiSelectionPath' -and $run -match '-RunMode Once' -and $run -match '-IgnoreRunGuard' -and $run -match '"C:\\Waves\\W 1"' }
    Test-Step 'Single PC command passes -ComputerName and no Wave' { $fictionalComputer = 'PC-1'; $c = Get-GuiRunCommand -Campaign HybridJoin -ComputerName $fictionalComputer -RunMode Once -IgnoreRunGuard $false -SelectionPath 'C:\s.json'; $c -match '-ComputerName PC-1' -and $c -notmatch '-WavePath' -and $c -notmatch '-IgnoreRunGuard' -and $c -match '-GuiSelectionPath' }
    $refreshHandler = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Start-InventoryRefresh' }, $true)
    Test-Step 'inventory refresh runs in a hidden process, never in a visible console' { $null -ne $refreshHandler -and $refreshHandler.Extent.Text -match '-WindowStyle Hidden' -and $guiText -notmatch '-NoExit' }
    $refreshAssignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$inventoryRefreshXaml' }, $true)
    $refreshXaml = [string]$refreshAssignment.Right.Expression.Value
    Test-Step 'inventory refresh window shows the output, stops and closes' { foreach ($name in 'InventoryRefreshStateText','InventoryRefreshADStateText','InventoryRefreshADOutputText','InventoryRefreshGraphStateText','InventoryRefreshGraphOutputText','InventoryRefreshStopButton','InventoryRefreshCloseButton','InventoryRefreshOpenFolderButton') { if ($refreshXaml -notmatch ('x:Name="{0}"' -f $name)) { return $false } }; $true }
    # Fictional export scripts: AD writes output, Intune fails; both must run and reach the log.
    $fakeRoot = Join-Path $tempRoot 'InventoryFake'
    New-Item -ItemType Directory -Path (Join-Path $fakeRoot 'Scripts') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $fakeRoot 'Scripts\SmartM365-EndpointMigration-Export-ADDevicesCsv.ps1') -Value "param([switch]`$ForceRefresh) Write-Host ('FAKE-AD force=' + `$ForceRefresh); Write-Host ('FAKE-AD-LONG ' + ('x' * 200)); Write-Output 'FAKE-AD output'" -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $fakeRoot 'Scripts\SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1') -Value "param([switch]`$ForceRefresh) `$ErrorActionPreference = 'Stop'; throw 'FAKE-INTUNE failure'" -Encoding ASCII
    $savedRoot = $script:PackageRoot; $script:PackageRoot = $fakeRoot
    $refreshLog = Join-Path $fakeRoot "Run's log\InventoryRefresh.log"
    New-Item -ItemType Directory -Path (Split-Path $refreshLog -Parent) -Force | Out-Null
    $refreshCommand = Get-InventoryRefreshCommandText -Exports @([pscustomobject]@{ Label = 'AD'; Script = 'SmartM365-EndpointMigration-Export-ADDevicesCsv.ps1'; Arguments = @('-ForceRefresh') },[pscustomobject]@{ Label = 'Intune'; Script = 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1'; Arguments = @('-ForceRefresh') }) -LogPath $refreshLog
    $script:PackageRoot = $savedRoot
    $refreshProcess = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($refreshCommand))) -WindowStyle Hidden -PassThru -Wait
    $refreshText = if (Test-Path -LiteralPath $refreshLog) { Get-Content -LiteralPath $refreshLog -Raw } else { '' }
    Test-Step 'inventory refresh passes -ForceRefresh and logs Write-Host and output' { $refreshText -match 'FAKE-AD force=True' -and $refreshText -match 'FAKE-AD output' }
    Test-Step 'inventory refresh runs every export after a failure and logs it with a timestamp' { $refreshText -match '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} Intune export failed: FAKE-INTUNE failure' -and $refreshText -match 'failed exports: 1\.' }
    Test-Step 'inventory refresh exit code is 1 when an export failed' { $refreshProcess.ExitCode -eq 1 }
    Test-Step 'inventory refresh log keeps long lines whole (no wrap at the hidden console width)' { $refreshText -match ('FAKE-AD-LONG ' + ('x' * 200)) }
    # Parallel plan: AD and Graph streams; Graph uses one sign-in (combined exporter for HybridJoin).
    $hjPlan = @(Get-InventoryRefreshPlan -Campaign HybridJoin)
    $w11Plan = @(Get-InventoryRefreshPlan -Campaign Windows11Upgrade)
    Test-Step 'inventory refresh runs two parallel streams: AD and Microsoft Graph' { ($hjPlan.Name -join ',') -eq 'AD,Graph' -and ($w11Plan.Name -join ',') -eq 'AD,Graph' -and @($hjPlan + $w11Plan | Where-Object { @($_.Exports).Count -ne 1 }).Count -eq 0 }
    Test-Step 'HybridJoin Graph stream exports Intune and Entra on one connection; Windows 11 exports Intune only' { $hjPlan[1].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-AutomaticGraphInventories.ps1' -and $w11Plan[1].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1' }
    # "Refresh old inventories only": only the old inventories are exported.
    $adOnly = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds AD)
    $entraOnly = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Entra)
    $intuneOnly = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Intune)
    $graphBoth = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Intune,Entra)
    $w11Intune = @(Get-InventoryRefreshPlan -Campaign Windows11Upgrade -Kinds Intune,Entra)
    Test-Step 'old inventories only: AD alone runs no Graph sign-in; one Graph inventory runs its own exporter' { ($adOnly.Name -join ',') -eq 'AD' -and ($entraOnly.Name -join ',') -eq 'Graph' -and $entraOnly[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-EntraDevicesCsv.ps1' -and $intuneOnly[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1' }
    Test-Step 'old inventories only: Intune and Entra share one sign-in; Windows 11 ignores Entra' { $graphBoth.Count -eq 1 -and $graphBoth[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-AutomaticGraphInventories.ps1' -and $w11Intune.Count -eq 1 -and $w11Intune[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1' -and @(Get-InventoryRefreshPlan -Campaign Windows11Upgrade -Kinds Entra).Count -eq 0 }
    Test-Step 'Automatic Wave view offers "Refresh old inventories only"' { $xaml -match 'x:Name="AutomaticRefreshStaleInventoryButton" Content="Refresh old inventories only"' -and $guiText -match 'Start-InventoryRefresh -Kinds @\(\$old' }
    foreach ($functionName in @('Get-AutomaticInventoryWarning')) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    $freshChecks = @([pscustomobject]@{ Kind = 'AD'; MaxAgeHours = 12; Usable = $true; Reason = 'Fresh'; AgeHours = 3 }, [pscustomobject]@{ Kind = 'Intune'; MaxAgeHours = 2; Usable = $true; Reason = 'FreshVerifiedProvenance'; AgeHours = 1 })
    $oldChecks = @([pscustomobject]@{ Kind = 'AD'; MaxAgeHours = 12; Usable = $false; Reason = 'Stale:16.4h>12h'; AgeHours = 16.4 }, [pscustomobject]@{ Kind = 'Entra'; MaxAgeHours = 2; Usable = $false; Reason = 'Missing'; AgeHours = $null }, [pscustomobject]@{ Kind = 'Intune'; MaxAgeHours = 2; Usable = $true; Reason = 'Fresh'; AgeHours = 1 })
    $oldWarning = Get-AutomaticInventoryWarning -Checks $oldChecks
    Test-Step 'no inventory warning when the inventories are fresh' { (Get-AutomaticInventoryWarning -Checks $freshChecks) -eq '' }
    Test-Step 'inventory warning lists only the old or missing inventories, with age and limit' { $oldWarning -match 'AD: 16[.,]4 hours old \(limit 12 hours\)' -and $oldWarning -match 'Entra: no shared inventory' -and $oldWarning -notmatch 'Intune' -and $oldWarning -match 'Refresh old inventories only' }
    Test-Step 'Preview and Create warn about old inventories before running (Windows 11 follows the Source)' { $guiText -match 'Get-AutomaticInventoryChecks -Campaign \$campaign -Source \$filters\.Source' -and $guiText -match "Run the Preview anyway\?" -and $guiText -match "'Create Wave - old inventories'" }
    $planProblems = foreach ($export in @($hjPlan + $w11Plan + $entraOnly + $intuneOnly + $adOnly | ForEach-Object { $_.Exports })) {
        $exportPath = Join-Path $packageRoot ('Scripts\' + $export.Script)
        if (-not (Test-Path -LiteralPath $exportPath -PathType Leaf)) { 'missing ' + $export.Script; continue }
        $exportParameters = @([System.Management.Automation.Language.Parser]::ParseFile($exportPath, [ref]$null, [ref]$null).ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        foreach ($argument in @($export.Arguments)) { if ($exportParameters -notcontains $argument.TrimStart('-')) { '{0} has no {1}' -f $export.Script,$argument } }
    }
    Test-Step 'every planned export exists and accepts its arguments' { @($planProblems).Count -eq 0 }
    $refreshEnd = [string]($ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Update-InventoryRefreshWindow' }, $true).Extent.Text)
    Test-Step 'inventory refresh ends with a message, then closes the window (errors: the operator may keep it open)' { $refreshEnd -match "Show-InventoryRefreshMessage[^\r\n]*-Buttons 'OK' -Icon 'Information'" -and $refreshEnd -match "-Buttons 'YesNo' -Icon 'Warning'" -and ([regex]::Matches($refreshEnd, '\$state\.Window\.Close\(\)')).Count -eq 2 }
    Test-Step 'every stream has its state and output areas in the window (case-sensitive names for FindName)' { foreach ($item in $hjPlan) { if ($refreshXaml -cnotmatch ('x:Name="InventoryRefresh{0}(Output|State)Text"' -f $item.Name)) { return $false } }; $true }
    $domainCsv = Join-Path $tempRoot 'DomainInventory.csv'
    [IO.File]::WriteAllText($domainCsv, ("ComputerName,ADDomain,DistinguishedName`r`nPC1,FR.example,""CN=PC1,DC=fr,DC=example""`r`nPC2,fr.example.,""CN=PC2,DC=fr""`r`nPC3,be.example,""x, y""`r`nPC4,,`r`n"), [Text.UTF8Encoding]::new($true))
    Test-Step 'AD domain list is read from the AD inventory (distinct, lower case, sorted, quoted commas)' { (@(Get-AdInventoryDomains -Path $domainCsv) -join ',') -eq 'be.example,fr.example' }
    Set-Content -LiteralPath $domainCsv -Value 'ComputerName,Enabled' -Encoding ASCII
    Test-Step 'AD domain list is empty without the ADDomain column or the file' { @(Get-AdInventoryDomains -Path $domainCsv).Count -eq 0 -and @(Get-AdInventoryDomains -Path (Join-Path $tempRoot 'missing.csv')).Count -eq 0 }
    Test-Step 'Automatic Wave view offers the AD domain choice' { $xaml -match 'x:Name="AutomaticAdDomainCombo"' -and $guiText -match "AllAdDomainsLabel = '\(All domains\)'" }
    # Preview/Create: the real engine in a hidden process, as started by the GUI (fictional inventories).
    $waitWindow = [string]$ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$automaticWaitXaml' }, $true).Right.Expression.Value
    Test-Step 'Preview and Create show a waiting window with progress and the engine stage' { $waitWindow -match 'IsIndeterminate="True"' -and $waitWindow -match 'x:Name="AutomaticWaitStageText"' -and $guiText -match 'Show-AutomaticWaveWait -Create:\$Create' }
    Test-Step 'a successful Preview keeps its evidence; Create passes it to the engine' { $guiText -match "AutomaticPreviewEvidence = \[string\]\(Get-SemtPropertyValue -InputObject \`$summary -Name 'EvidencePath'\)" -and $guiText -match '-PreviewEvidencePath \$\(if \(\$Create\) \{ \$script:AutomaticPreviewEvidence \}' }
    Test-Step 'Preview and Create end with a message (counts; Create states the Wave is not launched)' { $guiText -match "'Preview', 'OK', 'Information'" -and $guiText -match "The Wave was NOT launched\." -and $guiText -match "'Create Wave', 'OK', 'Information'" }
    $powershell51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $runPreview = {
        param([string]$Name)
        $evidenceRoot = Join-Path $tempRoot ("AutoEvidence-{0}" -f $Name)
        New-Item -ItemType Directory -Path $evidenceRoot -Force | Out-Null
        $logPath = Join-Path $evidenceRoot 'GuiAutomaticWave.log'
        $command = Get-AutomaticWaveCommandText -Campaign HybridJoin -Filters @{ WaveName = 'WAVE-AUTO-HJ-GUI'; Prefixes = @(); Contains = @(); ExcludeStaleAd = $false; LastLogonDays = 45; Source = 'Both'; ExcludeIntunePresent = $false; AdDomain = '' } -EvidenceRoot $evidenceRoot
        $process = Start-Process -FilePath $powershell51 -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes((Get-AutomaticWaveProcessText -CommandText $command -LogPath $logPath)))) -WindowStyle Hidden -PassThru -Wait
        [pscustomobject]@{ ExitCode = $process.ExitCode; Log = $(if (Test-Path -LiteralPath $logPath) { Get-Content -LiteralPath $logPath -Raw } else { '' }); Summary = @(Get-ChildItem -LiteralPath $evidenceRoot -Filter 'AutomaticWaveSummary.json' -File -Recurse) }
    }
    $missing = & $runPreview 'Missing'
    Test-Step 'Preview without inventory fails with exit code 1 and the error in the log' { $missing.ExitCode -eq 1 -and $missing.Log -match '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ERROR Required automatic Wave inventory not found' -and $missing.Summary.Count -eq 0 }
    $inventoryRoot = Join-Path $packageRoot 'Inventory'
    New-Item -ItemType Directory -Path $inventoryRoot -Force | Out-Null
    [pscustomobject]@{ ComputerName = 'GUI-PC01'; Name = 'GUI-PC01'; ADInventoryPresent = $true; ADDomain = 'contoso.test'; Enabled = $true; DNSHostName = 'gui-pc01.contoso.test'; DistinguishedName = 'CN=GUI-PC01,DC=contoso,DC=test'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).ToUniversalTime().ToString('o') } | Export-Csv -LiteralPath (Join-Path $inventoryRoot 'DevicesAD.csv') -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{ ComputerName = 'OTHER'; DeviceName = 'OTHER'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = 'x'; LastSyncDateTime = (Get-Date).ToUniversalTime().ToString('o') } | Export-Csv -LiteralPath (Join-Path $inventoryRoot 'DevicesIntune.csv') -NoTypeInformation -Encoding UTF8
    $ok = & $runPreview 'Ok'
    Test-Step 'Preview succeeds in the hidden process and writes its summary' { $ok.ExitCode -eq 0 -and $ok.Summary.Count -eq 1 }
    Test-Step 'engine stages reach the log for the waiting window' { (Get-AutomaticWaveLastStage -Text $ok.Log) -match '^Automatic Wave preview ready\.' -and $ok.Log -match ' STAGE Loading inventory records\.\.\.' }
    Test-Step 'end message counts read the HybridJoin and Windows 11 summaries' { (Get-AutomaticWaveResultText -Summary ([pscustomobject]@{ SelectedDevices = 3; ExcludedDevices = 2; FilterExcludedDevices = 1 })) -match 'Selected: 3' -and (Get-AutomaticWaveResultText -Summary ([pscustomobject]@{ SelectedDevices = 4; ExcludedDevices = 0; NameFilterExcludedDevices = 7 })) -match 'Filtered out: 7' }
    Remove-Item -LiteralPath $inventoryRoot -Recurse -Force
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
    $popupScript = Join-Path $tempRoot 'popup-check.ps1'
    @'
param($Gui)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Gui, [ref]$null, [ref]$null)
$assignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$inventoryRefreshXaml' }, $true)
$popup = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new([string]$assignment.Right.Expression.Value)))
'POPUP=' + (@('AD','Graph' | Where-Object { $box = $popup.FindName(('InventoryRefresh{0}OutputText' -f $_)); $null -ne $box -and $box.IsReadOnly }).Count -eq 2)
'@ | Set-Content -LiteralPath $popupScript -Encoding ASCII
    $popupCheck = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $popupScript -Gui $guiPath 2>&1 | ForEach-Object { [string]$_ })
    Test-Step 'inventory refresh window loads under Windows PowerShell 5.1 STA' { ($popupCheck -join "`n") -match 'POPUP=True' }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCQLwT1E7jz2Bp+
# +TpoOoEpF8t4Mu2GFZr0ijHtaszlF6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB7dqjqYAGIdxu2ZKujmKtez6izTq/W7Gh815e9CG2f4MA0GCSqG
# SIb3DQEBAQUABIIBgEcZT7QOuiyHfvZb/RFwlk19+V025ALIYm/Uv0Td+RvQ9f2s
# F5Bbv7dAdPET2UUMTKn4OyxU6JXsVfLikBs1E1+I6fI5rcNbkdiHPKkc5Ccd2AqP
# GC+aCy/azBErOW5oA2YnRF+co/A40rlNIqK86ZCM9BoN804KlbeSuKpEZU7K2cOw
# w9yJQqs9N+ptV44F+o2OmH4nYRHXfN73UrtxNFxWbPUuwpSIYe1gajO5g7x+nN+j
# bdBqnVSqY0wHt7InazhaF3ZNjI51WmQBNZEjYuQ4KUp0ZcIrSvS+WSO5YyaU8wms
# nDjez5ejikBlnTYjG5j3NnoQ49IyTtuUBEIKujLxJwkbxBNKGr+7OAiAWhXC0Eep
# YM6TnLbMg+o+Thzr/FncBh5NCchKghmwypu7szr19RGZ+JMx/rZ19zL2ry5o83or
# obDDw9ptlW5uiTnX1Pm/AS4mVY8np9RtbBVyKEz4X9CUa22yat/letYpcqjGhaVi
# ox4zDsfvPyZ3uJYdI6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkwOTI2
# MzBaMC8GCSqGSIb3DQEJBDEiBCDV6ZW4lAPSOuARsA6oHk//wWEo7L+T9Sgyt4fj
# fgdJ2zANBgkqhkiG9w0BAQEFAASCAgBmq79DJznDQvKr53NkwMYqc8Q83pL8c88p
# OwyvbPX1e2JIAygxlY1VlZsKs3SYQIPA15i9KVtxApzIaJauWOLls78P0PboGBrZ
# t2W6/ZLiHbp7nAsi0FX30Sa3R5JJsadxhRk1n0AqYXUVYjziLWWisiXtdhMha5+J
# F2et0TuUhOz+iCRlSQ1lK4zJ3NAel/0GvaOC0HNsVxCi3rv5DZIazAlE7FoBlSFm
# NyThqqVAys1O7n7dz1NgdqFzKlydTPoKTfwAK9a7ZLxFYDVzQr4z1apoEZMa6CCc
# v4hwID4kknTWD2u723qfiJ5JbOzM1HMdFXxarRayKa7mfBPjFcrofX6AuZ/ies06
# 2CwY/sbKut7L2a7EzqhhrLq0LFBUM1Lg4aMOA5BOB8zv9VqdWPu8FPMEHl37Tg6I
# sN00ukOmh6hx6rA/oIVcmN01mSjMqqEEg+ZHbT6e48gpSl+HdhR7h6bwCDnSxPTU
# o00IAMn6S0S9Gx06HzJZ0cN8q/subFoklobtyJbAErGOjVyCbWicliyuMi68tcOw
# HuEj4niqO8QtA8eNBn35xe+5MxHHzfc1hPcKqfv6C5qDi0+3R3TxTM1AIrBSFwrp
# cS9UBq4oIMGCLAmFEMMHYiMylLrCBVtt37RVHiYaN1WcSa9Tmb5pJ6+IsHgv5qTk
# lZhCP6D58A==
# SIG # End signature block
