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
    1.1.0
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
    foreach ($item in @('Runs','Waves\HybridJoin','Waves\Windows11Upgrade','Waves\Prerequisites')) { New-Item -ItemType Directory -Path (Join-Path $packageRoot $item) -Force | Out-Null }
    Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter '*.local.json.txt' -File | Remove-Item -Force
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
    # Dedicated windows of the buttons Single PC, New Wave and Automatic Wave (literal here-strings too).
    $childXaml = [ordered]@{}
    foreach ($variable in @('singlePcXaml','newWaveXaml','automaticWaveXaml')) {
        $childAssignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$' + $variable) }, $true)
        Test-Step ("{0} is a literal here-string with the shared resources placeholder" -f $variable) { $null -ne $childAssignment -and $childAssignment.Right.Expression -is [System.Management.Automation.Language.StringConstantExpressionAst] -and ([string]$childAssignment.Right.Expression.Value) -match '<!--SEMT_SHARED_RESOURCES-->' }
        $childXaml[$variable] = [string]$childAssignment.Right.Expression.Value
    }
    $allXaml = $xaml + ($childXaml.Values -join [Environment]::NewLine)
    foreach ($name in @('CampaignCombo','WaveCombo','WaveModeCombo','LaunchWaveButton','LaunchAllButton','StopAllButton','OpenSinglePcButton','OpenNewWaveButton','OpenAutomaticWaveButton','DeleteWaveButton','RunsGrid','OpenRunReportButton','StopSelectedRunButton','ForceStopSelectedRunButton','OptionsHost','ResetOptionsButton','HeaderLogoImage')) {
        Write-TestResult -Condition ($xaml -match ('x:Name="{0}"' -f $name)) -Message ("main window control present: {0}" -f $name)
    }
    foreach ($pair in @(@('singlePcXaml',@('SingleComputerText','SingleModeCombo','LaunchSingleButton','SingleRunFolderText','OpenSingleRunFolderButton')),@('newWaveXaml',@('NewWaveNameText','NewWaveComputersText','CreateWaveButton','NewWavePathText','OpenNewWaveComputersButton')),@('automaticWaveXaml',@('AutomaticWaveNameText','AutomaticNamePrefixText','AutomaticNameContainsText','AutomaticNameNotContainsText','AutomaticAdDomainCombo','AutomaticSplitByDomainCheck','AutomaticSourceCombo','AutomaticExcludeIntuneCheck','AutomaticExcludeStaleAdCheck','AutomaticLastLogonDaysText','AutomaticPreviewButton','AutomaticCreateButton','AutomaticRefreshInventoryButton','AutomaticRefreshStaleInventoryButton','AutomaticSummaryText','AutomaticEvidencePathText','AutomaticOpenEvidenceButton')))) {
        foreach ($name in $pair[1]) { Write-TestResult -Condition ($childXaml[$pair[0]] -match ('x:Name="{0}"' -f $name) -and $xaml -notmatch ('x:Name="{0}"' -f $name)) -Message ("{0} control present only in its window: {1}" -f $pair[0],$name) }
    }
    Test-Step 'all the action card buttons are on one line under the Campaign selector and the status' {
        $bar = [regex]::Match($xaml, '(?s)<Grid Grid\.Row="1" Grid\.ColumnSpan="3" x:Name="ActionButtonBar".*?\n                </Grid>').Value
        $xaml -match '(?s)x:Name="CampaignCombo".*?x:Name="StatusText".*?x:Name="ActionButtonBar".*?x:Name="MainTabs"' -and $bar -match '(?s)x:Name="OpenSinglePcButton" Content="Single PC".*?x:Name="OpenNewWaveButton" Content="New Wave".*?x:Name="OpenAutomaticWaveButton" Content="Automatic Wave".*?x:Name="RefreshButton".*?x:Name="StopAllButton".*?x:Name="LaunchAllButton".*?x:Name="CloseGuiButton" Content="Close"'
    }
    Test-Step 'the tabs Single PC, New Wave and Automatic Wave are removed (Waves, Runs, Options remain)' { $mainTabs = @([regex]::Matches($xaml, '(?m)^            <TabItem x:Name="(?<Name>[A-Za-z]+)" Header="(?<Header>[^"]+)"') | ForEach-Object { $_.Groups['Header'].Value }); ($mainTabs -join ',') -eq 'Waves,Runs,Options' -and $xaml -notmatch 'SinglePcTab|NewWaveTab|AutomaticWaveTab' }
    Test-Step 'no control name is used twice across the GUI windows' { $names = @([regex]::Matches($allXaml, 'x:Name="(?<Name>[A-Za-z][A-Za-z0-9_]*)"') | ForEach-Object { $_.Groups['Name'].Value } | Where-Object { $_ -ne 'TabBorder' }); @($names | Group-Object | Where-Object { $_.Count -gt 1 }).Count -eq 0 }
    Test-Step 'main window uses the full height of the screen work area and fits on small screens' { $sizeText = [IO.File]::ReadAllText($guiPath); $sizeText -match '\$window\.Height = \$workArea\.Height' -and $sizeText -match '\$window\.MinHeight = \[Math\]::Min\(\$window\.MinHeight, \$workArea\.Height\)' -and $sizeText -match '\$window\.Top = \$workArea\.Top' }
    Test-Step 'no Batch/LOT wording in the interface' { $allXaml -notmatch '\b(LOT|LOTs|Lot|Batch|batch)\b' }
    Test-Step 'Preview and Create are separate controls; Create states it never launches' { $childXaml['automaticWaveXaml'] -match 'AutomaticPreviewButton' -and $childXaml['automaticWaveXaml'] -match 'never launches the Wave' }
    Test-Step 'Launch All is labelled as limited to the selected Campaign in its confirmation' { [IO.File]::ReadAllText($guiPath) -match 'Launch All Waves of this Campaign only' }
    Test-Step 'stop states remote endpoints are reported separately' { [IO.File]::ReadAllText($guiPath) -match 'Remote endpoints may continue and must be verified' }
    $guiText = [IO.File]::ReadAllText($guiPath)
    foreach ($functionName in @('ConvertTo-PowerShellLiteral','Get-AutomaticWaveCommandText','Get-AutomaticFilterSignature','Split-GuiList','Get-GuiRunCommand','Assert-CmdLiteralValue','ConvertTo-CmdArgument','Get-GuiStatePath','New-GuiSelectionFile','Get-InventoryRefreshCommandText','Get-AdInventoryDomains','Get-InventoryRefreshPlan','Get-AutomaticWaveProcessText','Get-AutomaticWaveLastStage','Get-AutomaticWaveResultText','Get-GuiChildWindowMarkup','Test-AutomaticPreviewReusable','Test-AutomaticSplitByDomainAvailable','Remove-GuiWaveFolder')) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    $script:PackageRoot = $packageRoot
    Test-Step 'dedicated windows receive the styles and brushes of the main window' {
        $merged = Get-GuiChildWindowMarkup -Markup $childXaml['automaticWaveXaml'] -MainMarkup $xaml
        $merged -notmatch 'SEMT_SHARED_RESOURCES' -and $merged -match 'x:Key="PrimaryButtonStyle"' -and $merged -match 'x:Key="StepperButtonStyle"' -and $merged -match 'x:Key="BorderBrush"'
    }
    Test-Step 'dedicated windows are modal, hidden when closed, and the Campaign is in their title' { $guiText -match '\[void\]\$child\.ShowDialog\(\)' -and $guiText -match '\$cancelEventArgs\.Cancel = \$true; \$eventSource\.Hide\(\)' -and $guiText -match '\$child\.Title = ''\{0\} - \{1\}'' -f \$script:GuiChildTitles\[\$Name\],\$script:CampaignTitles\[\$script:CurrentCampaign\]' }
    Test-Step 'each button opens its own window' { $guiText -match 'OpenSinglePcButton\.Add_Click\(\{ try \{ Show-GuiChildWindow -Name SinglePc \}' -and $guiText -match 'OpenNewWaveButton\.Add_Click\(\{ try \{ Show-GuiChildWindow -Name NewWave \}' -and $guiText -match 'OpenAutomaticWaveButton\.Add_Click\(\{ try \{ Show-GuiChildWindow -Name AutomaticWave \}' }
    Test-Step 'New Wave ends with a message, then clears its fields and closes its window' {
        $handler = [regex]::Match($guiText, '(?s)\$controls\.CreateWaveButton\.Add_Click\(\{.*?\n\}\)').Value
        $handler -match "(?s)New-SemtWave .*?\[System\.Windows\.MessageBox\]::Show\(\(Get-GuiDialogOwner\), \(""Wave '\{0\}' was created for .*?The Wave was NOT launched\..*?'Create Wave', 'OK', 'Information'\).*?NewWaveNameText\.Text = ''.*?NewWaveComputersText\.Text = ''.*?\`$script:GuiChildWindows\['NewWave'\]\.Close\(\)"
    }
    Test-Step 'Single PC ends with a message after a confirmed launch, then closes its window (not after a cancel)' {
        $handler = [regex]::Match($guiText, '(?s)\$controls\.LaunchSingleButton\.Add_Click\(\{.*?\n\}\)').Value
        $handler -match '(?s)if \(-not \(Start-GuiRun [^\n]*\)\) \{ return \}.*?Update-RunsView.*?\[System\.Windows\.MessageBox\]::Show\(\(Get-GuiDialogOwner\), \("Single PC Run launched for .*?''Launch Single PC'', ''OK'', ''Information''\).*?SingleComputerText\.Text = ''''.*?\$script:GuiChildWindows\[''SinglePc''\]\.Close\(\)'
    }
    Test-Step 'Automatic Wave closes its window after the end message of a successful Create (not after Preview)' { $guiText -match "(?s)'Create Wave', 'OK', 'Information'\)\r?\n\s*# The Automatic Wave window closes after a successful Create[^\n]*\n\s*\`$script:GuiChildWindows\['AutomaticWave'\]\.Close\(\)\r?\n\s*\}\r?\n\s*else \{" -and @([regex]::Matches($guiText, "GuiChildWindows\['AutomaticWave'\]\.Close\(\)")).Count -eq 1 }
    Test-Step 'Open Computers.txt before any New Wave explains instead of "File not found"' { $guiText -match "No Wave was created from this window yet\." }
    Test-Step 'each dedicated window has a Close button (Esc too) at the bottom right' {
        $ok = $true
        foreach ($pair in @(@('singlePcXaml','SinglePcCloseButton'),@('newWaveXaml','NewWaveCloseButton'),@('automaticWaveXaml','AutomaticWaveCloseButton'))) { if ($childXaml[$pair[0]] -notmatch ('<Button Grid\.Row="1" x:Name="{0}" Content="Close" IsCancel="True"[^>]*HorizontalAlignment="Right"' -f $pair[1])) { $ok = $false } }
        $ok -and $guiText -match "\`$closeButton = \`$controls\[\`$entry\[0\] \+ 'CloseButton'\]" -and $guiText -match '\$closeButton\.Add_Click\(\{ param\(\$eventSource,\$routedEventArgs\) \$eventSource\.Tag\.Close\(\) \}\)'
    }
    Test-Step 'Prerequisites hides the Automatic Wave button (no inventory, no Automatic Wave)' { $guiText -match '(?s)if \(\$script:CurrentCampaign -eq ''Prerequisites''\) \{\s*\$controls\.OpenAutomaticWaveButton\.Visibility = ''Collapsed''' -and $guiText -match '\$controls\.OpenAutomaticWaveButton\.Visibility = ''Visible''' }
    Test-Step 'Create Wave is never greyed out waiting for a Preview (only while the engine runs)' {
        $disabled = @([regex]::Matches($guiText, 'AutomaticCreateButton\.IsEnabled = \$false')).Count
        $disabled -eq 1 -and $guiText -match '\$controls\.AutomaticPreviewButton\.IsEnabled = \$false\r?\n\s*\$controls\.AutomaticCreateButton\.IsEnabled = \$false' -and $guiText -notmatch 'Run Preview again before Create' -and $childXaml['automaticWaveXaml'] -notmatch 'AutomaticCreateButton"[^>]*IsEnabled="False"'
    }
    Test-Step 'a Preview is reused only with the same filters and its evidence' {
        (Test-AutomaticPreviewReusable -Signature 'A' -PreviewSignature 'A' -PreviewEvidence 'C:\E') -and -not (Test-AutomaticPreviewReusable -Signature 'A' -PreviewSignature 'B' -PreviewEvidence 'C:\E') -and -not (Test-AutomaticPreviewReusable -Signature 'A' -PreviewSignature '' -PreviewEvidence '') -and -not (Test-AutomaticPreviewReusable -Signature 'A' -PreviewSignature 'A' -PreviewEvidence '')
    }
    Test-Step 'Create without a matching Preview states that the selection is computed now with the same safety exclusions' { $guiText -match 'No Preview matches these filters: the selection is computed now with the same rules and safety exclusions' -and $guiText -match 'Show-AutomaticWaveWait -Create:\$Create -ReusePreview:\$reusePreview' }
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
    # Active Run detection (used when closing): only <Runs>\<Campaign>\<Wave>\<RunId>\State is listed.
    $activeRunsRoot = Join-Path $tempRoot 'ActiveRunsDetection'
    $selfStart = (Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')
    $activeRunFiles = [ordered]@{
        'HybridJoin\WAVE-A\RUN-1\State' = @{ Campaign = 'HybridJoin'; WaveName = 'WAVE-A'; RunId = 'RUN-1'; Status = 'Running'; ProcessId = $PID; ProcessStartUtc = $selfStart }
        'Windows11Upgrade\_SinglePC\RUN-2\State' = @{ Campaign = 'Windows11Upgrade'; WaveName = '_SinglePC'; RunId = 'RUN-2'; Status = 'Running'; ProcessId = $PID; ProcessStartUtc = $selfStart }
        'HybridJoin\WAVE-A\RUN-0\State' = @{ Campaign = 'HybridJoin'; WaveName = 'WAVE-A'; RunId = 'RUN-0'; Status = 'Running'; ProcessId = $PID; ProcessStartUtc = '2001-01-01T00:00:00.0000000Z' }
        'HybridJoin\WAVE-A\RUN-1\Logs\Deep\Copy\State' = @{ Campaign = 'HybridJoin'; WaveName = 'WAVE-A'; RunId = 'COPY'; Status = 'Running'; ProcessId = $PID; ProcessStartUtc = $selfStart }
    }
    foreach ($relative in $activeRunFiles.Keys) {
        $folder = Join-Path $activeRunsRoot $relative
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $record = $activeRunFiles[$relative]
        foreach ($field in 'SignalPath','ReportPath','UpdatedUtc') { $record[$field] = '' }
        [IO.File]::WriteAllText((Join-Path $folder ('ActiveRun_{0}.json.txt' -f $PID)), ([pscustomobject]$record | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    }
    Test-Step 'active Run detection reads only the Run State folders (Wave and Single PC), flags stale PIDs, ignores copies deeper in a Run' {
        $found = @(Get-SemtActiveRuns -RunsRoot $activeRunsRoot)
        $alive = @($found | Where-Object { $_.Alive } | ForEach-Object { $_.RunId } | Sort-Object)
        $found.Count -eq 3 -and ($alive -join ',') -eq 'RUN-1,RUN-2' -and @($found | Where-Object { $_.RunId -eq 'RUN-0' -and $_.Status -eq 'Stale' }).Count -eq 1
    }
    # Every layout written by a Run (Wave, Single PC, names with [ ]), 2 Campaigns, several Waves, past
    # Runs, Archives and the Automatic Wave evidence folder: same files as the former -Recurse scan.
    $equivalenceRoot = Join-Path $tempRoot 'ActiveRunsEquivalence'
    $layout = @(
        @('HybridJoin','WAVE-HJ-DE','20261113-090000-A',$true), @('HybridJoin','WAVE-HJ-DE','20261112-090000-B',$false), @('HybridJoin','WAVE-[FR] 1','20261113-091000-C',$true),
        @('HybridJoin','_SinglePC','20261113-092000-D',$true), @('Windows11Upgrade','WAVE-W11-PILOT','20261113-093000-E',$true), @('Windows11Upgrade','WAVE-W11-PILOT','20261110-093000-F',$false),
        @('Windows11Upgrade','_SinglePC','20261113-094000-G',$true), @('Windows11Upgrade','WAVE-W11-EMPTY','20261113-095000-H',$false)
    )
    foreach ($item in $layout) {
        $runFolder = [IO.Path]::Combine($equivalenceRoot, $item[0], $item[1], $item[2])
        foreach ($sub in 'State','Logs\PC-0001','Reports') { [void][IO.Directory]::CreateDirectory([IO.Path]::Combine($runFolder, $sub)) }
        [IO.File]::WriteAllText([IO.Path]::Combine($runFolder, 'Logs\PC-0001', 'worker.log'), 'x')
        if ($item[3]) { [IO.File]::WriteAllText([IO.Path]::Combine($runFolder, 'State', 'ActiveRun_4242.json.txt'), (@{ Campaign = $item[0]; WaveName = $item[1]; RunId = $item[2]; Status = 'Running'; ProcessId = 0; SignalPath = ''; ReportPath = ''; UpdatedUtc = '' } | ConvertTo-Json), [Text.UTF8Encoding]::new($false)) }
    }
    [void][IO.Directory]::CreateDirectory([IO.Path]::Combine($equivalenceRoot, 'HybridJoin', 'WAVE-HJ-DE', 'Archives'))
    [IO.File]::WriteAllText([IO.Path]::Combine($equivalenceRoot, 'HybridJoin', 'WAVE-HJ-DE', 'Archives', '20261001-090000.zip'), 'x')
    [void][IO.Directory]::CreateDirectory([IO.Path]::Combine($equivalenceRoot, '_AutomaticWave', '20261113-080000', 'Evidence'))
    $recursePaths = @(Get-ChildItem -LiteralPath $equivalenceRoot -Recurse -Filter 'ActiveRun_*.json.txt' -File | ForEach-Object { $_.FullName } | Sort-Object)
    Test-Step 'active Run detection finds exactly the files of the former -Recurse scan (2 Campaigns, Waves, Single PC, [ ] names)' {
        $fastPaths = @(Get-SemtActiveRuns -RunsRoot $equivalenceRoot | ForEach-Object { $_.StatePath } | Sort-Object)
        $recursePaths.Count -eq 5 -and ($fastPaths -join '|') -eq ($recursePaths -join '|')
    }
    Test-Step 'active Run detection with -Campaign matches the former scan filtered on that Campaign' {
        $ok = $true
        foreach ($campaignName in 'HybridJoin','Windows11Upgrade') {
            $fast = @(Get-SemtActiveRuns -RunsRoot $equivalenceRoot -Campaign $campaignName | ForEach-Object { $_.StatePath } | Sort-Object)
            $expected = @($recursePaths | Where-Object { $_ -like ('{0}\{1}\*' -f $equivalenceRoot,$campaignName) })
            if ($fast.Count -eq 0 -or ($fast -join '|') -ne ($expected -join '|')) { $ok = $false }
        }
        $ok
    }
    # Campaign change: RunHistory read once per refresh; the guards still read the active Runs at the click.
    Test-Step 'Campaign change and Refresh read RunHistory once (Get-GuiRecentRuns), the Wave and Runs views reuse it' {
        ([regex]::Matches($guiText, 'Get-SemtRuns ')).Count -eq 1 -and $guiText -match '(?s)function Invoke-GuiViewRefresh \{.*?\[void\]\(Get-GuiRecentRuns\)\s*\$script:RefreshUsesRunCache = \$true\s*try \{\s*Update-WaveList -SelectName \$SelectName\s*Update-RunsView\s*Update-AutomaticView\s*\}\s*finally \{ \$script:RefreshUsesRunCache = \$false \}' -and $guiText -match '(?s)function Switch-Campaign \{.*?Invoke-GuiViewRefresh\s*Add-Status' -and $guiText -match 'RefreshButton\.Add_Click\(\{ try \{ [^\n]*Invoke-GuiViewRefresh -SelectName \$name'
    }
    Test-Step 'Launch, Delete, Launch All, Stop and Close guards read the active Runs at the click (never the refresh cache)' {
        $guards = [regex]::Matches($guiText, 'Get-SemtActiveRuns -RunsRoot \(Get-SemtStorageRoot -PackageRoot \$script:PackageRoot -Kind Runs\)').Count
        $guards -ge 7 -and $guiText -notmatch 'RecentRuns[^\n]*Get-SemtActiveRuns'
    }
    $listPath = Join-Path $tempRoot 'ComputerListEquivalence.txt'
    [IO.File]::WriteAllLines($listPath, [string[]]@('PC-1','  pc-1  ','"PC-2"',' "PC-3." ','PC-4.contoso.example.','# comment','   ','','"#PC-5"','PC-6...','""','PC-7'))
    Test-Step 'computer list keeps the ConvertTo-SemtComputerName normalization (quotes, dots, comments, case-insensitive duplicates)' {
        $expected = New-Object System.Collections.Generic.List[string]
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($line in [IO.File]::ReadAllLines($listPath)) { $name = ConvertTo-SemtComputerName -Line $line; if (-not [string]::IsNullOrWhiteSpace($name) -and $seen.Add($name)) { $expected.Add($name) } }
        ((Get-SemtComputerList -Path $listPath) -join '|') -eq ($expected -join '|') -and $expected.Count -eq 6
    }
    Test-Step 'active Run detection keeps the Campaign filter and accepts a missing Runs folder' { (@(Get-SemtActiveRuns -RunsRoot $activeRunsRoot -Campaign Windows11Upgrade | ForEach-Object { $_.RunId }) -join ',') -eq 'RUN-2' -and @(Get-SemtActiveRuns -RunsRoot (Join-Path $tempRoot 'NoRuns')).Count -eq 0 }
    # Wave options tab (replaces the "Wave options" button): Options first, then Activity; editors read and
    # write Wave.local.json.txt with inheritance. A fictional Wave in a temporary folder; nothing is launched.
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
        $waveFile = Join-Path $optionsWave 'Wave.local.json.txt'
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
    $filters = @{ WaveName = 'WAVE-AUTO-TEST'; Prefixes = @('CTSO-'); Contains = @("O'Brien"); NotContains = @('-A-'); ExcludeStaleAd = $true; LastLogonDays = 30; Source = 'Both'; ExcludeIntunePresent = $true; AdDomain = 'fr.example' }
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
        $noExclusion = $filters.Clone(); $noExclusion.NotContains = @()
        Write-TestResult -Condition ($create -match "-ComputerNameNotContains @\('-A-'\)" -and (Get-AutomaticWaveCommandText -Campaign $campaign -Filters $noExclusion -EvidenceRoot 'C:\Evidence') -notmatch 'ComputerNameNotContains') -Message ("{0}: does-not-contain filter is passed only when set" -f $campaign)
        $reuseCreate = Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence' -Create -PreviewEvidencePath 'C:\Evidence\Preview'
        $reuseUsed = @([regex]::Matches($reuseCreate, '(?<=\s)-(?<Name>[A-Za-z]+)') | ForEach-Object { $_.Groups['Name'].Value } | Where-Object { $engineParameters -notcontains $_ })
        Write-TestResult -Condition ($reuseCreate -match "-Create -PreviewEvidencePath 'C:\\Evidence\\Preview'$" -and $reuseUsed.Count -eq 0 -and (Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence' -PreviewEvidencePath 'C:\Evidence\Preview') -notmatch 'PreviewEvidencePath') -Message ("{0}: Create passes the previewed evidence to the engine (Preview never does)" -f $campaign)
        $allDomains = $filters.Clone(); $allDomains.AdDomain = ''
        Write-TestResult -Condition ($preview -match "-AdDomainFilter @\('fr\.example'\)" -and (Get-AutomaticWaveCommandText -Campaign $campaign -Filters $allDomains -EvidenceRoot 'C:\Evidence') -notmatch 'AdDomainFilter') -Message ("{0}: AD domain is passed only when a domain is chosen" -f $campaign)
        $splitDomains = $allDomains.Clone(); $splitDomains.SplitByAdDomain = $true
        Write-TestResult -Condition ((Get-AutomaticWaveCommandText -Campaign $campaign -Filters $splitDomains -EvidenceRoot 'C:\Evidence' -Create) -match ' -SplitByAdDomain( |$)' -and (Get-AutomaticWaveCommandText -Campaign $campaign -Filters $splitDomains -EvidenceRoot 'C:\Evidence') -match ' -SplitByAdDomain( |$)' -and $create -notmatch 'SplitByAdDomain') -Message ("{0}: one Wave per AD domain is passed only when chosen, for Preview and Create" -f $campaign)
    }
    Test-Step 'filter signature changes when a filter changes (Create needs a matching Preview)' {
        $a = Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $filters
        $changed = $filters.Clone(); $changed.LastLogonDays = 31
        $domain = $filters.Clone(); $domain.AdDomain = 'be.example'
        $exclusion = $filters.Clone(); $exclusion.NotContains = @('-B-')
        $a -ne (Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $changed) -and $a -ne (Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $domain) -and $a -ne (Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $exclusion)
    }
    # Delete Wave: the folder goes to the Recycle Bin; other Waves and Runs stay; unsafe names are refused.
    $deleteWave = New-SemtWave -PackageRoot $packageRoot -Campaign HybridJoin -Name 'WAVE-DELETE-TEST' -Computers @('PC-DEL-1') -Source Manual
    $keepWave = New-SemtWave -PackageRoot $packageRoot -Campaign HybridJoin -Name 'WAVE-KEEP-TEST' -Computers @('PC-KEEP-1') -Source Manual
    $deletedPath = Remove-GuiWaveFolder -PackageRoot $packageRoot -Campaign HybridJoin -Name 'WAVE-DELETE-TEST'
    Test-Step 'Delete Wave removes only the selected Wave folder' { $deletedPath -eq $deleteWave.Path -and -not (Test-Path -LiteralPath $deleteWave.Path) -and (Test-Path -LiteralPath $keepWave.ComputersPath) }
    Test-Step 'Delete Wave refuses a name that leaves the Waves folder of the Campaign' { $refused = 0; foreach ($bad in @('..','..\Windows11Upgrade','WAVE-KEEP-TEST\..\..')) { try { [void](Remove-GuiWaveFolder -PackageRoot $packageRoot -Campaign HybridJoin -Name $bad) } catch { $refused++ } }; $refused -eq 3 -and (Test-Path -LiteralPath $keepWave.ComputersPath) }
    Remove-Item -LiteralPath $keepWave.Path -Recurse -Force
    $run = Get-GuiRunCommand -Campaign Windows11Upgrade -WavePath 'C:\Waves\W 1' -RunMode Once -IgnoreRunGuard $true -SelectionPath 'C:\Sel\s.json.txt'
    Test-Step 'Run command uses Windows PowerShell 5.1, the engine, the GUI selection and the run mode' { $run -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $run -match 'SmartM365-Invoke-EndpointMigrationWithPsExec\.ps1' -and $run -match '-GuiSelectionPath' -and $run -match '-RunMode Once' -and $run -match '-IgnoreRunGuard' -and $run -match '"C:\\Waves\\W 1"' }
    Test-Step 'Single PC command passes -ComputerName and no Wave' { $fictionalComputer = 'PC-1'; $c = Get-GuiRunCommand -Campaign HybridJoin -ComputerName $fictionalComputer -RunMode Once -IgnoreRunGuard $false -SelectionPath 'C:\s.json.txt'; $c -match '-ComputerName PC-1' -and $c -notmatch '-WavePath' -and $c -notmatch '-IgnoreRunGuard' -and $c -match '-GuiSelectionPath' }
    $refreshHandler = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Start-InventoryRefresh' }, $true)
    $planAssignment = if ($refreshHandler) { $refreshHandler.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$plan' }, $true) }
    Test-Step 'inventory refresh keeps a one-stream plan (Graph only) as an array under StrictMode' { $null -ne $planAssignment -and $planAssignment.Right.Extent.Text -match '^@\(if ' }
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
    Test-Step 'inventory refresh runs AD and Graph streams with one Graph export each' { ($hjPlan.Name -join ',') -eq 'AD,Graph' -and ($w11Plan.Name -join ',') -eq 'AD,Graph' -and @($hjPlan[1].Exports).Count -eq 1 -and @($w11Plan[1].Exports).Count -eq 1 }
    Test-Step 'HybridJoin Graph stream exports Intune, Entra and user licenses on one sign-in; Windows 11 exports Intune only' { $hjPlan[1].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-AutomaticGraphInventories.ps1' -and (@($hjPlan[1].Exports[0].Arguments) -contains '-IncludeUserLicenses') -and $w11Plan[1].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1' }
    $usersOnly = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Users)
    $entraUsers = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Entra,Users)
    Test-Step 'old inventories only: users alone run the license exporter; users with a device inventory share its sign-in' { @($usersOnly[0].Exports).Count -eq 1 -and $usersOnly[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-UserLicensesCsv.ps1' -and @($entraUsers[0].Exports).Count -eq 1 -and $entraUsers[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-AutomaticGraphInventories.ps1' -and (@($entraUsers[0].Exports[0].Arguments) -contains '-IncludeUserLicenses') }
    $graphExporterText = Get-Content -LiteralPath (Join-Path $script:PackageRoot 'Scripts\SmartM365-EndpointMigration-Export-AutomaticGraphInventories.ps1') -Raw
    Test-Step 'combined Graph exporter adds User.Read.All and runs the license export without a second connection' { $graphExporterText -match "if \(\`$IncludeUserLicenses\) \{ \`$scopes \+= 'User\.Read\.All' \}" -and $graphExporterText -match '\$userLicenseExporter -OutputPath \$UserLicensesOutputPath -NoConnect -ForceRefresh' }
    # "Refresh old inventories only": only the old inventories are exported.
    $adOnly = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds AD)
    $entraOnly = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Entra)
    $intuneOnly = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Intune)
    $graphBoth = @(Get-InventoryRefreshPlan -Campaign HybridJoin -Kinds Intune,Entra)
    $w11Intune = @(Get-InventoryRefreshPlan -Campaign Windows11Upgrade -Kinds Intune,Entra)
    Test-Step 'old inventories only: AD alone runs no Graph sign-in; one Graph inventory runs its own exporter' { ($adOnly.Name -join ',') -eq 'AD' -and ($entraOnly.Name -join ',') -eq 'Graph' -and $entraOnly[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-EntraDevicesCsv.ps1' -and $intuneOnly[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1' }
    Test-Step 'old inventories only: Intune and Entra share one sign-in; Windows 11 ignores Entra' { $graphBoth.Count -eq 1 -and $graphBoth[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-AutomaticGraphInventories.ps1' -and $w11Intune.Count -eq 1 -and $w11Intune[0].Exports[0].Script -eq 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1' -and @(Get-InventoryRefreshPlan -Campaign Windows11Upgrade -Kinds Entra).Count -eq 0 }
    Test-Step 'Automatic Wave view offers "Refresh old inventories only"' { $allXaml -match 'x:Name="AutomaticRefreshStaleInventoryButton" Content="Refresh old inventories only"' -and $guiText -match 'Start-InventoryRefresh -Kinds @\(\$old' }
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
    Test-Step 'one Wave per AD domain is offered only with (All domains) and the AD inventory' {
        (Test-AutomaticSplitByDomainAvailable -Campaign HybridJoin -AdDomain '' -Source '') -and (Test-AutomaticSplitByDomainAvailable -Campaign Windows11Upgrade -AdDomain '' -Source 'Both') -and (Test-AutomaticSplitByDomainAvailable -Campaign Windows11Upgrade -AdDomain '' -Source 'AD') -and
        -not (Test-AutomaticSplitByDomainAvailable -Campaign HybridJoin -AdDomain 'fr.example' -Source '') -and -not (Test-AutomaticSplitByDomainAvailable -Campaign Windows11Upgrade -AdDomain '' -Source 'Intune') -and -not (Test-AutomaticSplitByDomainAvailable -Campaign Prerequisites -AdDomain '' -Source '')
    }
    Test-Step 'Preview and Create show the Waves per AD domain' {
        $splitText = Get-AutomaticWaveResultText -Summary ([pscustomobject]@{ SelectedDevices = 5; ExcludedDevices = 0; FilterExcludedDevices = 0; SplitByAdDomain = $true; DomainWaves = @([pscustomobject]@{ ADDomain = 'be.example'; WaveName = 'WAVE-X-BE'; SelectedDevices = 1 }, [pscustomobject]@{ ADDomain = ''; WaveName = 'WAVE-X-NO-AD-DOMAIN'; SelectedDevices = 4 }) })
        $splitText -match 'One Wave per AD domain: 2 Wave\(s\)' -and $splitText -match '  - WAVE-X-BE: 1 \(be\.example\)' -and $splitText -match '  - WAVE-X-NO-AD-DOMAIN: 4 \(no AD domain\)' -and (Get-AutomaticWaveResultText -Summary ([pscustomobject]@{ SelectedDevices = 1; ExcludedDevices = 0; FilterExcludedDevices = 0; SplitByAdDomain = $false; DomainWaves = @() })) -notmatch 'AD domain'
    }
    Test-Step 'Automatic Wave view offers the AD domain choice' { $allXaml -match 'x:Name="AutomaticAdDomainCombo"' -and $guiText -match "AllAdDomainsLabel = '\(All domains\)'" }
    # Preview/Create: the real engine in a hidden process, as started by the GUI (fictional inventories).
    $waitWindow = [string]$ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$automaticWaitXaml' }, $true).Right.Expression.Value
    # Close: the button and the window X share one path; a waiting window covers the detection.
    $closeWaitWindow = [string]$ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$closeWaitXaml' }, $true).Right.Expression.Value
    $closingHandler = [regex]::Match($guiText, '(?s)\$window\.Add_Closing\(\{.*?\n\}\)').Value
    Test-Step 'Close button only closes the window, so it runs the same Closing handler as the window X' { $guiText -match '\$controls\.CloseGuiButton\.Add_Click\(\{ \$window\.Close\(\) \}\)' -and ([regex]::Matches($guiText, '\$window\.Add_Closing\(')).Count -eq 1 }
    Test-Step 'close waiting window: "Please wait...", indeterminate progress, no button' { $closeWaitWindow -match 'Text="Please wait\.\.\."' -and $closeWaitWindow -match 'IsIndeterminate="True"' -and $closeWaitWindow -notmatch '<Button' }
    Test-Step 'Closing shows the waiting window first and closes it before the Active Runs question' { $closingHandler -match '(?s)^\$window\.Add_Closing\(\{\s*param\(\$eventSource,\$routedEventArgs\)\s*\$closeWait = Show-GuiCloseWait.*?Get-SemtActiveRuns.*?Close-GuiCloseWait -Wait \$closeWait; \$closeWait = \$null.*?''Active Runs''.*?finally \{ Close-GuiCloseWait -Wait \$closeWait \}' }
    Test-Step 'Closing closes the waiting window before the inventory refresh question too' { $closingHandler -match '(?s)Test-InventoryRefreshRunning\) \{\s*Close-GuiCloseWait -Wait \$closeWait; \$closeWait = \$null\s*\$choice = \[System\.Windows\.MessageBox\]::Show\(\$window, "The shared inventory refresh' }
    Test-Step 'Preview and Create show a waiting window with progress and the engine stage' { $waitWindow -match 'IsIndeterminate="True"' -and $waitWindow -match 'x:Name="AutomaticWaitStageText"' -and $guiText -match 'Show-AutomaticWaveWait -Create:\$Create' }
    Test-Step 'a successful Preview keeps its evidence; Create passes it to the engine only when it matches the filters' { $guiText -match "AutomaticPreviewEvidence = \[string\]\(Get-SemtPropertyValue -InputObject \`$summary -Name 'EvidencePath'\)" -and $guiText -match '-PreviewEvidencePath \$\(if \(\$reusePreview\) \{ \$script:AutomaticPreviewEvidence \}' -and $guiText -match '\$reusePreview = \(\$Create -and \(Test-AutomaticPreviewReusable -Signature \$signature -PreviewSignature \$script:AutomaticPreviewSignature -PreviewEvidence \$script:AutomaticPreviewEvidence\)\)' }
    Test-Step 'Preview and Create end with a message (counts; Create states the Wave is not launched)' { $guiText -match "'Preview', 'OK', 'Information'" -and $guiText -match "The Wave was NOT launched\." -and $guiText -match "'Create Wave', 'OK', 'Information'" }
    $powershell51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $runPreview = {
        param([string]$Name,[switch]$Create)
        $evidenceRoot = Join-Path $tempRoot ("AutoEvidence-{0}" -f $Name)
        New-Item -ItemType Directory -Path $evidenceRoot -Force | Out-Null
        $logPath = Join-Path $evidenceRoot 'GuiAutomaticWave.log'
        $command = Get-AutomaticWaveCommandText -Campaign HybridJoin -Filters @{ WaveName = 'WAVE-AUTO-HJ-GUI'; Prefixes = @(); Contains = @(); ExcludeStaleAd = $false; LastLogonDays = 45; Source = 'Both'; ExcludeIntunePresent = $false; AdDomain = '' } -EvidenceRoot $evidenceRoot -Create:$Create
        $process = Start-Process -FilePath $powershell51 -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes((Get-AutomaticWaveProcessText -CommandText $command -LogPath $logPath)))) -WindowStyle Hidden -PassThru -Wait
        [pscustomobject]@{ ExitCode = $process.ExitCode; Log = $(if (Test-Path -LiteralPath $logPath) { Get-Content -LiteralPath $logPath -Raw } else { '' }); Summary = @(Get-ChildItem -LiteralPath $evidenceRoot -Filter 'AutomaticWaveSummary.json.txt' -File -Recurse); Selection = @(Get-ChildItem -LiteralPath $evidenceRoot -Filter 'AutomaticWaveSelection.csv' -File -Recurse) }
    }
    $missing = & $runPreview 'Missing'
    Test-Step 'Preview without inventory fails with exit code 1 and the error in the log' { $missing.ExitCode -eq 1 -and $missing.Log -match '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ERROR Required automatic Wave inventory not found' -and $missing.Summary.Count -eq 0 }
    $inventoryRoot = Join-Path $packageRoot 'Inventory'
    New-Item -ItemType Directory -Path $inventoryRoot -Force | Out-Null
    [pscustomobject]@{ ComputerName = 'GUI-PC01'; Name = 'GUI-PC01'; ADInventoryPresent = $true; ADDomain = 'contoso.test'; Enabled = $true; DNSHostName = 'gui-pc01.contoso.test'; DistinguishedName = 'CN=GUI-PC01,DC=contoso,DC=test'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).ToUniversalTime().ToString('o') } | Export-Csv -LiteralPath (Join-Path $inventoryRoot 'DevicesAD.csv') -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{ ComputerName = 'OTHER'; DeviceName = 'OTHER'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = 'x'; LastSyncDateTime = (Get-Date).ToUniversalTime().ToString('o') } | Export-Csv -LiteralPath (Join-Path $inventoryRoot 'DevicesIntune.csv') -NoTypeInformation -Encoding UTF8
    $ok = & $runPreview 'Ok'
    Test-Step 'Preview succeeds in the hidden process and writes its summary' { $ok.ExitCode -eq 0 -and $ok.Summary.Count -eq 1 }
    # Create without a Preview: the GUI passes no -PreviewEvidencePath; the engine computes the selection,
    # keeps its evidence and creates the Wave (not launched). The Wave is removed afterwards.
    $direct = & $runPreview 'CreateWithoutPreview' -Create
    $directWave = Join-Path $packageRoot 'Waves\HybridJoin\WAVE-AUTO-HJ-GUI'
    $directComputers = if (Test-Path -LiteralPath (Join-Path $directWave 'Computers.txt')) { @(Get-Content -LiteralPath (Join-Path $directWave 'Computers.txt') | Where-Object { $_ -and $_ -notmatch '^\s*#' }) } else { @() }
    $directOk = ($direct.ExitCode -eq 0 -and $direct.Summary.Count -eq 1 -and $direct.Selection.Count -eq 1 -and @(Import-Csv -LiteralPath $direct.Selection[0].FullName | Where-Object { $_.ComputerName -like 'GUI-PC01*' }).Count -eq 1 -and @($directComputers | Where-Object { $_ -match '^GUI-PC01' }).Count -eq 1)
    Write-TestResult -Condition $directOk -Message 'Create without a Preview computes the selection, writes AutomaticWaveSelection.csv and creates the Wave' -Detail ('exit={0}; summary={1}; selection={2}; computers={3}; log={4}' -f $direct.ExitCode,$direct.Summary.Count,$direct.Selection.Count,($directComputers -join '|'),((@($direct.Log -split '\r?\n' | Where-Object { $_.Trim() }) | Select-Object -Last 3) -join ' / '))
    Remove-Item -LiteralPath $directWave -Recurse -Force -ErrorAction SilentlyContinue
    Test-Step 'engine stages reach the log for the waiting window' { (Get-AutomaticWaveLastStage -Text $ok.Log) -match '^Automatic Wave preview ready\.' -and $ok.Log -match ' STAGE Loading inventory records\.\.\.' }
    Test-Step 'end message counts read the HybridJoin and Windows 11 summaries' { (Get-AutomaticWaveResultText -Summary ([pscustomobject]@{ SelectedDevices = 3; ExcludedDevices = 2; FilterExcludedDevices = 1 })) -match 'Selected: 3' -and (Get-AutomaticWaveResultText -Summary ([pscustomobject]@{ SelectedDevices = 4; ExcludedDevices = 0; NameFilterExcludedDevices = 7 })) -match 'Filtered out: 7' }
    $reasonSummary = '{"SelectedDevices":5,"ExcludedDevices":4,"FilterExcludedDevices":0,"ExclusionReasonCounts":{"AD_DISABLED":1,"INTUNE_ENROLLED_CONFIRMED_BY_DEVICE_ID":3}}' | ConvertFrom-Json
    $reasonText = Get-AutomaticWaveResultText -Summary $reasonSummary
    $nameFilterText = Get-AutomaticWaveResultText -Summary ([pscustomobject]@{ SelectedDevices = 1; ExcludedDevices = 0; FilterExcludedDevices = 2; ComputerNamePrefixes = 'FR'; ComputerNameContains = ''; ComputerNameNotContains = '-A-' })
    Test-Step 'end message shows the name filters that are set' { $nameFilterText -match '(?m)^  - Name prefix: FR\r?$' -and $nameFilterText -match '(?m)^  - Name does not contain: -A-\r?$' -and $nameFilterText -notmatch 'Name contains' }
    Test-Step 'end message lists the count per safety exclusion reason' { $reasonText -match '(?m)^  - INTUNE_ENROLLED_CONFIRMED_BY_DEVICE_ID: 3\r?$' -and $reasonText -match '(?m)^  - AD_DISABLED: 1\r?$' -and $reasonText -match '(?s)Safety exclusions: 4.*AD_DISABLED.*Filtered out: 0' }
    Remove-Item -LiteralPath $inventoryRoot -Recurse -Force
    $catalog = Import-SemtOptionCatalog -PackageRoot $packageRoot
    $values = @{ ThrottleLimit = 4; 'Windows11Upgrade.SetupSourcePath' = ''; DryRun = $true }
    $selection = New-GuiSelectionFile -Catalog $catalog -Campaign Windows11Upgrade -Values $values
    $captured = Read-SemtJsonFile -Path $selection
    Test-Step 'GUI selection is stored per Campaign under Runs\<Campaign>\_GuiSelections' { $selection -like (Join-Path $packageRoot 'Runs\Windows11Upgrade\_GuiSelections\*.json.txt') -and $captured.Campaign -eq 'Windows11Upgrade' }
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
    Test-Step 'GUI -ValidateOnly loads under Windows PowerShell 5.1 STA for both Campaigns' { $LASTEXITCODE -eq 0 -and @($validate | Where-Object { $_ -match '^ValidateOnly: GUI 1\.1\.0 loaded; campaign=(HybridJoin|Windows11Upgrade); options=\d+' }).Count -eq 2 }
    Test-Step 'GUI -ValidateOnly loads the three dedicated windows; Automatic Wave is offered for HybridJoin and Windows 11' { @($validate | Where-Object { $_ -match 'campaign=(HybridJoin|Windows11Upgrade); .*; windows=SinglePc,NewWave,AutomaticWave; automaticWaveButton=Visible\.$' }).Count -eq 2 }
    Test-Step 'GUI hides the Prerequisites Campaign while Prerequisites.Enabled is false (default)' { @($validate | Where-Object { $_ -match 'campaign=Prerequisites' }).Count -eq 0 }
    $globalConfig = Join-Path $packageRoot 'Config\SmartEndpointMigrationToolkit.local.json.txt'
    [IO.File]::WriteAllText($globalConfig, '{ "SchemaVersion": 1, "Options": { "Prerequisites.Enabled": true } }')
    try { $validateEnabled = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $guiPath -ValidateOnly 2>&1 | ForEach-Object { [string]$_ }) }
    finally { Remove-Item -LiteralPath $globalConfig -Force -ErrorAction SilentlyContinue }
    Test-Step 'GUI lists the Prerequisites Campaign when Prerequisites.Enabled is true in the global configuration' { $LASTEXITCODE -eq 0 -and @($validateEnabled | Where-Object { $_ -match '^ValidateOnly: GUI 1\.1\.0 loaded; campaign=(HybridJoin|Windows11Upgrade|Prerequisites); options=\d+' }).Count -eq 3 -and @($validateEnabled | Where-Object { $_ -match 'campaign=Prerequisites; options=\d+' }).Count -eq 1 }
    Test-Step 'Prerequisites hides the Automatic Wave button and keeps Single PC and New Wave' { @($validateEnabled | Where-Object { $_ -match 'campaign=Prerequisites; .*; windows=SinglePc,NewWave,AutomaticWave; automaticWaveButton=Collapsed\.$' }).Count -eq 1 }
    # The three dedicated windows load under Windows PowerShell 5.1 STA with the resources of the main window.
    $windowsScript = Join-Path $tempRoot 'windows-check.ps1'
    @'
param($Gui)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Gui, [ref]$null, [ref]$null)
function Get-Literal([string]$Name) { [string]$ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$' + $Name) }, $true).Right.Expression.Value }
. ([scriptblock]::Create($ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-GuiChildWindowMarkup' }, $true).Extent.Text))
$main = Get-Literal 'xaml'
foreach ($pair in @(@('singlePcXaml','LaunchSingleButton'),@('newWaveXaml','CreateWaveButton'),@('automaticWaveXaml','AutomaticCreateButton'),@('singlePcXaml','SinglePcCloseButton'),@('newWaveXaml','NewWaveCloseButton'),@('automaticWaveXaml','AutomaticWaveCloseButton'))) {
    $child = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new((Get-GuiChildWindowMarkup -Markup (Get-Literal $pair[0]) -MainMarkup $main))))
    $button = $child.FindName($pair[1])
    'WINDOW={0};NAME={4};BUTTON={1};ENABLED={2};STYLED={3}' -f $pair[0],($null -ne $button),$button.IsEnabled,($null -ne $button.Style),$pair[1]
}
# Layout: at the default size and at the minimum size, every button outside a ScrollViewer is whole and inside the window (no clipped Create Wave).
foreach ($name in 'singlePcXaml','newWaveXaml','automaticWaveXaml') {
    foreach ($size in 'Default','Minimum') {
        $child = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new((Get-GuiChildWindowMarkup -Markup (Get-Literal $name) -MainMarkup $main))))
        if ($size -eq 'Minimum') { $child.Width = $child.MinWidth; $child.Height = $child.MinHeight }
        $child.WindowStartupLocation = 'Manual'; $child.Left = -20000; $child.Top = -20000; $child.ShowActivated = $false
        $child.Show(); $child.UpdateLayout()
        $root = $child.Content
        $bad = @()
        $queue = New-Object System.Collections.Queue; $queue.Enqueue($root)
        while ($queue.Count -gt 0) {
            $node = $queue.Dequeue()
            if ($node -is [Windows.Controls.ScrollViewer]) { continue }
            if ($node -is [Windows.Controls.Button] -and $node.Visibility -eq 'Visible' -and $node.Name) {
                $box = $node.TransformToAncestor($root).TransformBounds([Windows.Rect]::new(0, 0, $node.ActualWidth, $node.ActualHeight))
                $clip = [Windows.Controls.Primitives.LayoutInformation]::GetLayoutClip($node)
                if ($node.ActualHeight -lt 20 -or $box.Bottom -gt ($root.ActualHeight + 0.5) -or $box.Right -gt ($root.ActualWidth + 0.5) -or $null -ne $clip) { $bad += $node.Name }
            }
            for ($i = 0; $i -lt [Windows.Media.VisualTreeHelper]::GetChildrenCount($node); $i++) { $queue.Enqueue([Windows.Media.VisualTreeHelper]::GetChild($node, $i)) }
        }
        $child.Close()
        'LAYOUT={0};SIZE={1};CLIPPED={2}' -f $name,$size,($bad -join ',')
    }
}
'@ | Set-Content -LiteralPath $windowsScript -Encoding ASCII
    $windowsCheck = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $windowsScript -Gui $guiPath 2>&1 | ForEach-Object { [string]$_ })
    Test-Step 'Single PC, New Wave and Automatic Wave windows load under Windows PowerShell 5.1 STA (styled main buttons)' { @($windowsCheck | Where-Object { $_ -match '^WINDOW=(singlePcXaml|newWaveXaml|automaticWaveXaml);NAME=[A-Za-z]+;BUTTON=True;ENABLED=True;STYLED=True$' }).Count -eq 6 }
    Test-Step 'no button of the three windows is clipped, at the default size and at the minimum size (Create Wave, Launch Single PC, Close)' { @($windowsCheck | Where-Object { $_ -match '^LAYOUT=(singlePcXaml|newWaveXaml|automaticWaveXaml);SIZE=(Default|Minimum);CLIPPED=$' }).Count -eq 6 }
    Test-Step 'Create Wave of the Automatic Wave window is enabled when the window opens (no Preview needed)' { @($windowsCheck | Where-Object { $_ -match '^WINDOW=automaticWaveXaml;NAME=AutomaticCreateButton;BUTTON=True;ENABLED=True' }).Count -eq 1 }
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
    $closeScript = Join-Path $tempRoot 'close-check.ps1'
    @'
param($Gui)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Gui, [ref]$null, [ref]$null)
function Get-Literal([string]$Name) { [string]$ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$' + $Name) }, $true).Right.Expression.Value }
$wait = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new((Get-Literal 'closeWaitXaml'))))
$wait.Show()
'WAIT=' + ($wait.IsVisible -and $wait.FindName('CloseWaitTitleText').Text -eq 'Please wait...')
$wait.Close()
$main = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new((Get-Literal 'xaml'))))
$main.Width = $main.MinWidth; $main.Height = $main.MinHeight; $main.WindowStartupLocation = 'Manual'; $main.Left = -20000; $main.ShowActivated = $false
$main.Show(); $main.UpdateLayout()
$root = $main.Content
foreach ($name in 'LaunchAllButton','CloseGuiButton') {
    $button = $main.FindName($name)
    $box = $button.TransformToAncestor($root).TransformBounds([Windows.Rect]::new(0, 0, $button.ActualWidth, $button.ActualHeight))
    'BUTTON={0};VISIBLE={1}' -f $name,($button.ActualWidth -gt 40 -and $box.Right -le ($root.ActualWidth + 0.5))
}
$launch = $main.FindName('LaunchAllButton'); $close = $main.FindName('CloseGuiButton')
'ORDER=' + ($close.TransformToAncestor($root).Transform([Windows.Point]::new(0,0)).X -gt $launch.TransformToAncestor($root).Transform([Windows.Point]::new(0,0)).X)
'STYLE=' + ($null -eq $close.ReadLocalValue([Windows.FrameworkElement]::StyleProperty) -or $close.ReadLocalValue([Windows.FrameworkElement]::StyleProperty) -eq [Windows.DependencyProperty]::UnsetValue)
$main.Close()
'@ | Set-Content -LiteralPath $closeScript -Encoding ASCII
    $closeCheck = (@(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $closeScript -Gui $guiPath 2>&1 | ForEach-Object { [string]$_ })) -join "`n"
    Test-Step 'close waiting window loads and shows under Windows PowerShell 5.1 STA' { $closeCheck -match 'WAIT=True' }
    Test-Step 'Close button is right of Launch All Waves, with the default bar style, visible at the minimum window size' { $closeCheck -match 'BUTTON=LaunchAllButton;VISIBLE=True' -and $closeCheck -match 'BUTTON=CloseGuiButton;VISIBLE=True' -and $closeCheck -match 'ORDER=True' -and $closeCheck -match 'STYLE=True' }
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
'WINDOWS=' + (@($state.Data.SemtWindows.Keys) -join ',')
foreach ($name in 'OpenSinglePcButton','OpenNewWaveButton','OpenAutomaticWaveButton','DeleteWaveButton','LaunchWaveButton','LaunchAllButton','LaunchSingleButton','CreateWaveButton','AutomaticPreviewButton','AutomaticCreateButton','StopAllButton','ForceStopSelectedRunButton','OpenRunReportButton','ResetOptionsButton','RefreshButton','CloseGuiButton') {
    $state.Controls[$name].RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
}
'FAILED=' + @($state.Activity | Where-Object { $_ -match 'Demo action failed' }).Count
'OUTPUT=' + $state.LastOutput
foreach ($name in 'SinglePcCloseButton','NewWaveCloseButton','AutomaticWaveCloseButton') { $state.Controls[$name].RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
'CLOSED=' + @($state.Activity | Where-Object { $_ -match '(Single PC|New Wave|Automatic Wave) window closed' }).Count
'OPENED=' + @($state.Activity | Where-Object { $_ -match '(Single PC|New Wave|Automatic Wave) window opened' }).Count
'NOLAUNCH=' + (@($state.Activity | Where-Object { $_ -match 'No launcher, PsExec, worker' }).Count -gt 0)
'DEMOCLOSE=' + (@($state.Activity | Where-Object { $_ -match 'Close simulated' }).Count -eq 1)
$state.Window.Close()
'@ | Set-Content -LiteralPath $demoScript -Encoding ASCII
    $demo = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $demoScript -Gui $guiPath 2>&1 | ForEach-Object { [string]$_ })
    $demoText = $demo -join "`n"
    Test-Step 'demo loads every view' { $demoText -match 'VIEWS=Waves,Runs,Options' }
    Test-Step 'demo loads the three dedicated windows and their buttons open them' { $demoText -match 'WINDOWS=SinglePc,NewWave,AutomaticWave' -and $demoText -match 'OPENED=3' -and $demoText -match 'CLOSED=3' }
    Test-Step 'demo shows the real catalog options for both Campaigns' { $demoText -match 'OPTIONS1=\d{2}' -and $demoText -match 'CAMPAIGN=Windows11Upgrade;OPTIONS2=\d{2}' }
    Test-Step 'every demo action completes without error' { $demoText -match 'FAILED=0' }
    Test-Step 'demo report is generated in the demo temp folder only' { $demoText -match 'OUTPUT=.*SmartM365-Demo.*Run-Report-DEMO\.html' }
    Test-Step 'demo states that nothing was launched' { $demoText -match 'NOLAUNCH=True' }
    Test-Step 'demo Close button is simulated and keeps the demo open' { $demoText -match 'DEMOCLOSE=True' }
    Test-Step 'demo wrote no Wave, Run or configuration into the package' { @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Waves') -Directory -Recurse | Where-Object { $_.Parent.Name -in @('HybridJoin','Windows11Upgrade') }).Count -eq 0 -and @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter '*.local.json.txt' -File).Count -eq 0 }
    # Real start (not the demo) of the temp package copy: the window renders, the Close button runs
    # the close path of the window X (waiting window, active Run detection) and the GUI exits.
    $smoke = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $guiPath -SmokeTest 2>&1 | ForEach-Object { [string]$_ })
    $smokeText = $smoke -join "`n"
    Test-Step 'GUI -SmokeTest starts the real window and the Close button closes it through the waiting window' { $LASTEXITCODE -eq 0 -and $smokeText -match 'SmokeTest: GUI 1\.1\.0 rendered and closed by the Close button; campaign=HybridJoin; waves=\d+; closeWaitShown=1; activeRuns=0\.' }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxPVhMVsvuf0Z5
# F7VUvM1A2PPoKTKfUb4TSSqM21EOUaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGPLCDN6ZskWIoqD63jA6VyyyhidYYSguuv5W4EGieR6MA0GCSqG
# SIb3DQEBAQUABIIBgH4FKIx9Qx8dDZqZqvwG/h/KxWXSt8WPHxWLtxp/dGxwxaj+
# ZNkQdAPd8uBubOIMQw5+m43YC9irqy8CgJSNwVMD3fsCVIR07PjNSacQLTvoLsHc
# nsWGptruDfCnrhNQuRvY6vinGXmm9nz8FJxLbRvdEHOzq8qkk/wIdEWmdmhMlyt0
# ikRyVwTe6E2EIW+/fRZdt4HBkquV61DlX/XQEzwEUAgqwRfGC9Y8rVTh5ogVjCXi
# ndkrKv3BY0mfu4b+YEMl3osEkhEovsitTMDBmS1isWxRc1Si2Da6IygJp56QZNti
# i1SWt+9OEAurUZakDuxbg/qFgA82T6/5M6vDZ/U4qyIc8KNkSAbxeRhBmwS9I8G0
# dRDekLYN9I8bMIYPhLlXY1LuvVTnKu26R/fGbtbt5eZrBJEXITIpND0z8G7rwSeB
# yoR9RrRAncr3oSEBXKy6eOcELdPfgoGchN4uAJg/BL77Zh8q+ZHfPszbB8uale+g
# /bOLt6vfUgTgu08WL6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMzMx
# MTdaMC8GCSqGSIb3DQEJBDEiBCDkwirwZYpuallBVt0cbA50ARMYQa7fl4kkgQ3N
# 7o/kWzANBgkqhkiG9w0BAQEFAASCAgBY4DnRXWy6/7I9z6sg1cU536eIUVSKiVzp
# rTyhfP6t1DnQeh3eDQtqc9aTBExZCz/LlbdGSGSuhlVQy6NZg8C819BWtcFzd8LB
# 4MMRxKr0Jy8nj9VAxYrhfePyW45hpkbmb3OxdXxSsQQNh9xPcWyieXl7qhxSyiGS
# i60MARiAAX1eMEzN8jrlKsioHs98CdGWidz4dFKml2JO3PUMtYWH6Nf7YyN9MHEN
# VO9jv+hJyBVTHSOMDJ27eetzAeuaQin4iz98LfXoZr+rUB4LJVUGMJKfDLOFj4P7
# ejuoPF8m4a5Jb4J/S0MV9/u8F8GaCDDq9GxutfpBg2iWixmBlhEZ/WnihxSnFH5y
# VUihgiXLNeH2neoAHsCmBxjKHT8qKLfH/UWaCKInN26Jv/RUFFGtrmsdlkHxgvAL
# bAvnHrOfBA32NFSbPYKiulH7Xv2l4Sc3cCqO7KmFxjzF2jW3cpk6wnYzc+2Zt9Ep
# rSkUl24tTqBvApH8RYZhX0BoSZEFXDviM1Eoy/OOUlrFcnoVOB5ZTVbnVHJWlMNT
# Cgm/SRgB0tT3UHj6td8YdaHOoZl8Gtt4hl1z0bDVH+Sva8cmOD4lwUG/Ui6zy1mg
# apiWZssAYxZRokHj77N2FLrowp8kLm635NfXEzXRUCmcDchVZyHh0Vmv1t6k9TFM
# kOJH3G+ouQ==
# SIG # End signature block
