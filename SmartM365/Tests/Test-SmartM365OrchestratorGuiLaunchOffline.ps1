#Requires -Version 7.0
<#
.SYNOPSIS
Offline tests of the Operations launch features of the orchestrator GUI.
.VERSION
1.0.0
.NOTES
GUI functions are extracted by AST and bound to the real XAML. The catalog is the committed
template (read only); the published configuration is a copy of its jobs. No orchestrator, request
or collector is started: the request submission is mocked and the only process started is a
fictional script in a hidden window.
Checks: job and manual lists, resolved command preview, confirmation and interactive flags,
execution mode of on-premises jobs in the draft (set, reset, apply to all, validation), orchestrator
actions identical to the former Launchers\Orchestrator .cmd files, request with the selected job,
console start with quoted arguments.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$inventoryRoot = (Resolve-Path (Join-Path $folder '..')).Path
Import-Module (Join-Path $folder 'SmartM365.Inventory.Launch.psm1') -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Management.psm1') -Force
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'GUI parsing failed.' }
foreach ($name in @('ConvertFrom-OrchestratorGuiXaml', 'Get-OrchestratorGuiPropertyValue', 'Set-JsonProperty', 'Copy-JsonDocument', 'Get-ComboText', 'Select-ComboText', 'Get-GuiLaunchCatalog', 'Refresh-LaunchView',
        'Get-GuiLaunchPlanText', 'Update-SelectedLaunchPreview', 'Start-GuiConsoleProcess', 'Get-GuiOrchestratorActionCommand', 'Set-GuiJobExecutionMode', 'Request-SelectedJobRun')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Missing GUI function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
$assignment = $ast.Find({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml' }, $true)
$window = ConvertFrom-OrchestratorGuiXaml -Text $assignment.Right.Expression.Value
$script:Controls = @{}
foreach ($name in @('OperationsTabs', 'OnPremModeText', 'JobExecutionModeCombo', 'ApplyJobExecutionModeButton', 'ApplyModeToAllOnPremButton', 'LaunchJobsGrid', 'RunJobNowButton', 'RequestJobButton',
        'LaunchIncludeDependenciesCheck', 'LaunchJobCommandBox', 'ManualLaunchesGrid', 'RunManualLaunchButton', 'ManualLaunchCommandBox', 'OrchestratorStartButton', 'OrchestratorStopButton',
        'OrchestratorRestartButton', 'FullCollectButton', 'SendSummaryButton', 'StatusText', 'PlanningGrid', 'IncludeDependenciesCheck', 'OperationsRunningGrid')) {
    $script:Controls[$name] = $window.FindName($name)
    if (-not $script:Controls[$name]) { throw "Missing XAML control: $name" }
}

$script:Cases = 0
function Assert-Case { param([bool]$Condition, [string]$Message) $script:Cases++; if (-not $Condition) { throw "FAIL: $Message" } }
$script:Activity = [Collections.Generic.List[string]]::new()
function Write-GuiActivity { param($Message, $Level = 'INFO', [switch]$DisplayOnly) $script:Activity.Add("$Level|$Message") }
function Write-GuiException { param($Context, $ErrorRecord) throw "$Context : $($ErrorRecord.Exception.Message)" }

$Tenant = 'demo'
$previousOrchestratorTenant = $env:SMARTM365_ORCHESTRATOR_TENANT
$env:SMARTM365_ORCHESTRATOR_TENANT = ''
$script:SmartInventoryRoot = $inventoryRoot
$template = Get-Content -LiteralPath (Join-Path $folder 'Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json -Depth 100
$published = [pscustomobject]@{ Jobs = @($template.Jobs) }
$script:Snapshot = [pscustomobject]@{ Jobs = $published }
$script:DraftJobs = Copy-JsonDocument -Document $published
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiLaunch-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root | Out-Null
try {
    # Lists, global mode and previews.
    Refresh-LaunchView
    Assert-Case (@($script:LaunchJobRows).Count -eq @($template.Jobs).Count) 'Every catalog job must be listed.'
    Assert-Case (@($script:ManualLaunchRows).Count -eq @($template.Launches).Count) 'Every manual entry must be listed separately.'
    Assert-Case ($script:Controls.OnPremModeText.Text -in @('LocalCache', 'Direct', 'Auto')) 'The global on-premises mode must be shown.'
    $mailboxes = @($script:LaunchJobRows | Where-Object Name -eq 'Exchange2016-Local-Mailboxes-Inventory')[0]
    Assert-Case ($mailboxes.OnPremises -and $mailboxes.Runtime -eq 'WindowsPowerShell' -and $mailboxes.DraftMode -eq '(default)' -and $mailboxes.ModeSource -eq 'OnPremExecutionMode') 'An on-premises job must show its runtime, default mode and source.'
    $cloud = @($script:LaunchJobRows | Where-Object Name -eq 'M365-ActiveUsers-Inventory')[0]
    Assert-Case (-not $cloud.OnPremises -and $cloud.Mode -eq 'Direct' -and $cloud.DraftMode -eq '') 'A cloud job runs Direct and has no mode to set.'

    $script:Controls.LaunchJobsGrid.SelectedItem = $cloud
    Update-SelectedLaunchPreview -Kind Job
    Assert-Case ($script:Controls.RunJobNowButton.IsEnabled -and $script:Controls.RequestJobButton.IsEnabled -and -not $script:Controls.JobExecutionModeCombo.IsEnabled) 'Cloud job: run and request enabled, mode disabled.'
    Assert-Case ($script:Controls.LaunchJobCommandBox.Text -match 'Tenant: demo \(source: Parameter\)' -and $script:Controls.LaunchJobCommandBox.Text -match 'SmartM365-ActiveUsers-Inventory\.ps1 -Tenant demo -Connect') 'The resolved command of the job must be previewed.'
    $script:Controls.LaunchJobsGrid.SelectedItem = $mailboxes
    Update-SelectedLaunchPreview -Kind Job
    Assert-Case ($script:Controls.JobExecutionModeCombo.IsEnabled -and (Get-ComboText -Combo $script:Controls.JobExecutionModeCombo) -eq '(default)') 'On-premises job: the mode can be set.'

    $write = @($script:ManualLaunchRows | Where-Object Name -eq 'Exchange2016-ProxyAddresses-Write')[0]
    Assert-Case ($write.RequiresConfirmation -and $write.Interactive -and $write.Runtime -eq 'WindowsPowerShell') 'The write entry must be flagged for confirmation.'
    $script:Controls.ManualLaunchesGrid.SelectedItem = $write
    Update-SelectedLaunchPreview -Kind Manual
    Assert-Case ($script:Controls.RunManualLaunchButton.IsEnabled -and $script:Controls.ManualLaunchCommandBox.Text -match 'Confirmation: the user must type WRITE' -and $script:Controls.ManualLaunchCommandBox.Text -match '-AllOrganizationalUnit -AddMissingAddress') 'The manual preview must show the command and the confirmation.'

    # Execution mode in the draft: set, reset, apply to all, validation.
    [void](Set-GuiJobExecutionMode -JobName 'Exchange2016-Local-Mailboxes-Inventory' -Mode Direct)
    $draftJob = @($script:DraftJobs.Jobs | Where-Object Name -eq 'Exchange2016-Local-Mailboxes-Inventory')[0]
    Assert-Case ($draftJob.Execution.Mode -eq 'Direct' -and @($script:LaunchJobRows | Where-Object Name -eq 'Exchange2016-Local-Mailboxes-Inventory')[0].DraftMode -eq 'Direct') 'The mode must be written to the draft and shown.'
    Assert-Case ((Test-SmartM365OrchestratorJobsDocument -Document $script:DraftJobs).Valid) 'A draft with an execution mode must stay valid.'
    Assert-Case ($draftJob.Execution.PSObject.Properties['CacheName'] -and $draftJob.Execution.CacheName -eq 'ExchangeLocalMailboxes\WithADPermissions') 'Other Execution fields of the draft job must be kept.'
    [void](Set-GuiJobExecutionMode -JobName 'Exchange2016-Local-Mailboxes-Inventory' -Mode '(default)')
    Assert-Case (-not $draftJob.Execution.PSObject.Properties['Mode']) '(default) must remove the job mode.'
    $onPremNames = @($script:LaunchJobRows | Where-Object { $_.OnPremises -and $_.InManifest } | ForEach-Object Name)
    $changed = @(Set-GuiJobExecutionMode -JobName $onPremNames -Mode Auto)
    Assert-Case ($changed.Count -eq $onPremNames.Count -and $changed.Count -ge 7 -and @($script:LaunchJobRows | Where-Object { $_.OnPremises -and $_.DraftMode -ne 'Auto' }).Count -eq 0) 'Apply to all must set every on-premises job.'
    $bad = Copy-JsonDocument -Document $script:DraftJobs
    $bad.Jobs[0] | Add-Member -NotePropertyName Execution -NotePropertyValue ([pscustomobject]@{ Mode = 'Remote' }) -Force
    $badValidation = Test-SmartM365OrchestratorJobsDocument -Document $bad
    Assert-Case (-not $badValidation.Valid -and (@($badValidation.Errors) -join ' ') -match 'Execution.Mode must be Direct, LocalCache or Auto') 'An invalid execution mode must be rejected before publication.'
    $previewOfPublished = (Get-GuiLaunchPlanText -Name 'Exchange2016-Local-Mailboxes-Inventory').Plan
    Assert-Case ($previewOfPublished.Mode -eq 'LocalCache') 'Previews follow the published configuration, not the unpublished draft.'

    # Orchestrator actions: same commands as the former Launchers\Orchestrator .cmd files (frozen reference).
    $reference = @((Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'SmartInventory-LegacyLaunchers.psd1')).Launchers)
    $legacy = @{
        Start = 'Start-SmartM365-Inventory-Orchestrator.cmd'; Stop = 'Stop-SmartM365-Inventory-Orchestrator.cmd'; Restart = 'Restart-SmartM365-Inventory-Orchestrator.cmd'
        FullCollect = 'Start-SmartM365-Inventory-Full-Collect.cmd'; SendSummary = 'Send-SmartM365-Inventory-Orchestrator-ExecutionSummary.cmd'
    }
    foreach ($action in $legacy.Keys) {
        $command = Get-GuiOrchestratorActionCommand -Action $action
        Assert-Case (Test-Path -LiteralPath $command.ScriptPath -PathType Leaf) "$action script exists."
        $record = @($reference | Where-Object { $_.Launcher -eq ('Launchers\Orchestrator\' + $legacy[$action]) }) | Select-Object -First 1
        Assert-Case ($null -ne $record -and $record.Runtime -eq 'PowerShell7' -and $record.Tenant -eq 'prod') "$action is in the frozen reference."
        Assert-Case ([IO.Path]::GetFullPath($command.ScriptPath) -ieq [IO.Path]::GetFullPath((Join-Path $inventoryRoot $record.Script))) "$action uses the script of $($legacy[$action])."
        $arguments = @($command.Arguments)
        Assert-Case ($arguments[0] -eq '-Tenant' -and $arguments[1] -eq 'demo') "$action passes the GUI tenant."
        Assert-Case ((@($arguments | Select-Object -Skip 2) -join ' ') -eq (@($record.Arguments) -join ' ')) "$action arguments [$(@($arguments | Select-Object -Skip 2) -join ' ')] <> [$(@($record.Arguments) -join ' ')]."
    }

    # Request through the orchestrators with the selected job (submission mocked).
    $script:RequestCalls = [Collections.Generic.List[object]]::new()
    $script:RequestRows = @()
    function Refresh-RequestsView {}
    function Get-PublishedJobsPath { 'published.json' }
    function Test-JobDraftChanged { param($JobName) $false }
    function New-SmartM365OrchestratorJobRunRequest { param($SharedDataFolderPath, $JobsPath, $JobName, $Tenant, [switch]$IncludeDependencies) $script:RequestCalls.Add([pscustomobject]@{ Jobs = @($JobName); Tenant = $Tenant; Dependencies = [bool]$IncludeDependencies }); [pscustomobject]@{ BatchId = 'Batch1'; TotalCount = 1 } }
    $script:SharedDataFolderPath = $root
    $script:Controls.IncludeDependenciesCheck.IsChecked = $false
    [void](Request-SelectedJobRun -SkipConfirmation -JobName 'AD-Inventory' -IncludeDependencies $true)
    Assert-Case ($script:RequestCalls.Count -eq 1 -and $script:RequestCalls[0].Jobs[0] -eq 'AD-Inventory' -and $script:RequestCalls[0].Dependencies -and $script:RequestCalls[0].Tenant -eq 'demo') 'Request must submit the selected job with its dependency choice.'

    # Console start: arguments with spaces reach the script (hidden window, fictional script).
    $probe = Join-Path $root 'Probe Folder\SmartM365-Fixture-Probe.ps1'
    New-Item -ItemType Directory -Path (Split-Path $probe -Parent) | Out-Null
    [IO.File]::WriteAllText($probe, "param([string]`$Value, [string]`$Marker) [IO.File]::WriteAllText(`$Marker, (`$Value + '|' + `$PSVersionTable.PSEdition))", [Text.UTF8Encoding]::new($false))
    $marker = Join-Path $root 'console marker.txt'
    $process = Start-GuiConsoleProcess -Runtime WindowsPowerShell -ScriptPath $probe -Arguments @('-Value', 'two words "quoted"', '-Marker', $marker) -WindowStyle Hidden
    if (-not $process.WaitForExit(60000)) { throw 'Console probe did not exit.' }
    Assert-Case ((Get-Content -LiteralPath $marker -Raw) -eq 'two words "quoted"|Desktop') 'The console must receive the exact arguments with the requested runtime.'

    # GUI shortcut of the scheduled task installer: hidden PowerShell 7, minimized, no console window.
    $installerAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'), [ref]$null, [ref]$null)
    foreach ($name in @('ConvertTo-QuotedArgument', 'Set-OrchestratorGuiShortcut')) {
        . ([scriptblock]::Create($installerAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true).Extent.Text))
    }
    $shortcutPath = Join-Path $root 'Start Menu\WorkplaceCloudHub\SmartM365 Orchestrator - demo.lnk'
    $guiPath = Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
    $pwshPath = Get-SmartM365InventoryEnginePath -Runtime PowerShell7
    Set-OrchestratorGuiShortcut -ShortcutPath $shortcutPath -PowerShell7Path $pwshPath -GuiPath $guiPath -TenantName 'demo' -IconPath (Join-Path $folder 'WorkplaceCloudHub.ico') -WorkingDirectory $root
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    Assert-Case ($shortcut.TargetPath -ieq $pwshPath -and $shortcut.WindowStyle -eq 7 -and $shortcut.IconLocation -like '*WorkplaceCloudHub.ico*') 'The GUI shortcut must start PowerShell 7 minimized with the WorkplaceCloudHub icon.'
    Assert-Case ($shortcut.Arguments -ceq ('-NoProfile -NoLogo -WindowStyle Hidden -STA -ExecutionPolicy Bypass -File "{0}" -Tenant demo' -f $guiPath)) "The GUI shortcut arguments must hide the console: $($shortcut.Arguments)"
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
}
finally {
    $env:SMARTM365_ORCHESTRATOR_TENANT = $previousOrchestratorTenant
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: orchestrator GUI launch features, $($script:Cases) offline checks."
