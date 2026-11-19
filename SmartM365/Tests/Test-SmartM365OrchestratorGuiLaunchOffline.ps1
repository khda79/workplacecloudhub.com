#Requires -Version 7.0
<#
.SYNOPSIS
Offline tests of the Operations launch features of the orchestrator GUI.
.VERSION
1.1.8
.NOTES
GUI functions are extracted by AST and bound to the real XAML. The catalog is the committed
template (read only); the published configuration is a copy of its jobs. No orchestrator, request
or collector is started: the request submission is mocked and the only process started is a
fictional script in a hidden window.
Checks: job and manual lists, resolved command preview, confirmation and interactive flags,
execution mode of on-premises jobs in the draft (set, reset, apply to all, validation), orchestrator
actions identical to the former Launchers\Orchestrator .cmd files, request with the selected job,
console start with quoted arguments, Run now with the orchestrator service account (runas command,
account read from the scheduled task, managed account refused; runas itself is never started).
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
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.ReportCadence.psm1') -Force
$script:ReportMailProducersByPath = Get-SmartM365OrchestratorReportMailProducers -InventoryRoot $inventoryRoot
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'GUI parsing failed.' }
foreach ($name in @('ConvertFrom-OrchestratorGuiXaml', 'Get-OrchestratorGuiPropertyValue', 'Set-JsonProperty', 'Copy-JsonDocument', 'Get-ComboText', 'Select-ComboText', 'Get-GuiLaunchCatalog', 'Get-GuiJobScriptVersion', 'Get-GuiJobReportMailProducer', 'Get-GuiSwiRow', 'Refresh-LaunchView',
        'Get-GuiLaunchPlanText', 'Update-SelectedLaunchPreview', 'Update-GuiPrepareRequestButton', 'Start-GuiConsoleProcess', 'Get-GuiOrchestratorActionCommand', 'Set-GuiJobExecutionMode', 'Request-SelectedJobRun',
        'Connect-OrchestratorJobReportReplay', 'Invoke-GuiPreparedReportReplay', 'New-GuiWaitDialog', 'Close-GuiWaitDialog', 'Start-GuiSendSummary', 'Receive-GuiSendSummary', 'Get-GuiOrchestratorServiceAccount', 'Get-GuiRunAsArgumentList')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Missing GUI function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
$assignment = $ast.Find({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml' }, $true)
$window = ConvertFrom-OrchestratorGuiXaml -Text $assignment.Right.Expression.Value
$script:Controls = @{}
foreach ($name in @('OperationsTabs', 'OnPremModeText', 'JobExecutionModeCombo', 'ApplyJobExecutionModeButton', 'ApplyModeToAllOnPremButton', 'LaunchJobsGrid', 'RunJobNowButton', 'RunJobAsServiceAccountButton', 'RequestJobButton', 'RequestPrepareButton',
        'LaunchIncludeDependenciesCheck', 'LaunchJobCommandBox', 'ManualLaunchesGrid', 'RunManualLaunchButton', 'ManualLaunchCommandBox', 'OrchestratorStartButton', 'OrchestratorStopButton',
        'OrchestratorRestartButton', 'FullCollectButton', 'SendSummaryButton', 'StatusText', 'PlanningGrid', 'PlanningReportReplayButton', 'OperationsReportReplayButton', 'IncludeDependenciesCheck', 'OperationsRunningGrid')) {
    $script:Controls[$name] = $window.FindName($name)
    if (-not $script:Controls[$name]) { throw "Missing XAML control: $name" }
}

$script:Cases = 0
$script:RequestSubmitBusy = $false; $script:CancellationBusy = $false; $script:ScriptVersions = @{}
$script:HealthByName = @{}; $script:SwiByName = @{}; $script:SwiValidityError = ''
$prepareName = $ast.Find({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$script:PrepareJobName' }, $true)
$script:PrepareJobName = $prepareName.Right.Expression.Value
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
    Assert-Case ($script:PrepareJobName -eq 'WorkplaceEvidence-Prepare' -and $script:Controls.RequestPrepareButton.IsEnabled) 'Request Prepare is enabled when Prepare is in the published configuration.'
    $script:RequestSubmitBusy = $true; Update-GuiPrepareRequestButton
    Assert-Case (-not $script:Controls.RequestPrepareButton.IsEnabled) 'Request Prepare is disabled while a request is submitted.'
    $script:RequestSubmitBusy = $false; Update-GuiPrepareRequestButton
    $visibleLaunches = @($template.Launches | Where-Object { -not ($_.PSObject.Properties['HideInGui'] -and $_.HideInGui) })
    Assert-Case (@($script:ManualLaunchRows).Count -eq $visibleLaunches.Count -and $visibleLaunches.Count -lt @($template.Launches).Count) 'Every visible manual entry must be listed separately.'
    # The server acceptance test is the one orchestrator entry also listed in Manual launches (GUI 1.6.32).
    Assert-Case (@($script:ManualLaunchRows | Where-Object { ($_.Name -like 'Orchestrator-*' -and $_.Name -ne 'Orchestrator-ServerAcceptance') -or $_.Name -like '*-Installer' }).Count -eq 0) 'Orchestrator entries and installers must stay out of the GUI.'
    Assert-Case (@($script:ManualLaunchRows | Where-Object { $_.Name -eq 'Orchestrator-ServerAcceptance' }).Count -eq 1) 'The server acceptance test is listed in Manual launches.'
    Assert-Case ($script:Controls.OnPremModeText.Text -in @('LocalCache', 'Direct', 'Auto')) 'The global on-premises mode must be shown.'
    $mailboxes = @($script:LaunchJobRows | Where-Object Name -eq 'ExchangeOnPrem-Local-Mailboxes-Inventory')[0]
    Assert-Case ($mailboxes.OnPremises -and $mailboxes.Runtime -eq 'WindowsPowerShell' -and $mailboxes.DraftMode -eq 'Auto' -and $mailboxes.ModeSource -like 'Job*, zone *' -and $mailboxes.Mode -like 'Auto (*)') 'An on-premises job must show its runtime, the Auto mode of the template with its resolved mode, and its source.'
    $cloud = @($script:LaunchJobRows | Where-Object Name -eq 'M365-ActiveUsers-Inventory')[0]
    Assert-Case (-not $cloud.OnPremises -and $cloud.Mode -eq 'Direct' -and $cloud.DraftMode -eq '') 'A cloud job runs Direct and has no mode to set.'

    Assert-Case (@($script:LaunchJobRows | Where-Object ReportMailProducer).Count -eq 12) 'The published job list must expose every supported report producer once.'
    Assert-Case ($mailboxes.ReportMailProducer -eq 'SmartM365-Exchange-Local-Mailboxes-Inventory' -and -not $cloud.ReportMailProducer) 'Report eligibility must come from the known producer path.'
    $draftMailbox = $script:DraftJobs.Jobs | Where-Object Name -EQ $mailboxes.Name
    $originalPath = $draftMailbox.ScriptPath
    $draftMailbox.ScriptPath = ($script:DraftJobs.Jobs | Where-Object Name -EQ $cloud.Name).ScriptPath
    Refresh-LaunchView
    Assert-Case (($script:LaunchJobRows | Where-Object Name -EQ $mailboxes.Name).ReportMailProducer -eq $mailboxes.ReportMailProducer) 'An unpublished change must not reroute report mail.'
    $draftMailbox.ScriptPath = $originalPath
    Refresh-LaunchView
    $mailboxes = $script:LaunchJobRows | Where-Object Name -EQ $mailboxes.Name
    $cloud = $script:LaunchJobRows | Where-Object Name -EQ $cloud.Name

    $script:Controls.LaunchJobsGrid.SelectedItem = $cloud
    Update-SelectedLaunchPreview -Kind Job
    Assert-Case ($script:Controls.RunJobNowButton.IsEnabled -and $script:Controls.RequestJobButton.IsEnabled -and -not $script:Controls.JobExecutionModeCombo.IsEnabled) 'Cloud job: run and request enabled, mode disabled.'
    Assert-Case ($script:Controls.RunJobAsServiceAccountButton.IsEnabled) 'Run now with service account follows the selection.'
    # Service account: read from the orchestrator scheduled task (fictional tasks, nothing queried or started).
    $fakeTask = { param($Name, $User) [pscustomobject]@{ TaskName = $Name; Principal = [pscustomobject]@{ UserId = $User } } }
    $account = Get-GuiOrchestratorServiceAccount -Tasks @((& $fakeTask 'SmartM365 Inventory Orchestrator - other' 'CONTOSO\svc_other'), (& $fakeTask 'SmartM365 Inventory Orchestrator - demo' 'CONTOSO\svc_inventory'))
    Assert-Case ($account.Account -eq 'CONTOSO\svc_inventory' -and $account.TaskName -eq 'SmartM365 Inventory Orchestrator - demo') 'The task of the GUI tenant gives the service account.'
    $account = Get-GuiOrchestratorServiceAccount -Tasks @((& $fakeTask 'SmartM365 Inventory Orchestrator - demo' 'svc_inventory'))
    Assert-Case ($account.Account -eq ('{0}\svc_inventory' -f $env:USERDOMAIN)) 'An account without domain gets the user domain.'
    foreach ($bad in @(@(), @((& $fakeTask 'SmartM365 Inventory Orchestrator - demo' 'CONTOSO\gmsa_inventory$')))) {
        $refused = $false; try { [void](Get-GuiOrchestratorServiceAccount -Tasks $bad) } catch { $refused = $true }
        Assert-Case $refused 'No task or a managed account (no password) is refused with an error.'
    }
    $runAs = Get-GuiRunAsArgumentList -Account 'CONTOSO\svc_inventory' -Engine 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Arguments @('-NoProfile', '-File', 'C:\Smart Inventory\Invoke-SmartM365InventoryJob.ps1', 'Job-A', '-Now', '-Tenant', 'demo')
    Assert-Case ($runAs -ceq '/user:CONTOSO\svc_inventory "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -File \"C:\Smart Inventory\Invoke-SmartM365InventoryJob.ps1\" Job-A -Now -Tenant demo"') "runas gets the whole command as one argument with escaped inner quotes: $runAs"
    Assert-Case ($script:Controls.LaunchJobCommandBox.Text -match 'Tenant: demo \(source: Parameter\)' -and $script:Controls.LaunchJobCommandBox.Text -match 'SmartM365-ActiveUsers-Inventory\.ps1 -Tenant demo -Connect') 'The resolved command of the job must be previewed.'
    $script:Controls.LaunchJobsGrid.SelectedItem = $mailboxes
    Update-SelectedLaunchPreview -Kind Job
    Assert-Case ($script:Controls.JobExecutionModeCombo.IsEnabled -and (Get-ComboText -Combo $script:Controls.JobExecutionModeCombo) -eq 'Auto') 'On-premises job: the mode can be set.'

    $write = @($script:ManualLaunchRows | Where-Object Name -eq 'ExchangeOnPrem-ProxyAddresses-Write')[0]
    Assert-Case ($write.RequiresConfirmation -and $write.Interactive -and $write.Runtime -eq 'WindowsPowerShell') 'The write entry must be flagged for confirmation.'
    $script:Controls.ManualLaunchesGrid.SelectedItem = $write
    Update-SelectedLaunchPreview -Kind Manual
    Assert-Case ($script:Controls.RunManualLaunchButton.IsEnabled -and $script:Controls.ManualLaunchCommandBox.Text -match 'Confirmation: the user must type WRITE' -and $script:Controls.ManualLaunchCommandBox.Text -match '-AllOrganizationalUnit -AddMissingAddress') 'The manual preview must show the command and the confirmation.'

    # Execution mode in the draft: set, reset, apply to all, validation.
    [void](Set-GuiJobExecutionMode -JobName 'ExchangeOnPrem-Local-Mailboxes-Inventory' -Mode Direct)
    $draftJob = @($script:DraftJobs.Jobs | Where-Object Name -eq 'ExchangeOnPrem-Local-Mailboxes-Inventory')[0]
    Assert-Case ($draftJob.Execution.Mode -eq 'Direct' -and @($script:LaunchJobRows | Where-Object Name -eq 'ExchangeOnPrem-Local-Mailboxes-Inventory')[0].DraftMode -eq 'Direct') 'The mode must be written to the draft and shown.'
    Assert-Case ((Test-SmartM365OrchestratorJobsDocument -Document $script:DraftJobs).Valid) 'A draft with an execution mode must stay valid.'
    Assert-Case ($draftJob.Execution.PSObject.Properties['CacheName'] -and $draftJob.Execution.CacheName -eq 'ExchangeLocalMailboxes\WithADPermissions') 'Other Execution fields of the draft job must be kept.'
    [void](Set-GuiJobExecutionMode -JobName 'ExchangeOnPrem-Local-Mailboxes-Inventory' -Mode '(default)')
    Assert-Case (-not $draftJob.Execution.PSObject.Properties['Mode']) '(default) must remove the job mode.'
    $onPremNames = @($script:LaunchJobRows | Where-Object { $_.OnPremises -and $_.InManifest } | ForEach-Object Name)
    $changed = @(Set-GuiJobExecutionMode -JobName $onPremNames -Mode Auto)
    Assert-Case ($changed.Count -eq $onPremNames.Count -and $changed.Count -ge 7 -and @($script:LaunchJobRows | Where-Object { $_.OnPremises -and $_.DraftMode -ne 'Auto' }).Count -eq 0) 'Apply to all must set every on-premises job.'
    $bad = Copy-JsonDocument -Document $script:DraftJobs
    $bad.Jobs[0] | Add-Member -NotePropertyName Execution -NotePropertyValue ([pscustomobject]@{ Mode = 'Remote' }) -Force
    $badValidation = Test-SmartM365OrchestratorJobsDocument -Document $bad
    Assert-Case (-not $badValidation.Valid -and (@($badValidation.Errors) -join ' ') -match 'Execution.Mode must be Direct, LocalCache or Auto') 'An invalid execution mode must be rejected before publication.'
    [void](Set-GuiJobExecutionMode -JobName 'ExchangeOnPrem-Local-Mailboxes-Inventory' -Mode LocalCache)
    $previewOfPublished = (Get-GuiLaunchPlanText -Name 'ExchangeOnPrem-Local-Mailboxes-Inventory').Plan
    Assert-Case ($previewOfPublished.ModeInfo.Requested -eq 'Auto') 'Previews follow the published configuration, not the unpublished draft.'

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
    function New-SmartM365OrchestratorJobRunRequest { param($SharedDataFolderPath, $JobsPath, $JobName, $Tenant, [switch]$IncludeDependencies) $script:RequestCalls.Add([pscustomobject]@{ Jobs = @($JobName); Tenant = $Tenant; Dependencies = [bool]$IncludeDependencies }); [pscustomobject]@{ BatchId = 'Batch1'; TotalCount = 1; QueuePosition = 0 } }
    $script:SharedDataFolderPath = $root
    $script:Controls.IncludeDependenciesCheck.IsChecked = $false
    [void](Request-SelectedJobRun -SkipConfirmation -JobName 'AD-Inventory' -IncludeDependencies $true)
    Assert-Case ($script:RequestCalls.Count -eq 1 -and $script:RequestCalls[0].Jobs[0] -eq 'AD-Inventory' -and $script:RequestCalls[0].Dependencies -and $script:RequestCalls[0].Tenant -eq 'demo') 'Request must submit the selected job with its dependency choice.'
    $prepareClick = $ast.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression.Extent.Text -eq '$script:Controls.RequestPrepareButton' -and $node.Member.Extent.Text -eq 'Add_Click' }, $true)
    Assert-Case ($null -ne $prepareClick -and $prepareClick.Extent.Text -match 'Request-SelectedJobRun -JobName \$script:PrepareJobName -IncludeDependencies \$false') 'Request Prepare must request Prepare alone, without its dependencies.'

    # Console start: arguments with spaces reach the script (hidden window, fictional script).
    $probe = Join-Path $root 'Probe Folder\SmartM365-Fixture-Probe.ps1'
    New-Item -ItemType Directory -Path (Split-Path $probe -Parent) | Out-Null
    [IO.File]::WriteAllText($probe, "param([string]`$Value, [string]`$Marker) [IO.File]::WriteAllText(`$Marker, (`$Value + '|' + `$PSVersionTable.PSEdition))", [Text.UTF8Encoding]::new($false))
    $marker = Join-Path $root 'console marker.txt'
    $process = Start-GuiConsoleProcess -Runtime WindowsPowerShell -ScriptPath $probe -Arguments @('-Value', 'two words "quoted"', '-Marker', $marker) -WindowStyle Hidden
    if (-not $process.WaitForExit(60000)) { throw 'Console probe did not exit.' }
    Assert-Case ((Get-Content -LiteralPath $marker -Raw) -eq 'two words "quoted"|Desktop') 'The console must receive the exact arguments with the requested runtime.'

    # Send summary: hidden process, wait window, result from the exit code (fictional script).
    $script:Results = [Collections.Generic.List[object]]::new()
    function Show-GuiJobRunRequestResult { param($Message, $Title, $Icon) $script:Results.Add([pscustomobject]@{ Message = $Message; Title = $Title; Icon = $Icon }) }
    $script:SendSummaryRun = $null; $script:SendSummaryTimeoutMinutes = 10
    $summaryAction = Get-GuiOrchestratorActionCommand -Action SendSummary
    Assert-Case ($summaryAction.Arguments -contains '-SendExecutionSummary') 'Send summary runs the orchestrator with -SendExecutionSummary.'
    foreach ($case in @(@{ Code = 0; Icon = 'Information' }, @{ Code = 3; Icon = 'Error' })) {
        $fixture = Join-Path $root ('Send summary {0}.ps1' -f $case.Code)
        [IO.File]::WriteAllText($fixture, ('Start-Sleep -Milliseconds 300; exit {0}' -f $case.Code), [Text.UTF8Encoding]::new($false))
        Start-GuiSendSummary -Command ([pscustomobject]@{ ScriptPath = $fixture; Arguments = @() })
        $run = $script:SendSummaryRun
        Assert-Case ($null -ne $run -and $run.Dialog.IsVisible -and -not $script:Controls.SendSummaryButton.IsEnabled) 'Send summary shows the wait window and disables its button.'
        Assert-Case ($run.Dialog.FindName('WaitTitleText').Text -eq 'Send summary') 'The wait window is titled Send summary.'
        Receive-GuiSendSummary
        Assert-Case ($null -ne $script:SendSummaryRun -and $script:Results.Count -eq 0) 'The wait window stays while the process runs.'
        if (-not $run.Process.WaitForExit(60000)) { throw 'Send summary fixture did not exit.' }
        Receive-GuiSendSummary
        $result = $script:Results[$script:Results.Count - 1]
        Assert-Case ($null -eq $script:SendSummaryRun -and -not $run.Dialog.IsVisible -and $script:Controls.SendSummaryButton.IsEnabled) 'The wait window closes and the button returns when the process ends.'
        Assert-Case ($result.Icon -eq $case.Icon -and $result.Title -eq 'Send summary') "Exit code $($case.Code) gives a $($case.Icon) result."
        $script:Results.Clear()
    }

    # Exercise the real confirmation path from both job menus, with a recorded dialog and no worker process.
    $script:OrchestratorDialog = [pscustomobject]@{ Answer = 'Yes'; Message = ''; Title = ''; Buttons = ''; Icon = '' }
    $script:OrchestratorDialog | Add-Member -MemberType ScriptMethod -Name Show -Value {
        param($Message,$Title,$Buttons,$Icon)
        $this.Message = $Message
        $this.Title = $Title
        $this.Buttons = $Buttons
        $this.Icon = $Icon
        return $this.Answer
    }
    $script:ReportWorkerCalls = [Collections.Generic.List[object]]::new()
    $originalStart = (Get-Command Start-GuiSendSummary).ScriptBlock
    Set-Item -Path Function:script:Start-GuiSendSummary -Value {
        param($Command,[switch]$PreparedReport)
        $script:ReportWorkerCalls.Add([pscustomobject]@{ Command = $Command; PreparedReport = [bool]$PreparedReport })
    }
    try {
        $invokeReportMail = Get-Command Invoke-GuiPreparedReportReplay -ErrorAction Stop
        Connect-OrchestratorJobReportReplay -Controls $script:Controls -Replay {
            param($producer,$audience)
            & $invokeReportMail -Producer $producer -Audience $audience
        }.GetNewClosure()
        $ad = $script:LaunchJobRows | Where-Object Name -EQ 'AD-Inventory'
        $script:Controls.PlanningGrid.ItemsSource = @($ad)
        foreach ($pair in @(
            @($script:Controls.PlanningGrid,$script:Controls.PlanningReportReplayButton),
            @($script:Controls.LaunchJobsGrid,$script:Controls.OperationsReportReplayButton)
        )) {
            $grid,$button = $pair
            $grid.SelectedItem = $ad
            foreach ($item in $button.ContextMenu.Items) {
                $before = $script:ReportWorkerCalls.Count
                $item.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.MenuItem]::ClickEvent))
                $audience = [string]$item.Tag
                $recipient = if ($audience -eq 'Administrators') { 'Smart365Administrators only' } else { 'the configured recipients' }
                $dialog = $script:OrchestratorDialog
                Assert-Case ($dialog.Message -ceq ('Rebuild the SmartM365-ActiveDirectory-Inventory report from the current published exports and send it to {0}? No inventory will run.' -f $recipient) -and
                    $dialog.Title -eq 'Send mail only' -and $dialog.Buttons -eq 'YesNo' -and $dialog.Icon -eq 'Question') 'Report confirmation did not identify the selected producer and audience.'
                Assert-Case ($script:ReportWorkerCalls.Count -eq ($before + 1)) 'Confirming report delivery must start one mail-only worker.'
                $call = $script:ReportWorkerCalls[$before]
                Assert-Case ($call.PreparedReport -and $call.Command.ScriptPath -eq (Join-Path $inventoryRoot 'Orchestrator/SmartM365-SendPreparedReportMail.ps1') -and
                    ($call.Command.Arguments -join ' ') -ceq "-Tenant demo -Producer SmartM365-ActiveDirectory-Inventory -Audience $audience") 'Confirmed delivery must call the report worker with the expected tenant, producer and audience.'
            }
        }
        $script:OrchestratorDialog.Answer = 'No'
        $script:Controls.OperationsReportReplayButton.ContextMenu.Items[0].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.MenuItem]::ClickEvent))
        Assert-Case ($script:ReportWorkerCalls.Count -eq 4) 'Declining report delivery must not start a worker.'
    } finally {
        Set-Item -Path Function:script:Start-GuiSendSummary -Value $originalStart
    }

    # GUI shortcut of the scheduled task installer: hidden PowerShell 7, minimized, no console window.
    $installerAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'), [ref]$null, [ref]$null)
    foreach ($name in @('ConvertTo-QuotedArgument', 'Set-OrchestratorGuiShortcut')) {
        . ([scriptblock]::Create($installerAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true).Extent.Text))
    }
    Assert-Case ($installerAst.Extent.Text -match "GetFolderPath\('CommonDesktopDirectory'\)" -and $installerAst.Extent.Text -match 'foreach \(\$guiShortcutPath in \$guiShortcutPaths\)') 'The installer must create and remove the GUI shortcut on the all-users desktop as well as in the Start menu.'
    $shortcutPath = Join-Path $root 'Start Menu\WorkplaceCloudHub\SmartM365 Orchestrator - demo.lnk'
    $guiPath = Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
    $pwshPath = Get-SmartM365InventoryEnginePath -Runtime PowerShell7
    Set-OrchestratorGuiShortcut -ShortcutPath $shortcutPath -PowerShell7Path $pwshPath -GuiPath $guiPath -TenantName 'demo' -IconPath (Join-Path $folder 'WorkplaceCloudHub.ico') -WorkingDirectory $root -LocalIconDirectory (Join-Path $root 'ProgramData\WorkplaceCloudHub\SmartM365')
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $localIconPath = Join-Path $root 'ProgramData\WorkplaceCloudHub\SmartM365\WorkplaceCloudHub.ico'
    Assert-Case ($shortcut.TargetPath -ieq $pwshPath -and $shortcut.WindowStyle -eq 7) 'The GUI shortcut must start PowerShell 7 minimized.'
    Assert-Case ((Test-Path -LiteralPath $localIconPath -PathType Leaf) -and $shortcut.IconLocation -ieq ('{0},0' -f $localIconPath)) "The GUI shortcut icon must be the local WorkplaceCloudHub.ico copy, not the code share: $($shortcut.IconLocation)"
    Assert-Case ($shortcut.Arguments -ceq ('-NoProfile -NoLogo -WindowStyle Hidden -STA -ExecutionPolicy Bypass -File "{0}" -Tenant demo' -f $guiPath)) "The GUI shortcut arguments must hide the console: $($shortcut.Arguments)"
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
}
finally {
    $env:SMARTM365_ORCHESTRATOR_TENANT = $previousOrchestratorTenant
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: orchestrator GUI launch features, $($script:Cases) offline checks."

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB3wPYN1yN4C64f
# nF6gu1q/z6nCRqxU3KYeJNNt2ShqZaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgppxgVgmPPXie
# 0p6L6dMbEo8joCz60XGZM7KwSxqhIHEwDQYJKoZIhvcNAQEBBQAEggGAUL+hueDM
# KugmgBVjhdklpuSkCXWBQSENT3045G2AB8O64Wir5CkegS8kX8wnSyOsn3rLPXK6
# 4uqkp/MZm6PUfhzrKuuuqcdEfmCmQW5Pq9juTrsEyHNP+pSi7qkgSZoTPE2bzhPJ
# mB9U76IaRKmitM3Tc1HQgGrK8qL3oRPilpEWfW2HzOlETq0ObKCB20mv/pdhTp6V
# dq34wPMMtCx7QJyzMMNoZ85oPNnISFuRMHIhCi0ZgfKYKeepTDT4FiZlgHmYCCWc
# nndBQklb94+p16AmNE58k5u+vTsT5tVpPLbFsOykxAzJDQMWR9fjm+lITbbVVreU
# YDDnhg0VsFYmJ+cxjQ5+GTXSQE2evoY8IeXftgojcV4Z2lieJG+3HpliiFh2HtGh
# mNfvxoxcuIHP2Ux8PxD61sysNKOwXLZdy/z9MqMbBhDH9K9MhGIQWPwQOwN1kZqZ
# 4X4DzxXJU4y/6cgR0w32qgIbE6Mn6Yo4kfUevgl1LMTLxVrsMQ6AWCspoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MTAwNjE0WjA/BgkqhkiG9w0BCQQxMgQw0mIj/4AoXalslsX43ju/DLyP
# 6oLM+DLyW7HbdRi8dyg2J+gvsHrB4ogPB0yI2SifMA0GCSqGSIb3DQEBAQUABIIC
# AAFNQGSqhKflLkPjO4SAPoVYmwax/DGC7lHNdY8eAW3jtVNPW4FC9I+bxs2+Vmyp
# BeqoIxfGIcq7fwIdX3mi77EDzK87PB7ljuOxCY4tZt1H6AwhcXzLDn/UHqFoo8l0
# JdDQ8/S+CeIVF837bSDcf90xpCHa02IIny/0Xx82N0UFnWBv9Rg2i75YR/HxzjHZ
# nggJhLKMXx6ooiiaD/kWDXoyRz2cDaDfGoUM81MifT6piK8xlJnArYwFMTcv/OVz
# 78AXrsT0Vgdq9WmR5MW8IWfdhJqlgiLJAWUMrB+T5dQxOzQ+WtX2iDuoiWdIwoj9
# lQu/zBMMHBkHuRaTkmvtn+ZlqFbimkoBR6TaIi0q0XKsXZGUdYrr4SBzu/HOaFRg
# j2eJpOh50ngzQ32wLYYIdnzY5e7DEEvneiDnzabHseWHBBBrOsO54Zd/5+jarMym
# MlAMfPt3PeF2BQRiqTj4sQrNSSeImc20LRqsZv6ksBEWJHgZXDeSGmbklSFgd63M
# xshR6VVp3j1/Yi8T/cbB27krdBYDceWD/lYUz2yOccDBrZA9UzfyyHdUd4AjpN8V
# R66RE8HzxGm4VWQYmP0vyn39+xLuFhkBIJFBdTbODuHOZsbMdM9aDeCvO0CoKZyo
# A+Vlucp4OQ6c33X6cfVKPa0aCLso/uMTHFmCOJQ/Y3L6
# SIG # End signature block
