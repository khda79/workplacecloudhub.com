#Requires -Version 7.0
<#
.SYNOPSIS
Offline tests of the Operations launch features of the orchestrator GUI.
.VERSION
1.1.0
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
    $visibleLaunches = @($template.Launches | Where-Object { -not ($_.PSObject.Properties['HideInGui'] -and $_.HideInGui) })
    Assert-Case (@($script:ManualLaunchRows).Count -eq $visibleLaunches.Count -and $visibleLaunches.Count -lt @($template.Launches).Count) 'Every visible manual entry must be listed separately.'
    Assert-Case (@($script:ManualLaunchRows | Where-Object { $_.Name -like 'Orchestrator-*' -or $_.Name -like '*-Installer' }).Count -eq 0) 'Orchestrator entries and installers must stay out of the GUI.'
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCj2M5RX9gdUxR/
# CXn5PRwHrC7LUEfqsEFXf0aE8z1mk6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOcu2zqIv7HZwLrw9WWwAWO5HlmA7uZrdzo8u+sRoqQvMA0GCSqG
# SIb3DQEBAQUABIIBgJ1J/2EzAVwFxp2GOs3jzsPY/mv/M7tf0EKhM7r8aD4GzCqt
# OPxyDjpqelws2mnl1UpFUFV85M/I+uRnwKX4QTGxCrkQYzIOxHLMDesJWaXMYq1g
# rJblrMcsSXihkmRigyDf+hEGPlhSVuMrNyJcNdZ2OXCsDnZqYAvhBtfZxzEH3ec4
# fDm76sq7cC5RJrPZ6WgwjsorMfEPyvnK3mOAi6+YWdEGSQih/S6oCCJlM8dLCtH8
# yIOYDB7KSCHhAn9oLEfZhEFJgfB6LDndx+ZwD4gAJElgn1L5yX9IhXtyud3w1Shw
# RDK2rPapBMXp6RxO526OEXc20dnIsWAE84W8hNB43OvpkdBlpOP0yukCgNwbkUGd
# 9zU+iwiSiCxPY5V+dY25Dpt2u0iuz3YlNX3no/dREMx0IVRS1XRrZWbhqQSpMb9g
# HuLlye5PaOWsszqzmNuaWW2Jefm1zR4OXIMPjjpndZc1PK8YAlYQ0EJKEn7DTVlx
# aTzIA2Ur6HEdJA9UQqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYwOTU2
# MTNaMC8GCSqGSIb3DQEJBDEiBCBKbUCiURkP7ctvC9j78tLrzvZLHaY9VWuIEoWC
# zGirBzANBgkqhkiG9w0BAQEFAASCAgAEd6YIxQPCwNCX+XzaPfrbWP1bOiXdRyaM
# iASKV8ufXiyxW4EBg3atI5IAc9Yoa5KmQ1RpBOhJGJ7DKZ4oFr1BQ3wuDXvqlUEu
# N/aAl14BUKmZnxCDj4+n/i+cwc6s5RgHMJwcczeiYCCdF9JJosRsaPXpU+msyZhL
# fupt9HqWsYZKX3z05OwbMnBoQAD8MvslEzPscS5LUl/kaku3nRyfzAw7fQocRiSW
# mdp2isfKoVaGL5hFe8CRm+Y/To2leUFr46JSBblOhR8RAXFDWN1q3pblw8pk6xYp
# Fb50dLXjwgxI8hUbwU72+5GArkCW11SORWX2JcHUN+HLScKX44UEiI9q+/cjFH5g
# P4i6Mju7uDJfyQzAdlJU+nSN4+Gk8XDBY3DQvv8nrnsMHW4yrUQ6Qgar903NIS25
# ijMwr5ZxgSrww/HK4yGjgcFRnjcD1q/ey25UghMXPo1exDio25EiKyBNO8C3tfFn
# KvgGcIt3ix+ghLh8tOzTYLJIpX18W30JwCOVkY6OiiIrsbzpVZDkEjYlBePqzqB2
# rVqmch/vZWw5NXUnpmU066K/EmqocuE9SDwyopib9Au88rAZs1jEFy1u+VMxKAsk
# wi11gYFgEF78pGQ4tF0CHS0lBCxSCkB9BJasvWLVDRXBpQ33nzTQnpdbbDab5d+E
# babltiDYUg==
# SIG # End signature block
