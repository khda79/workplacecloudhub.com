<#
.SYNOPSIS
    Runs a Smart Endpoint Migration Toolkit Wave (or a Single PC) for one Campaign through PsExec.

.DESCRIPTION
    Technician-side engine of the Smart Endpoint Migration Toolkit (Hybrid Join, Intune
    Enrollment and Windows 11 Upgrade).

    Campaign  : HybridJoin or Windows11Upgrade.
    Wave      : group of computers, Waves\<Campaign>\<Wave>\Computers.txt.
    Run       : one execution of a Wave, Runs\<Campaign>\<Wave>\<RunId>.
    Cycle     : one iteration inside a Run (Loop mode repeats cycles).

    Options are resolved from Config\SmartEndpointMigrationToolkit.options.json.txt with this
    precedence: explicit CLI (-Option, -DryRun, -AuditOnly, -IgnoreRunGuard) > GUI choices
    captured for the Run (-GuiSelectionPath) > SEMT_* environment variables > Wave.local.json >
    Config\SmartEndpointMigrationToolkit.local.json > campaign default. The effective options
    and their source are written to the Run folder and shown in the reports.

    Dry run, remote audit and repair are distinct:
    - dry run (-DryRun): DNS, ping, SMB 445 and administrative shares only; nothing is copied or
      executed and no inventory is refreshed;
    - remote audit (-AuditOnly): the endpoint script runs diagnostics only;
    - repair: the endpoint applies the actions allowed by the campaign options.

.PARAMETER Campaign
    HybridJoin or Windows11Upgrade.

.PARAMETER WavePath
    Wave folder under Waves\<Campaign>. Mutually exclusive with -ComputerName.

.PARAMETER ComputerName
    Single PC: creates an ephemeral one-computer Wave inside the Run folder.

.PARAMETER RunMode
    Loop (default) repeats cycles; Once runs a single cycle.

.PARAMETER Option
    Explicit option values, for example @{ ThrottleLimit = 5; 'Windows11Upgrade.AllowReboot' = $false }.

.PARAMETER GuiSelectionPath
    JSON file of the GUI choices captured for this Run.

.PARAMETER ValidateOnly
    Resolves options, rules and component integrity, writes nothing on targets and exits.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $true)][ValidateSet('HybridJoin','Windows11Upgrade')][string]$Campaign,
    [string]$WavePath,
    [string]$ComputerName,
    [ValidateSet('Loop','Once')][string]$RunMode = 'Loop',
    [hashtable]$Option = @{},
    [string]$GuiSelectionPath,
    [string]$PsExecPath,
    [switch]$DryRun,
    [switch]$AuditOnly,
    [switch]$IgnoreRunGuard,
    [switch]$ValidateOnly,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$UnexpectedArguments
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:LauncherVersion = '1.0.0'

if ($UnexpectedArguments -and $UnexpectedArguments.Count -gt 0) {
    throw ("Unexpected argument(s): {0}. Pass options with -Option @{{ Name = Value }} and PsExec with -PsExecPath." -f ($UnexpectedArguments -join ' '))
}
if ([string]::IsNullOrWhiteSpace($WavePath) -eq [string]::IsNullOrWhiteSpace($ComputerName)) {
    throw 'Specify exactly one of -WavePath or -ComputerName.'
}

$script:PackageRoot = Split-Path -Path $PSScriptRoot -Parent
Import-Module (Join-Path -Path $script:PackageRoot -ChildPath 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
Import-Module (Join-Path -Path $script:PackageRoot -ChildPath ("Campaigns\{0}\SmartEndpointMigration.Campaign.{0}.psd1" -f $Campaign)) -MinimumVersion 1.0.0 -Force
$script:CampaignProfile = & ("Get-Semt{0}Profile" -f $Campaign)
$script:Product = Get-SemtProductInfo
$script:RunId = New-SemtRunId

if (-not [string]::IsNullOrWhiteSpace($ComputerName)) {
    $ephemeral = New-SemtEphemeralWave -PackageRoot $script:PackageRoot -Campaign $Campaign -ComputerName $ComputerName -RunId $script:RunId
    $script:RunContext = New-SemtRunContext -PackageRoot $script:PackageRoot -Campaign $Campaign -WavePath $ephemeral.WavePath -RunId $script:RunId -SinglePc -RunRoot $ephemeral.RunRoot
}
else {
    $waveRoot = Get-SemtWaveRoot -PackageRoot $script:PackageRoot -Campaign $Campaign
    $waveFull = [System.IO.Path]::GetFullPath($WavePath).TrimEnd('\')
    if ((Split-Path -Path $waveFull -Parent) -ine [System.IO.Path]::GetFullPath($waveRoot).TrimEnd('\')) {
        throw ("SEMT_WAVE_PATH_INVALID: a {0} Wave must be a direct child of {1}. Received {2}." -f $Campaign,$waveRoot,$waveFull)
    }
    if (-not (Test-SemtWaveName -Name (Split-Path -Path $waveFull -Leaf))) { throw ("SEMT_WAVE_NAME_INVALID: '{0}' is not a valid Wave name." -f (Split-Path -Path $waveFull -Leaf)) }
    if (-not (Test-Path -LiteralPath (Join-Path $waveFull 'Computers.txt') -PathType Leaf)) { throw "SEMT_WAVE_EMPTY: Computers.txt not found in $waveFull" }
    $script:RunContext = New-SemtRunContext -PackageRoot $script:PackageRoot -Campaign $Campaign -WavePath $waveFull -RunId $script:RunId
}
$script:LogPaths = @($script:RunContext.LauncherLogPath)

function Write-Host {
    param(
        [Parameter(Position = 0)][object]$Object = '',
        [switch]$NoNewline,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )

    $text = [string]$Object
    foreach ($line in @($text -split "`r?`n")) {
        $prefix = if ([string]::IsNullOrEmpty($line)) { '' } else { '[{0}] ' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
        $parameters = @{ Object = ($prefix + $line) }
        if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters['ForegroundColor'] = $ForegroundColor }
        if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $parameters['BackgroundColor'] = $BackgroundColor }
        if ($PSBoundParameters.ContainsKey('NoNewline')) { $parameters['NoNewline'] = $NoNewline }
        Microsoft.PowerShell.Utility\Write-Host @parameters
        if (-not [string]::IsNullOrEmpty($line)) { Write-SemtLogLine -Path $script:LogPaths -Message $line }
    }
}

$script:Log = {
    param([string]$Message, [string]$Color = 'Gray')
    Write-Host $Message -ForegroundColor ([System.ConsoleColor]$Color)
}

#region Options and rules

$catalog = Import-SemtOptionCatalog -PackageRoot $script:PackageRoot
$globalConfig = Read-SemtGlobalConfig -PackageRoot $script:PackageRoot -CreateWhenMissing
if (@($globalConfig.AddedKeys).Count -gt 0) { Write-Host ("Global configuration completed from template: {0}" -f (@($globalConfig.AddedKeys) -join ', ')) -ForegroundColor DarkGray }
$waveLayer = $null
if (-not $script:RunContext.SinglePc) {
    $waveConfig = Read-SemtWaveConfig -PackageRoot $script:PackageRoot -Campaign $Campaign -WavePath $script:RunContext.WavePath -CreateWhenMissing
    $waveCampaign = [string](Get-SemtPropertyValue -InputObject $waveConfig.Data -Name 'Campaign' -Default $Campaign)
    if ($waveCampaign -ne $Campaign) { throw ("SEMT_WAVE_CAMPAIGN_MISMATCH: Wave.local.json declares Campaign '{0}', not '{1}'." -f $waveCampaign,$Campaign) }
    $waveLayer = $waveConfig.Data
}
$cliLayer = @{}
foreach ($key in @($Option.Keys)) { $cliLayer[[string]$key] = $Option[$key] }
if ($DryRun) { $cliLayer['DryRun'] = $true }
if ($AuditOnly) { $cliLayer['AuditOnly'] = $true }
if ($IgnoreRunGuard) { $cliLayer['IgnoreEndpointRunGuard'] = $true }
$guiSelection = Read-SemtGuiSelection -Path $GuiSelectionPath -Campaign $Campaign
$script:Options = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $Campaign -GlobalLayer $globalConfig.Data -WaveLayer $waveLayer -EnvironmentLayer (Get-SemtEnvironmentOptionLayer -Catalog $catalog -Campaign $Campaign) -GuiLayer $guiSelection -CliLayer $cliLayer
if ($script:RunContext.SinglePc -and (Get-SemtOption -Options $script:Options -Name 'ThrottleLimit') -ne 1) { Set-SemtRuleOption -Options $script:Options -Name 'ThrottleLimit' -Value 1 -Rule 'SinglePcOneWorker' }
if ($RunMode -eq 'Once') { Set-SemtRuleOption -Options $script:Options -Name 'MaxCycles' -Value 1 -Rule 'RunModeOnce' }
$script:OptionWarnings = @(& $script:CampaignProfile.Hooks.ResolveOptionRules $script:Options)
function Opt([string]$Name) { return (Get-SemtOption -Options $script:Options -Name $Name) }

#endregion

#region Integrity, identity and manifest

$componentPaths = @(
    'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1',
    'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psm1'
) + @(Get-ChildItem -LiteralPath (Join-Path $script:PackageRoot 'Modules\SmartEndpointMigration.Core\Functions') -Filter '*.ps1' -File | ForEach-Object { 'Modules\SmartEndpointMigration.Core\Functions\{0}' -f $_.Name }) + @(
    ("Campaigns\{0}\SmartEndpointMigration.Campaign.{0}.psd1" -f $Campaign),
    ("Campaigns\{0}\SmartEndpointMigration.Campaign.{0}.psm1" -f $Campaign),
    'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1',
    'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1',
    'Config\SmartEndpointMigrationToolkit.options.json.txt',
    $script:CampaignProfile.EndpointScriptRelativePath
)
$script:Integrity = Test-SemtComponentIntegrity -PackageRoot $script:PackageRoot -RelativePaths $componentPaths -Mode (Opt 'ComponentIntegrityMode')
$script:Technician = Get-SemtTechnicianIdentity
$endpointLocalPath = Join-Path $script:PackageRoot $script:CampaignProfile.EndpointScriptRelativePath
$script:EndpointSha256 = Get-SemtFileSha256 -Path $endpointLocalPath
$script:EndpointVersion = Get-SemtScriptVersion -Path $endpointLocalPath
$script:TargetDataRoot = Get-SemtTargetCampaignRoot -Campaign $Campaign
$script:TargetProgramRoot = Get-SemtTargetProgramRoot -Campaign $Campaign

Write-SemtJsonFileAtomic -Path $script:RunContext.EffectiveOptionsPath -InputObject ([pscustomobject]@{
    SchemaVersion = 1
    Campaign = $Campaign
    WaveName = $script:RunContext.WaveName
    RunId = $script:RunId
    RunMode = $RunMode
    ResolvedUtc = Get-SemtUtcNowText
    CatalogSha256 = $catalog.Sha256
    Warnings = @($script:OptionWarnings)
    Options = @(Get-SemtOptionRows -Options $script:Options)
})

Write-Host ("{0} v{1} - {2}" -f $script:Product.DisplayName,$script:LauncherVersion,$script:CampaignProfile.Title) -ForegroundColor Cyan
Write-Host ("Campaign={0}; Wave={1}; RunId={2}; RunMode={3}; SinglePc={4}" -f $Campaign,$script:RunContext.WaveName,$script:RunId,$RunMode,$script:RunContext.SinglePc)
Write-Host ("Run folder: {0}" -f $script:RunContext.RunRoot)
Write-Host ("Technician: Account={0}; UPN={1}; SID={2}; Elevated={3}; Session={4}; Computer={5}" -f $script:Technician.Account,$script:Technician.UserPrincipalName,$script:Technician.Sid,$script:Technician.IsElevatedAdministrator,$script:Technician.SessionId,$script:Technician.ComputerName)
Write-Host ("PowerShell: {0} {1}" -f $PSVersionTable.PSEdition,$PSVersionTable.PSVersion)
Write-Host ("Mode: DryRun={0}; AuditOnly={1}; Endpoint={2} v{3}; Target data root={4}" -f (Opt 'DryRun'),(Opt 'AuditOnly'),(Split-Path $endpointLocalPath -Leaf),$script:EndpointVersion,$script:TargetDataRoot)
Write-Host ("Component integrity: {0}" -f $script:Integrity.Summary) -ForegroundColor $(if ($script:Integrity.Blocking) { 'Red' } else { 'DarkGray' })
Write-Host 'Effective options:'
foreach ($row in @(Get-SemtOptionRows -Options $script:Options)) { Write-Host ("  [{0}] {1}={2} {3} (source: {4})" -f $row.Category,$row.Option,$row.Value,$row.Unit,$row.Source) -ForegroundColor DarkGray }
foreach ($warning in @($script:OptionWarnings)) { Write-Host ("WARNING: {0}" -f $warning) -ForegroundColor Yellow }

if ($script:Integrity.Blocking) {
    foreach ($file in @($script:Integrity.Files | Where-Object { $_.Blocking })) { Write-Host ("  BLOCKED {0}: {1}" -f $file.Path,$file.Detail) -ForegroundColor Red }
    throw ('SEMT_COMPONENT_INTEGRITY_BLOCKED: {0}' -f $script:Integrity.Summary)
}

if ($ValidateOnly) {
    Write-Host 'ValidateOnly: options, rules and component integrity are valid. No target was contacted.' -ForegroundColor Green
    Update-SemtRunHistory -RunContext $script:RunContext -Status 'Validated' -Detail 'ValidateOnly'
    return
}

#endregion

#region Run setup

$script:ResolvedPsExec = ''
$script:PsExecEvidence = Get-SemtPsExecSecurityEvidence -Path ''
if (-not (Opt 'DryRun')) {
    $script:ResolvedPsExec = Resolve-SemtPsExecPath -Path $PsExecPath -PackageRoot $script:PackageRoot
    $script:PsExecEvidence = Get-SemtPsExecSecurityEvidence -Path $script:ResolvedPsExec
    Assert-SemtPsExecSecurityEvidence -Evidence $script:PsExecEvidence
}
Write-Host ("PsExec: Status={0}; Path={1}; SHA256={2}; Signer={3}" -f $script:PsExecEvidence.CheckStatus,$script:PsExecEvidence.Path,$script:PsExecEvidence.SHA256,$script:PsExecEvidence.SignerSubject)

$script:WaveLock = $null
if (-not $script:RunContext.SinglePc) { $script:WaveLock = Enter-SemtWaveLock -Campaign $Campaign -WavePath $script:RunContext.WavePath }
[void](Initialize-SemtCancellation -StateRoot $script:RunContext.StateRoot -Campaign $Campaign -WaveName $script:RunContext.WaveName -RunId $script:RunId)
Set-SemtActiveRunState -Status 'Starting'

$leaseTimeout = Opt 'GlobalWorkerLeaseTimeoutMinutes'
if ($leaseTimeout -lt 1) {
    $base = if ((Opt 'PsExecTimeoutMinutes') -gt 0) { Opt 'PsExecTimeoutMinutes' } else { 240 }
    $waitOption = [string]$script:CampaignProfile.CommunicationLostWaitOption
    $extra = if ([string]::IsNullOrWhiteSpace($waitOption)) { 0 } else { Opt $waitOption }
    $leaseTimeout = [math]::Max(30, $base + $extra + 30)
}
$script:Gate = Initialize-SemtWorkerGate -Limit (Opt 'GlobalWorkerLimit') -LeaseTimeoutMinutes $leaseTimeout -HeartbeatStaleMinutes (Opt 'WorkerHeartbeatStaleMinutes')
Write-Host ("Shared worker cap: Limit={0}; ReservationMaxAge={1} min; HeartbeatStale={2} min; Scope=this technician computer (all sessions and accounts); Root={3}; Acl={4}" -f $script:Gate.Limit,$script:Gate.LeaseTimeoutMinutes,$script:Gate.HeartbeatStaleMinutes,$script:Gate.Root,$script:Gate.AclStatus)

Write-SemtJsonFileAtomic -Path $script:RunContext.RunManifestPath -InputObject ([pscustomobject]@{
    SchemaVersion = 1
    Product = $script:Product
    LauncherVersion = $script:LauncherVersion
    Campaign = $Campaign
    ProfileVersion = $script:CampaignProfile.ProfileVersion
    WaveName = $script:RunContext.WaveName
    WavePath = $script:RunContext.WavePath
    RunId = $script:RunId
    RunMode = $RunMode
    StartedUtc = Get-SemtUtcNowText
    Technician = $script:Technician
    PowerShell = ('{0} {1}' -f $PSVersionTable.PSEdition,$PSVersionTable.PSVersion)
    Endpoint = [pscustomobject]@{ RelativePath = $script:CampaignProfile.EndpointScriptRelativePath; Version = $script:EndpointVersion; Sha256 = $script:EndpointSha256; TargetDataRoot = $script:TargetDataRoot }
    PsExec = $script:PsExecEvidence
    Integrity = $script:Integrity
    WorkerGate = [pscustomobject]@{ Root = $script:Gate.Root; Limit = $script:Gate.Limit; LeaseTimeoutMinutes = $script:Gate.LeaseTimeoutMinutes }
})
Update-SemtRunHistory -RunContext $script:RunContext -Status 'Running'

foreach ($archived in @(Invoke-SemtRunArchive -RunContext $script:RunContext -OlderThanDays (Opt 'RunArchiveAfterDays'))) {
    Write-Host ("Run archive: {0} -> {1}: {2} {3}" -f $archived.Run,$archived.Archive,$archived.Status,$archived.Detail) -ForegroundColor DarkGray
}

$computerListPath = $script:RunContext.ComputerListPath
if (-not (Opt 'DryRun') -and -not $script:RunContext.SinglePc) {
    try {
        $dedupe = Remove-SemtComputerListDuplicates -Path $computerListPath
        if ($dedupe.Changed) { Write-Host ("Computers.txt de-duplicated before the Run: RemovedLines={0}; Samples={1}; Backup={2}" -f $dedupe.DuplicateLines,$dedupe.DuplicateSamples,$dedupe.BackupPath) -ForegroundColor Yellow }
    }
    catch { Write-Host ("WARNING: Computers.txt de-duplication failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
}

$securityRows = @('CheckStatus','Path','SHA256','SignatureStatus','SignerSubject','FileVersion','IsCompliant','Detail' | ForEach-Object { [pscustomobject]@{ Field = "PsExec.$_"; Value = [string](Get-SemtPropertyValue -InputObject $script:PsExecEvidence -Name $_ -Default '') } })
$securityRows += [pscustomobject]@{ Field = 'Endpoint.Sha256'; Value = $script:EndpointSha256 }
$securityRows += [pscustomobject]@{ Field = 'Technician'; Value = ('{0} ({1})' -f $script:Technician.Account,$script:Technician.ComputerName) }
$script:Report = New-SemtReportState -RunContext $script:RunContext -Columns $script:CampaignProfile.ReportColumns -CampaignTitle $script:CampaignProfile.Title -SeparatedStatuses $script:CampaignProfile.SeparatedStatuses -LauncherVersion $script:LauncherVersion -OptionRows @(Get-SemtOptionRows -Options $script:Options) -SecurityRows $securityRows -IntegrityRows @($script:Integrity.Files) -RemoteLogsPathTemplate $script:CampaignProfile.RemoteLogsPathTemplate
Write-Host ("Live report: {0}" -f $script:Report.HtmlPath) -ForegroundColor DarkCyan
Set-SemtActiveRunState -Status 'Running' -ReportPath $script:Report.HtmlPath

$script:Inventory = @{}
foreach ($kind in @($script:CampaignProfile.InventoryKinds)) {
    $script:Inventory[$kind] = [pscustomobject]@{ Kind = $kind; Map = @{}; CsvPath = ''; Source = 'NotLoaded'; SharedPath = (Get-SemtSharedInventoryPath -PackageRoot $script:PackageRoot -Kind $kind); LastRefreshUtc = [datetime]::MinValue; ScopeKeys = @() }
}
$script:ExportScripts = [pscustomobject]@{
    AD = Join-Path $script:PackageRoot 'Scripts\SmartM365-EndpointMigration-Export-ADDevicesCsv.ps1'
    Intune = Join-Path $script:PackageRoot 'Scripts\SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1'
    Entra = Join-Path $script:PackageRoot 'Scripts\SmartM365-EndpointMigration-Export-EntraDevicesCsv.ps1'
}

#endregion

#region Helpers

function Update-CycleInventory {
    param([Parameter(Mandatory = $true)][string]$Kind,[Parameter(Mandatory = $true)][int]$CycleNumber)

    $state = $script:Inventory[$Kind]
    $maxAge = if ($Kind -eq 'AD') { Opt 'AdInventoryFreshnessHours' } else { Opt 'CloudInventoryFreshnessHours' }
    $shared = Test-SemtSharedInventoryUsable -Path $state.SharedPath -Kind $Kind -MaxAgeHours $maxAge -TenantId (Opt 'IntuneTenantId')
    if ($shared.Usable) {
        if ($state.Source -ne 'Shared') {
            $state.Map = Get-SemtInventoryMap -Path $state.SharedPath -Kind $Kind
            $state.CsvPath = $state.SharedPath
            $state.Source = 'Shared'
            Write-Host ("Cycle {0}: {1} inventory: shared cache reused read-only ({2}; age {3} h; {4} device(s))." -f $CycleNumber,$Kind,$shared.Reason,$shared.AgeHours,$state.Map.Count) -ForegroundColor DarkGray
        }
        return
    }
    $refreshOption = [string]$script:CampaignProfile.InventoryRefreshOptions[$Kind]
    $refreshAllowed = [string]::IsNullOrWhiteSpace($refreshOption) -or [bool](Opt $refreshOption)
    $currentKeys = @(Get-SemtComputerList -Path $computerListPath | ForEach-Object { Get-SemtComputerKey -ComputerName $_ } | Select-Object -Unique)
    if ($state.LastRefreshUtc -gt [datetime]::MinValue -and ((Get-Date).ToUniversalTime() - $state.LastRefreshUtc).TotalHours -lt $maxAge -and @($currentKeys | Where-Object { $state.ScopeKeys -notcontains $_ }).Count -eq 0) { return }
    if ((Opt 'DryRun') -or -not $refreshAllowed) {
        if (Test-Path -LiteralPath $state.SharedPath -PathType Leaf) {
            $state.Map = Get-SemtInventoryMap -Path $state.SharedPath -Kind $Kind
            $state.CsvPath = $state.SharedPath
            $state.Source = 'SharedNotFresh'
        }
        Write-Host ("Cycle {0}: {1} inventory refresh skipped ({2}); shared cache {3} ({4}); devices loaded: {5}." -f $CycleNumber,$Kind,$(if (Opt 'DryRun') { 'dry run' } else { 'disabled by option' }),$shared.Reason,$state.SharedPath,$state.Map.Count) -ForegroundColor Yellow
        return
    }
    $scopePath = Join-Path $script:RunContext.InventoryRoot ('Cycle{0}_{1}Scope.txt' -f $CycleNumber,$Kind)
    Write-SemtTextLinesAtomic -Path $scopePath -Lines @(Get-SemtComputerList -Path $computerListPath)
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    Write-Host ("Cycle {0}: refreshing {1} inventory scoped to the Wave ({2} computer(s)); shared cache {3}." -f $CycleNumber,$Kind,$currentKeys.Count,$shared.Reason) -ForegroundColor Yellow
    $export = Invoke-SemtInventoryExport -Kind $Kind -ExportScriptPath $script:ExportScripts.$Kind -OutputPath (Join-Path $script:RunContext.InventoryRoot ('Devices{0}_Cycle{1}_{2}.csv' -f $Kind,$CycleNumber,$stamp)) -LogPath (Join-Path $script:RunContext.InventoryRoot ('Devices{0}_Cycle{1}_{2}.log' -f $Kind,$CycleNumber,$stamp)) -ComputerListPath $scopePath -TenantId (Opt 'IntuneTenantId') -PageSize (Opt 'IntuneInventoryPageSize') -SharedInventoryPath $state.SharedPath
    if ($export.Success) {
        $state.Map = $export.Map
        $state.CsvPath = $export.CsvPath
        $state.Source = 'RunScopedRefresh'
        $state.LastRefreshUtc = (Get-Date).ToUniversalTime()
        $state.ScopeKeys = $currentKeys
        Write-Host ("Cycle {0}: {1} inventory refreshed: {2} device(s); CSV={3}" -f $CycleNumber,$Kind,$state.Map.Count,$export.CsvPath) -ForegroundColor Green
    }
    else {
        Write-Host ("WARNING: Cycle {0}: {1} inventory refresh failed: {2}" -f $CycleNumber,$Kind,$export.Error) -ForegroundColor Yellow
        if ($state.Map.Count -eq 0 -and (Test-Path -LiteralPath $state.SharedPath -PathType Leaf)) {
            $state.Map = Get-SemtInventoryMap -Path $state.SharedPath -Kind $Kind
            $state.CsvPath = $state.SharedPath
            $state.Source = 'SharedNotFreshAfterRefreshFailure'
        }
    }
}

function Add-InventoryFields {
    param($Row)
    foreach ($kind in @($script:Inventory.Keys)) { [void](Add-SemtInventoryFieldsToResult -Result $Row -Kind $kind -Map $script:Inventory[$kind].Map -CsvPath $script:Inventory[$kind].CsvPath) }
    return $Row
}

function New-EngineResult {
    param([string]$Computer,[int]$CycleNumber,[string]$Status,[string]$NextAction,[string]$Detail)
    $row = New-SemtResult -Campaign $Campaign -ComputerName $Computer -CycleNumber $CycleNumber -WaveName $script:RunContext.WaveName -RunId $script:RunId -LauncherStatus $Status -NextAction $NextAction -Detail $Detail -ExtensionFields $script:CampaignProfile.ExtensionFields
    [void](Complete-SemtResultStatus -Result $row)
    $row.PsExecLogPath = $script:RunContext.LauncherLogPath
    return (Add-InventoryFields -Row ([pscustomobject]$row))
}

function Invoke-CompletedRules {
    param([object[]]$Rows,[int]$CycleNumber)
    if ((Opt 'DryRun') -or $script:RunContext.SinglePc) { return }
    foreach ($rule in @($script:CampaignProfile.CompletedRules)) {
        $names = @($Rows | Where-Object { $_ -and (& $rule.Test $_) } | ForEach-Object { [string]$_.ComputerName } | Select-Object -Unique)
        if ($names.Count -eq 0) { continue }
        try {
            $moved = Move-SemtComputersFromList -ComputerListPath $computerListPath -ComputerNames $names -ArchiveFileName $rule.ArchiveFileName -Label $rule.Label
            if ($moved.Moved -gt 0) { Write-Host ("Cycle {0}: {1}" -f $CycleNumber,$moved.Detail) -ForegroundColor Green }
        }
        catch { Write-Host ("WARNING: Cycle {0}: {1} move failed: {2}" -f $CycleNumber,$rule.Label,$_.Exception.Message) -ForegroundColor Yellow }
    }
}

function Start-WorkerJobWithRetry {
    param([string]$JobName,[string]$WorkerPath,[string]$ContractPath,[string]$Computer)
    $diagnostics = New-Object System.Collections.Generic.List[string]
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $job = Start-Job -Name $JobName -ScriptBlock { param($w,$c) & $w -ContractPath $c } -ArgumentList $WorkerPath,$ContractPath
            if ($null -eq $job) { throw 'Start-Job returned no job.' }
            if ($attempt -gt 1) { Write-Host ("LOCAL_WORKER_START_RECOVERED: {0} attempt {1}/3." -f $Computer,$attempt) -ForegroundColor Green }
            return [pscustomobject]@{ Succeeded = $true; Job = $job; Detail = ($diagnostics -join ' | ') }
        }
        catch {
            $diagnostics.Add(('Attempt={0}/3; PSEdition={1}; PSVersion={2}; Error={3}' -f $attempt,$PSVersionTable.PSEdition,$PSVersionTable.PSVersion,$_.Exception.Message))
            Write-Host ("LOCAL_WORKER_START_RETRY: {0}: {1}" -f $Computer,$diagnostics[$diagnostics.Count - 1]) -ForegroundColor Yellow
            Start-Sleep -Seconds 3
        }
    }
    return [pscustomobject]@{ Succeeded = $false; Job = $null; Detail = ($diagnostics -join ' | ') }
}

function Get-RunGuardEntry {
    param([string]$Fqdn)
    $found = @{}
    Invoke-SemtRunGuardHistoryTransaction -Campaign $Campaign -SchemaVersion $script:CampaignProfile.RunGuardSchemaVersion -Prune $script:CampaignProfile.Hooks.RunGuardPrune -ArgumentList @($Fqdn,$found) -Update {
        param($Entries,$LockedFqdn,$FoundRef)
        foreach ($entry in $Entries) { if ([string]$entry.ComputerFqdn -eq $LockedFqdn) { $FoundRef['Value'] = $entry; break } }
    } | Out-Null
    if ($found.ContainsKey('Value')) { return $found['Value'] }
    return $null
}

function Update-RunGuardEntry {
    param([string]$Fqdn,[string]$Computer,[ValidateSet('Started','Result')][string]$State,$Result,[string]$JobId,[int]$CycleNumber)
    $meta = @{ WaveName = $script:RunContext.WaveName; RunId = $script:RunId; TechnicianAccount = $script:Technician.Account; LauncherVersion = $script:LauncherVersion }
    $hours = Opt 'TechnicianRunGuardHours'
    $entryBuilder = $script:CampaignProfile.Hooks.RunGuardEntry
    Invoke-SemtRunGuardHistoryTransaction -Campaign $Campaign -SchemaVersion $script:CampaignProfile.RunGuardSchemaVersion -Prune $script:CampaignProfile.Hooks.RunGuardPrune -ArgumentList @($Fqdn,$Computer,$State,$Result,$JobId,$CycleNumber,$hours,$meta,$entryBuilder) -Update {
        param($Entries,$LockedFqdn,$LockedComputer,$LockedState,$LockedResult,$LockedJobId,$LockedCycle,$LockedHours,$LockedMeta,$LockedBuilder)
        $previous = $null
        foreach ($entry in @($Entries.ToArray())) { if ([string]$entry.ComputerFqdn -eq $LockedFqdn) { $previous = $entry; [void]$Entries.Remove($entry) } }
        $new = & $LockedBuilder $previous $LockedFqdn $LockedComputer $LockedState $LockedResult $LockedJobId $LockedCycle $LockedHours $LockedMeta
        if ($null -ne $new) { $Entries.Add($new) }
    } | Out-Null
}

#endregion

#region Cycles

$workerPath = Join-Path $script:PackageRoot 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
$baseEndpointArguments = @(ConvertTo-SemtEndpointArguments -Catalog $catalog -Campaign $Campaign -Options $script:Options) + @('-DataRoot',$script:TargetDataRoot)
if ($Campaign -eq 'Windows11Upgrade') { $baseEndpointArguments += @('-SetupCacheRoot',(Join-Path $script:TargetDataRoot 'SetupMedia')) }
$useRunGuard = [bool]((Opt 'UseTechnicianRunGuard') -and -not (Opt 'IgnoreEndpointRunGuard') -and -not (Opt 'DryRun') -and (Opt 'TechnicianRunGuardHours') -gt 0)
Write-Host ("Technician run guard: Effective={0}; Hours={1}; History={2}" -f $useRunGuard,(Opt 'TechnicianRunGuardHours'),(Get-SemtRunGuardHistoryPath -Campaign $Campaign)) -ForegroundColor DarkCyan
$cycle = 0
$stopReason = 'Finished'
$lastCycleRows = @()
$runStatus = 'Finished'

try {
    while ($true) {
        if ((Get-SemtCancellationState).Requested) { Write-Host ("Controlled stop requested before cycle {0}; no new cycle starts." -f ($cycle + 1)) -ForegroundColor Yellow; $stopReason = 'StoppedByOperator'; break }
        if (Opt 'NightPauseEnabled') {
            $resume = Get-SemtNightPauseResumeTime -Now (Get-Date) -StartHour (Opt 'NightPauseStartHour') -EndHour (Opt 'NightPauseEndHour')
            if ($null -ne $resume) {
                Write-Host ("Night pause: cycle {0} will start after {1:yyyy-MM-dd HH:mm:ss}. Cycles already running are never interrupted." -f ($cycle + 1),$resume) -ForegroundColor Yellow
                Set-SemtActiveRunState -Status 'NightPause' -ReportPath $script:Report.HtmlPath
                [void](Wait-SemtCancellationAware -Seconds ([int][math]::Max(1, [math]::Ceiling(($resume - (Get-Date)).TotalSeconds))))
                if ((Get-SemtCancellationState).Requested) { $stopReason = 'StoppedByOperator'; break }
                Set-SemtActiveRunState -Status 'Running' -ReportPath $script:Report.HtmlPath
            }
        }
        if (-not (Test-SemtComputerListPresent -Path $computerListPath)) { Write-Host ("[WAVE_INPUT_REMOVED] Computers.txt disappeared; reports are kept and the Run stops. Path={0}" -f $computerListPath) -ForegroundColor Yellow; $stopReason = 'StoppedInputRemoved'; break }
        $cycle++
        $cycleStart = Get-Date
        $computers = @(Get-SemtComputerList -Path $computerListPath)
        $listStats = Get-SemtComputerListStats -Path $computerListPath
        if ($computers.Count -eq 0) { Write-Host 'No computer left in Computers.txt. The Run is complete.' -ForegroundColor Green; break }
        Write-Host ("Cycle {0}: {1} computer(s)." -f $cycle,$computers.Count) -ForegroundColor Cyan

        foreach ($kind in @($script:CampaignProfile.InventoryKinds)) {
            # A stop requested during an inventory refresh does not start the next refresh.
            if ((Get-SemtCancellationState).Requested) { Write-Host ("Cycle {0}: {1} inventory refresh skipped: stop requested." -f $cycle,$kind) -ForegroundColor Yellow; continue }
            try { Update-CycleInventory -Kind $kind -CycleNumber $cycle } catch { Write-Host ("WARNING: Cycle {0}: {1} inventory unavailable: {2}" -f $cycle,$kind,$_.Exception.Message) -ForegroundColor Yellow }
        }
        $hookContext = [pscustomobject]@{ Options = $script:Options; RunContext = $script:RunContext; Inventory = $script:Inventory; CycleNumber = $cycle; DryRun = [bool](Opt 'DryRun'); Log = $script:Log; ExportScripts = $script:ExportScripts; PackageRoot = $script:PackageRoot }
        $cycleRows = New-Object System.Collections.ArrayList
        $preRows = @(& $script:CampaignProfile.Hooks.PreCycle $hookContext $computers | ForEach-Object { Add-InventoryFields -Row $_ })
        if ($preRows.Count -gt 0) {
            foreach ($row in $preRows) { [void]$cycleRows.Add($row) }
            Invoke-CompletedRules -Rows $preRows -CycleNumber $cycle
            $preKeys = @{}
            foreach ($row in $preRows) { $preKeys[(Get-SemtComputerKey -ComputerName ([string]$row.ComputerName))] = $true }
            $computers = @($computers | Where-Object { -not $preKeys.ContainsKey((Get-SemtComputerKey -ComputerName $_)) })
            Write-Host ("Cycle {0}: {1} computer(s) resolved from inventory before PsExec; remaining {2}." -f $cycle,$preRows.Count,$computers.Count) -ForegroundColor Green
        }

        $dns = if ($computers.Count -gt 0) { Test-SemtDnsSample -ComputerNames $computers -DomainSuffix (Opt 'AdDomain') -SampleSize (Opt 'DnsPreflightSampleSize') } else { $null }
        if ($dns -and $dns.AllFailed -and (Opt 'DnsPreflightBlocksCycle')) {
            Write-Host ("*** DNS preflight: all {0} sampled computer(s) failed DNS resolution. Check VPN and DNS on this technician computer. The cycle stops before PsExec. ***" -f $dns.Tested) -ForegroundColor Red
            foreach ($computer in $computers) { [void]$cycleRows.Add((New-EngineResult -Computer $computer -CycleNumber $cycle -Status 'DNS_PREFLIGHT_ALL_SAMPLES_FAILED' -NextAction 'CHECK_VPN_DNS_BEFORE_WAVE' -Detail ('DNS preflight failed for every sampled computer: {0}' -f ((@($dns.Samples) | ForEach-Object { '{0}->{1}' -f $_.Computer,$_.ConnectionTarget }) -join '; ')))) }
            $computers = @()
        }
        elseif ($dns -and $dns.AllFailed) { Write-Host ("WARNING: DNS preflight failed for all {0} sample(s); continuing because DnsPreflightBlocksCycle is disabled." -f $dns.Tested) -ForegroundColor Yellow }

        $running = New-Object System.Collections.ArrayList
        $forced = @{}
        $cancelObservedAt = $null
        $next = 0
        $lastProgressLog = Get-Date
        $lastHtml = [datetime]::MinValue
        $cycleIgnoreGuard = [bool]((Opt 'IgnoreEndpointRunGuard') -and ($cycle -eq 1 -or (Opt 'IgnoreEndpointRunGuardEveryCycle')))
        while ($next -lt $computers.Count -or $running.Count -gt 0) {
            $cancel = Get-SemtCancellationState
            if ($cancel.Requested) {
                if ($null -eq $cancelObservedAt) {
                    $cancelObservedAt = Get-Date
                    Set-SemtActiveRunState -Status 'StopRequested' -ReportPath $script:Report.HtmlPath
                    Write-Host ("Controlled stop requested by {0}: no new computer starts; running workers drain for up to {1} minute(s). A second stop forces the local stop." -f $cancel.Source,(Opt 'CancellationDrainTimeoutMinutes')) -ForegroundColor Yellow
                }
                while ($next -lt $computers.Count) { [void]$cycleRows.Add((New-EngineResult -Computer $computers[$next] -CycleNumber $cycle -Status 'CANCELLED_NOT_STARTED' -NextAction 'SAFE_TO_RELAUNCH' -Detail 'Controlled stop requested before this computer was started.')); $next++ }
                $drainExpired = ((Opt 'CancellationDrainTimeoutMinutes') -eq 0 -or ((Get-Date) - $cancelObservedAt).TotalMinutes -ge (Opt 'CancellationDrainTimeoutMinutes'))
                if ($running.Count -gt 0 -and ($cancel.Force -or $drainExpired)) {
                    foreach ($item in @($running | Where-Object { $_.Job.State -eq 'Running' })) {
                        if (-not $forced.ContainsKey([string]$item.Job.Id)) {
                            $forced[[string]$item.Job.Id] = $true
                            Write-Host ("Forcing the local stop of {0}: {1}. The remote endpoint may already be running." -f $item.Computer,(Stop-SemtWorkerProcessTree -LeasePath $item.LeasePath)) -ForegroundColor Red
                        }
                        Stop-Job -Job $item.Job -ErrorAction SilentlyContinue
                    }
                }
            }

            while ($next -lt $computers.Count -and $running.Count -lt (Opt 'ThrottleLimit') -and -not (Get-SemtCancellationState).Requested) {
                $computer = $computers[$next]
                $next++
                $target = Resolve-SemtConnectionTarget -ComputerName $computer -DomainSuffix (Opt 'AdDomain')
                $fqdn = Get-SemtRunGuardFqdn -ComputerName $computer -AdInventoryMap $(if ($script:Inventory.ContainsKey('AD')) { $script:Inventory['AD'].Map } else { $null })
                $perComputer = New-Object System.Collections.Generic.List[string]
                if ($cycleIgnoreGuard) { $perComputer.Add('-IgnoreRunGuard') }
                if ($useRunGuard) {
                    $decision = $null
                    try { $decision = & $script:CampaignProfile.Hooks.RunGuardDecision (Get-RunGuardEntry -Fqdn $fqdn) $script:Options }
                    catch {
                        [void]$cycleRows.Add((New-EngineResult -Computer $computer -CycleNumber $cycle -Status 'TECH_RUN_GUARD_HISTORY_UNAVAILABLE' -NextAction 'CHECK_TECHNICIAN_STATE_ACCESS' -Detail ('Technician run guard history is unavailable; the computer was not launched. Error={0}' -f $_.Exception.Message)))
                        continue
                    }
                    if ($decision.Block) {
                        $skip = New-EngineResult -Computer $computer -CycleNumber $cycle -Status $decision.Status -NextAction $decision.NextAction -Detail $decision.Detail
                        foreach ($fieldName in @($decision.Fields.Keys)) { Set-SemtResultField -Result $skip -Name $fieldName -Value $decision.Fields[$fieldName] }
                        [void]$cycleRows.Add($skip)
                        # Why and until when: the previous status and the next eligible local time.
                        $priorStatus = [string]$(if ($decision.Fields -and $decision.Fields.ContainsKey('BackoffStatus')) { $decision.Fields['BackoffStatus'] } else { '' })
                        $untilUtc = ConvertTo-SemtUtcDateTime -Value $(if ($decision.Fields -and $decision.Fields.ContainsKey('BackoffUntilUtc')) { $decision.Fields['BackoffUntilUtc'] } else { '' })
                        $untilText = if ($null -ne $untilUtc) { ' next try after {0:yyyy-MM-dd HH:mm} (local);' -f $untilUtc.ToLocalTime() } else { '' }
                        Write-Host ("  [{0}] {1} - previous status {2};{3} a run mode 'ignore run guard' bypasses it." -f $decision.Status,$computer,$(if ($priorStatus) { $priorStatus } else { 'unknown' }),$untilText) -ForegroundColor DarkYellow
                        continue
                    }
                    if ($decision.BypassEndpointGuard -and -not $perComputer.Contains('-IgnoreRunGuard')) {
                        $perComputer.Add('-IgnoreRunGuard')
                        Write-Host ("  [ENDPOINT_RUN_GUARD_RETRY_BYPASS] {0}: adaptive retry is due; the bypass applies to this attempt only." -f $computer) -ForegroundColor DarkCyan
                    }
                }
                foreach ($argument in @(& $script:CampaignProfile.Hooks.PerComputerArguments $hookContext $computer)) { $perComputer.Add([string]$argument) }

                $leasePath = Request-SemtWorkerLease -Gate $script:Gate -Campaign $Campaign -RunId $script:RunId -WaveName $script:RunContext.WaveName -Computer $computer -CycleNumber $cycle -OnWaitLog { param($message) Write-Host $message -ForegroundColor DarkYellow }
                if ((Get-SemtCancellationState).Requested) { [void](Remove-SemtWorkerLease -Gate $script:Gate -LeasePath $leasePath); $next--; break }

                $correlationId = [guid]::NewGuid().ToString('N')
                $arguments = @($baseEndpointArguments) + @($perComputer.ToArray()) + @('-CorrelationId',$correlationId)
                $waitOption = [string]$script:CampaignProfile.CommunicationLostWaitOption
                $pollOption = [string]$script:CampaignProfile.CommunicationLostPollOption
                $contract = New-SemtWorkerContract -RunContext $script:RunContext -Computer $computer -ConnectionTarget $target.ConnectionTarget -CycleNumber $cycle -CorrelationId $correlationId `
                    -Endpoint @{ RelativePath = $script:CampaignProfile.EndpointScriptRelativePath; ExpectedSha256 = $script:EndpointSha256; ExpectedVersion = $script:EndpointVersion; RemoteDataRoot = $script:TargetDataRoot; RemoteProgramRoot = $script:TargetProgramRoot; Arguments = @($arguments) } `
                    -PsExec @{ Path = $script:ResolvedPsExec; Sha256 = $script:PsExecEvidence.SHA256; TimeoutMinutes = (Opt 'PsExecTimeoutMinutes'); UseElevatedToken = [bool]$script:CampaignProfile.PsExecUseElevatedToken; CommunicationLostWaitMinutes = $(if ($waitOption) { Opt $waitOption } else { 0 }); CommunicationLostPollMinutes = $(if ($pollOption) { Opt $pollOption } else { 10 }) } `
                    -Evidence @{ Enabled = [bool](Opt 'CentralLogCollection'); Mode = (Opt 'CentralLogCollectionMode'); KeepHistory = [bool](Opt 'KeepCentralLogHistory'); LogRoot = $script:RunContext.PsExecLogRoot; CentralLogRoot = $script:RunContext.CentralLogRoot } `
                    -Flags @{ DryRun = [bool](Opt 'DryRun'); ComponentIntegrityMode = (Opt 'ComponentIntegrityMode') } `
                    -Gate @{ LeasePath = $leasePath }
                $jobName = 'SEMT_{0}_C{1}_{2}' -f $Campaign,$cycle,$computer
                $start = Start-WorkerJobWithRetry -JobName $jobName -WorkerPath $workerPath -ContractPath $contract.Path -Computer $computer
                if (-not $start.Succeeded) {
                    [void](Remove-SemtWorkerLease -Gate $script:Gate -LeasePath $leasePath)
                    $failed = New-EngineResult -Computer $computer -CycleNumber $cycle -Status 'LOCAL_WORKER_START_FAILED' -NextAction 'RETRY_LAUNCH' -Detail ('The local worker could not start after 3 attempts; the endpoint was not contacted. {0}' -f $start.Detail)
                    [void]$cycleRows.Add($failed)
                    continue
                }
                [void](Update-SemtWorkerLease -LeasePath $leasePath -LockPath $script:Gate.LockPath -Properties @{ JobId = [string]$start.Job.Id; JobName = $jobName })
                if ($useRunGuard) { try { Update-RunGuardEntry -Fqdn $fqdn -Computer $computer -State Started -Result $null -JobId ([string]$start.Job.Id) -CycleNumber $cycle } catch { Write-Host ("TECH_RUN_GUARD_HISTORY_WRITE_DEFERRED: Started; {0}; {1}" -f $computer,$_.Exception.Message) -ForegroundColor Yellow } }
                [void]$running.Add([pscustomobject]@{ Job = $start.Job; Computer = $computer; Target = $target.ConnectionTarget; LeasePath = $leasePath; ContractPath = $contract.Path; CorrelationId = $correlationId; StartedAt = Get-Date; Fqdn = $fqdn })
                Write-Host ("Queued {0} ({1}/{2}); running={3}" -f $computer,$next,$computers.Count,$running.Count) -ForegroundColor DarkCyan
                if ((Opt 'DelayBetweenComputersSeconds') -gt 0 -and $next -lt $computers.Count) { [void](Wait-SemtCancellationAware -Seconds (Opt 'DelayBetweenComputersSeconds')) }
            }

            $finished = @($running | Where-Object { $_.Job.State -ne 'Running' -and $_.Job.State -ne 'NotStarted' })
            if ($finished.Count -eq 0) {
                if (((Get-Date) - $lastProgressLog).TotalMinutes -ge 5 -and $running.Count -gt 0) {
                    Write-Host ("Waiting for {0} worker(s); elapsed {1:N1} min; running: {2}" -f $running.Count,((Get-Date) - $cycleStart).TotalMinutes,((@($running) | ForEach-Object { '{0} ({1:N1}m)' -f $_.Computer,((Get-Date) - $_.StartedAt).TotalMinutes }) -join ', '))
                    $lastProgressLog = Get-Date
                }
                if (((Get-Date) - $lastHtml).TotalSeconds -ge 60) {
                    try {
                        $runningRows = @($running | ForEach-Object { [pscustomobject]@{ ComputerName = $_.Computer; ConnectionTarget = $_.Target; JobId = [string]$_.Job.Id; State = [string]$_.Job.State; Started = $_.StartedAt.ToString('yyyy-MM-dd HH:mm:ss'); ElapsedMinutes = [math]::Round(((Get-Date) - $_.StartedAt).TotalMinutes, 1) } })
                        [void](Write-SemtRunHtmlReport -State $script:Report -CycleNumber $cycle -IsLive -CurrentCycleRows @($cycleRows.ToArray()) -CurrentProgress (New-SemtProgressRow -CycleNumber $cycle -CycleStart $cycleStart -Total ($computers.Count + $preRows.Count) -Queued $next -Completed $cycleRows.Count -Running $running.Count -ComputerListStats $listStats) -RunningRows $runningRows)
                        $lastHtml = Get-Date
                    }
                    catch { Write-Host ("WARNING: live report update failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
                }
                Start-Sleep -Seconds 2
                continue
            }

            foreach ($item in $finished) {
                $jobErrors = New-Object System.Collections.Generic.List[string]
                $received = @()
                try {
                    $receiveErrors = @()
                    $received = @(Receive-Job -Job $item.Job -ErrorAction SilentlyContinue -ErrorVariable receiveErrors | Where-Object { $_ -and $_.PSObject.Properties['CorrelationId'] -and [string]$_.CorrelationId -eq $item.CorrelationId })
                    foreach ($errorRecord in @($receiveErrors)) { $jobErrors.Add($errorRecord.ToString()) }
                    foreach ($child in @($item.Job.ChildJobs)) {
                        foreach ($childError in @($child.Error)) { $jobErrors.Add($childError.ToString()) }
                        if ($child.JobStateInfo.Reason) {
                            $reason = [string]$child.JobStateInfo.Reason.Message
                            $jobErrors.Add($reason)
                            if ($reason -match '\bBroken\b') { $jobErrors.Add('RUNSPACE_BROKEN: the job runspace broke; consider DelayBetweenComputersSeconds=1.') }
                        }
                    }
                }
                catch { $jobErrors.Add($_.Exception.Message) }
                $row = $null
                if ($received.Count -gt 0) { $row = $received[-1] }
                elseif ($forced.ContainsKey([string]$item.Job.Id)) {
                    $remote = $null
                    try {
                        $share = ConvertTo-SemtAdminSharePath -ComputerName $item.Target -LocalPath $script:TargetDataRoot
                        for ($check = 1; $check -le 3; $check++) {
                            $remote = Read-SemtRemoteAttemptResult -RemoteDataShare $share -CorrelationId $item.CorrelationId
                            if ($remote.Completed) { break }
                            if ($check -lt 3) { Start-Sleep -Seconds 5 }
                        }
                    }
                    catch { $jobErrors.Add(('Forced-stop recheck failed: {0}' -f $_.Exception.Message)) }
                    if ($remote -and $remote.Completed) {
                        $row = New-EngineResult -Computer $item.Computer -CycleNumber $cycle -Status 'CANCELLED_BY_OPERATOR' -NextAction ([string](Get-SemtPropertyValue -InputObject $remote.Result -Name 'NextAction' -Default '')) -Detail ('Local worker forcibly stopped; the endpoint attempt completed: {0}' -f (Get-SemtPropertyValue -InputObject $remote.Result -Name 'Detail' -Default ''))
                        Set-SemtResultField -Result $row -Name 'RemoteStatus' -Value ([string](Get-SemtPropertyValue -InputObject $remote.Result -Name 'Status' -Default ''))
                        Set-SemtResultField -Result $row -Name 'Status' -Value ([string](Get-SemtPropertyValue -InputObject $remote.Result -Name 'Status' -Default ''))
                    }
                    elseif ($remote -and $remote.State -eq 'Running') {
                        $row = New-EngineResult -Computer $item.Computer -CycleNumber $cycle -Status 'CANCELLED_BY_OPERATOR_REMOTE_RUNNING' -NextAction 'WAIT_REMOTE_ENDPOINT_COMPLETION' -Detail 'Local worker forcibly stopped; the endpoint attempt is still running on the target.'
                    }
                    else {
                        $row = New-EngineResult -Computer $item.Computer -CycleNumber $cycle -Status 'CANCELLED_BY_OPERATOR_REMOTE_STATE_UNCONFIRMED' -NextAction 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH' -Detail 'Local worker forcibly stopped; no current-attempt endpoint result was found after three checks.'
                    }
                }
                else {
                    $status = if (@($jobErrors | Where-Object { $_ -like 'RUNSPACE_BROKEN*' }).Count -gt 0) { 'RUNSPACE_BROKEN' } else { 'JOB_ERROR' }
                    $row = New-EngineResult -Computer $item.Computer -CycleNumber $cycle -Status $status -NextAction 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH' -Detail ($jobErrors -join ' | ')
                }
                if ($jobErrors.Count -gt 0) { Set-SemtResultField -Result $row -Name 'JobErrorMessage' -Value ((@(Get-SemtPropertyValue -InputObject $row -Name 'JobErrorMessage' -Default '') + @($jobErrors)) -join ' | ').Trim(' ','|') }
                $row = Add-InventoryFields -Row $row
                if ($useRunGuard) { try { Update-RunGuardEntry -Fqdn $item.Fqdn -Computer $item.Computer -State Result -Result $row -JobId ([string]$item.Job.Id) -CycleNumber $cycle } catch { Write-Host ("TECH_RUN_GUARD_HISTORY_WRITE_DEFERRED: Result; {0}; {1}" -f $item.Computer,$_.Exception.Message) -ForegroundColor Yellow } }
                [void]$cycleRows.Add($row)
                Invoke-CompletedRules -Rows @($row) -CycleNumber $cycle
                $color = if ([string]$row.Status -match 'SUCCESS|ALREADY|COMPLETED|STARTED|DETECTED') { 'Green' } elseif ([string]$row.Status -match 'GUARD|PENDING|SKIPPED|DRYRUN|INSUFFICIENT|TIMEOUT|LEGACY|TARGET') { 'Yellow' } else { 'Red' }
                $detailText = [string]$row.Detail
                if ($detailText.Length -gt 180) { $detailText = $detailText.Substring(0, 177) + '...' }
                Write-Host ("Completed {0}; Elapsed={1:N1} min; Status={2}; NextAction={3}; ExitCode={4}; Detail={5}" -f $item.Computer,((Get-Date) - $item.StartedAt).TotalMinutes,$row.Status,$row.NextAction,$row.ExitCode,$detailText) -ForegroundColor $color
                [void](Remove-SemtWorkerLease -Gate $script:Gate -LeasePath $item.LeasePath)
                Remove-Job -Job $item.Job -Force -ErrorAction SilentlyContinue
                $running.Remove($item)
            }
        }

        $rows = @($cycleRows.ToArray())
        try { & $script:CampaignProfile.Hooks.PostCycle $hookContext $rows } catch { Write-Host ("WARNING: Cycle {0}: post-cycle step failed: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
        Invoke-CompletedRules -Rows $rows -CycleNumber $cycle
        $csvPath = Join-Path $script:RunContext.ReportRoot ('SEMT_{0}_cycle{1}_{2}.csv' -f $Campaign,$cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
        Export-SemtReportCsv -Rows $rows -Columns $script:CampaignProfile.ReportColumns -Path $csvPath
        Add-SemtReportAttemptRows -State $script:Report -Rows $rows
        [void]$script:Report.ProgressRows.Add((New-SemtProgressRow -CycleNumber $cycle -CycleStart $cycleStart -Total ($computers.Count + $preRows.Count) -Queued $computers.Count -Completed $rows.Count -Running 0 -ComputerListStats $listStats))
        # The HTML report is a view of the results already saved in the cycle CSV: its failure never fails the Run.
        try { [void](Write-SemtRunHtmlReport -State $script:Report -CycleNumber $cycle -IsLive) }
        catch { Write-Host ("WARNING: Cycle {0}: HTML report update failed: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
        Write-Host ("Cycle {0} report: {1}" -f $cycle,$csvPath) -ForegroundColor Green
        foreach ($group in @($rows | Group-Object -Property Status | Sort-Object Count -Descending)) { Write-Host ("  {0,-48} {1,5}" -f $group.Name,$group.Count) }
        $lastCycleRows = $rows

        if ((Get-SemtCancellationState).Requested) { $stopReason = 'StoppedByOperator'; break }
        if ((Opt 'MaxCycles') -gt 0 -and $cycle -ge (Opt 'MaxCycles')) { break }
        $minimumDelay = (Opt 'DelayBetweenCyclesMinutes') * 60
        $delay = [int](& $script:CampaignProfile.Hooks.AdaptiveDelaySeconds $lastCycleRows $minimumDelay)
        if ($delay -gt 0) {
            Write-Host ("Waiting {0:N1} minute(s) before cycle {1}." -f ($delay / 60.0),($cycle + 1)) -ForegroundColor DarkGray
            [void](Wait-SemtCancellationAware -Seconds $delay)
        }
    }
}
catch {
    $runStatus = 'Failed'
    Write-Host ("RUN FAILED: {0}" -f $_.Exception.Message) -ForegroundColor Red
    throw
}
finally {
    try {
        $finalPath = Join-Path $script:RunContext.ReportRoot ('SEMT_{0}_{1}_final.html' -f $Campaign,$script:RunId)
        if ($script:Report.AttemptRowCount -gt 0) {
            # A report failure must not prevent the Run history and active state updates below.
            try {
                [void](Write-SemtRunHtmlReport -State $script:Report -CycleNumber $cycle -Path $finalPath)
                [void](Write-SemtRunHtmlReport -State $script:Report -CycleNumber $cycle)
                Write-Host ("Final report: {0}" -f $finalPath) -ForegroundColor Green
            }
            catch { Microsoft.PowerShell.Utility\Write-Host ("Final report failed: {0}" -f $_.Exception.Message) }
        }
        $historyStatus = if ($runStatus -eq 'Failed') { 'Failed' } else { $stopReason }
        Update-SemtRunHistory -RunContext $script:RunContext -Status $historyStatus -ReportPath $finalPath -Cycles $cycle
        Set-SemtActiveRunState -Status $historyStatus -ReportPath $finalPath
    }
    catch { Microsoft.PowerShell.Utility\Write-Host ("Final report or history update failed: {0}" -f $_.Exception.Message) }
    Complete-SemtCancellation
    Exit-SemtWaveLock -WaveLock $script:WaveLock
    Write-Host ("{0} Run finished. Status={1}; Cycles={2}; Log={3}" -f $script:Product.DisplayName,$stopReason,$cycle,$script:RunContext.LauncherLogPath)
}

#endregion

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAL+iqeJoF4BdRs
# uwEn0zZJ+dI5rw3TbSw6wFs/gfji8KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICAwFemtLRDD+xZ+p8linli7dNPcA78W55rqWLewzo5CMA0GCSqG
# SIb3DQEBAQUABIIBgIbQU6XtpBwd/Yu4BMXKkKAjRHTUE93J6KmSsNpayY2PlH4G
# X7JUaHjAANMJqUwEH+aIdeAz5+k5S64j6ffZZoYOOnc+RjNzKD7dAq/Q3qVYcZF4
# wIajAcWIov+6Q18URtVZUwkyYu+Gt19DbpdQzSteFBhLw+dXllALRWJC3tjsDdsX
# a86+6AbGnDZy2dZ5KooOA9UHSE5ZjwRBYJ8cmIF0jeQ4y43Da38u2Cztz0WzBISV
# lUbFeLfvEvreFFWD+9hj7pf5rAETZ2qoO+u3L78jPa2XUNwHpMz1PbD0u2IgeCv0
# zW/lz3TtylgHY+uVxXersvtYJsMTaI816lqR+ospambltE7LvOFK12X8oKoQqQBa
# W8J+7mzP+kxseYiQ1El5FL3mKncn6udhUclVG7jlrGC/N5kIBJGI8JmTbE5iHl0E
# c5efBKlgSUZEvois+e+HO4ZInpY5r8roo4/qTfkjpdpveId1c1Fn+aeNcE3gCoi0
# b88LRQD6jwGRO0y7s6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxMDU5
# MjFaMC8GCSqGSIb3DQEJBDEiBCAjQR0ZwxoC/Kn3bjCiusxnTo5dTleJENNDSt1j
# MCNYXDANBgkqhkiG9w0BAQEFAASCAgBetLynW2V9f7/LBmK7NxN6onAraChXGYqU
# CNFGDHLrU4NINNjuUNz9UdNs92WdTKjW0iDc/jgY3NGs2OiQbiB5KkZjgat80uNs
# OX4N1kpBk4fYGUPxk3/kGaOZaFg6mtJwo8K+jT+Mxje7HKnQ3/s1cixTHj5qDHZh
# UWMZnxk3TmOqIDcWvptQI4+MexWIRCnvsEnff9lQYnuW6DWNP/g3WwXwSPyNCyQN
# 0j3nZ2g/63m4BZdkaVuyVAQUviSeCKhuCx8hBPthItor6flPFDyLiZGljitviOkw
# /TdjcnSRoWCMWi4TCQrlGamLngHjK573kkj8nw4tF6X+jPVXVwtiDMd5F6dX4lum
# Ajdo0eIryfV12QUIb7dpMl88kQkrVrlSIfVlEReh6nbTQvc1JBM68KgSNy321m6q
# vaWNM391n42jW+CejYRFUpQxTjIFLdaKVR4TZeytqhO4I98faNsmoXRdMWHJlJ5q
# fIenFv/Kz+Fv86QBy8+7u8bDTuKK+p750wjgcxh0laK0WcUQ6//VHA5aNmjSGNzw
# R8bEERY9u1q/dxyFYhFX8pOvpZsfSMa0OZiwiglNqKS3jxr0B88Iv6hYEAbmkoeC
# hThjUXSzlROzvIWhvLxGRIPLbd4CXs8BMW5PP0lA0SLAmAgdxrJUJnKfGP+BC60F
# dLzJCvwz2Q==
# SIG # End signature block
