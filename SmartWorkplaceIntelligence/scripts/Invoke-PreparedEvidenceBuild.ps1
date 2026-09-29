[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputRoot,
    [ValidateSet('All','Workforce','Devices','Applications','Collaboration','Content','Mailbox','Licensing','Security','Backup','Trust','Executive')]
    [string]$Only = 'All',
    [switch]$SkipWorkforceHistory,
    [switch]$ReuseStagedWorkforce,
    [string]$AccountClassificationConfigPath
)

# Offline preparation only. No collection, upload, Desktop control or promotion.
# Each child process releases its memory before the next domain starts.
$ErrorActionPreference = 'Stop'
$DataRoot = (Resolve-Path -LiteralPath $DataRoot).ProviderPath
$last = Join-Path $DataRoot 'DATA-LAST'
if (-not (Test-Path -LiteralPath $last -PathType Container)) { throw 'DATA-LAST is required.' }
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
if ($OutputRoot.TrimEnd('\') -eq $DataRoot.TrimEnd('\') -or $OutputRoot.TrimEnd('\') -eq $last.TrimEnd('\')) {
    throw 'Outputs must use a dedicated staging directory, never the raw source root.'
}
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
$shell = (Get-Process -Id $PID).Path
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
function OutPath([string]$Name) { Join-Path $OutputRoot ($Name + '.csv') }
function Run([string]$Domain, [string]$Script, [object[]]$Arguments) {
    if ($Only -ne 'All' -and $Only -ne $Domain) { return }
    $started = [datetime]::UtcNow
    Write-Host ('[{0:O}] Starting {1}' -f $started, $Domain)
    & $shell -NoProfile -File (Join-Path $PSScriptRoot $Script) @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Preparation failed: $Domain (exit $LASTEXITCODE). Nothing was published." }
    Write-Host ('[{0:O}] Completed {1}: {2:n1} seconds' -f [datetime]::UtcNow, $Domain, ([datetime]::UtcNow - $started).TotalSeconds)
}
$wf = @('-DataRoot',$DataRoot,'-UserOutputPath',(OutPath 'UserInventoryEvidence'),'-HistoryOutputPath',(OutPath 'UserActivityHistoryEvidence'),'-TrendOutputPath',(OutPath 'WorkforceTrendEvidence'),'-IdentityOutputPath',(OutPath 'IdentityReconciliationEvidence'),'-SignalsOutputPath',(OutPath 'WorkforceOperationalSignals'))
if ($AccountClassificationConfigPath) { $wf += @('-AccountClassificationConfigPath',$AccountClassificationConfigPath) }
if ($SkipWorkforceHistory) { $wf += '-SkipHistory' }
if ($ReuseStagedWorkforce) {
    foreach ($name in @('UserInventoryEvidence','IdentityReconciliationEvidence','WorkforceOperationalSignals')) {
        if (-not (Test-Path -LiteralPath (OutPath $name))) { throw "Missing staged workforce output: $name" }
    }
} else { Run 'Workforce' 'New-WorkforceIdentityEvidence.ps1' $wf }
Run 'Devices' 'New-DeviceInventoryEvidence.ps1' @('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'DeviceInventoryEvidence'),'-SignalsOutputPath',(OutPath 'DeviceRiskEvidenceSummary'),'-DirectorySummaryOutputPath',(OutPath 'DeviceDirectorySummary'))
Run 'Devices' 'New-DeviceLifecycleExperienceTrendEvidence.ps1' @('-DataRoot',$DataRoot,'-WindowsOutputPath',(OutPath 'WindowsLifecycleTrendEvidence'),'-EndpointOutputPath',(OutPath 'EndpointExperienceTrendEvidence'))
Run 'Applications' 'New-ApplicationInventoryEvidence.ps1' @('-DataRoot',$DataRoot,'-InventoryOutputPath',(OutPath 'ApplicationInventoryEvidence'),'-TrendOutputPath',(OutPath 'ApplicationTrendEvidence'))
Run 'Collaboration' 'New-CollaborationEvidence.ps1' @('-DataRoot',$DataRoot,'-AdoptionOutputPath',(OutPath 'CollaborationAdoptionEvidence'),'-ObjectOutputPath',(OutPath 'CollaborationObjectEvidence'),'-TrendOutputPath',(OutPath 'CollaborationTrendEvidence'))
Run 'Content' 'New-ContentStorageEvidence.ps1' @('-DataRoot',$DataRoot,'-ObjectOutputPath',(OutPath 'ContentStorageEvidence'),'-TrendOutputPath',(OutPath 'ContentStorageTrendEvidence'))
Run 'Mailbox' 'New-MailboxEvidence.ps1' @('-DataRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-OutputPath',(OutPath 'MailboxEvidence'))
Run 'Licensing' 'New-LicensingEvidence.ps1' @('-DataRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-MailboxEvidencePath',(OutPath 'MailboxEvidence'),'-OutputPath',(OutPath 'LicenseEvidence'),'-OptimizationOutputPath',(OutPath 'LicenseOptimizationEvidence'))
Run 'Security' 'New-SecurityEvidence.ps1' @('-DataRoot',$last,'-ExtendedEvidenceRoot',$last,'-UserEvidencePath',(OutPath 'UserInventoryEvidence'),'-DeviceEvidencePath',(OutPath 'DeviceInventoryEvidence'),'-OutputPath',(OutPath 'SecurityControlEvidence'))
Run 'Backup' 'New-BackupResilienceEvidence.ps1' @('-DataRoot',$last,'-OutputRoot',$OutputRoot)
Run 'Trust' 'New-DataTrustEvidence.ps1' @('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'DataTrustEvidence'))
Run 'Executive' 'New-ExecutiveTrendEvidence.ps1' @('-DataRoot',$DataRoot,'-OutputPath',(OutPath 'ExecutiveKPITrends'))
Write-Host 'Preparation finished. Outputs are staging only; schema/data validation and promotion are separate steps.'
