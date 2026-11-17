<#
.SYNOPSIS
Uploads to SharePoint the DATA-ALL files written while DATA-ALL uploads were disabled.

.DESCRIPTION
When EnableSharePointUploadDataAll (or the legacy EnableSharePointUpload) is off, the collectors keep
writing DATA-ALL on the server but nothing reaches SharePoint. Turning the setting back on only
uploads the next runs. This script uploads the gap once, with the same SmartM365.Core upload as the
collectors (large-file sessions, time limits, JSON name transition, WeeklyHistory manifest check).

A file is planned for upload when all of these hold:
- it is under DATA-ALL, outside the Orchestrator folder (mirrored by the orchestrator itself) and
  outside hidden folders (.staging, .locks...);
- it is not a lock, temporary or partial file;
- it was written on or after -Since and at least -MinimumAgeMinutes ago (a file still being written
  by a running collector is left to that collector);
- for a timestamped run copy (<prefix>_yyyyMMdd_HHmmss.csv or .xlsx), its run timestamp is within
  -TimestampedDays: the collectors delete older copies from SharePoint at their next upload, so
  uploading them would be useless;
- SharePoint does not already hold a file of the same name and size in the target folder.
Every other file written since -Since is reported with its reason.

Without -Execute the script only reads: it lists the server files and the SharePoint folders and
prints the plan. With -Execute it uploads the planned files, oldest first, and writes a CSV report of
every decision. Execute is refused when the tenant does not enable DATA-ALL uploads.

Run it from a collection server, with the account of the orchestrator scheduled task, so the tenant
profile, the certificate and the DATA-ALL share are the ones the collectors use.

.PARAMETER Tenant
Tenant profile key. Default prod.

.PARAMETER Since
Oldest last-write time considered. Use the day DATA-ALL uploads stopped.

.PARAMETER TimestampedDays
Age limit of the timestamped run copies, from their file-name timestamp. Default 7, the SharePoint
retention applied by the collectors.

.PARAMETER MinimumAgeMinutes
Files written more recently are skipped. Default 30.

.PARAMETER DataAllRootPath
DATA-ALL root. Defaults to DataAllRootPath of the tenant profile.

.PARAMETER MaximumFiles
Uploads at most this many files (0 = all). Lets a long backfill run in several passes.

.PARAMETER ReportPath
CSV report of every decision. Defaults to %TEMP%\SmartM365-DataAllBackfill_<timestamp>.csv.

.PARAMETER Execute
Uploads the planned files. Omit for preview only.

.EXAMPLE
pwsh -NoProfile -File .\Publish-SmartM365-DataAllBackfill.ps1 -Since 2026-10-24
Preview of the files written since 24 October.

.EXAMPLE
pwsh -NoProfile -File .\Publish-SmartM365-DataAllBackfill.ps1 -Since 2026-10-24 -Execute
Uploads them.

.VERSION
1.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$Tenant = 'prod',

    [Parameter(Mandatory = $true)]
    [datetime]$Since,

    [ValidateRange(1, 3650)]
    [int]$TimestampedDays = 7,

    [ValidateRange(0, 1440)]
    [int]$MinimumAgeMinutes = 30,

    [string]$DataAllRootPath = '',

    [ValidateRange(0, 1000000)]
    [int]$MaximumFiles = 0,

    [string]$ReportPath = '',

    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0'

function Get-DataAllBackfillPlan {
    # Pure selection of the local files; SharePoint presence is checked afterwards.
    param(
        [Parameter(Mandatory = $true)][string]$DataAllRoot,
        [Parameter(Mandatory = $true)][datetime]$Since,
        [Parameter(Mandatory = $true)][datetime]$TimestampedCutoff,
        [Parameter(Mandatory = $true)][datetime]$NewestAllowed
    )
    $namePattern = '^(?<Prefix>.+)_(?<Stamp>\d{8}[-_]\d{6})\.(csv|xlsx)$'
    $stampFormats = [string[]]@('yyyyMMdd_HHmmss', 'yyyyMMdd-HHmmss')
    $transientPattern = '(?i)\.(lock|tmp|partial|part)$|^~\$'
    $plan = [Collections.Generic.List[object]]::new()
    foreach ($file in @(Get-ChildItem -LiteralPath $DataAllRoot -Recurse -File -Force -ErrorAction Stop)) {
        if ($file.LastWriteTime -lt $Since) { continue }
        $relativeFolder = [IO.Path]::GetRelativePath($DataAllRoot, $file.DirectoryName)
        $segments = @($relativeFolder -split '[\\/]' | Where-Object { $_ -and $_ -ne '.' })
        $reason = ''
        if ($segments.Count -and $segments[0] -ieq 'Orchestrator') { $reason = 'OrchestratorFolder' }
        elseif (@($segments | Where-Object { $_.StartsWith('.') }).Count) { $reason = 'HiddenFolder' }
        elseif ($file.Name -match $transientPattern -or $file.Name.StartsWith('.')) { $reason = 'TransientFile' }
        elseif ($file.LastWriteTime -gt $NewestAllowed) { $reason = 'TooRecent' }
        else {
            $match = [regex]::Match($file.Name, $namePattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $stamp = [datetime]::MinValue
            if ($match.Success -and [datetime]::TryParseExact($match.Groups['Stamp'].Value, $stampFormats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$stamp) -and $stamp -lt $TimestampedCutoff) {
                $reason = 'BeyondSharePointRetention'
            }
        }
        $plan.Add([pscustomobject]@{
            Action = if ($reason) { 'Skip' } else { 'Upload' }; Reason = $reason; Folder = $relativeFolder; Name = $file.Name
            SizeBytes = $file.Length; LastWriteTime = $file.LastWriteTime; Result = ''; FullName = $file.FullName
        })
    }
    return @($plan | Sort-Object LastWriteTime, FullName)
}

# Lists one SharePoint folder (name -> size); runs in the SmartM365.Core scope to reuse its Graph helpers.
$listRemoteFolder = {
    param([Parameter(Mandatory = $true)][string]$FolderPath)
    if (-not $script:SmartM365BackfillDriveId) {
        if (-not (Connect-SmartM365GraphForSharePointUpload -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $global:Thumbprint)) { throw 'Microsoft Graph connection for SharePoint failed.' }
        $site = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri ('https://graph.microsoft.com/v1.0/sites/{0}:{1}' -f $global:SharePointSiteHostname, $global:SharePointSitePath) -Operation 'Resolve SharePoint site'
        $drives = @((Invoke-SmartM365GraphRestWithRetry -Method GET -Uri ('https://graph.microsoft.com/v1.0/sites/{0}/drives' -f $site.id) -Operation 'Resolve SharePoint document libraries').value)
        $normalize = { param($Text) if ($null -eq $Text) { '' } else { ([string]$Text).Normalize([Text.NormalizationForm]::FormD) -replace '\p{M}', '' } }
        $wanted = & $normalize $global:SharePointLibraryDisplayName
        $drive = @($drives | Where-Object { (& $normalize $_.name) -ieq $wanted } | Select-Object -First 1)
        if (-not $drive.Count) { throw ("Document library '{0}' not found." -f $global:SharePointLibraryDisplayName) }
        $script:SmartM365BackfillDriveId = [string]$drive[0].id
    }
    $sizes = @{}
    $uri = 'https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/children?$select=name,size,file' -f $script:SmartM365BackfillDriveId, (ConvertTo-GraphDrivePath $FolderPath)
    while ($uri) {
        try { $page = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri $uri -Operation 'List SharePoint folder' }
        catch {
            if ([string]$_.Exception.Message -match '(?i)Status=404|itemNotFound|404 \(Not Found\)') { return $sizes }
            throw
        }
        foreach ($item in @($page.value)) { if ($item.PSObject.Properties['file']) { $sizes[[string]$item.name] = [long]$item.size } }
        $next = $page.PSObject.Properties['@odata.nextLink']
        $uri = if ($next) { [string]$next.Value } else { '' }
    }
    return $sizes
}

#region Main

$smartM365Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
. (Join-Path $smartM365Root 'Config\SmartM365-TenantContext.ps1')
$tenantConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
Import-Module (Join-Path $smartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.117' -ErrorAction Stop
$core = Get-Module -Name SmartM365.Core

function Get-BackfillConfigValue {
    param([Parameter(Mandatory = $true)][string]$Name, $Default = '')
    $value = $null
    if ($tenantConfig -is [Collections.IDictionary]) { if ($tenantConfig.Contains($Name)) { $value = $tenantConfig[$Name] } }
    elseif ($null -ne $tenantConfig -and $tenantConfig.PSObject.Properties[$Name]) { $value = $tenantConfig.$Name }
    if ($null -eq $value -or ($value -is [string] -and ([string]::IsNullOrWhiteSpace($value) -or $value.Trim() -in @('__USE_GLOBAL__', 'USE_GLOBAL')))) { return $Default }
    return $value
}

$scriptName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$workFolder = Join-Path ([IO.Path]::GetTempPath()) ('{0}_{1}' -f $scriptName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
$null = InitializeScriptEnvironment -OutputPathInit $workFolder -LogFileName $scriptName -CallerScriptPath $PSCommandPath

$global:AppId = [string](Get-BackfillConfigValue -Name 'AppId')
$global:TenantId = [string](Get-BackfillConfigValue -Name 'TenantId')
$global:Thumbprint = [string](Get-BackfillConfigValue -Name 'Thumbprint' -Default (Get-BackfillConfigValue -Name 'Thumb'))
$global:SharePointSiteHostname = [string](Get-BackfillConfigValue -Name 'SharePointSiteHostname')
$global:SharePointSitePath = [string](Get-BackfillConfigValue -Name 'SharePointSitePath')
$global:SharePointLibraryDisplayName = [string](Get-BackfillConfigValue -Name 'SharePointLibraryDisplayName' -Default 'Documents')
$global:SharePointTargetFolderPath = [string](Get-BackfillConfigValue -Name 'SharePointTargetFolderPath')
foreach ($required in @('AppId', 'TenantId', 'Thumbprint', 'SharePointSiteHostname', 'SharePointSitePath', 'SharePointTargetFolderPath')) {
    if ([string]::IsNullOrWhiteSpace([string](Get-Variable -Name $required -Scope Global -ValueOnly))) { throw "Tenant profile '$Tenant' has no $required." }
}

$legacyUpload = Get-BackfillConfigValue -Name 'EnableSharePointUpload' -Default $false
$policy = & $core {
    param($Legacy)
    Resolve-SmartM365SharePointUploadPolicy -Category DataAll -LocalConfig $null -GlobalConfig (Get-SmartM365EffectiveModuleGlobalConfig) -LegacyEnabled ([bool]($Legacy -eq $true -or [string]$Legacy -in @('true', 'True', '1')))
} $legacyUpload
$policyText = 'DATA-ALL uploads for tenant {0}: {1} ({2}).' -f $Tenant, $(if ($policy.Enabled) { 'enabled' } else { 'disabled' }), $policy.Setting
WriteLog -Message $policyText -Level $(if ($policy.Enabled) { 'INFO' } else { 'WARNING' })
if ($Execute -and -not $policy.Enabled) { throw "$policyText Enable EnableSharePointUploadDataAll before -Execute, otherwise the next runs would not keep SharePoint current either." }

if ([string]::IsNullOrWhiteSpace($DataAllRootPath)) {
    $DataAllRootPath = [string](& $core { Get-ModuleLocalConfigValue -Config (Get-ModuleLocalConfig) -Name 'DataAllRootPath' -DefaultValue '' })
}
if ([string]::IsNullOrWhiteSpace($DataAllRootPath)) { throw 'DataAllRootPath is not configured; pass -DataAllRootPath.' }
$dataAllRoot = [IO.Path]::GetFullPath($DataAllRootPath).TrimEnd('\', '/')
if ((Split-Path -Path $dataAllRoot -Leaf) -ne 'DATA-ALL') { throw "DataAllRootPath must be a DATA-ALL folder: $dataAllRoot" }
if (-not (Test-Path -LiteralPath $dataAllRoot -PathType Container)) { throw "DATA-ALL root not found: $dataAllRoot" }
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DataAllBackfill_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

$now = Get-Date
WriteLog -Message ('Scanning {0}: written since {1:yyyy-MM-dd HH:mm}, timestamped copies since {2:yyyy-MM-dd HH:mm}, skip files newer than {3} min (script v{4}).' -f $dataAllRoot, $Since, $now.AddDays(-$TimestampedDays), $MinimumAgeMinutes, $ScriptVersion)
$plan = @(Get-DataAllBackfillPlan -DataAllRoot $dataAllRoot -Since $Since -TimestampedCutoff $now.AddDays(-$TimestampedDays) -NewestAllowed $now.AddMinutes(-$MinimumAgeMinutes))

# SharePoint presence: one listing per folder, same name and size means already uploaded.
$targetRoot = ConvertTo-SmartM365SharePointDataRootPath -TargetFolderPath $global:SharePointTargetFolderPath
$listInCore = $core.NewBoundScriptBlock($listRemoteFolder)
$remoteFolders = @{}
foreach ($entry in @($plan | Where-Object Action -eq 'Upload')) {
    $relativeFile = Get-SmartM365SharePointRelativeFilePath -LocalFilePath $entry.FullName
    $remoteFolder = $targetRoot.TrimEnd('/') + '/' + (($relativeFile -replace '/[^/]+$', '').Trim('/'))
    if (-not $remoteFolders.ContainsKey($remoteFolder)) { $remoteFolders[$remoteFolder] = & $listInCore $remoteFolder }
    $remote = $remoteFolders[$remoteFolder]
    if ($remote.ContainsKey($entry.Name) -and [long]$remote[$entry.Name] -eq [long]$entry.SizeBytes) {
        $entry.Action = 'Skip'
        $entry.Reason = 'AlreadyOnSharePoint'
    }
}

$toUpload = @($plan | Where-Object Action -eq 'Upload')
if ($MaximumFiles -gt 0 -and $toUpload.Count -gt $MaximumFiles) {
    foreach ($entry in @($toUpload | Select-Object -Skip $MaximumFiles)) { $entry.Action = 'Skip'; $entry.Reason = 'NextPass' }
    $toUpload = @($toUpload | Select-Object -First $MaximumFiles)
}
$skipped = @($plan | Where-Object Action -eq 'Skip')
Write-Host ''
Write-Host ('Files written since {0:yyyy-MM-dd}: {1}; to upload: {2} ({3:N0} MB) from {4} SharePoint folder(s); skipped: {5}.' -f $Since, $plan.Count, $toUpload.Count, (($toUpload | Measure-Object SizeBytes -Sum).Sum / 1MB), $remoteFolders.Count, $skipped.Count)
if ($toUpload.Count) {
    $toUpload | Group-Object Folder | ForEach-Object {
        [pscustomobject]@{ Folder = $_.Name; Files = $_.Count; MB = [math]::Round(($_.Group | Measure-Object SizeBytes -Sum).Sum / 1MB, 1) }
    } | Sort-Object MB -Descending | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
}
if ($skipped.Count) {
    Write-Host 'Skipped, by reason:'
    $skipped | Group-Object Reason | ForEach-Object {
        [pscustomobject]@{ Reason = $_.Name; Files = $_.Count; MB = [math]::Round(($_.Group | Measure-Object SizeBytes -Sum).Sum / 1MB, 1) }
    } | Sort-Object Files -Descending | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
}

$failed = 0
if ($Execute) {
    $index = 0
    foreach ($entry in $toUpload) {
        $index++
        Write-Host ('[{0}/{1}] {2}\{3} ({4:N1} MB)' -f $index, $toUpload.Count, $entry.Folder, $entry.Name, ($entry.SizeBytes / 1MB))
        $record = $null
        try { $record = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $entry.FullName -Enabled $true -EnsureParentFolders }
        catch { $entry.Result = 'Failed: ' + $_.Exception.Message }
        if (-not $entry.Result) { $entry.Result = if ($null -ne $record -and $record.PSObject.Properties['SharePointPath']) { 'Uploaded' } else { 'Failed: see the run log' } }
        if ($entry.Result -ne 'Uploaded') { $failed++ }
    }
    WriteLog -Message ('DATA-ALL backfill: uploaded {0}; failed {1}.' -f ($toUpload.Count - $failed), $failed) -Level $(if ($failed) { 'WARNING' } else { 'INFO' })
}
else {
    Write-Host 'Preview only. Run again with -Execute to upload the files listed above.'
}
$plan | Select-Object Action, Reason, Result, Folder, Name, SizeBytes, LastWriteTime | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding utf8
Write-Host ('Report: {0} (script v{1})' -f $ReportPath, $ScriptVersion)
if ($failed -gt 0) { exit 1 }

#endregion
