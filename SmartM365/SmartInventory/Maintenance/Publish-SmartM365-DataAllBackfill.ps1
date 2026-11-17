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
- it is not in a folder that no collector uploads (NotPublishedByCollector): the Intune remediation
  script exports (Intune\Remediations\Exports, script sources stay on the server), the Windows Update
  status work and archive folders (Intune\WindowsUpdate\Status\Work and \Archive), the Discovered
  Apps checkpoints (Intune\Applications\DiscoveredApps\ResumeHistory) and the daily capture metadata
  (DailyCaptureRecords, any depth);
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
1.2

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
1.2 - NotPublishedByCollector also covers the Windows Update status Archive and the Discovered Apps
      ResumeHistory; .VERSION holds the number only.
1.1 - NotPublishedByCollector: remediation exports, Windows Update Work folder, DailyCaptureRecords.
1.0 - First version.
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
$ScriptVersion = '1.2'

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
    # Folders that no collector uploads: copying them would publish what the pipeline keeps local.
    $notPublishedPrefixes = @(
        @('Intune', 'Remediations', 'Exports'),
        @('Intune', 'WindowsUpdate', 'Status', 'Work'),
        @('Intune', 'WindowsUpdate', 'Status', 'Archive'),
        @('Intune', 'Applications', 'DiscoveredApps', 'ResumeHistory')
    )
    $plan = [Collections.Generic.List[object]]::new()
    foreach ($file in @(Get-ChildItem -LiteralPath $DataAllRoot -Recurse -File -Force -ErrorAction Stop)) {
        if ($file.LastWriteTime -lt $Since) { continue }
        $relativeFolder = [IO.Path]::GetRelativePath($DataAllRoot, $file.DirectoryName)
        $segments = @($relativeFolder -split '[\\/]' | Where-Object { $_ -and $_ -ne '.' })
        $reason = ''
        if ($segments.Count -and $segments[0] -ieq 'Orchestrator') { $reason = 'OrchestratorFolder' }
        elseif (@($segments | Where-Object { $_.StartsWith('.') }).Count) { $reason = 'HiddenFolder' }
        elseif (@($segments | Where-Object { $_ -ieq 'DailyCaptureRecords' }).Count -or @($notPublishedPrefixes | Where-Object {
                $prefix = $_
                $segments.Count -ge $prefix.Count -and -not @(0..($prefix.Count - 1) | Where-Object { $segments[$_] -ine $prefix[$_] }).Count
            }).Count) { $reason = 'NotPublishedByCollector' }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCw27FMtdeIyHQA
# WkXWvKEbiPXQdW/hza18qz+5vF5PE6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKUgCvmEoBJ4XdYsR1jAu5K+ttkrS1YAJSUqHrtxoZ56MA0GCSqG
# SIb3DQEBAQUABIIBgEgNdTpKYfz38FIzD56Dw7FqzX/jWeyPZBBxZXZJNFVytaal
# x7oVs+XlcmmHqrCxWuni1phgTx50Jce1h+zKjUy1W5SqYt8o2/ZWxTbGScV4hWqi
# viu64FyVCXdWesSRDUSg2jFbNMFrtn9Vp7UiaK0Ug3vQxeF6uxrd4aYDFgUELsNc
# vng1u5soLDZBugXHmPlw5DiOEVigUzQkkrlZGHFreJ3OVVicptfMZmWTvfmyB8tq
# 3g38bMifzQWJUIh5BkNOUysFkNed4Uw4Byj5PkPhG/HPxeg2e3btUf2l75yCHUDC
# LHRzU9YPC811z4qOWZsBaO78kpBPaXTFgQ89+nLc3DVpXHF9oUcyP9vxIQc29mwY
# /Pt0cuj+KjsRoEyCvWGk4EdsqQMSyCD0XkoJO+6/ghbmaeCDSAP+InGtqVwypUI9
# qKRywKrLDmObPUm/IYM0f21uoKJlVBX/1yjgKi/4nKJS6jeL6hqBnEBOBN7AgggE
# ECatoPQID2tcEdexWKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTcxMjA4
# NTBaMC8GCSqGSIb3DQEJBDEiBCD756j2WACnZ7D3YBGKRtkohHNcBuIL3TLh7grl
# nfuJNjANBgkqhkiG9w0BAQEFAASCAgAQYkWfRyn8fXs0YvktA0IAc6izu7Mql7rl
# nO/lXARhJpZTccygEL4Cvw93dEe20XvYp0urbREyacYlffNzc2Bw7dvt1NlNxsne
# 48USpGUSsa83P4JIkY2W0DaRsLVxInfBTIusrhFIWEP5rk+qiryjv+nI7ia0lyQh
# ewbym2kS0NZKenwuAp8goo30V/HokCAQQJY2PpGFIsIysd7rA9QL9op4PLGP5sS+
# KrJnCflLF+IqQOZq3qUz2io1lvfIERdVWyNvjxk70fsztxkmZdFvSRKv5wOWeuER
# wktLL02LVF6DyB6JQQwkjfwSYxKN0GmXCE/O7/aDXGXq9OucCTt3ixQnpJS6wVJm
# XS5ymJ5l2W8TU261UfDJ9vH2dOn4flyWBRm66iynzYFhseahMtz0sOA7Aov+YxPK
# kLBFe+b0KH4I1pAGCEh7ynQ+E9Mf+/f3kZZ8HkVfI6mJvaFZcK+XQhWbOZZRdslu
# 6ulSEOevzCr12/CEJkcR8TYtBw8HEHsQkWOfiUCSvxEtAmI8PjVKeekVMVIlnLDj
# 6KwHAJAaAZjeyQsKtKRdGpmCgay5PAIalzT8mXqcGoCrEOKEshqLrLBJEngYutda
# v47d/JCDba2HpDVu8zzU3GYiIqUJPCeQjImdTTvUxfsXOmQQEnh6KcPnOnIAlruN
# YuC/QQswvg==
# SIG # End signature block
