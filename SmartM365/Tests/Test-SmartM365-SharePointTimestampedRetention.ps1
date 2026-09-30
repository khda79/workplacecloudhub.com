<#
.SYNOPSIS
Offline tests for the seven-day SharePoint retention of timestamped CSV copies (SmartM365.Core and
the Windows PowerShell 5 compatibility module).
.VERSION
1.0.0
#>
[CmdletBinding()]
param([switch]$WindowsPowerShell5)

$ErrorActionPreference = 'Stop'
$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$modulePath = if ($WindowsPowerShell5) {
    Join-Path -Path $smartM365Root -ChildPath 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
}
else {
    Join-Path -Path $smartM365Root -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
}
$minimumVersion = if ($WindowsPowerShell5) { '1.0.45' } else { '1.0.61' }
$module = Import-Module -Name $modulePath -MinimumVersion $minimumVersion -Force -PassThru -ErrorAction Stop

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

$retention = & $module {
    param($Ps5)
    $script:Deletes = New-Object System.Collections.ArrayList
    $script:Lists = New-Object System.Collections.ArrayList
    $script:Logs = New-Object System.Collections.ArrayList
    function script:Connect-SmartM365GraphForSharePointUpload { param($AppId, $TenantId, $Thumbprint) return $true }
    function script:ConvertTo-SmartM365SharePointDataRootPath { param($TargetFolderPath) return 'SMART-M365/DATA' }
    function script:Get-SmartM365SharePointRelativeFilePath { param($LocalFilePath) return 'DATA-ALL/Intune/Alerts/Report_20260930_120000.csv' }
    function script:WriteLog { param($Message, $Level) [void]$script:Logs.Add(("{0}|{1}" -f $Level, $Message)) }
    $file = { param($Id, $Name) [pscustomobject]@{ id = $Id; name = $Name; file = [pscustomobject]@{ mimeType = 'text/csv' } } }
    $page1 = [pscustomobject]@{
        value = @(
            (& $file 'current' 'Report_20260930_120000.csv'),
            (& $file 'recent' 'Report_20260925_010000.csv'),
            (& $file 'expired1' 'Report_20260920_010000.csv'),
            (& $file 'latest' 'Report.csv'),
            (& $file 'otherprefix' 'Report_Detail_20260901_010000.csv'),
            (& $file 'otherext' 'Report_20260901_010000.xlsx'),
            [pscustomobject]@{ id = 'folder'; name = 'Report_20260901_010000.csv'; folder = [pscustomobject]@{ childCount = 1 } }
        )
        '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/next-page'
    }
    $page2 = [pscustomobject]@{ value = @((& $file 'expired2' 'Report_20260801-230000.csv'), (& $file 'failing' 'Report_20260810_010000.csv')) }
    $script:Pages = @{ first = $page1; second = $page2 }
    function script:Invoke-SmartM365GraphRestWithRetry {
        param($Method, $Uri, $Body, $ContentType, $Operation, $AdditionalHeaders)
        if ($Method -eq 'DELETE') {
            if ($Uri -like '*/items/failing') { throw 'Graph request failed. Status=403' }
            [void]$script:Deletes.Add(($Uri -replace '^.*/items/', ''))
            return $null
        }
        if ($Uri -like '*/sites/*/drives') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'drive1'; name = 'Documents' }) } }
        if ($Uri -like '*/sites/*') { return [pscustomobject]@{ id = 'site1' } }
        [void]$script:Lists.Add($Uri)
        if ($Uri -eq 'https://graph.microsoft.com/v1.0/next-page') { return $script:Pages.second }
        return $script:Pages.first
    }
    function script:Invoke-SmartM365GraphDeleteQuietly {
        param($Uri, $Operation)
        if ($Uri -like '*/items/failing') { return [pscustomobject]@{ Success = $false; NotFound = $false; Message = 'HTTP 403' } }
        [void]$script:Deletes.Add(($Uri -replace '^.*/items/', ''))
        return [pscustomobject]@{ Success = $true; NotFound = $false; Message = '' }
    }
    $result = Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath 'C:\Data\DATA-ALL\Intune\Alerts\Report_20260930_120000.csv' -RetentionDays 7 -ReferenceTime ([datetime]'2026-09-30T13:00:00') -Enabled $true -SiteHostname 'tenant.sharepoint.test' -SitePath '/sites/S' -LibraryDisplayName 'Documents' -TargetFolderPath 'SMART-M365'
    $notTimestamped = Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath 'C:\Data\DATA-LAST\Report.csv' -Enabled $true -SiteHostname 'h' -SitePath '/s' -LibraryDisplayName 'Documents' -TargetFolderPath 'T'
    New-Object psobject -Property @{
        Result = $result; NotTimestamped = $notTimestamped
        Deletes = @($script:Deletes | Sort-Object); Lists = @($script:Lists); Logs = @($script:Logs)
    }
} $WindowsPowerShell5.IsPresent

Assert-True (($retention.Deletes -join ',') -eq 'expired1,expired2') ("expired copies deleted: expected expired1,expired2, got {0}" -f ($retention.Deletes -join ','))
Assert-True ($retention.Result.Deleted -eq 2 -and $retention.Result.Failed -eq 1 -and $retention.Result.Kept -eq 1) ("counters Deleted/Failed/Kept = {0}/{1}/{2}" -f $retention.Result.Deleted, $retention.Result.Failed, $retention.Result.Kept)
Assert-True ($retention.Lists.Count -eq 2 -and $retention.Lists[0] -like '*/drives/drive1/root:/SMART-M365/DATA/DATA-ALL/Intune/Alerts:/children*') ("folder listing followed the next link from the CSV folder: {0}" -f ($retention.Lists -join ' | '))
Assert-True (@($retention.Logs | Where-Object { $_ -like 'WARNING|*Report_20260810_010000.csv*' }).Count -eq 1) 'a failed deletion is reported as a warning'
Assert-True ($retention.NotTimestamped.Deleted -eq 0 -and $retention.NotTimestamped.Failed -eq 0) 'a latest (non-timestamped) path never triggers retention'

$publish = & $module {
    $script:RetentionCalls = New-Object System.Collections.ArrayList
    $script:UploadTimestamped = $true
    function script:Invoke-SmartM365SharePointCsvUpload {
        param($LocalFilePath)
        if (-not $script:UploadTimestamped -and $LocalFilePath -match '_\d{8}_\d{6}\.csv$') { return $null }
        return [pscustomobject]@{ LocalFilePath = $LocalFilePath; SharePointPath = 'x'; WebUrl = '' }
    }
    function script:Remove-SmartM365SharePointTimestampedCsvOlderThan { param($TimestampedPath, $RetentionDays) [void]$script:RetentionCalls.Add(("{0}|{1}" -f [IO.Path]::GetFileName($TimestampedPath), $RetentionDays)) }
    function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param($SourceFiles, $TimestampedPath) }
    function script:WriteLog { param($Message, $Level) }
    $root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SpRetention-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path (Join-Path $root 'all'), (Join-Path $root 'last') -Force | Out-Null
    $global:RetentionMaxCSV = 0
    try {
        $data = @([pscustomobject]@{ Name = 'a'; Value = '1' })
        $run = {
            param($Stamp, [int]$Days = 7)
            # The PS5 module has no -NoWeeklyHistory; the history publication is stubbed in both modules.
            $extra = @{}
            if ((Get-Command Publish-SmartM365Csv).Parameters.ContainsKey('NoWeeklyHistory')) { $extra['NoWeeklyHistory'] = $true }
            Publish-SmartM365Csv -Data $data -TimestampedPath (Join-Path $root "all\Report_$Stamp.csv") -LatestPath (Join-Path $root 'last\Report.csv') -Columns @('Name', 'Value') -NoTenantKey -RetentionMaxCsv 0 -SharePointRetentionDays $Days @extra | Out-Null
        }
        & $run '20260930_120000'
        $script:UploadTimestamped = $false
        & $run '20260930_130000'
        $script:UploadTimestamped = $true
        & $run '20260930_140000' 0
        @($script:RetentionCalls)
    }
    finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
}
Assert-True ((@($publish) -join ',') -eq 'Report_20260930_120000.csv|7') ("Publish-SmartM365Csv retention calls: expected only the uploaded run with 7 days, got '{0}'" -f (@($publish) -join ','))

$edition = if ($WindowsPowerShell5) { 'WindowsPowerShell5' } else { 'Core' }
Write-Host ("SHAREPOINT_TIMESTAMPED_RETENTION_TEST_OK Module={0}; Deleted={1}; Failed={2}; Kept={3}" -f $edition, $retention.Result.Deleted, $retention.Result.Failed, $retention.Result.Kept)
if (-not $WindowsPowerShell5 -and $PSVersionTable.PSEdition -eq 'Core') {
    $ps5 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $ps5) {
        & $ps5 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -WindowsPowerShell5
        if ($LASTEXITCODE -ne 0) { throw "Windows PowerShell 5 compatibility test failed (exit $LASTEXITCODE)." }
    }
}
