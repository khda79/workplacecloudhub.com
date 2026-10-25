<#
.SYNOPSIS
    CSV and HTML reports of Runs.

.DESCRIPTION
    A report state keeps the latest result per computer and the attempt counters of the Run.
    Live and final HTML reports share one renderer with WorkplaceCloudHub branding; the
    per-cycle CSV files keep the complete attempt history.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

function ConvertTo-SemtHtmlText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ConvertTo-SemtFileUri {
    [CmdletBinding()]
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $value = ([string]$Path).Trim()
    try {
        if ($value.StartsWith('\\')) {
            $parts = @($value.TrimStart('\') -split '\\' | Where-Object { $_ -ne '' })
            if ($parts.Count -ge 2) {
                $escaped = @($parts | ForEach-Object { [System.Uri]::EscapeDataString($_) })
                return ('file://{0}/{1}' -f $escaped[0],(($escaped | Select-Object -Skip 1) -join '/'))
            }
        }
        return ([System.Uri]::new([System.IO.Path]::GetFullPath($value))).AbsoluteUri
    }
    catch { return '' }
}

function New-SemtHtmlLink {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Path,
        [string]$Label = 'Open'
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return '-' }
    $uri = ConvertTo-SemtFileUri -Path $Path
    if ([string]::IsNullOrWhiteSpace($uri)) { return (ConvertTo-SemtHtmlText $Path) }
    return ('<a href="{0}" title="{1}">{2}</a>' -f (ConvertTo-SemtHtmlText $uri),(ConvertTo-SemtHtmlText $Path),(ConvertTo-SemtHtmlText $Label))
}

function ConvertTo-SemtHtmlTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [string[]]$Columns,
        [string[]]$RawHtmlColumns = @()
    )

    if (-not $Rows -or $Rows.Count -eq 0) { return "<p class='empty'>No rows.</p>" }
    if (-not $Columns -or $Columns.Count -eq 0) {
        $set = New-Object System.Collections.Generic.List[string]
        foreach ($row in $Rows) { foreach ($property in $row.PSObject.Properties) { if (-not $set.Contains($property.Name)) { $set.Add($property.Name) } } }
        $Columns = @($set.ToArray())
    }
    $html = New-Object System.Text.StringBuilder
    [void]$html.Append('<table><tr>')
    foreach ($column in $Columns) { [void]$html.Append(('<th>{0}</th>' -f (ConvertTo-SemtHtmlText $column))) }
    [void]$html.Append('</tr>')
    foreach ($row in $Rows) {
        [void]$html.Append('<tr>')
        foreach ($column in $Columns) {
            $value = Get-SemtPropertyValue -InputObject $row -Name $column -Default ''
            if ($RawHtmlColumns -contains $column) { [void]$html.Append(('<td>{0}</td>' -f $value)) }
            else { [void]$html.Append(('<td>{0}</td>' -f (ConvertTo-SemtHtmlText $value))) }
        }
        [void]$html.Append('</tr>')
    }
    [void]$html.Append('</table>')
    return $html.ToString()
}

function ConvertTo-SemtReportRows {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$Items,
        [Parameter(Mandatory = $true)][string[]]$Columns
    )

    return @(foreach ($item in @($Items)) {
        if ($null -eq $item) { continue }
        $row = [ordered]@{}
        foreach ($column in $Columns) { $row[$column] = [string](Get-SemtPropertyValue -InputObject $item -Name $column -Default '') }
        [pscustomobject]$row
    })
}

function Export-SemtReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [Parameter(Mandatory = $true)][string]$Path
    )

    New-SemtDirectory -Path (Split-Path -Path $Path -Parent)
    $normalized = @(ConvertTo-SemtReportRows -Items $Rows -Columns $Columns)
    if ($normalized.Count -gt 0) {
        $normalized | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
        return
    }
    $header = ($Columns | ForEach-Object { '"{0}"' -f $_ }) -join ','
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
}

function Get-SemtBrandLogoDataUri {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$PackageRoot)

    $candidate = Join-Path -Path $PackageRoot -ChildPath 'Assets\WorkplaceCloudHub-lockup-WPF.png'
    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { return '' }
    try { return ('data:image/png;base64,' + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($candidate))) } catch { return '' }
}

function New-SemtReportState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$RunContext,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [Parameter(Mandatory = $true)][string]$CampaignTitle,
        [string[]]$SeparatedStatuses = @(),
        [string]$LauncherVersion = '',
        [AllowEmptyCollection()][object[]]$OptionRows = @(),
        [AllowEmptyCollection()][object[]]$SecurityRows = @(),
        [AllowEmptyCollection()][object[]]$IntegrityRows = @(),
        [string]$RemoteLogsPathTemplate = ''
    )

    return [pscustomobject]@{
        RunContext = $RunContext
        Columns = @($Columns)
        CampaignTitle = $CampaignTitle
        SeparatedStatuses = @($SeparatedStatuses)
        LauncherVersion = $LauncherVersion
        OptionRows = @($OptionRows)
        SecurityRows = @($SecurityRows)
        IntegrityRows = @($IntegrityRows)
        RemoteLogsPathTemplate = $RemoteLogsPathTemplate
        LatestByComputer = @{}
        StatusCounter = @{}
        NextActionCounter = @{}
        AdminShareFailureCounter = @{}
        AttemptRowCount = 0
        ProgressRows = (New-Object System.Collections.ArrayList)
        LogoUri = Get-SemtBrandLogoDataUri -PackageRoot $RunContext.PackageRoot
        HtmlPath = Join-Path -Path $RunContext.ReportRoot -ChildPath ('SEMT_{0}_{1}_{2}.html' -f $RunContext.Campaign,(Get-SemtSafeName -Value $RunContext.WaveName),$RunContext.RunId)
    }
}

function Add-SemtCounterValue {
    [CmdletBinding()]
    param([hashtable]$Counter,[AllowNull()][string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    if (-not $Counter.ContainsKey($Key)) { $Counter[$Key] = 0 }
    $Counter[$Key] = [int]$Counter[$Key] + 1
}

function Add-SemtReportAttemptRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$State,
        [AllowEmptyCollection()][object[]]$Rows
    )

    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $status = [string](Get-SemtPropertyValue -InputObject $row -Name 'Status' -Default '')
        Add-SemtCounterValue -Counter $State.StatusCounter -Key $status
        Add-SemtCounterValue -Counter $State.NextActionCounter -Key ([string](Get-SemtPropertyValue -InputObject $row -Name 'NextAction' -Default ''))
        if ($status -like '*ADMIN_SHARE_UNREACHABLE') { Add-SemtCounterValue -Counter $State.AdminShareFailureCounter -Key ([string](Get-SemtPropertyValue -InputObject $row -Name 'AdminShareFailureType' -Default 'UNKNOWN')) }
        $key = Get-SemtComputerKey -ComputerName ([string](Get-SemtPropertyValue -InputObject $row -Name 'ComputerName' -Default ''))
        if (-not [string]::IsNullOrWhiteSpace($key)) { $State.LatestByComputer[$key] = $row }
        $State.AttemptRowCount++
    }
}

function New-SemtProgressRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][datetime]$CycleStart,
        [int]$Total,
        [int]$Queued,
        [int]$Completed,
        [int]$Running,
        [AllowNull()]$ComputerListStats
    )

    return [pscustomobject]@{
        Cycle = $CycleNumber
        Started = $CycleStart.ToString('yyyy-MM-dd HH:mm:ss')
        ElapsedMinutes = [math]::Round(((Get-Date) - $CycleStart).TotalMinutes, 1)
        ComputerListLines = if ($ComputerListStats) { [int]$ComputerListStats.RawLines } else { $Total }
        TotalUnique = $Total
        Queued = $Queued
        Completed = $Completed
        Running = $Running
        Remaining = [math]::Max(0, $Total - $Completed - $Running)
        DuplicateLines = if ($ComputerListStats) { [int]$ComputerListStats.DuplicateLines } else { 0 }
        DuplicateSamples = if ($ComputerListStats) { [string]$ComputerListStats.DuplicateSamples } else { '' }
    }
}

function ConvertTo-SemtCounterRows {
    [CmdletBinding()]
    param([hashtable]$Counter,[string]$NameColumn)
    return @($Counter.GetEnumerator() | ForEach-Object { [pscustomobject][ordered]@{ $NameColumn = [string]$_.Key; Count = [int]$_.Value } } | Sort-Object Count,$NameColumn -Descending)
}

function Write-SemtRunHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [switch]$IsLive,
        [AllowEmptyCollection()][object[]]$CurrentCycleRows = @(),
        $CurrentProgress,
        [AllowEmptyCollection()][object[]]$RunningRows = @(),
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { $Path = $State.HtmlPath }
    $latest = @{}
    foreach ($entry in $State.LatestByComputer.GetEnumerator()) { $latest[$entry.Key] = $entry.Value }
    foreach ($row in @($CurrentCycleRows)) {
        $key = Get-SemtComputerKey -ComputerName ([string](Get-SemtPropertyValue -InputObject $row -Name 'ComputerName' -Default ''))
        if (-not [string]::IsNullOrWhiteSpace($key)) { $latest[$key] = $row }
    }
    $latestRows = @($latest.GetEnumerator() | Sort-Object Name | ForEach-Object { $_.Value })
    $decorated = @(foreach ($row in $latestRows) {
        $item = [ordered]@{}
        foreach ($column in $State.Columns) {
            $item[$column] = [string](Get-SemtPropertyValue -InputObject $row -Name $column -Default '')
            if ($column -eq 'Detail') {
                $item['Local log'] = New-SemtHtmlLink -Path ([string](Get-SemtPropertyValue -InputObject $row -Name 'PsExecLogPath' -Default ''))
                $item['Collected evidence'] = New-SemtHtmlLink -Path ([string](Get-SemtPropertyValue -InputObject $row -Name 'RemoteLogsPath' -Default ''))
                $target = [string](Get-SemtPropertyValue -InputObject $row -Name 'ConnectionTarget' -Default '')
                if ([string]::IsNullOrWhiteSpace($target)) { $target = [string](Get-SemtPropertyValue -InputObject $row -Name 'ComputerName' -Default '') }
                $item['Target logs'] = if ([string]::IsNullOrWhiteSpace($State.RemoteLogsPathTemplate) -or [string]::IsNullOrWhiteSpace($target)) { '-' } else { New-SemtHtmlLink -Path ($State.RemoteLogsPathTemplate -f $target) }
            }
        }
        [pscustomobject]$item
    })
    $htmlColumns = New-Object System.Collections.Generic.List[string]
    foreach ($column in $State.Columns) {
        $htmlColumns.Add($column)
        if ($column -eq 'Detail') { $htmlColumns.Add('Local log'); $htmlColumns.Add('Collected evidence'); $htmlColumns.Add('Target logs') }
    }
    $raw = @('Local log','Collected evidence','Target logs')
    $mainRows = @($decorated | Where-Object { $State.SeparatedStatuses -notcontains [string]$_.Status })
    $separatedRows = @($decorated | Where-Object { $State.SeparatedStatuses -contains [string]$_.Status })
    $latestStatus = @($decorated | Group-Object -Property Status | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ Status = $_.Name; Count = $_.Count } })
    $progress = @($State.ProgressRows.ToArray())
    if ($CurrentProgress) { $progress += $CurrentProgress }
    $mode = if ($IsLive) { 'LIVE' } else { 'FINAL' }
    $context = $State.RunContext
    $titleWave = if ($context.SinglePc) { 'Single PC' } else { $context.WaveName }
    $style = @"
<style>
body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 24px; background: #F5F8FB; color: #1F2937; }
.header-card { display: flex; align-items: center; justify-content: space-between; background: #fff; border: 1px solid #DDE7F0; border-radius: 8px; padding: 18px 22px; margin-bottom: 18px; }
.title { font-size: 22px; font-weight: 600; }
.subtitle { font-size: 13px; color: #5F6B7A; margin-top: 2px; }
.wave { font-size: 18px; font-weight: 700; color: #005A9E; margin-top: 8px; }
.meta { font-size: 12px; color: #5F6B7A; margin-top: 6px; }
.badge { display: inline-block; background: #E6F4FF; color: #005A9E; border: 1px solid #B9DDF7; border-radius: 10px; padding: 1px 10px; font-size: 11px; font-weight: 600; margin-left: 8px; vertical-align: middle; }
.logo { height: 46px; }
.card { background: #fff; border: 1px solid #DDE7F0; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; overflow-x: auto; }
h2 { font-size: 15px; margin: 0 0 8px 0; }
table { border-collapse: collapse; width: 100%; font-size: 12px; }
th { background: #0078D4; color: #fff; text-align: left; font-weight: 600; }
th, td { border: 1px solid #DDE7F0; padding: 6px 8px; vertical-align: top; }
tr:nth-child(even) td { background: #F5F8FB; }
.empty { color: #5F6B7A; font-style: italic; }
.footer { font-size: 11px; color: #5F6B7A; margin-top: 8px; }
.footer a { color: #0078D4; text-decoration: none; }
</style>
"@
    $html = New-Object System.Text.StringBuilder
    [void]$html.Append(("<html><head><meta charset='utf-8'><title>{0} - {1} - {2}</title>{3}</head><body>" -f (ConvertTo-SemtHtmlText $State.CampaignTitle),(ConvertTo-SemtHtmlText $titleWave),(ConvertTo-SemtHtmlText $context.RunId),$style))
    [void]$html.Append("<div class='header-card'><div>")
    [void]$html.Append(("<div class='title'>{0} - cycle {1}<span class='badge'>{2}</span></div>" -f (ConvertTo-SemtHtmlText $State.CampaignTitle),$CycleNumber,$mode))
    [void]$html.Append(("<div class='subtitle'>{0} v{1} - {2}</div>" -f (ConvertTo-SemtHtmlText $script:SemtProductDisplayName),(ConvertTo-SemtHtmlText $State.LauncherVersion),(ConvertTo-SemtHtmlText $script:SemtProductSubtitle)))
    [void]$html.Append(("<div class='wave' title='{1}'>Wave: {0}</div>" -f (ConvertTo-SemtHtmlText $titleWave),(ConvertTo-SemtHtmlText $context.WavePath)))
    [void]$html.Append(("<div class='meta'>Run: {0} | Generated: {1} | Attempts: {2} | Unique computers: {3}</div>" -f (ConvertTo-SemtHtmlText $context.RunId),(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$State.AttemptRowCount,$latestRows.Count))
    [void]$html.Append(("<div class='meta'>Run folder: {0} | Launcher log: {1}</div>" -f (New-SemtHtmlLink -Path $context.RunRoot -Label $context.RunRoot),(New-SemtHtmlLink -Path $context.LauncherLogPath)))
    [void]$html.Append('</div>')
    if (-not [string]::IsNullOrWhiteSpace($State.LogoUri)) { [void]$html.Append(("<img class='logo' src='{0}' alt='WorkplaceCloudHub' />" -f $State.LogoUri)) }
    [void]$html.Append('</div>')
    [void]$html.Append("<div class='card'><h2>Cycle progress</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows $progress)); [void]$html.Append('</div>')
    if (@($RunningRows).Count -gt 0) { [void]$html.Append("<div class='card'><h2>Running workers</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows @($RunningRows))); [void]$html.Append('</div>') }
    [void]$html.Append("<div class='card'><h2>Latest status by unique computer</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows $latestStatus -Columns @('Status','Count'))); [void]$html.Append('</div>')
    [void]$html.Append("<div class='card'><h2>Status summary by attempts</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows @(ConvertTo-SemtCounterRows -Counter $State.StatusCounter -NameColumn 'Status') -Columns @('Status','Count'))); [void]$html.Append('</div>')
    if ($State.AdminShareFailureCounter.Count -gt 0) { [void]$html.Append("<div class='card'><h2>Administrative share failures by attempts</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows @(ConvertTo-SemtCounterRows -Counter $State.AdminShareFailureCounter -NameColumn 'FailureType') -Columns @('FailureType','Count'))); [void]$html.Append('</div>') }
    [void]$html.Append("<div class='card'><h2>Next action summary by attempts</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows @(ConvertTo-SemtCounterRows -Counter $State.NextActionCounter -NameColumn 'NextAction') -Columns @('NextAction','Count'))); [void]$html.Append('</div>')
    [void]$html.Append("<div class='card'><h2>Latest computer details</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows $mainRows -Columns @($htmlColumns.ToArray()) -RawHtmlColumns $raw)); [void]$html.Append('</div>')
    if ($State.SeparatedStatuses.Count -gt 0) {
        [void]$html.Append("<div class='card'><h2>Latest guards, skips and unreachable computers</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows $separatedRows -Columns @($htmlColumns.ToArray()) -RawHtmlColumns $raw))
        [void]$html.Append(("<div class='footer'>Statuses shown here: {0}. Per-cycle CSV files keep every attempt.</div></div>" -f (ConvertTo-SemtHtmlText ($State.SeparatedStatuses -join ', '))))
    }
    [void]$html.Append("<div class='card'><h2>Effective Run options</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows $State.OptionRows -Columns @('Category','Option','Value','Unit','Source'))); [void]$html.Append('</div>')
    [void]$html.Append("<div class='card'><h2>Security and component integrity</h2>"); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows $State.SecurityRows -Columns @('Field','Value'))); [void]$html.Append((ConvertTo-SemtHtmlTable -Rows $State.IntegrityRows -Columns @('Path','Status','Signature','Sha256','Detail')))
    [void]$html.Append(("<div class='footer'>{0} - <a href='https://workplacecloudhub.com'>workplacecloudhub.com</a></div></div>" -f (ConvertTo-SemtHtmlText $script:SemtProductDisplayName)))
    [void]$html.Append('</body></html>')
    New-SemtDirectory -Path (Split-Path -Path $Path -Parent)
    $temp = '{0}.{1}.tmp' -f $Path,([guid]::NewGuid().ToString('N'))
    [System.IO.File]::WriteAllText($temp, $html.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    Move-Item -LiteralPath $temp -Destination $Path -Force
    return $Path
}
