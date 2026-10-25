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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDAdG2XC6AiJ4tf
# sZbxwy41+bqJfWOnmGq0c6vv6bnYSKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFLVhaRelK4ZQnz8UYpTArNbMvwyGz/mBqEeaz0WeDTYMA0GCSqG
# SIb3DQEBAQUABIIBgJ7ifNbAkMBOgfTS1f+Zps142Oigc7/Onx2eFyqOMWiXTCF4
# 37xw/7bZqSooob2bLeq2SiiakvRUA5C9Ugh4J6a3alWHlRWToo+2R3xncNlnFbvw
# NQj26//YW2aeDJ6PxJS647tMk515WUTe9KhEFqLOVjbvbgzIUW0sh1MsC+9eidlH
# 4/SCPVrGL0voydbviknWK6Y2fR5TV1F+LgZzKuY6XhITPKudb17l6USwB79EU6wQ
# vYum2PqIAt2MFb9LMDW5mJjQNHQQhMpnnSg2ZaQ0czpKQyiSFK1wgCqmRlTV5XlX
# 0JEWRoteIX7BC0iOhkYkeyyzB2yz1mQOGtf3s5IVP6TrBFeEGNEZkiMtsqXJJckZ
# h00yF+AoGB9i1yQe2tlM5fr0HhGEHPcehcU0bmjqURNuX3cQqpyOkzkr0jCOa6BN
# u3CiBFKcpX4804wceOamwQQM+A0wvJJVOzY4U9pZuXEY4AdUFF1jO1SanAbpIlvZ
# nGZiPDGEK/ykZVpY1KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTNaMC8GCSqGSIb3DQEJBDEiBCAvoEXnc9/wxKz2c0J6DBKiDG7KeJom8fYs07AA
# ItaS0TANBgkqhkiG9w0BAQEFAASCAgA1OFjAlk9gzNe9BkIpNJNSLmp0WDCf2dbg
# ufwgdOg3jssy6JsPN1qQX2MZp01jziwGHjUqjb3ZJO/FWcOH6nVmeAJw+uSOepXL
# 5jjpbLlfiYo40qESyn0M9Hw4afabTR51Lxh9E8/aC07/7JVaSz3xnUlj43d501V7
# H3ab04sh/qweupjVEDVr1L3A0jHpjHz7kYcF7FC74OC9Tsy7NYGpdoQ0lPeMRQFt
# uI+rU4vekrk2ARaiWe0RBtC3mXO6HtAz+VauVPf0AK1YG3gSW8lNMTryfqU9Qp0I
# AvgRrLJu0K2aXBfEcRbmZsQ1/4wBBavrVokPB7Uf0n+UKoRMkVo5k9XNZecJiE1n
# i00vvOzzAqmx5aKjJ3bfd+vUtFEh0cpXiPugjUYLHEWySb6fOVARTbsRtrU+6Wcs
# GdRUkb6+r3kNphcBTvDOwb0v3dyTGWOS9WXsE3OS2HUbGZ3yPNEU7+DixAzI6EVd
# AjAYBfQrvFhLhyW1yRIcLcrxQsmH3uk8tZJbCQi/N2K0Z0wrY+DWOXu+ULgGLDa+
# V82L5MSRCTnHD8ACUEnpULvY83oBQkR9ggWL9N96WTMlMwFJUSiIhDbhjsGNT6xe
# YpvtCU6A7ZRddim0o9/PiFIJcobnlOalhXCUV8/xNyd5fzjaTRxSht+eTVIfa+T6
# CH9FQa2meg==
# SIG # End signature block
