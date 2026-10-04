<#
.SYNOPSIS
    Offline GUI checks for newest ShareGate report selection and cached analysis.
.VERSION
    1.0.4
#>
#Requires -Version 7.4
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$guiPath = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors | ForEach-Object Message) }
foreach ($name in @('Clear-DiagnosticResult','Get-CurrentDiagnosticAnalysis','Refresh-DiagnosticReportState','Update-DiagnosticSummaryCrossCheck','Get-DiagnosticReportDisplayPath')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name },$true)
    if (-not $definition) { throw "Missing GUI function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
function Load-DiagnosticResult { param([string]$Directory) $script:LoadedAnalysis = $Directory }

$root = Join-Path $PSScriptRoot ('.diagnostics-gui-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }
try {
    $script:ScriptRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
    $project = Join-Path $root 'Synthetic'
    $reports = Join-Path $project 'ShareGate\MigrationReport'
    [void](New-Item -ItemType Directory -Path $reports -Force)
    $script:CurrentMigration = [pscustomobject]@{ Name='Synthetic'; Root=$project }
    $script:DiagProcess = $null
    $script:DiagReportSignature = ''
    $script:DiagReportHash = ''
    $script:LoadedAnalysis = ''
    $lblDiagInputPath = [pscustomobject]@{ Text='' }
    $lblDiagMigration = [pscustomobject]@{ Text='' }
    $lblDiagLatestReport = [pscustomobject]@{ Text=''; ToolTip='' }
    $lblDiagProgress = [pscustomobject]@{ Text='' }
    $lblDiagKpis = [pscustomobject]@{ Text=''; Visibility='Visible' }
    $panelDiagMetrics = [pscustomobject]@{ Visibility='Visible' }
    foreach ($name in @('lblSummaryShareGateValue','lblSummaryShareGateDetail',
        'lblSummaryFilesValue','lblSummaryFilesDetail','lblSummaryPermissionsValue',
        'lblSummaryPermissionsDetail','lblSummaryCrossCheckValue','lblSummaryCrossCheckDetail')) {
        Set-Variable -Name $name -Value ([pscustomobject]@{ Text=''; ToolTip='' })
    }
    $btnDiagAnalyze = [pscustomobject]@{ IsEnabled=$true }
    $btnDiagOpenReport = [pscustomobject]@{ IsEnabled=$true }
    $cardDiagSummary = [pscustomobject]@{ IsEnabled=$true }
    $panelDiagReview = [pscustomobject]@{ IsEnabled=$true }
    $btnFarmRun = [pscustomobject]@{ IsEnabled=$true }
    $btnFarmCheck = [pscustomobject]@{ IsEnabled=$true }
    $txtFarmDryRun = [pscustomobject]@{ Text='' }
    $txtFarmRun = [pscustomobject]@{ Text='' }
    $gridDiagPatterns = [pscustomobject]@{ ItemsSource=$null }
    $gridDiagRows = [pscustomobject]@{ ItemsSource=$null }
    $txtDiagRaw = [pscustomobject]@{ Text='' }
    $cmbDiagSession = [pscustomobject]@{ Items=[System.Collections.ArrayList]::new(); SelectedIndex=0 }

    Refresh-DiagnosticReportState -Force
    if (-not $lblDiagInputPath.Text.StartsWith('SharePointMigration\Tests\',[StringComparison]::Ordinal) -or
        $lblDiagInputPath.Text -notlike '*\ShareGate\MigrationReport' -or
        $lblDiagInputPath.Text.Contains($script:ScriptRoot)) { throw 'Report folder display leaked an absolute path or omitted the SharePointMigration prefix.' }
    if ($btnDiagAnalyze.IsEnabled -or -not $cardDiagSummary.IsEnabled -or $panelDiagReview.IsEnabled -or $btnFarmCheck.IsEnabled -or
        $lblDiagKpis.Visibility -ne 'Visible' -or $panelDiagMetrics.Visibility -ne 'Collapsed' -or
        $lblSummaryShareGateValue.Text -ne '—') { throw 'Empty folder did not show summary placeholders or left report actions enabled.' }
    $syntheticCrossCheck = [pscustomobject]@{
        Evidence=@(
            [pscustomobject]@{Evidence='ShareGate';Rate='—';State='Analysis missing'},
            [pscustomobject]@{Evidence='Files';Rate='80.00 %';Coverage='8 / 10 source keys';State='Newer scans available'},
            [pscustomobject]@{Evidence='Permissions';Rate='70.00 %';Coverage='7 / 10 source keys';State='Legacy scan dates'}
        )
        Scopes=@([pscustomobject]@{Site='https://target.example/a'},[pscustomobject]@{Site='https://target.example/b'})
        Ambiguous=1; FilesSummary='files/Summary.csv'; PermissionsSummary='permissions/Summary.csv'
    }
    Update-DiagnosticSummaryCrossCheck $syntheticCrossCheck
    if ($lblSummaryFilesValue.Text -ne '80.00 %' -or $lblSummaryPermissionsValue.Text -ne '70.00 %' -or
        $lblSummaryCrossCheckValue.Text -ne '2' -or $lblSummaryFilesDetail.Text -notmatch 'Newer scans available' -or
        $lblSummaryFilesDetail.ToolTip -ne 'files/Summary.csv') { throw 'Summary indicators did not retain separate measures and evidence states.' }
    if ($lblDiagLatestReport.Text -notmatch 'Place the latest') { throw 'Missing deposit instruction.' }
    if ($lblDiagMigration.Text -notmatch 'Synthetic') { throw 'Selected migration is not visible.' }

    $old = Join-Path $reports 'old.csv'
    $latest = Join-Path $reports 'latest.csv'
    'old' | Set-Content -LiteralPath $old
    Set-StrictMode -Version Latest
    Refresh-DiagnosticReportState -Force
    if (-not $btnDiagAnalyze.IsEnabled -or $script:DiagLatestReport.Name -ne 'old.csv') { throw 'A single report was not selected.' }
    $otherProject = Join-Path $root 'OtherMigration'
    $otherReports = Join-Path $otherProject 'ShareGate\MigrationReport'
    [void](New-Item -ItemType Directory -Path $otherReports -Force)
    'other' | Set-Content -LiteralPath (Join-Path $otherReports 'other.csv')
    $script:CurrentMigration = [pscustomobject]@{ Name='OtherMigration'; Root=$otherProject }
    Refresh-DiagnosticReportState -Force
    if ($script:DiagLatestReport.Name -ne 'other.csv' -or $lblDiagMigration.Text -notmatch 'OtherMigration') { throw 'Selecting another migration did not update the diagnostic report.' }
    $script:CurrentMigration = [pscustomobject]@{ Name='Synthetic'; Root=$project }
    Refresh-DiagnosticReportState -Force
    if ($script:DiagLatestReport.Name -ne 'old.csv' -or $lblDiagMigration.Text -notmatch 'Synthetic') { throw 'Returning to the original migration did not update the diagnostic report.' }
    'latest' | Set-Content -LiteralPath $latest
    (Get-Item -LiteralPath $old).LastWriteTimeUtc = [datetime]::UtcNow.AddHours(-2)
    (Get-Item -LiteralPath $latest).LastWriteTimeUtc = [datetime]::UtcNow.AddHours(-1)
    Refresh-DiagnosticReportState -Force
    if (-not $btnDiagAnalyze.IsEnabled -or $script:DiagLatestReport.Name -ne 'latest.csv') { throw 'Newest CSV was not selected.' }
    if ($lblDiagLatestReport.Text -notmatch 'latest.csv') { throw 'Latest report details are missing.' }

    $analysis = Join-Path $project 'ShareGate\Diagnostics\20261004-120000'
    [void](New-Item -ItemType Directory -Path $analysis -Force)
    '<html></html>' | Set-Content -LiteralPath (Join-Path $analysis 'MigrationDiagnostics-Report.html')
    'Status' | Set-Content -LiteralPath (Join-Path $analysis 'ClassifiedRows.csv')
    $item = Get-Item -LiteralPath $latest
    @{ Project='Synthetic'; Inputs=@($latest); SelectedSessionId='test-session'; GeneratedAtUtc=[datetime]::UtcNow.ToString('o');
       ReportPath=(Join-Path $analysis 'MigrationDiagnostics-Report.html'); RowsPath=(Join-Path $analysis 'ClassifiedRows.csv');
       InputEvidence=@(@{ Path=$latest; Size=$item.Length; Sha256=(Get-FileHash -LiteralPath $latest -Algorithm SHA256).Hash })
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $analysis 'Summary.json.txt')
    Refresh-DiagnosticReportState -Force
    if ($script:LoadedAnalysis -ne $analysis -or -not $script:DiagAnalysisVerified -or $lblDiagProgress.Text -notmatch 'SHA256 verified') { throw "Current analysis was not restored: loaded=$($script:LoadedAnalysis); verified=$($script:DiagAnalysisVerified); progress=$($lblDiagProgress.Text)" }

    'change' | Set-Content -LiteralPath $latest
    if ((Get-Item -LiteralPath $latest).Length -ne $item.Length) { throw 'The hash-change test did not preserve report size.' }
    $script:LoadedAnalysis = ''
    Refresh-DiagnosticReportState -Force
    if ($script:LoadedAnalysis -or $script:DiagAnalysisVerified -or $lblDiagProgress.Text -notmatch 'has not been analyzed') { throw 'Stale analysis was accepted after input changed.' }

    'latest' | Set-Content -LiteralPath $latest
    $legacySummary = @{
        Project='Synthetic'; Inputs=@('\\server\share\latest.csv'); SelectedSessionId='legacy-session'
        GeneratedAtUtc=[datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    $legacySummary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $analysis 'Summary.json.txt')
    $originalCulture = [Globalization.CultureInfo]::CurrentCulture
    try {
        [Globalization.CultureInfo]::CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
        $script:LoadedAnalysis = ''
        Refresh-DiagnosticReportState -Force
        if ($script:LoadedAnalysis -ne $analysis -or $script:DiagAnalysisVerified -or
            $lblDiagProgress.Text -notmatch 'legacy analysis') {
            throw "A matching legacy HTML analysis was hidden by locale-dependent date parsing: loaded=$($script:LoadedAnalysis); verified=$($script:DiagAnalysisVerified); progress=$($lblDiagProgress.Text)"
        }
    }
    finally { [Globalization.CultureInfo]::CurrentCulture = $originalCulture }
    Write-Output 'Diagnostics GUI report-state offline test passed.'
}
finally {
    if ((Test-Path -LiteralPath $root) -and $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCjSVEDsVlR0YkQ
# iK9fz1FxRbMpgryjl3ZR7k4XbcDsg6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC11OBIJ6Bdmfg5G+GNl7t9
# LTD8c4j+oJVXwW+7XCLMrTANBgkqhkiG9w0BAQEFAASCAYBcTAlBByHqvB7q+L3m
# lCmb1m8PDUN/+SL3yxjblN0tcFAbC0sX3u6jSz6AbkuiBxUjvCoONHLekDzBKJ7Y
# D9mjpKKo/fn/Wghbo4R0H5Q6oYz5t3RRj33UsKQMm3lWHlmvwpZcX5BdFRdM8/rZ
# Dcwq56x1Z91LfmVYRdCeKuavJvXR0MJNiwSTevz9hwuL1LNYapa/w0ZoD++1qZRg
# ssc+EZpD2UZA1/0tPM8VAT2KuWU/5UEBX9OS2P7ChbznGrt5FLbP3uo5z4AvcbU4
# QvRHO+Khh9A10QbeGhBgNobJChzWAlbSS9CAnmikm+LBn9+VqWOvx/l2TMq2aH2m
# FylwRMos9ANwpX9mj93n37z/bVLU1arY73Al+6JPeyC63us8OVQ4RqypcMl/GFhw
# YOd8MiCNg96KVrMbfd4yYxv+DidOoZl1mAdKLPB9wavQewS4Yz2SW/YX3G8vsF+A
# /QcGKQMQVeDycfwqWsVclMyvd8nAmmsAQgqjZOwQCjAW08w=
# SIG # End signature block
