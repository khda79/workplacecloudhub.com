<#
.SYNOPSIS
    Offline GUI checks for newest ShareGate report selection and cached analysis.
.VERSION
    1.0.11
#>
#Requires -Version 7.4
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$guiPath = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors | ForEach-Object Message) }
foreach ($name in @('Clear-DiagnosticResult','Get-CurrentDiagnosticAnalysis','Refresh-DiagnosticReportState','Update-DiagnosticReportStatus','Update-DiagnosticSummaryCrossCheck','Get-DiagnosticReportDisplayPath','Update-DiagnosticAnalysisProgress','Get-DiagnosticReportKey')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name },$true)
    if (-not $definition) { throw "Missing GUI function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
. (Join-Path $PSScriptRoot '..\Scripts\Launchers\SmartM365-SharePointMigration-Storage.ps1')
function Start-AutomaticDiagnosticAnalysis { } # Automatic launch policy has its own worker-free tests.
function Load-DiagnosticResult {
    param([string]$Directory)
    $script:LoadedAnalysis = $Directory
    $script:DiagSummary = @{ ReportPath = (Join-Path $Directory 'MigrationDiagnostics-Report.html') }
    $btnDiagOpenReport.IsEnabled = $true
}

$root = Join-Path $PSScriptRoot ('.diagnostics-gui-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }
try {
    $script:ScriptRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
    $project = Join-Path $root 'Synthetic'
    $reports = Join-Path $project 'ShareGate\MigrationReport'
    [void](New-Item -ItemType Directory -Path $reports -Force)
    $script:CurrentMigration = [pscustomobject]@{ Name='Synthetic'; Root=$project }
    $script:DiagProcess = $null
    $script:DiagActiveReportKey = ''
    $script:DiagReportSignature = ''
    $script:DiagReportHash = ''
    $script:LoadedAnalysis = ''
    $lblDiagInputPath = [pscustomobject]@{ Text='' }
    $lblDiagMigration = [pscustomobject]@{ Text='' }
    $lblDiagLatestReport = [pscustomobject]@{ Text=''; ToolTip='' }
    $lblDiagProgress = [pscustomobject]@{ Text='' }
    $panelDiagAnalysisLoading = [pscustomobject]@{ Visibility='Collapsed' }
    foreach ($name in @('lblDiagReportState','lblDiagReportEvidence','lblDiagAnalysisState',
        'lblDiagAnalysisEvidence','lblDiagHtmlState','lblDiagHtmlEvidence','lblDiagNextAction')) {
        Set-Variable -Name $name -Value ([pscustomobject]@{ Text='' })
    }
    $lblDiagKpis = [pscustomobject]@{ Text=''; Visibility='Visible' }
    $panelDiagMetrics = [pscustomobject]@{ Visibility='Visible' }
    foreach ($name in @('lblSummaryShareGateValue','lblSummaryShareGateDetail',
        'lblSummaryFilesValue','lblSummaryFilesDetail','lblSummaryPermissionsValue',
        'lblSummaryPermissionsDetail','lblSummaryCrossCheckValue','lblSummaryCrossCheckDetail')) {
        Set-Variable -Name $name -Value ([pscustomobject]@{ Text=''; ToolTip='' })
    }
    $btnDiagAnalyze = [pscustomobject]@{ IsEnabled=$true }
    $btnDiagOpenReport = [pscustomobject]@{ IsEnabled=$true }
    $btnDiagCopyMessage = [pscustomobject]@{ IsEnabled=$true; Content=''; ToolTip='' }
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
    if ($lblDiagKpis.Text -notmatch 'No ShareGate migration report') { throw 'Absent input was described as an unanalyzed report.' }
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
    if ($lblDiagReportState.Text -ne 'Missing' -or $lblDiagAnalysisState.Text -ne 'Unavailable' -or
        $lblDiagHtmlState.Text -ne 'Unavailable' -or $lblDiagNextAction.Text -notmatch 'Place the latest') {
        throw 'Empty report folder did not show the correct report status and next action.'
    }

    $old = Join-Path $reports 'old.csv'
    $latest = Join-Path $reports 'latest.csv'
    'old' | Set-Content -LiteralPath $old
    Set-StrictMode -Version Latest
    Refresh-DiagnosticReportState -Force
    if (-not $btnDiagAnalyze.IsEnabled -or $script:DiagLatestReport.Name -ne 'old.csv') { throw 'A single report was not selected.' }
    if ($lblDiagReportState.Text -ne 'Detected' -or $lblDiagAnalysisState.Text -ne 'Not analyzed' -or
        $lblDiagHtmlState.Text -ne 'Unavailable') { throw 'Unanalyzed report status is incorrect.' }
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
    if ($lblDiagAnalysisState.Text -ne 'SHA256 verified' -or $lblDiagHtmlState.Text -ne 'Available' -or
        $lblDiagNextAction.Text -notmatch 'Review issue patterns') { throw 'Verified analysis status is incorrect.' }

    'change' | Set-Content -LiteralPath $latest
    if ((Get-Item -LiteralPath $latest).Length -ne $item.Length) { throw 'The hash-change test did not preserve report size.' }
    $script:LoadedAnalysis = ''
    Refresh-DiagnosticReportState -Force
    if ($script:LoadedAnalysis -or $script:DiagAnalysisVerified -or $lblDiagProgress.Text -notmatch 'Waiting for automatic analysis') { throw 'Stale analysis was accepted after input changed.' }
    if ($lblDiagAnalysisState.Text -ne 'Not analyzed' -or $lblDiagHtmlState.Text -ne 'Unavailable') { throw 'Stale analysis remained visible in the report status.' }

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
        if ($lblDiagAnalysisState.Text -ne 'Legacy match' -or $lblDiagNextAction.Text -notmatch 'Reanalyze') {
            throw 'Legacy analysis status did not explain the weaker evidence.'
        }
    }
    finally { [Globalization.CultureInfo]::CurrentCulture = $originalCulture }
    # A latest XLSX is actionable even when no ImportExcel is installed in the GUI.
    function Get-Module { throw 'The GUI must not gate XLSX analysis on module availability.' }
    $xlsx=Join-Path $reports 'newest.xlsx'
    'Synthetic workbook placeholder' | Set-Content -LiteralPath $xlsx
    (Get-Item -LiteralPath $xlsx).LastWriteTimeUtc=[datetime]::UtcNow.AddMinutes(5)
    Refresh-DiagnosticReportState -Force
    if(-not $btnDiagAnalyze.IsEnabled -or $script:DiagLatestReport.FullName -ne $xlsx -or $lblDiagProgress.Text -notmatch 'installs automatically' -or $lblDiagAnalysisState.Text -ne 'Ready to analyze'){
        throw 'An XLSX report was blocked instead of offering automatic module installation.'
    }
    if ($script:DiagSummary -or $btnDiagOpenReport.IsEnabled -or
        $lblDiagKpis.Text -notmatch 'newest.xlsx.*start automatically') {
        throw 'A new ShareGate input was presented as an analyzed report or left its missing indicators unexplained.'
    }
    $script:DiagProcess=[pscustomobject]@{HasExited=$false}
    $script:DiagActiveReportKey=Get-DiagnosticReportKey
    Refresh-DiagnosticReportState -Force
    if($btnDiagAnalyze.IsEnabled){throw 'Analysis button was enabled during an active analysis.'}
    if($panelDiagAnalysisLoading.Visibility -ne 'Visible'){throw 'Running analysis did not display its progress bar.'}
    if ($lblDiagKpis.Text -notmatch 'Analyzing ShareGate report: newest.xlsx') { throw 'Running analysis left an incorrect empty detail.' }
    $script:DiagProjectRoot=$project;$script:DiagOutputDirectory=$analysis
    @{State='InstallingImportExcel';Message='Installing ImportExcel from PSGallery for the current user…'} | ConvertTo-Json |
        Set-Content -LiteralPath (Join-Path $analysis 'analysis.phase.json.txt')
    Update-DiagnosticAnalysisProgress
    if($lblDiagAnalysisState.Text -ne 'Installing module' -or $lblDiagProgress.Text -notmatch 'Installing ImportExcel' -or $lblDiagAnalysisEvidence.Text -notmatch 'CurrentUser'){
        throw 'Background module installation was not visible in the GUI.'
    }
    Update-DiagnosticReportStatus -Failure 'Synthetic analysis failure'
    if($panelDiagAnalysisLoading.Visibility -ne 'Collapsed'){throw 'Failed analysis kept its progress bar visible.'}
    Update-DiagnosticReportStatus
    $script:CurrentMigration=[pscustomobject]@{Name='Other';Root=(Join-Path $root 'Other')}
    Update-DiagnosticReportStatus
    if($panelDiagAnalysisLoading.Visibility -ne 'Collapsed'){throw 'Another migration displayed the previous analysis progress.'}
    $lblDiagProgress.Text='Other migration'
    Update-DiagnosticAnalysisProgress
    if($lblDiagProgress.Text -ne 'Other migration'){throw 'A background installation overwrote another migration status.'}
    $script:DiagProcess=$null
    Update-DiagnosticReportStatus
    if($panelDiagAnalysisLoading.Visibility -ne 'Collapsed'){throw 'Completed analysis kept its progress bar visible.'}
    Write-Output 'Diagnostics GUI report-state offline test passed, including XLSX installation availability and progress.'
}
finally {
    if ((Test-Path -LiteralPath $root) -and $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAtPFDDdTNcECnT
# cwGOGmkZDuJc0CSS8M8bK850/xvOVaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILs2+G8dfY+ocsf4lxkDp0vi3MkK8mUsynoOXAxO3oVeMA0GCSqG
# SIb3DQEBAQUABIIBgI8cqmKtcF204dKMa8NTl0XatswBXPh3wnmz26ZQAL51iOas
# XPsIc+QzMc1Fj/B3hLQOknYwfCPYEC7iI3AHIJphzuMWwFu8mEJ1pkesiKU+crXQ
# xFw6dWAqew8xXsdThROB4BQq1ZMjvI58fDqCpEyavSjxg1Bsk/KVWkYGOa4mo5la
# gPUVz07lC9DdfNBlOAe10EPXuSCl+RFgJH1TUk8k5bOzIDKhQd9AQN/b4moUWuOu
# Vd+//2w4Sm4QO/kIONrwMMQ5x/HTN43CFC//u3vqO1OzqnTqzbHTP/yoJaaiG5Xl
# 2nROwJmDwqsQc9jdrG3RrpAAR3+kFSlyvNSr2/y1tW7WFB1XnnHcv4M/xWKCBMPl
# ojJtfGnom2CP0xRqhsX1WkWR0nn7aKyLlSHgczBAwqlsOXOReKh2EXFU7QWhbTq9
# sD4xI5V53QsHybeM2ohf3f7bcmRZYjd3+5UaBUWZQ7T2C7GQBxmudlxURIt3ITdP
# hr2RSJWHFIBz/aEH/6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzdaMC8GCSqGSIb3DQEJBDEiBCC/hGn4LAXwNWKc1rnU26F8V6omUiNp+l0QoJKM
# ZXisjDANBgkqhkiG9w0BAQEFAASCAgCnB2VmWX6dUSGhVvuI2aJFdZa/4U2t9DZT
# kc+MKxW5yvhjtY2ZFLEJe49AWSNDz1bmmE85qIqJAabgiP76uwqoOfupHFE/xEfB
# 7VZyPUrKbYkmToJ2AGP1t6P9Ws9ysaDys8IWTjsUEEuUrCGxokqNd1r1FivzskvO
# tZQmXyh8kMsaapvh0MmLLkwbZmZgAnwObPaxQua8XHvlP0ljpRhM1sLKF5GNNhmc
# sLmJ5QmxZBupROMoZkXbTXiTseRGaNbQ6MP4aSnXIdwNdobQaw4FR40R1U8WOLNm
# DLLnXzqd4D1L59A2e8zk5NwpQjs8EJgB4szika5wLLn3anNoUiB0oLlqBk8kWW+V
# PhzJZGcoJDf8H0FXR92VbvyXT9nTedYFQxypEAl0LOAbVcMebZIXo5PiuCoXEVhS
# HLJe30BbcCyqtx8j51w/fc1pvn1lQtLoZ8i8qv01b9JkzWbKFi/+p6jmI+3tXLO1
# yuqxgAhO71vvNuOyKLwEGH80RKkpVMQdDRlQ/qW9rmBDRasr6ZaCwY2ipDY13Y3S
# 0bMpHAuXoEqZ18lwUs7uvdfXB2FXRhd9L8VcWln7w6MK7T3Ruhp5VPhxopntp6Ym
# r6w4fH+x6ItCyxBk0mQ3YaZF1YGskjLje5m94Vhhv8DLlX1Jd3+IrgvjA0ais0Qj
# TasMohhG6Q==
# SIG # End signature block
