<#
.SYNOPSIS
    Offline Migration Diagnostics farm-result and command-generation test.
.VERSION
    1.0.5
#>
#Requires -Version 7.4
[CmdletBinding()]
param([string]$PreviewDirectory='')

$ErrorActionPreference = 'Stop'
$guiPath = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors | ForEach-Object Message) }
$function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Refresh-FarmDiagnostics' },$true)
if (-not $function) { throw 'Farm refresh function is missing from the GUI.' }
. ([scriptblock]::Create($function.Extent.Text))
$preflight = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-FarmDiagnosticsPrerequisites' },$true)
if (-not $preflight) { throw 'Farm prerequisite check is missing from the GUI.' }
. ([scriptblock]::Create($preflight.Extent.Text))
$constructor = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-FarmDiagnosticsWindow' },$true)
. ([scriptblock]::Create($constructor.Extent.Text))
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$layoutWindow=$null; $farmWindow=$null
function Refresh-DiagnosticReportState { param([switch]$Force) }
$root = Join-Path $PSScriptRoot ('.farm-gui-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }
try {
    $project = Join-Path $root 'Synthetic'
    $diagnostics = Join-Path $project 'ShareGate\Diagnostics'
    $farm = Join-Path $diagnostics 'Farm-20261003-000000'
    $analysis = Join-Path $diagnostics 'Analysis-20261002'
    [void](New-Item -ItemType Directory -Path $farm,$analysis -Force)
    $report = Join-Path $farm 'Farm-Report.html'
    '<html></html>' | Out-File -LiteralPath $report -Encoding utf8
    [pscustomobject]@{ SchemaVersion=1; Project='Synthetic'; Status='Partial'; GeneratedAtUtc='2026-10-03T00:00:00Z'; Coverage=@([pscustomobject]@{ Server='WFE02'; Source='IIS logs'; Status='Missing'; Detail='Share unavailable' }); ReportPath=$report } | ConvertTo-Json -Depth 6 | Out-File -LiteralPath (Join-Path $farm 'Farm-Summary.json.txt') -Encoding utf8
    @'
WindowUtc,Lines,Items,Source,Destination,Undetermined
2026-10-02 22:00 UTC,9,4,9,0,0
2026-10-02 22:05 UTC,4,2,4,0,0
'@ | Out-File -LiteralPath (Join-Path $analysis 'AccessFailures-5min.csv') -Encoding utf8
    $script:CurrentMigration = [pscustomobject]@{ Name='Synthetic'; Root=$project }
    $script:DiagLoadedDirectory = $analysis
    $script:DiagAnalysisVerified = $true
    $script:FarmReportPath = ''
    $script:FarmToolkitRoot = '\\fileserver.example\toolkit\SharePointMigration'
    $lblFarmResult = [pscustomobject]@{ Text='' }
    $lblFarmPeaks = [pscustomobject]@{ Text='' }
    $btnFarmOpenReport = [pscustomobject]@{ IsEnabled=$false }
    $btnFarmCheck = [pscustomobject]@{ IsEnabled=$false }
    $btnFarmRun = [pscustomobject]@{ IsEnabled=$false }
    $lblFarmPrerequisites = [pscustomobject]@{ Text='' }
    $txtFarmDryRun = [pscustomobject]@{ Text='' }
    $txtFarmRun = [pscustomobject]@{ Text='' }
    Set-StrictMode -Version Latest
    Refresh-FarmDiagnostics
    if ($lblFarmResult.Text -notmatch 'Partial; 1 missing') { throw 'Farm coverage was not displayed.' }
    if (-not $btnFarmOpenReport.IsEnabled) { throw 'Farm HTML report was not detected.' }
    if (-not $btnFarmCheck.IsEnabled) { throw 'Farm prerequisite check was not offered after current analysis.' }
    if ($btnFarmRun.IsEnabled -or $lblFarmPrerequisites.Text -notmatch 'Check prerequisites') { throw 'Farm Run was enabled before prerequisite checks.' }
    if ($lblFarmPeaks.Text -notmatch '2 in') { throw 'ShareGate peak count was not displayed.' }
    foreach($command in @($txtFarmDryRun.Text,$txtFarmRun.Text)) {
        if ($command -notmatch 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File') { throw 'Windows PowerShell command flags are missing.' }
        if ($command -notmatch '-Project "Synthetic"') { throw 'Project was not included.' }
        if ($command -notmatch '-Around "2026-10-02T22:00:00Z"') { throw 'Largest UTC peak was not included.' }
        if (-not $command.Contains('-ShareGatePeaksCsv "' + $script:FarmToolkitRoot)) { throw 'UNC peak CSV was not included.' }
        if (-not $command.Contains('-ToolkitRoot "' + $script:FarmToolkitRoot + '"')) { throw 'UNC toolkit root was not included.' }
        if ($command -match "`n") { throw 'A command spans multiple lines.' }
    }
    if ($txtFarmDryRun.Text -notmatch ' -DryRun$' -or $txtFarmRun.Text -match ' -DryRun$') { throw 'DryRun and real command separation failed.' }
    $script:FarmInvocation = $null
    if (Test-FarmDiagnosticsPrerequisites) { throw 'Farm Run passed without an analysis.' }
    if ($btnFarmRun.IsEnabled -or $lblFarmPrerequisites.Text -notmatch 'Run unavailable') { throw 'Farm Run did not remain disabled.' }
    # Use the real main XAML and dedicated window without showing a GUI or running a probe.
    $guiText=[IO.File]::ReadAllText($guiPath)
    $match=[regex]::Match($guiText,"(?s)\[xml\]\`$xaml\s*=\s*@'\r?\n(.*?)\r?\n'@")
    [xml]$markup=$match.Groups[1].Value
    $layoutWindow=[Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($markup))
    $diagnosticsPanel=$layoutWindow.FindName('panelDiagnostics')
    $review=$layoutWindow.FindName('panelDiagReview')
    $farmHost=$layoutWindow.FindName('farmDiagnosticsHost')
    $farmCard=$farmHost.Child
    $loading=$layoutWindow.FindName('panelCrossCheckLoading')
    $analysisLoading=$layoutWindow.FindName('panelDiagAnalysisLoading')
    if ($analysisLoading.Parent -ne $loading.Parent -or -not $analysisLoading.Children[1].IsIndeterminate) {
        throw 'ShareGate analysis progress is not at the top alongside cross-check progress.'
    }
    if ([Windows.Controls.Grid]::GetRow($loading.Parent) -ne 0 -or
        $loading.Parent.Parent -ne $diagnosticsPanel) { throw 'Diagnostics loading progress is not at the top of the tab.' }
    if ([Windows.Controls.Grid]::GetColumnSpan($review) -ne 3 -or
        $farmHost.Visibility -ne 'Collapsed' -or -not $layoutWindow.FindName('btnFarmWindow')) {
        throw 'Farm diagnostics remained in the review layout or review did not span all columns.'
    }
    $farmHost.Child=$null
    $farmWindow=New-FarmDiagnosticsWindow -Owner $layoutWindow -MigrationName 'Synthetic' -Content $farmCard
    if ($farmWindow.Title -notmatch 'Synthetic' -or $farmCard.Parent -eq $farmHost -or
        $layoutWindow.FindName('btnFarmRun').IsEnabled) { throw 'Dedicated farm window lost migration context or prerequisite gating.' }
    $farmWindow.Content.Measure([Windows.Size]::new(1000,440))
    $farmWindow.Content.Arrange([Windows.Rect]::new(0,0,1000,440)); $farmWindow.Content.UpdateLayout()
    if($PreviewDirectory){
        [void](New-Item -ItemType Directory -Path $PreviewDirectory -Force)
        $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new(1000,440,96,96,[Windows.Media.PixelFormats]::Pbgra32)
        $bitmap.Render($farmWindow.Content)
        $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new();$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $stream=[IO.File]::Create((Join-Path $PreviewDirectory 'farm-diagnostics-window.png'));try{$encoder.Save($stream)}finally{$stream.Dispose()}
    }
    $farmWindow.Content.Children[1].Content=$null; $farmWindow.Close(); $farmWindow=$null
    $farmHost.Child=$farmCard
    $diagnosticsPanel.Visibility='Visible'; $layoutWindow.FindName('panelSummary').Visibility='Collapsed'
    $review.IsEnabled=$true
    $content=$layoutWindow.Content; $layoutWindow.Content=$null
    $content.Resources=$layoutWindow.Resources; $content.Background=$layoutWindow.Background
    foreach($width in @(1280,1920)){
        $content.Measure([Windows.Size]::new($width,2200));$content.Arrange([Windows.Rect]::new(0,0,$width,2200));$content.UpdateLayout()
        if ($review.ActualWidth -lt $width-60 -or $review.ActualWidth -lt $diagnosticsPanel.ActualWidth-2 -or
            $layoutWindow.FindName('gridDiagPatterns').ActualWidth -lt $review.ActualWidth-40 -or
            $layoutWindow.FindName('gridDiagRows').ActualWidth -lt $review.ActualWidth-40) { throw "Review tables did not fill the available width at $width." }
        if($PreviewDirectory){
            $height=[int][math]::Ceiling($review.ActualHeight)
            $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new([int]$review.ActualWidth,$height,96,96,[Windows.Media.PixelFormats]::Pbgra32)
            $visual=[Windows.Media.DrawingVisual]::new(); $drawing=$visual.RenderOpen()
            $drawing.DrawRectangle($layoutWindow.Background,$null,[Windows.Rect]::new(0,0,$review.ActualWidth,$height))
            $drawing.DrawRectangle([Windows.Media.VisualBrush]::new($review),$null,[Windows.Rect]::new(0,0,$review.ActualWidth,$height)); $drawing.Close()
            $bitmap.Render($visual)
            $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new();$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $stream=[IO.File]::Create((Join-Path $PreviewDirectory "diagnostics-review-$width.png"));try{$encoder.Save($stream)}finally{$stream.Dispose()}
        }
    }
    Write-Output 'Farm diagnostics GUI offline test passed: result routing, prerequisite gates, dedicated window and full-width review at 1280/1920.'
}
finally {
    if ($farmWindow) { $farmWindow.Close() }
    if ($layoutWindow) { $layoutWindow.Close() }
    if ((Test-Path -LiteralPath $root) -and $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6F+sZrZqwlSA+
# 1KA4y9XojmVXuPH/u798faqGY2i1n6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA2tr7Ihouneq0w1e0NzOKm
# oZt+RQHzOSErwjHEQxZulzANBgkqhkiG9w0BAQEFAASCAYCuT/HQVGTHYxHQygyv
# qvE63Llqh/lZV/8UQzMs7X6QHa57LeqiL3QY59F873jMQcat5yKHaOaf1FHkxeI+
# za+tnUUL3Vapk/Rh/3ZaG3RSQEaoWIBylAjAxGhV64J6kMZ+F21UrZ9YA7Fstdap
# DGRmWxevnEJK7ez/RuBQ1NQ1A7VpLRA/qRWatSv89rokn0FkxWYNprKF9Hc2kNDO
# GfIvLhub8NkDdH75jw49WRextHybzgITE2wrtwwB1Aw9Z2qOSp6IcjKnzKK2nJU1
# KXiijRG5Woza3Tjg2RXRB17Oceb0WsmfSoSj1SFuoFlDxwJgAb/qlWw5dx4ifjFT
# 7GU17EG8fCVZuRply80yosBa01aysbf2RZByqy+K0wJGW2ponoBEv9WfM3aoLTLZ
# j6CxQvumAiC0w6LOWEf8OwGDZzVD7VEKYK86Wxfp9XkCVeAr3WKLKEKK1DNi6L/m
# KIM3x/oXYu9XV7hGMDuWJ/zMo/i/lYMlWXK+BjLPnai9CDg=
# SIG # End signature block
