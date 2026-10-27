<#
.SYNOPSIS
    Offline Migration Diagnostics farm-result and command-generation test.
.VERSION
    1.0.6
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
. (Join-Path $PSScriptRoot '..\Scripts\Launchers\SmartM365-SharePointMigration-Storage.ps1')
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVYy3POkc1JDrb
# pJddeuTOv2eLBAyaACSlZXckdxZiqqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHzqi0MopZ6/HS/XhBBsDbtcgJm87BvXPIqoi0H0xhlCMA0GCSqG
# SIb3DQEBAQUABIIBgHD6AeG89Lh+QJCOPZzO6pWGfzhIs8dDyr1SsuDXGpR9FTLY
# mmTPwy6HGgpcnP+dPb3xBRGV4s3FiH899r6QcHqAMQahLxK53ONJBO2TCoEP7QkR
# P/Fu/T/clcIYTl8MmTpDy4uVyzeX/9ybWraNIwbICJmKVcACnr8azPz2bktuHrHM
# BcYlfShbeoQTZCVRwu2f7HOmADfmGeh128bajgGlH3fonBykiQYbT6ZS4O63hnjb
# rFIHC/K70CEwHqRhXBkME0dI15QYk01uFkXcPax/2mJ0jGuJdps0g4pI8GhHRzVQ
# GIcXUXnDh3+D/AbDtoyVbRbAMjSsh2TGQpl42TpfPz5uE2+27MGzoWDchvrUAPox
# iKhEisnNOT47BzIAoZr2nGR0w2vo69/D3njv0hvQUKLDhaATgceCla/5oNQr3Mwc
# 8cVYVomH1ti5r3ajW3Wsc980mT6hxG9Eln4QavxXMooiqvFsD/WSr3rTDwTMYHYU
# JLdl8axunR4vA9MgCKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzhaMC8GCSqGSIb3DQEJBDEiBCCpdlNlJ/ZaDGVXjZyZKSc20gvWldxHdWcDgHw3
# 8VD/GTANBgkqhkiG9w0BAQEFAASCAgBqMuFm1eD9OfyfD8aoTX+FwXgNfsWxHBmH
# r7z4ul64yyT9rax0mpC25AB3OHaqjTxwOvKA4mHuiz7V3/47BZbbQokcLv7As6g0
# Uu0P/xyIMu4yqIPcUnSql3J6NN9wpY70zo9LmwnM4pHJwBYrU/KPZo7WypbyhJx0
# dR8itPaWTxHcfIBrUjGZq3M2YUVQGPxFQixoyTM6d2iKdYOJ4oAJhhmWoV+hJxsg
# XthTS6Brv9O3Jcv+jVOBBVRaNH7aPUkZoaM7dYeEUdXgu1SKFR8cfHtRkJHUyYnp
# BNxUEfHRRBhspRIgfTKO7ZtpgIcdvAUijTAw4ElsisOuWt159By6mfDT213Hi88b
# TkS6KgGVPcEOu9BUf88eEtDbiPQ4EzQGj6n0WQhR0/iVqUuTAiZQagfgprwAAeLy
# O+D2AB8RdW1pT3wcKr2XbLQ6BhO4zmdLfdetP201cFC1y/1AIZWIA7QBAGHyuRtg
# VWLj/1xg4yv0cpCoWmAP2E5eLsVZlkOF0FZ2k7pM3+7om8fBcwbZrOWgU8H1+l1L
# WvaLJLW+ozdfnpyC7nF2YNQ8Oq95nGSwrr8Rp1mFasIjBYHKw93o5gdh1aH6VvL4
# WoNw//B+vv/uL2z8YrU4gdbinJ3bAxYrmTeRTfllKR+z/w+cg5++MO1dUq1YOWSp
# 94Rq6ij6oQ==
# SIG # End signature block
