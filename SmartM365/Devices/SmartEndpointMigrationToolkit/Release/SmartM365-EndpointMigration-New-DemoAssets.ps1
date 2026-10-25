#Requires -Version 5.1

<#
.SYNOPSIS
    Regenerates the local DemoAssets of Smart Endpoint Migration Toolkit from the offline demo.

.DESCRIPTION
    Starts the GUI demonstration (fictional Contoso data, nothing launched or queried), renders
    every view for both Campaigns to PNG, exports the fictional Run reports, and writes
    DemoAssets\<date>\ with index.html, SHA256SUMS.txt and manifest.json.txt (capture date,
    package and GUI versions, GUI SHA256, file sizes, image sizes and SHA256). The top-level
    DemoAssets\index.html and manifest.json.txt point to the new set; older sets are kept.
    DemoAssets is ignored by Git. Run it after every GUI change.
    Must run in Windows PowerShell 5.1 with -STA.

.EXAMPLE
    powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Release\SmartM365-EndpointMigration-New-DemoAssets.ps1

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$PackageRoot,
    [string]$CaptureDate = (Get-Date -Format 'yyyy-MM-dd'),
    [ValidateRange(1000,3840)][int]$Width = 1500,
    [ValidateRange(650,2160)][int]$Height = 950,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) { throw 'Run this script with powershell.exe -STA.' }
if ([string]::IsNullOrWhiteSpace($PackageRoot)) { $PackageRoot = Split-Path -Path $PSScriptRoot -Parent }
$PackageRoot = (Resolve-Path -LiteralPath $PackageRoot).Path
if ($CaptureDate -notmatch '^\d{4}-\d{2}-\d{2}$') { throw "Invalid CaptureDate: $CaptureDate" }
$assetRoot = Join-Path $PackageRoot 'DemoAssets'
$setRoot = Join-Path $assetRoot $CaptureDate
if ((Test-Path -LiteralPath $setRoot) -and -not $Force) { throw "DemoAssets set already exists (use -Force to replace): $setRoot" }
if (Test-Path -LiteralPath $setRoot) { Remove-Item -LiteralPath $setRoot -Recurse -Force }
foreach ($folder in @('GUI','Exports','HTML')) { New-Item -ItemType Directory -Path (Join-Path $setRoot $folder) -Force | Out-Null }

$guiPath = Join-Path $PackageRoot 'Scripts\SmartM365-EndpointMigration-GUI.ps1'
$previousUpdate = $env:SEMT_GUI_UPDATE_CHECK
$env:SEMT_GUI_UPDATE_CHECK = '0'
try {
    $state = . $guiPath -Demo -ValidateOnly -DemoWindowWidth $Width -DemoWindowHeight $Height
}
finally { $env:SEMT_GUI_UPDATE_CHECK = $previousUpdate }

function Save-ViewPng {
    param($Window, [string]$Path)
    $visual = $Window.Content
    $visual.Measure([Windows.Size]::new($Window.Width, $Window.Height))
    $visual.Arrange([Windows.Rect]::new(0, 0, $Window.Width, $Window.Height))
    $visual.UpdateLayout()
    $visual.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new([int]$Window.Width, [int]$Window.Height, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $background = [Windows.Media.DrawingVisual]::new()
    $drawing = $background.RenderOpen()
    $drawing.DrawRectangle([Windows.Media.Brushes]::White, $null, [Windows.Rect]::new(0, 0, $Window.Width, $Window.Height))
    $drawing.Close()
    $bitmap.Render($background)
    $bitmap.Render($visual)
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create($Path)
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
}

$campaigns = @(@{ Index = 0; Name = 'HybridJoin' }, @{ Index = 1; Name = 'Windows11Upgrade' })
$tabs = @($state.Tabs | ForEach-Object { @($_.Items) } | Where-Object { $_ -is [Windows.Controls.TabItem] })
foreach ($campaign in $campaigns) {
    $state.Controls.CampaignCombo.SelectedIndex = $campaign.Index
    foreach ($tab in $tabs) {
        $owner = [Windows.Controls.ItemsControl]::ItemsControlFromItemContainer($tab)
        $owner.SelectedItem = $tab
        $safe = ([string]$tab.Header) -replace '[^A-Za-z0-9]+', '-'
        Save-ViewPng -Window $state.Window -Path (Join-Path $setRoot ('GUI\SEMT-{0}-{1}.png' -f $campaign.Name,$safe.Trim('-')))
    }
    foreach ($output in @(Export-SmartM365DemoProfileOutputs)) {
        Copy-Item -LiteralPath $output -Destination (Join-Path $setRoot 'Exports') -Force
        if ($output -like '*.html') { Copy-Item -LiteralPath $output -Destination (Join-Path $setRoot 'HTML') -Force }
    }
}
$state.Window.Close()

Add-Type -AssemblyName PresentationCore
$files = foreach ($file in @(Get-ChildItem -LiteralPath $setRoot -Recurse -File | Sort-Object FullName)) {
    $relative = $file.FullName.Substring($setRoot.Length + 1).Replace('\','/')
    $entry = [ordered]@{ file = ('{0}/{1}' -f $CaptureDate,$relative); bytes = $file.Length; sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
    if ($file.Extension -eq '.png') {
        $stream = [IO.File]::OpenRead($file.FullName)
        try { $frame = [Windows.Media.Imaging.BitmapFrame]::Create($stream, [Windows.Media.Imaging.BitmapCreateOptions]::None, [Windows.Media.Imaging.BitmapCacheOption]::OnLoad); $entry.width = $frame.PixelWidth; $entry.height = $frame.PixelHeight } finally { $stream.Dispose() }
    }
    [pscustomobject]$entry
}
$version = (Get-Content -LiteralPath (Join-Path $PackageRoot 'SmartEndpointMigrationToolkit.version.json.txt') -Raw | ConvertFrom-Json)
$commit = ''
try { $commit = (& git -C $PackageRoot rev-parse --short HEAD 2>$null) } catch { $commit = '' }
$dirty = ''
try { $porcelain = @(& git -C $PackageRoot status --porcelain -- . 2>$null); if ($porcelain.Count -gt 0) { $dirty = ' (working tree with uncommitted changes)' } } catch { }
$manifest = [ordered]@{
    application = 'SmartEndpointMigrationToolkit'
    fictional = $true
    captured_date = $CaptureDate
    package_version = [string]$version.PackageVersion
    gui_version_at_storage = [string]$version.Components.GUI
    gui_sha256_at_storage = (Get-FileHash -LiteralPath $guiPath -Algorithm SHA256).Hash.ToLowerInvariant()
    source_commit = ('{0}{1}' -f $commit,$dirty)
    capture_method = 'WPF render tree of the offline demo (Scripts\SmartM365-EndpointMigration-Demo.ps1) for every view of both Campaigns; fictional demo exports.'
    live_operations = $false
    files = @($files)
    current_gallery = ('{0}/index.html' -f $CaptureDate)
    historical_assets_preserved = $true
}
$json = $manifest | ConvertTo-Json -Depth 5
[IO.File]::WriteAllText((Join-Path $setRoot 'manifest.json.txt'), ($json -replace ('"{0}/' -f [regex]::Escape($CaptureDate)), '"'), [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllLines((Join-Path $setRoot 'SHA256SUMS.txt'), @($files | ForEach-Object { '{0}  {1}' -f $_.sha256.ToUpperInvariant(),($_.file.Substring($CaptureDate.Length + 1)) }), [Text.Encoding]::ASCII)
$style = '<style>body{font:17px Segoe UI;background:#f5f8fb;color:#1f2937;max-width:1500px;margin:auto;padding:24px}img{max-width:100%;border:1px solid #dde7f0;border-radius:8px}section{background:white;padding:20px;margin:20px 0}a{color:#0078d4}</style>'
$sections = foreach ($file in $files) {
    $relative = $file.file.Substring($CaptureDate.Length + 1)
    $encoded = [Net.WebUtility]::HtmlEncode($relative)
    if ($relative -like '*.png') { '<section><a href="{0}">{0}</a><p><img loading="lazy" src="{0}"></p></section>' -f $encoded } else { '<section><a href="{0}">{0}</a></section>' -f $encoded }
}
$title = 'Smart Endpoint Migration Toolkit DEMO'
[IO.File]::WriteAllText((Join-Path $setRoot 'index.html'), ('<!doctype html><meta charset="utf-8"><title>{0}</title>{1}<h1>Smart Endpoint Migration Toolkit {2}</h1><p>DEMO - Sample data. Captured {3}. Fictional Contoso data; no tenant, device or endpoint operation.</p>{4}' -f $title,$style,$version.PackageVersion,$CaptureDate,($sections -join '')), [Text.UTF8Encoding]::new($false))
$history = @(Get-ChildItem -LiteralPath $assetRoot -Directory | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' -and $_.Name -ne $CaptureDate } | Sort-Object Name -Descending | ForEach-Object { '<p><a href="{0}/index.html">Previous gallery - {0}</a></p>' -f $_.Name })
[IO.File]::WriteAllText((Join-Path $assetRoot 'index.html'), ('<!doctype html><meta charset="utf-8"><title>{0}</title><style>body{{font:20px Segoe UI;background:#f5f8fb;color:#1f2937;margin:48px}}a{{color:#0078d4}}</style><h1>Smart Endpoint Migration Toolkit {1}</h1><p>DEMO - Sample data. Captured {2}.</p><p><a href="{2}/index.html">Open current demo gallery</a></p>{3}' -f $title,$version.PackageVersion,$CaptureDate,($history -join '')), [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $assetRoot 'manifest.json.txt'), $json, [Text.UTF8Encoding]::new($false))
Write-Output ('DemoAssets {0}: {1} file(s) in {2}' -f $CaptureDate,@($files).Count,$setRoot)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBBaWYvR/4Vs1lq
# rM2xJ5didMkTp1WeZtsnRijusJU9cKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMkMioQC8aUjXwEIKQPpwlVtvanGRPvt7juGuFuiS+3GMA0GCSqG
# SIb3DQEBAQUABIIBgCNGYPsX39JZB5iK4CZ/Mz1s6u/BBqK8ZUtrwrkwDrGs2bvT
# aL+88pKJEoPgh5I4eJSBJ6b455v92GV8rMXO2ZkEQKc3v763F76rpvezQrogjXJA
# 2fQ+U0SSOAq1dQrN3fC8x2Br5d+0O6ussvD8w6cWHPAIUsdhK1Lp3zkVS7utGF98
# 80ciEMw5xERghg/CI2mPqBeN1JZd3NrbSH0HjyWIQG9Nc5C6vGxmBGKk/2yKP4Rj
# yqEKGCuUG0XlPZ3T0gTLYIWMfG4AJQ3qxY9qkEJb5CDsevrr9JrrO1vPEq0MLvLT
# vrE0fLEem37TDXUkJEX+gwteu1YRZkkg5zx05o8SyxrDj8q0nVuVkVxG78CQUb0w
# ljG3Z0GTIHvFCElPp2yX9PYzk8az93u3j7XMFnYNu378Ic46TGIR1bJvGuax6PfV
# EZe7fBJ3+ccR2+43PWvq4Quk3zstkqglYja5wBGFifs645sNjT0P9iMvRaZBtsty
# We7XDACLLs/dpqPB7aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTRaMC8GCSqGSIb3DQEJBDEiBCD1d1S7gZUHt455U7wvYwFbcveOIzDxv4zMsi92
# 2ygQmDANBgkqhkiG9w0BAQEFAASCAgBvqHsoK3GggP35whajDDAIkQ0LRZDfxUVL
# 0BGA9JslztixXsTekl2uL+HxLkSWH4J1asLA34vKbusn55zuAmpfpjMeeUWrwjeL
# mvpVF9SIwvFc51gByxHanGDIGR05LfpcfvO2WIUPARVmTJqx7IpiY7BDWwOoXVvk
# 7baexY/4JobFDvsm840fFOK7iZ8a96JxorTRg3LxKbbAUY+K16e2VPwADULCqFE4
# NJlUSSYGCnanmaYc02iyXcrKzSXnVx5H0WvzUe+HrLOgcFEHTfzE3StjMNXmvmjE
# BMftztumUO9AX/eW0IINPMmCHWo0qbTChVHpZ4J/s6xq0Rbli1eE/R0Gl9qvz+qc
# 3m/Q/cvfdl27j0Ki0dAYg1jtKoSzUmYcY1pAPpl9F6UMcuK1Z1ObgFFy9828UOQh
# kDRIvpzS/ybTdg/dWfSJ4ei4NW4t+ZXWtr97G90M412s9f4SZ55fGoytFhNvBZkE
# 884NCS84oUAhCVZ6ZWk2CvzQ2gxRIERym147vJSqbTEfu5VgeRqPhsr0LKAhm6BV
# V6rnxIIAniaCd7ZpEcv7oLR5i4QJ5op7zaZTV5BvuuvsK6aVGcbt/Gsf4K3zvMkD
# rU7hVPbVIp59NP6DlE8zghGVwI9u5dtbj7QtXEqpDtidIt6M9R5ADqayoPuwxrmq
# eJpC1lINQQ==
# SIG # End signature block
