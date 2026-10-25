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
