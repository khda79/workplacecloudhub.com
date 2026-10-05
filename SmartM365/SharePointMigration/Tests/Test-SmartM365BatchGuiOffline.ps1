<#
.SYNOPSIS
    Offline validation of batch dashboard controls, CMD routing and receipts.
.DESCRIPTION
    Uses synthetic launchers and summaries only. No SharePoint, Graph, real scans
    or prerequisite probes are invoked. Renders the WPF tab at two window widths.
#>
#requires -Version 7.4
[CmdletBinding()]
param([string]$PreviewDirectory='')
$ErrorActionPreference='Stop'
$projectRoot=Split-Path $PSScriptRoot -Parent
$generic=Join-Path $projectRoot 'Scripts\Launchers\Generic'
. (Join-Path $generic 'SmartM365-SharePointMigration-BatchGui.ps1')
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
function Assert-BatchGui { param([bool]$Condition,[string]$Message) if(-not $Condition){throw $Message} }
$temp=Join-Path $env:TEMP ('SmartM365-BatchGuiTest-'+[guid]::NewGuid().ToString('N'))
$root=Join-Path $temp 'Tool kit & migration'
[void](New-Item -ItemType Directory -Path $root -Force)
$script:Clicked=@()
try {
    foreach($kind in @('Source','Target','Comparison')){
        $c=Get-SmartM365BatchCommand -Kind $kind -Root $root -Names @('A','B') -Mode PermissionsOnly -AuthMode Certificate -PlanOnly
        Assert-BatchGui ($c.Command -match '-MigrationNames "A,B"' -and $c.Command.EndsWith('-PlanOnly')) 'Subset or preview flag missing.'
        Assert-BatchGui (($c.Command -match '-MaxParallel 2') -eq ($kind -eq 'Target')) 'Destination parallel limit missing or applied to another batch.'
        Assert-BatchGui (($c.Command -match '-AuthMode Certificate') -eq ($kind -ne 'Source')) 'Incorrect authentication routing.'
    }
    foreach($bad in @('Bad%Path','Bad"Path','Bad!Path','Bad^Path')){
        $blocked=$false;try{$null=Get-SmartM365BatchCommand Target -Root $bad}catch{$blocked=$true}
        Assert-BatchGui $blocked 'Unsafe CMD path was accepted.'
    }
    $stub=Join-Path $root 'fake.ps1'
    @'
param([string]$InventoryMode,[string]$ComparisonMode,[string]$AuthMode,[int]$MaxParallel,[string]$MigrationNames,[string]$BatchId,[switch]$PlanOnly,[string]$Kind,[string]$Root)
$ErrorActionPreference='Stop'
$argsRecord=@{InventoryMode=$InventoryMode;ComparisonMode=$ComparisonMode;AuthMode=$AuthMode;MaxParallel=$MaxParallel;MigrationNames=$MigrationNames;BatchId=$BatchId;PlanOnly=[bool]$PlanOnly;Kind=$Kind}
$argsRecord | ConvertTo-Json | Set-Content (Join-Path $Root 'arguments.json')
if($PlanOnly){exit 0}
$folder=@{Source='source-scan-batches';Target='target-scan-batches';Comparison='comparison-batches'}[$Kind]
$output=Join-Path $Root "Migrations\logs\$folder\$BatchId"
[void](New-Item -ItemType Directory -Path $output)
$rows=@([pscustomobject]@{Migration='A';Action='Files';Status='SUCCESS';ExitCode=0},[pscustomobject]@{Migration='B';Action='Files';Status='FAILED';ExitCode=9})
$rows | Export-Csv (Join-Path $output 'summary.csv') -NoTypeInformation
Set-Content (Join-Path $output 'batch.log') '[2026-10-05 10:00:00] Batch finished: 1 successful; 1 failed.'
exit 9
'@ | Set-Content -LiteralPath $stub -Encoding utf8
    foreach($kind in @('Source','Target','Comparison')){
        $file=Join-Path $root (Get-SmartM365BatchDefinition $kind).Launcher
        $cmd="@echo off`r`n`"$PSHOME\pwsh.exe`" -NoProfile -ExecutionPolicy Bypass -File `"$stub`" -Kind $kind -Root `"$root`" %*`r`nexit /b %ERRORLEVEL%`r`n"
        [IO.File]::WriteAllText($file,$cmd,[Text.UTF8Encoding]::new($false))
        foreach($preview in @($true,$false)){
            $id='20261005-100000-'+[guid]::NewGuid().ToString('N').Substring(0,8)
            $requestPath=Join-Path $temp "$id.json.txt"
            @{Kind=$kind;Root=$root;Mode='FilesOnly';AuthMode='Interactive';Names=@('A','B');PlanOnly=$preview;BatchId=$id;Status='Starting'} | ConvertTo-Json | Set-Content $requestPath
            & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File (Join-Path $generic 'SmartM365-SharePointMigration-BatchGuiRun.ps1') -RequestPath $requestPath
            $receipt=Get-Content $requestPath -Raw | ConvertFrom-Json
            Assert-BatchGui ($receipt.Status -eq $(if($preview){'Previewed'}else{'Failed'})) 'Worker completion status is incorrect.'
            Assert-BatchGui ($receipt.ExitCode -eq $(if($preview){0}else{9})) 'Child exit code was lost.'
            Assert-BatchGui ([bool]$receipt.FinishedUtc) 'Completion timestamp missing.'
            $actual=Get-Content (Join-Path $root 'arguments.json') -Raw | ConvertFrom-Json
            Assert-BatchGui ($actual.MigrationNames -eq 'A,B' -and $actual.BatchId -eq $id -and $actual.Kind -eq $kind) 'CMD arguments or exact batch identity lost.'
            if(-not $preview){
                $r=Get-SmartM365BatchResult $root $kind -BatchId $id
                Assert-BatchGui ($r.Total -eq 2 -and $r.Success -eq 1 -and $r.Failed -eq 1 -and $r.State -eq 'Completed with errors.') 'Summary counts or terminal status incorrect.'
                Assert-BatchGui ($r.Date -eq '2026-10-05 10:00') 'Date depends on culture.'
                Set-Content (Join-Path $r.Directory 'batch.log') 'CSV was written but work continues.'
                $r=Get-SmartM365BatchResult $root $kind -BatchId $id
                Assert-BatchGui ($r.State -like 'Running or incomplete*') 'A progressive CSV was treated as completion.'
            }
        }
    }
    # WPF controls and events use the actual GUI resource dictionary.
    $guiText=[IO.File]::ReadAllText((Join-Path $projectRoot 'SmartM365-SharePointMigration-GUI.ps1'))
    $match=[regex]::Match($guiText,"(?s)\[xml\]\`$xaml\s*=\s*@'\r?\n(.*?)\r?\n'@")
    Assert-BatchGui $match.Success 'Main GUI XAML not found.'
    [xml]$xml=$match.Groups[1].Value
    $window=[Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xml))
    $panel=$window.FindName('panelBatches')
    $panel.Visibility='Visible';$window.FindName('panelSummary').Visibility='Collapsed'
    Initialize-SmartM365BatchGui -Panel $panel -Root $root -SourceRoot $root
    Assert-BatchGui ($script:BatchGui.HelperRoot -eq $generic) 'GUI worker path is not relative to the helper library.'
    Sync-SmartM365BatchMigrations @('A','B')
    $v=$script:BatchGui.View
    Assert-BatchGui (-not $v.FindName('batchSourceRun').IsEnabled) 'Source execution enabled before prerequisite verification.'
    Assert-BatchGui ($v.FindName('batchTargetRun').IsEnabled -and $v.FindName('batchComparisonRun').IsEnabled) 'Available destination/comparison actions disabled.'
    $script:CurrentMigration=[pscustomobject]@{Name='UnrelatedHeaderMigration'}
    $selection=Get-SmartM365BatchSelection Target
    Assert-BatchGui ($selection.Names.Count -eq 0 -and $selection.Mode -eq 'Both') 'Default batch scope depends on migration header.'
    $command=Get-SmartM365BatchCommand @selection
    Assert-BatchGui ($command.Command -notmatch '-MigrationNames|-BatchId') 'All-migrations copy/preview command contains an empty scope or identity.'
    $v.FindName('batchAll').IsChecked=$false
    [void]$v.FindName('batchNames').SelectedItems.Add('B')
    Update-SmartM365BatchControls
    Sync-SmartM365BatchMigrations @('A','B','C')
    $selection=Get-SmartM365BatchSelection Comparison
    Assert-BatchGui ($selection.Names.Count -eq 1 -and $selection.Names[0] -eq 'B') 'Subset selection lost during GUI refresh.'
    $script:BatchGui.SourceReady=$true;Update-SmartM365BatchControls
    Assert-BatchGui $v.FindName('batchSourceRun').IsEnabled 'Source ready state did not enable execution.'
    $v.FindName('batchSourceRoot').Text=Join-Path $temp 'DifferentRoot'
    Assert-BatchGui (-not $v.FindName('batchSourceRun').IsEnabled -and -not $script:BatchGui.SourceReady) 'Root change retained stale prerequisite approval.'
    $v.FindName('batchSourceRoot').Text=$root
    # Exercise GUI request creation and completion without opening a console.
    # Only the synthetic launcher above runs; capture the intended visible-window options.
    function Start-Process {
        param($FilePath,$ArgumentList,$WindowStyle,[switch]$PassThru)
        Assert-BatchGui ($WindowStyle -eq 'Normal' -and $ArgumentList -like '-NoExit *') 'Batch console is not configured as a visible independent console.'
        $path=[regex]::Match($ArgumentList,'-RequestPath "([^"]+)"').Groups[1].Value
        Assert-BatchGui ([bool]$path) 'GUI request path missing.'
        & $FilePath -NoProfile -File (Join-Path $generic 'SmartM365-SharePointMigration-BatchGuiRun.ps1') -RequestPath $path | Out-Null
        [pscustomobject]@{HasExited=$false}
    }
    Start-SmartM365BatchFromGui Target -PlanOnly
    Assert-BatchGui $script:BatchGui.Active.ContainsKey('Target') 'GUI did not register its preview request.'
    Assert-BatchGui (-not $v.FindName('batchTargetRun').IsEnabled) 'A second destination batch is allowed while the GUI request is active.'
    Refresh-SmartM365BatchGui -Force
    Assert-BatchGui (-not $script:BatchGui.Active.ContainsKey('Target') -and $v.FindName('batchTargetStatus').Text -like 'Previewed*') 'GUI did not consume the completion receipt.'
    Start-SmartM365BatchFromGui Source
    Assert-BatchGui (-not $script:BatchGui.Active.ContainsKey('Source')) 'Source request launched without verified prerequisites.'
    # Verify actual button handlers route each card, without launching real processes.
    function Start-SmartM365BatchFromGui {param([string]$Kind,[switch]$PlanOnly) $script:Clicked+= "$Kind/$([bool]$PlanOnly)"}
    foreach($kind in @('Source','Target','Comparison')){
        $v.FindName("batch${kind}Preview").RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
        Assert-BatchGui ($script:Clicked[-1] -eq "$kind/True") 'Preview button routes to the wrong card.'
    }
    $v.FindName('batchAll').IsChecked=$true;Update-SmartM365BatchControls
    Refresh-SmartM365BatchGui -Force
    if($PreviewDirectory){
        [void](New-Item -ItemType Directory -Path $PreviewDirectory -Force)
        $window.FindName('tabSummary').IsChecked=$false;$window.FindName('tabBatches').IsChecked=$true
        $content=$window.Content;$window.Content=$null;$content.Resources=$window.Resources;$content.Background=$window.Background
        foreach($width in @(1280,1920)){
            $content.Measure([Windows.Size]::new($width,900));$content.Arrange([Windows.Rect]::new(0,0,$width,900));$content.UpdateLayout()
            $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new($width,900,96,96,[Windows.Media.PixelFormats]::Pbgra32)
            $bitmap.Render($content)
            $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new();$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $stream=[IO.File]::Create((Join-Path $PreviewDirectory "batch-runs-$width.png"));try{$encoder.Save($stream)}finally{$stream.Dispose()}
        }
    }
    Write-Host '[PASS] Batch GUI offline: CMD routing, previews, failure receipts, summaries, scope, prerequisites gate and WPF events.'
}finally{
    Stop-SmartM365BatchGui
    if($window){$window.Close()}
    # Delete only the unique synthetic directory created above.
    if([IO.Path]::GetFullPath($temp).StartsWith([IO.Path]::GetFullPath($env:TEMP),[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $temp -Recurse -Force}
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB36+yRlENHu5tY
# pBfP0u/L+dQTfdHt9fcx8mUH8jNz56CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDaMqIrHIAdGI8LXFatrcHr
# IxMRZbb/bLwCd/cGVYEiKzANBgkqhkiG9w0BAQEFAASCAYBVhJ+dUq04ZZY2wdIM
# 9qMHu8kkAVQriYYu1iyL7kN1oPGruyuwlGPnXP2wgR7NrAMrfSyVx6h2kKwmjNLQ
# qEI3M5D1EsU2eGFprUCS1Z9gJv/aaNPKJZckbvI09lifHfvQswQ/IvPUBQuM8Wu0
# yEQM/MUNtP68b/VajdoOMqDqvM5jktItYnhcsGHQMC3f7L2M9+hlFeGaCzdmdqdd
# TaWjoxM5M17KwTyiRJ46IbkEY6Rd0pMtgBaro/fFhEmTFyiOdlZwzDUXlpshf37s
# hItuqFIR1h6EuamwNLFCrCwdeeBJHhkhTRQih1/5OQ3b1vOboDW9I8GpU51+DUr9
# fH0w7uqWm9qdx85YBM+48+hWe+u27X2DX+niT4V9GvvVc8YoGo1BMfdR84M6GoV+
# PaYLySkrmHeKUvz0eFkYTx1CWYvqtwXLBhlhuCLF0jrrxeT86PRUdkHwAsjbZj0I
# 5oI3TZDLUXQBGezxWGsy5DelKPRWsktvL2w5n1mUl2Gza2c=
# SIG # End signature block
