<#
.SYNOPSIS
    Offline validation of batch dashboard controls, CMD routing and receipts.
.DESCRIPTION
    Uses synthetic launchers and summaries only. No SharePoint, Graph, real scans
    or prerequisite probes are invoked. Renders the WPF tab at two window widths.
.VERSION
    1.0.2
#>
#requires -Version 7.4
[CmdletBinding()]
param([string]$PreviewDirectory='')
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$window=$null
$script:BatchGui=$null
$projectRoot=Split-Path $PSScriptRoot -Parent
$generic=Join-Path $projectRoot 'Scripts\Launchers\Generic'
. (Join-Path $generic 'SmartM365-SharePointMigration-BatchGui.ps1')
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
function Assert-BatchGui { param([bool]$Condition,[string]$Message) if(-not $Condition){throw $Message} }
$temp=Join-Path $env:TEMP ('SmartM365-BatchGuiTest-'+[guid]::NewGuid().ToString('N'))
$root=Join-Path $temp 'Tool kit & migration'
[void](New-Item -ItemType Directory -Path $root -Force)
$script:Clicked=@()
# Worker consoles wait for Enter after a failure or preview; the offline test must never block.
$env:SPMIG_NO_CLOSE_PROMPT='1'
try {
    # Missing, empty and single-row summaries must remain arrays under StrictMode.
    $summaryFixtureRoot=Join-Path $temp 'summary fixtures'
    foreach($kind in @('Source','Target','Comparison')){
        $parent=Join-Path $summaryFixtureRoot ('Migrations\logs\'+(Get-SmartM365BatchDefinition $kind).LogFolder)
        $r=Get-SmartM365BatchResult $summaryFixtureRoot $kind
        Assert-BatchGui ($r.Total -eq 0 -and $r.State -eq 'No batch result yet.') 'Absent batch was treated as a result.'
        $fixtureId='20261005-120000-abcdef01'
        $fixture=Join-Path $parent $fixtureId
        [void](New-Item -ItemType Directory -Path $fixture -Force)
        $summaryPath=Join-Path $fixture 'summary.csv'
        $logPath=Join-Path $fixture 'batch.log'
        Set-Content -LiteralPath $logPath -Value 'START: work continues.'
        foreach($summaryState in @('Missing','Empty','HeaderOnly')){
            if($summaryState -eq 'Empty'){[IO.File]::WriteAllText($summaryPath,'')}
            if($summaryState -eq 'HeaderOnly'){Set-Content -LiteralPath $summaryPath -Value '"Migration","Action","Status","ExitCode"'}
            $r=Get-SmartM365BatchResult $summaryFixtureRoot $kind -BatchId $fixtureId
            Assert-BatchGui ($r.Total -eq 0 -and $r.Success -eq 0 -and $r.Failed -eq 0 -and
                $r.State -like 'Running or incomplete*' -and $r.Directory -eq $fixture) "$kind/$summaryState summary failed."
            Assert-BatchGui (($r.Summary -eq '') -eq ($summaryState -eq 'Missing')) 'Summary button path disagrees with file availability.'
        }
        $oneRow=[pscustomobject]@{Migration='A';Action='Files';Status='SUCCESS';ExitCode=0}
        $oneRow | Export-Csv -LiteralPath $summaryPath -NoTypeInformation
        $r=Get-SmartM365BatchResult $summaryFixtureRoot $kind
        Assert-BatchGui ($r.Total -eq 1 -and $r.Success -eq 1 -and $r.Failed -eq 0 -and
            $r.State -like 'Running or incomplete*') 'One progressive result was lost or marked completed.'
        Set-Content -LiteralPath $logPath -Value 'Batch finished: 1 successful; 0 failed.'
        $r=Get-SmartM365BatchResult $summaryFixtureRoot $kind
        Assert-BatchGui ($r.Total -eq 1 -and $r.State -eq 'Completed.') 'One completed result failed.'
        $oneRow.Status='FAILED';$oneRow.ExitCode=9
        $oneRow | Export-Csv -LiteralPath $summaryPath -NoTypeInformation
        $r=Get-SmartM365BatchResult $summaryFixtureRoot $kind
        Assert-BatchGui ($r.Total -eq 1 -and $r.Success -eq 0 -and $r.Failed -eq 1 -and
            $r.State -eq 'Completed with errors.') 'One failed result failed.'
    }
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
    $blocked=$false;try{$null=Get-SmartM365BatchSelection Target}catch{$blocked=$true}
    Assert-BatchGui $blocked 'An empty manual selection was accepted.'
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
        # The console closes by itself on success: no -NoExit, BatchGuiRun pauses only on failure or preview.
        Assert-BatchGui ($WindowStyle -eq 'Normal' -and $ArgumentList -like '-NoProfile *' -and $ArgumentList -notlike '*-NoExit*') 'Batch console is not configured as a visible self-closing console.'
        $path=[regex]::Match($ArgumentList,'-RequestPath "([^"]+)"').Groups[1].Value
        Assert-BatchGui ([bool]$path) 'GUI request path missing.'
        & $FilePath -NoProfile -File (Join-Path $generic 'SmartM365-SharePointMigration-BatchGuiRun.ps1') -RequestPath $path | Out-Null
        [pscustomobject]@{HasExited=$false}
    }
    Start-SmartM365BatchFromGui Target -PlanOnly
    Assert-BatchGui $script:BatchGui.Active.ContainsKey('Target') 'GUI did not register its preview request.'
    Assert-BatchGui (-not $v.FindName('batchTargetRun').IsEnabled) 'A second destination batch is allowed while the GUI request is active.'
    Refresh-SmartM365BatchGui -Force
    # Latest destination batch can exist while summary.csv has not been published.
    $pendingId='20261005-130000-abcdef02'
    $pending=Join-Path $root "Migrations\logs\target-scan-batches\$pendingId"
    [void](New-Item -ItemType Directory -Path $pending -Force)
    Set-Content -LiteralPath (Join-Path $pending 'batch.log') -Value 'START: scan continues.'
    Refresh-SmartM365BatchGui -Force
    Assert-BatchGui ($v.FindName('batchTargetLatest').Text -like '*Running or incomplete*' -and
        $v.FindName('batchTargetCounts').Text -eq 'Successful actions: —   Failed actions: —' -and
        $v.FindName('batchTargetLogs').IsEnabled -and -not $v.FindName('batchTargetSummary').IsEnabled) 'Pending target batch card is incorrect.'
    [pscustomobject]@{Migration='A';Action='Files';Status='SUCCESS';ExitCode=0} |
        Export-Csv -LiteralPath (Join-Path $pending 'summary.csv') -NoTypeInformation
    Set-Content -LiteralPath (Join-Path $pending 'batch.log') -Value 'Batch finished: 1 successful; 0 failed.'
    Refresh-SmartM365BatchGui -Force
    Assert-BatchGui ($v.FindName('batchTargetLatest').Text -like '*Completed.' -and
        $v.FindName('batchTargetCounts').Text -eq 'Successful actions: 1   Failed actions: 0' -and
        $v.FindName('batchTargetLogs').IsEnabled -and $v.FindName('batchTargetSummary').IsEnabled) 'Completed single-row target batch did not refresh the card.'
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
    Remove-Item Env:SPMIG_NO_CLOSE_PROMPT -ErrorAction SilentlyContinue
    Stop-SmartM365BatchGui
    if($window){$window.Close()}
    # Delete only the unique synthetic directory created above.
    if([IO.Path]::GetFullPath($temp).StartsWith([IO.Path]::GetFullPath($env:TEMP),[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $temp -Recurse -Force}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBpzHJPOPeCOdVD
# CNCCDs66GP+ihsEHSP6gwVzpB5rUpaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGrylyzpIlOOU5s7kUe4tzgGP25XVpH4gi+aqx8ZP54tMA0GCSqG
# SIb3DQEBAQUABIIBgFO8YiVoD9zo8+PSuZTrm0hKA112ash067dUEhPZgPisNsvo
# QBp2Cc+sfhwebKwUtNIHPzPjIMMvVXpmOrszYj38ANefA3j9/R5jb9ugEtyQ9AZx
# N8Nck7RJYbNPqEnH1rF1xU0Ak0/j7u1+OSGqrLBz2LA/Q+jcOYZQrwl+wei7dUeM
# 8OqYO0yEileFE6fiNJgy4qaKf2EtV6Wtii8UCIoj960p8Bg+iv9RCzEI9Xu3O+o9
# kFr4H5/muBOmPtNwvaPWzCEAz6K9Uqp0tHycyWvEjPRAGOIeg+gjr48o88RzZu/0
# V2vpm214G0h4jrV9jl5Bsa84bJyhYLzJTvtCfqccvLqUA43FYTnx/a5KHUoDhRwx
# SOywV6CqgU0XCoJrp4IAjTehhQcOzNsruBRSCHZxyFmG1r9+/9wNwJwFQsgI0DOk
# KQCWH/AdRFj5J+yI7tnnBfVNersftEMruLhkNEOaSMVnT/ffzJ5xOyn99ryhlQDP
# 4PEfGrn2pShhqDC/7KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjE0
# MjVaMC8GCSqGSIb3DQEJBDEiBCDCcfGDUov6S36Vb72lQ8dUfVOc/Ln0nteZQM5Q
# EY4QzzANBgkqhkiG9w0BAQEFAASCAgAbs96WvkYDXWDl6GupvOVcJAWMmRBvPWfl
# rtA46aKSffXsH46Sq62er7+Yd/L1ycxe37/PFCC37+/qGwB+ZP6R1slG34MJAiON
# kzzqoBCXv+Ww+CUEZ3n1BZHeEGTbod2BaIpL+NLeufRC+MYKiT4bWete4oGUlg5d
# xEfZHHr59GRssxYew8MvagBoHDFwNIhsH6Y6aPhvmojs/IyH756y3nRF6CfKlxFZ
# NQo65dGm6D1lmqgVvLMnKclYv2L5pl/4o65ofwsGleSSO8tuG7Abo6gBaZ7Ic4xq
# rOJneu3pcSeQ3IyetiC3/RwXyAldM3jI0xXPdZJ2g1RmgJbFKzuexmRQOiyXFeWB
# r/5BYw3+e29Qr2Ep0k+MX6NbJz+4/VmyzhoUp4tffRZcMxjau/Z8h+MZiX/DiEpf
# 09ButbZfs7maKXu/yjTZL7KQJJ/CayOO0dpDMlxjt2iTj+LS+V7UV8i8qEPbNtNZ
# zJCMCWMfWnKKc1OZ0fWdrDiQA3FN3gc9OPtIPewpsolf8jXhEz3n/5eoZgcIb7T3
# lhqMoq8WsjAeaWivhmS6oD7AqO1CWIDRrhut4ymzYWHKdUK/zGUQPl8L/9eGN/04
# LIWzihT5yOh0fTmck9GmtFguhTHh5Y3okLnsfgHyzxInQB3FRhQAEZXpWUPewrkC
# AO5i4yb1Ag==
# SIG # End signature block
