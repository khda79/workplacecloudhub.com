#Requires -Version 7.4
<#
.SYNOPSIS
Offline WPF loading lifecycle and production navigation regression checks.
.VERSION
1.0.0
#>
param([string]$PreviewFolder='')
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$gui=Join-Path $root 'SmartM365-SharePointMigration-GUI.ps1'
$script:checks=0
function Assert-Loading([bool]$Condition,[string]$Message) {
    if(-not $Condition){throw $Message};$script:checks++
}
$s=. $gui -Demo -ValidateOnly
try {
    $overlay=$s.Controls.panelGuiLoading
    Assert-Loading ($overlay.Visibility -eq 'Collapsed') 'Loading shield starts hidden'
    $s.Window.Cursor=[Windows.Input.Cursors]::Arrow
    Invoke-SmartM365SharePointLoading -Window $s.Window -Message 'Chargement de la migration : Contoso-HR (DEMO)' -Action {
        Assert-Loading ($overlay.Visibility -eq 'Visible') 'Shield visible before work'
        Assert-Loading ($s.Controls.lblGuiLoadingDetail.Text -like '*Contoso-HR*') 'Migration context displayed'
        Assert-Loading ($s.Window.Cursor -eq [Windows.Input.Cursors]::Wait) 'Wait cursor during work'
        Invoke-SmartM365SharePointLoading -Window $s.Window -Message 'Nested' -Action {
            Assert-Loading ($overlay.Visibility -eq 'Visible') 'Nested loading remains visible'
        }
        Assert-Loading ($overlay.Visibility -eq 'Visible') 'Nested completion preserves caller shield'
        if($PreviewFolder){
            [void][IO.Directory]::CreateDirectory($PreviewFolder)
            $visual=$s.Window.Content
            $visual.Measure([Windows.Size]::new($s.Window.Width,$s.Window.Height))
            $visual.Arrange([Windows.Rect]::new(0,0,$s.Window.Width,$s.Window.Height));$visual.UpdateLayout()
            $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new([int]$s.Window.Width,[int]$s.Window.Height,96,96,[Windows.Media.PixelFormats]::Pbgra32)
            $bitmap.Render($visual)
            $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new();$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $stream=[IO.File]::Create((Join-Path $PreviewFolder 'SharePoint-Loading-DEMO.png'))
            try{$encoder.Save($stream)}finally{$stream.Dispose()}
        }
    }
    Assert-Loading ($overlay.Visibility -eq 'Collapsed') 'Shield cleared after success'
    Assert-Loading ($s.Window.Cursor -eq [Windows.Input.Cursors]::Arrow) 'Cursor restored after success'
    $caught=$false
    try {Invoke-SmartM365SharePointLoading -Window $s.Window -Message 'Failure' -Action {throw 'synthetic failure'}}
    catch {$caught=$_.Exception.Message -eq 'synthetic failure'}
    Assert-Loading $caught 'Original failure propagates'
    Assert-Loading ($overlay.Visibility -eq 'Collapsed') 'Shield cleared after failure'
    Assert-Loading ($s.Window.Cursor -eq [Windows.Input.Cursors]::Arrow) 'Cursor restored after failure'
    # Extract only production navigation functions; no operational startup or transports.
    . ([scriptblock]::Create((Get-SmartM365DemoFunctionText $gui @('Open-GuiTab','Switch-Tab')) -join [Environment]::NewLine))
    $script:Window=$s.Window
    foreach($name in @('Summary','ArchivedOverview','Files','Operations','Diagnostics','Batches','Logs','Config')){
        Set-Variable -Name ('tab'+$name) -Value $s.Controls[('tab'+$name)]
    }
    foreach($name in @('panelSummary','panelArchivedOverview','panelWorkflows','panelOperations','panelDiagnostics','panelBatches','panelLogs','panelConfig')){
        Set-Variable -Name $name -Value $s.Controls[$name]
    }
    $script:refreshCalls=0
    function Assert-Refresh {Assert-Loading ($overlay.Visibility -eq 'Visible') 'Production refresh runs inside shield';$script:refreshCalls++}
    function Refresh-PortfolioSummary {Assert-Refresh}
    function Refresh-ArchivedOverview {Assert-Refresh}
    function Refresh-DiagnosticInventoryMetrics {Assert-Refresh}
    function Refresh-DiagnosticCrossCheck {Assert-Refresh}
    function Refresh-SmartM365BatchGui {param([switch]$Force) Assert-Refresh}
    foreach($name in @('Summary','ArchivedOverview','Files','Operations','Diagnostics','Batches','Logs','Config')){
        Open-GuiTab $name
        Assert-Loading ([bool]$s.Controls[('tab'+$name)].IsChecked) ('Production tab selected: '+$name)
        Assert-Loading ($overlay.Visibility -eq 'Collapsed') ('Production tab shield cleared: '+$name)
    }
    Assert-Loading ($script:refreshCalls -eq 5) 'All five production refreshes exercised'
    function Refresh-PortfolioSummary {throw 'synthetic tab read failure'}
    $caught=$false
    try {Open-GuiTab Summary} catch {$caught=$_.Exception.Message -eq 'synthetic tab read failure'}
    Assert-Loading $caught 'Production tab refresh failure propagates'
    Assert-Loading ($overlay.Visibility -eq 'Collapsed') 'Production tab refresh failure clears shield'
    Assert-Loading ($s.Window.Cursor -eq [Windows.Input.Cursors]::Arrow) 'Production tab failure restores cursor'
    # Execute the actual production migration-selection event against synthetic objects.
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($gui,[ref]$tokens,[ref]$errors)
    $event=$ast.Find({param($n) $n -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Expression.Extent.Text -eq '$cmbMigration' -and $n.Member.Value -eq 'Add_SelectionChanged'},$true)
    Assert-Loading ($null -ne $event) 'Production migration event located'
    $cmbMigration=[Windows.Controls.ComboBox]::new()
    [void]$cmbMigration.Items.Add('Demo-A');[void]$cmbMigration.Items.Add('Demo-B')
    $script:Migrations=@([pscustomobject]@{Name='Demo-A'},[pscustomobject]@{Name='Demo-B'})
    $script:CurrentMigration=$script:Migrations[0];$script:CurrentStatus=$null
    $script:selectionObserved=$false
    function Set-CurrentMigration {param($Migration)
        $script:selectionObserved=($overlay.Visibility -eq 'Visible' -and $s.Controls.lblGuiLoadingDetail.Text -like '*Demo-B*' -and $Migration.Name -eq 'Demo-B')
    }
    $handler=[scriptblock]::Create($event.Arguments[0].ScriptBlock.Extent.Text.Trim().Substring(1).TrimEnd().TrimEnd('}'))
    $cmbMigration.Add_SelectionChanged($handler)
    $cmbMigration.SelectedIndex=1
    Assert-Loading $script:selectionObserved 'Production migration selection loads inside shield with correct context'
    Assert-Loading ($overlay.Visibility -eq 'Collapsed') 'Migration selection clears shield'
    'PASS: {0} loading/navigation checks; synthetic WPF only.' -f $script:checks
}
finally {$s.Window.Close()}
# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxPRa8SegIhAGp
# QoryDqg5Q6/6uFMKQCDqtyhO6Y292KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDAe1PPwPvTc3zv81AJ/TfZ
# hvt0ChGxGbnsSfdXH9rN8DANBgkqhkiG9w0BAQEFAASCAYAIBV0cqvDMZdiiqrei
# UnF+Hay77m6IlI/EWkGwfFoDpM4411Dl2bhXcVrS+2Il+s/Eygj4trbdAPnVDe1j
# z7eL40ApIZQhqkK7r+OksTGInCoypWk13NvSJzXMIMr99q8t182ssUIqIZCvOMLD
# B0YQeQ0OFk5Sc35RyQdEWqbYt/ne+JvElCH8XHohGz0D5Lm0ydse/CR02/B96jux
# aW/HfTXqtndNEafYFLXA+WFEwRenG1eXPHY/EjkDU6QFT1zFDF1wOkEeQRMBFy5S
# i2d9JbabR4lg81YNT3hNmlMFBLGfdMkQooz6Y2u9CPJ1BfN+BWXLpG6ORJ9F4vJq
# 8dl1I3iiC78KNl3urxRU/ZVFtyO55BF047EcYP51LUCR6yJ4XaRpXkkozO5sgOhv
# o+zsZGntLM+Qtzl1BwPzB//2aDqeynUs2cM/HDp0SIG475OXxb4qsjPBFpwmVrF4
# T/ZrKkE8L2yzUW0EuXHnZz7EvO6YEfKBE/NIw2u1KJCU52Y=
# SIG # End signature block
