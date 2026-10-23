#Requires -Version 5.1
<#
.SYNOPSIS
Offline interaction, view coverage and export checks for the nine demo applications.
.VERSION
1.0.2
.NOTES
Run one application per STA process. No operational handler is loaded.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Registration','Analyzer','HybridJoin','Windows11','Readiness','Notifications','Remediation','SharePoint','IntunePackager')][string]$Application,
    [string]$PreviewFolder='',
    [switch]$InteractiveSmoke
)
$ErrorActionPreference='Stop'
Import-Module Microsoft.PowerShell.Utility
Import-Module Microsoft.PowerShell.Management
$root=Split-Path $PSScriptRoot -Parent
$paths=@{
    Registration='Devices/DeviceRegistrationTool/SmartM365-DeviceRegistration-Tool.ps1'
    Analyzer='Devices/EndpointDiagnosticsAnalyzer/SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
    HybridJoin='Devices/IntuneHybridJoinToolkit/Scripts/SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1'
    Windows11='Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1'
    Readiness='Exchange/MigrationReadiness/SmartM365-ExchangeMigrationReadiness-GUI.ps1'
    Notifications='Exchange/ExchangeUserNotifications/SmartM365-ExchangeUserNotifications-GUI.ps1'
    Remediation='Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1'
    SharePoint='SharePointMigration/SmartM365-SharePointMigration-GUI.ps1'
    IntunePackager='Intune/IntunePackager/SmartM365-IntunePackager-GUI.ps1'
}
$demoEntryPath=Join-Path $root $paths[$Application]
$script:checks=0
function Assert-Demo {param([bool]$Condition,[string]$Message)if(-not $Condition){throw $Message};$script:checks++}
function Get-DemoProtectedState {
    $folder=Split-Path $demoEntryPath -Parent
    if((Split-Path $folder -Leaf) -eq 'Scripts'){$folder=Split-Path $folder -Parent}
    $files=@(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File)
    $config=Join-Path $folder 'Config'
    if(Test-Path -LiteralPath $config){$files+=@(Get-ChildItem -LiteralPath $config -Filter '*.json' -File -Recurse)}
    $result=@{}
    $sha=[Security.Cryptography.SHA256]::Create()
    try{foreach($file in $files){$result[$file.FullName]=[Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($file.FullName)))+'|'+$file.LastWriteTimeUtc.Ticks}}
    finally{$sha.Dispose()}
    return $result
}
$protectedBefore=Get-DemoProtectedState
# Fail closed if a demo accidentally reaches an operational command.
foreach($command in @('Import-Module','Install-Module','Connect-MgGraph','Connect-ExchangeOnline','Connect-PnPOnline','Connect-SPOService','Invoke-MgGraphRequest','Invoke-RestMethod','Invoke-WebRequest','Invoke-Command','New-PSSession','Get-CimInstance','Get-WmiObject','Restart-Computer','Send-MailMessage','Send-MgUserMail','Start-Process','Set-Clipboard')){
    Set-Item -Path ('Function:'+ $command) -Value {throw 'An operational command was reached by an offline demo.'}
}
function Save-DemoPreview {
    param($State,[string]$Name)
    if(-not $PreviewFolder){return}
    [void][IO.Directory]::CreateDirectory($PreviewFolder)
    $window=$State.Window
    $visual=$window.Content
    $visual.Measure([Windows.Size]::new($window.Width,$window.Height))
    $visual.Arrange([Windows.Rect]::new(0,0,$window.Width,$window.Height))
    $visual.UpdateLayout()
    $visual.Dispatcher.Invoke([action]{},[Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new([int]$window.Width,[int]$window.Height,96,96,[Windows.Media.PixelFormats]::Pbgra32)
    $background=[Windows.Media.DrawingVisual]::new()
    $drawing=$background.RenderOpen()
    $drawing.DrawRectangle([Windows.Media.Brushes]::White,$null,[Windows.Rect]::new(0,0,$window.Width,$window.Height))
    $drawing.Close()
    $bitmap.Render($background)
    $bitmap.Render($visual)
    $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $safe=$Name -replace '[^A-Za-z0-9-]','-'
    $stream=[IO.File]::Create((Join-Path $PreviewFolder ($Application+'-'+$safe+'.png')))
    try{$encoder.Save($stream)}finally{$stream.Dispose()}
}
try{
    $s=. $demoEntryPath -Demo -ValidateOnly
    if($Application -in @('Analyzer','Windows11')){Assert-Demo ($s.Controls.ProductVersionText.Text -eq 'v1.0.0') 'Visible product version is not 1.0.0.'}
    Assert-Demo ($s.Mode -eq 'Demo') 'Demo startup did not return its state.'
    Assert-Demo ($s.Window.Title -like '*DEMO*') 'Window is not visibly marked as demo.'
    $markup=Get-SmartM365DemoMarkup $demoEntryPath
    $original=Get-SmartM365DemoControls $s.Window $markup
    foreach($name in $original.Keys){Assert-Demo ($s.Controls.ContainsKey($name)) ('Original control missing: '+$name)}
    # Verify both copies used by standalone apps match the maintained source.
    foreach($helper in @('SmartM365.GuiDemo.ps1','SmartM365.GuiDemo.Reports.ps1')){
        $local=Join-Path (Split-Path $demoEntryPath -Parent) $helper
        $canonical=Join-Path (Join-Path $root 'Modules/SmartM365.UI') $helper
        Assert-Demo ([IO.File]::ReadAllText($local) -ceq [IO.File]::ReadAllText($canonical)) ('Helper copy differs: '+$local)
    }
    foreach($scenario in @('Overview','Issues','Completed')){
        $s.Controls.DemoScenario.SelectedItem=$scenario
        Assert-Demo ($s.Scenario -eq $scenario) ('Scene did not change: '+$scenario)
    }
    $s.Controls.DemoScenario.SelectedItem='Overview'
    Save-DemoPreview $s 'Overview'
    foreach($control in $s.Tabs){
            foreach($tab in @($control.Items)){
                if($tab -isnot [Windows.Controls.TabItem]){continue}
                $control.SelectedItem=$tab
                Assert-Demo ($control.SelectedItem -eq $tab) ('Tab unavailable: '+$tab.Header)
                Save-DemoPreview $s ([string]$tab.Header)
            }
    }
    if($Application -eq 'SharePoint'){
        foreach($tab in @('tabSummary','tabArchivedOverview','tabFiles','tabDiagnostics','tabBatches','tabOperations','tabLogs','tabConfig')){
            Switch-SharePointDemoPanel $tab
            Assert-Demo ($s.Data.SelectedPanel -eq $tab) ('Panel unavailable: '+$tab)
            Save-DemoPreview $s $tab
        }
    }
    if($Application -eq 'Notifications'){
        Assert-Demo ($s.Data.Templates.Count -ge 15) 'Language template coverage is incomplete.'
        Assert-Demo (@($s.Data.Templates.Campaign|Select-Object -Unique).Count -eq 3) 'A campaign is missing.'
        foreach($template in $s.Data.Templates){
            $html=Get-NotificationsDemoHtml $template
            Assert-Demo ($html -match 'DEMO - Sample data' -and $html -notmatch '\{\{[^}]+\}\}') ('Email tokens unresolved: '+$template.FileName)
        }
    }
    if($Application -eq 'Readiness'){
        Assert-Demo ($s.Data.Options.Count -eq 59) 'Assessment check catalogue is incomplete.'
        $s.Controls.MailboxFilterBox.Text='user03@'
        Assert-Demo ($s.Controls.FindingsGrid.Items.Count -eq 1) 'Mailbox filter did not isolate the sample user.'
        $s.Controls.MailboxFilterBox.Clear()
    }
    if($Application -eq 'Analyzer'){
        $s.Controls.AppsSearch.Text='Microsoft'
        Assert-Demo ($s.Controls.AppsGrid.Items.Count -eq 1) 'Application filter did not match its sample row.'
        $s.Controls.AppsSearch.Clear()
        $s.Controls.ZipCategoryFilter.SelectedItem='Logs'
        Assert-Demo ($s.Controls.ZipFilesGrid.Items.Count -eq 21) 'ZIP category filter did not isolate synthetic log files.'
        $s.Controls.ZipCategoryFilter.SelectedItem='All'
    }
    if($Application -in @('HybridJoin','Windows11')){Assert-Demo ($s.Controls.LotModeCombo.Items.Count -eq 4) 'An original LOT launch mode is missing.'}
    if($Application -eq 'Windows11'){Assert-Demo ($s.Controls.SetupLanguageCombo.Items.Count -eq 11) 'Setup language choices are incomplete.'}
    if($Application -eq 'Remediation'){Assert-Demo ($s.Data.Packages.Count -eq 22) 'The shipped package categories are incomplete.'}
    # Raise the real WPF events: this checks routing as well as profile actions.
    foreach($name in @($s.Controls.Keys|Sort-Object)){
        $control=$s.Controls[$name]
        if($control -is [Windows.Controls.Button] -and $name -notmatch '^Demo|Close|Cancel'){
            $control.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
            Assert-Demo ($s.LastAction -eq $name) ('Button is not routed: '+$name)
        }
    }
    Assert-Demo (@($s.Activity|Where-Object {$_ -match 'Demo action failed:'}).Count -eq 0) ('A demo action failed: '+($s.Activity -join ';'))
    $outputs=@(Export-SmartM365DemoProfileOutputs)
    Assert-Demo ($outputs.Count -gt 0) 'No demo output was generated.'
    foreach($path in $outputs){
        Assert-Demo (Test-Path -LiteralPath $path -PathType Leaf) ('Missing export: '+$path)
        Assert-Demo ($path.StartsWith($s.OutputRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) ('Export escaped its demo directory: '+$path)
        if($path -like '*.html'){
            $html=[IO.File]::ReadAllText($path)
            if($Application -eq 'Analyzer'){Assert-Demo ($html -match '<h1>[^<]* v1[.]0[.]0</h1>') 'Analyzer report version is not 1.0.0.'}
            if($Application -eq 'Windows11'){Assert-Demo ($html -match 'Windows 11 Upgrade Toolkit v1[.]0[.]0</div>') 'Windows 11 report version is not 1.0.0.'}
            Assert-Demo ($html -match 'DEMO' -and $html -match '<html|<!doctype') ('Invalid or unmarked HTML: '+$path)
            Assert-Demo ($html -notmatch [regex]::Escape($env:USERPROFILE)) ('HTML contains a real Windows profile: '+$path)
            foreach($identity in @($env:USERNAME,$env:COMPUTERNAME)){
                if($identity){Assert-Demo ($html -notmatch ('(?i)\b'+[regex]::Escape($identity)+'\b')) ('HTML contains a real machine/operator name: '+$path)}
            }
            Assert-Demo ($html -notmatch '\{\{[^}]+\}\}') ('Unresolved HTML token: '+$path)
        }
    }
    if($PreviewFolder){
        $reportFolder=Join-Path $PreviewFolder ($Application+'-outputs')
        [void][IO.Directory]::CreateDirectory($reportFolder)
        foreach($path in $outputs){Copy-Item -LiteralPath $path -Destination $reportFolder -Force}
        $result=[pscustomobject]@{Application=$Application;Checks=$script:checks;Controls=$s.Controls.Count;Views=$s.ViewNames.ToArray();Actions=$s.ActionCounts.Keys.Count;Outputs=$outputs;Mode='Offline/synthetic'}
        $result|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PreviewFolder ($Application+'-validation.json.txt')) -Encoding UTF8
    }
    foreach($child in $s.ChildWindows){$child.Close()}
    $s.Window.Close()
    $protectedAfter=Get-DemoProtectedState
    Assert-Demo ($protectedAfter.Count -eq $protectedBefore.Count) 'A runtime configuration file was created or removed.'
    foreach($path in $protectedBefore.Keys){Assert-Demo ($protectedAfter[$path] -ceq $protectedBefore[$path]) 'An operational configuration changed during demo validation.'}
    if($InteractiveSmoke){
        # A dispatcher timer closes only the demo window created by this test.
        $script:DemoSmokeProblem=''
        $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[timespan]::FromSeconds(9)
        $timer.Add_Tick({if($script:SmartM365Demo.Window.IsVisible){
            if($Application -eq 'Notifications'){
                try{if([string]$script:SmartM365Demo.Controls.TemplatePreviewBrowser.Document.body.innerText -notmatch 'DEMO - Sample data'){$script:DemoSmokeProblem='The integrated browser did not display the fictional email preview.'}}
                catch{$script:DemoSmokeProblem=$_.Exception.Message}
            }
            $script:SmartM365Demo.Window.Close();$timer.Stop()
        }})
        $timer.Start()
        $smokeView=if($Application -eq 'Notifications'){'Templates'}else{''}
        . $demoEntryPath -Demo -DemoView $smokeView
        $timer.Stop()
        Assert-Demo (-not $script:DemoSmokeProblem) $script:DemoSmokeProblem
    }
    Write-Output ('PASS '+$Application+': '+$script:checks+' checks; '+$outputs.Count+' outputs; original WPF events exercised.')
}catch{
    Write-Output ('FAIL '+$Application+': '+$_.Exception.Message)
    Write-Output $_.ScriptStackTrace
    exit 1
}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxXapLc1Uz56QJ
# xb7mw3xWDKmvEBGYwvHxYE+Lp/tgfqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHxZ+Qx2JmXMfiZSRTmMpo/XwGRTrFTddSmFJJ5pGAZKMA0GCSqG
# SIb3DQEBAQUABIIBgKsZbQMzBExeMCcw1o5UWKMaI5r3eRp93S6qRMbHf5CNMgmM
# iCRWn/q624EFH/9C4p7KzNLLULUUv4RtqUQ0Q/c2jXtcT25oRajF5TmdBKQtoRZI
# PtnDizfxBomCZDyS6PCQoK+9H9lZzsYY+6/bJAT3qHWFkaYm3h9KtleDCR8kTTVs
# tLL2BKG6qLt3sfp8lKLZjvcKxmdWht29NvWWuAuO7jbfw7BRhwwHEm9x9u02BHN3
# juxyQS/leTffiClX66wI2QlmoDBf28IsAX2yY8P8vOVUesOxYX2wV+MDbEV8LyqR
# xqa6vpYQ3FEJP0wJjP6GARaKuxSoNgENHFB8x2g735AGLSPu2CkcIEk+nZ3rn/PO
# VUxw7zWpHFy2WP95VOFAhor3svVfMVorFIQR9W7f2UElCEGPokCJDM1Ige5TNsaf
# fhFu6Y+Aa+dl8ZmliiUbcXYjj8pAP9ebptWTqRSr8Yf1HJ6ZK/MVwsWIpt5QAZHN
# ijkkAKeOX9RsOflNCqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxODIy
# MzFaMC8GCSqGSIb3DQEJBDEiBCADBj3GknmVyRlCp5e/mruKFq7Ka0OyjYGGjM6G
# 2mCW+jANBgkqhkiG9w0BAQEFAASCAgBGJuMpNEeR8Z2IHtumXoctICu3etX7URKi
# 199+SoUBRdVEV6lZr4P7Ifvz/OZq6f0IMM8C1k42B9vjoXE9rtfEgFpIVqjfN5Dy
# aAJvGhqSr4JI2qDmrHOvGgVFZP1CSqo6o7Nyf8nOoKN/ZemRoD9hJBMPl7VorxxT
# 9sinNGnERdPGwrjcpKtpOstgSJoqpt+jb+P/du9JKMKbK3A7WjAOU2XaKWkYP94D
# 6Kgqkse3xW62T+FdWgD1hoTDl2ydqilhWVITVIzOhFNumVaZB9LrX9ROuLhZOkqc
# lJJljX3vkHNIpC4aG98MmHGcGiE9tNQRewj8f4NURut3GVbJZySP8NmuzYKKlzd6
# YcSrKIjWu36TQooRlhyHfor0bRCvfN3wAmSPJeGkxB9Xi5gf2cTgZDQy0/twHfWp
# q7xi3e9KBe/JUnZ6Du38y0ml/vH74FusVUS4DMJJGVJyBirsuDHXy9ECLB5G/f32
# kiBbcs2UuE9FZTSsqrQqZ3I3YG31EEprPWjJvlTsmsTEcdYABKHV2+OYIsu2zlMX
# YDiKCBcWIx+RtuTehE4LeNtsxwEZtz/5nNCEFPNzfDeOYTyp3acUNy+vtTYSGD4s
# iWE4UbhURk53Dkq0qRXXmOggUpOhR9My8aiSINrJ2zZZ9I+YhlyHs07Ls+XNxB7I
# zO8I3U8GJA==
# SIG # End signature block
