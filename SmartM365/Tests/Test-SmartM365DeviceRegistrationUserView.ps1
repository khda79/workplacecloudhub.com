<# .VERSION 1.4.1
Offline presentation scenarios; never executes native diagnostics or repairs.
#>
[CmdletBinding()]
param([string]$BaselinePath='', [string]$CaptureDirectory='')
$ErrorActionPreference='Stop'
$app=Join-Path (Split-Path $PSScriptRoot -Parent) 'Devices\DeviceRegistrationTool'
$source=Join-Path $app 'SmartM365-DeviceRegistration-Tool.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors){throw ($errors|Out-String)}
$functions=@($ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.FunctionDefinitionAst]})
foreach($fn in $functions|Where-Object {$_.Name -in @('Get-DrtUserViewStrings','Get-DrtUserViewModel')}){Invoke-Expression $fn.Extent.Text}
$script:checks=0
function Assert-View($Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
if($BaselinePath){
 $old=[Management.Automation.Language.Parser]::ParseFile($BaselinePath,[ref]$tokens,[ref]$errors)
 foreach($fn in $old.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -ne 'Show-SmartM365DeviceRegistrationGui'}){
  $current=@($functions|Where-Object Name -eq $fn.Name)
  Assert-View ($current.Count -eq 1 -and ($current[0].Extent.Text -replace '\r\n',"`n") -ceq ($fn.Extent.Text -replace '\r\n',"`n")) ('Existing diagnostic/support function changed: '+$fn.Name)
 }
}
$healthy=[pscustomobject]@{Status='HEALTHY';UserSsoState='Present';RegistrationHealth='Healthy';LocalMdmState='Detected';DiagnosticCoverage='Complete';MdmNonIntuneEnrollmentDetected=$false}
$before=$healthy|ConvertTo-Json -Depth 12
foreach($language in 'en','fr'){
 $t=Get-DrtUserViewStrings $language
 $v=Get-DrtUserViewModel $healthy -Language $language
 Assert-View ($v.Tone -eq 'Good' -and $v.Title -eq $t.Good) 'Fully observed healthy state should be clear'
 Assert-View ($v.Account -eq $t.Connected -and $v.Device -eq $t.Connected -and $v.Management -eq $t.Detected) 'Three independent labels'
 $v=Get-DrtUserViewModel $healthy -Findings @('[WARNING] policy or MDM trace') -Language $language
 Assert-View ($v.Tone -eq 'Review' -and $v.Title -eq $t.Review) 'HEALTHY must not hide a warning'
 $trace=$healthy.PSObject.Copy();$trace.MdmNonIntuneEnrollmentDetected=$true
 Assert-View ((Get-DrtUserViewModel $trace -Findings @('[WARNING] trace') -Language $language).Detail -eq $t.ExtraManagement) 'Explain additional trace without claiming a fault'
 foreach($field in 'UserSsoState','RegistrationHealth','LocalMdmState'){
  $partial=$healthy.PSObject.Copy();$partial.$field='Unknown'
  $v=Get-DrtUserViewModel $partial -Language $language
  Assert-View ($v.Tone -eq 'Review') ('Unknown must not look healthy: '+$field)
  $label=switch($field){'UserSsoState'{'Account'}'RegistrationHealth'{'Device'}'LocalMdmState'{'Management'}}
  Assert-View ($v.$label -eq $t.Unknown) ('Unknown must not be represented as absent: '+$field)
 }
 $partial=$healthy.PSObject.Copy();$partial.DiagnosticCoverage='Partial'
 Assert-View ((Get-DrtUserViewModel $partial -Language $language).Tone -eq 'Review') 'Partial coverage needs review'
 $partial=$healthy.PSObject.Copy();$partial|Add-Member NoteProperty UserDiagnosticsError 'denied'
 Assert-View ((Get-DrtUserViewModel $partial -Language $language).Tone -eq 'Review') 'Unavailable diagnostics cannot look healthy'
 Assert-View ((Get-DrtUserViewModel ([pscustomobject]@{Status='ERROR'}) -Language $language).Tone -eq 'Error') 'Failure must guide to support'
 Assert-View ((Get-DrtUserViewModel $healthy -Findings @('[ERROR] unavailable') -Language $language).Tone -eq 'Error') 'Errors have precedence'
 Assert-View ((Get-DrtUserViewModel -Language $language).Tone -eq 'Neutral') 'Empty initial view is not success'
}
Assert-View (($healthy|ConvertTo-Json -Depth 12) -ceq $before) 'Presentation must not mutate the result'
Assert-View ((Get-DrtUserViewStrings 'de').Check -eq (Get-DrtUserViewStrings 'en').Check) 'Existing other languages fall back to English for new labels'
$gui=$functions|Where-Object Name -eq 'Show-SmartM365DeviceRegistrationGui'
$copy=$gui.FindAll({param($n) $n -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Expression.Extent.Text -eq '$userCopyButton' -and $n.Member.Value -eq 'Add_Click'},$true)|Select-Object -First 1
Assert-View ($copy.Extent.Text -match 'Clipboard\]::SetText' -and $copy.Extent.Text -notmatch 'Send-|Start-Process|Invoke-DeviceRegistrationPrtRefresh') 'Simple copy must remain clipboard-only'
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$xamlAssignment=$gui.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$xaml'},$true)|Select-Object -First 1
$markup=Invoke-Expression $xamlAssignment.Right.Extent.Text
$w=[Windows.Markup.XamlReader]::Parse($markup)
$logo=$w.FindName('HeaderLogoImage');$logo.Source=[Windows.Media.Imaging.BitmapImage]::new([uri](Join-Path $app 'WorkplaceCloudHub-lockup-WPF.png'));$logo.Visibility='Visible';$w.FindName('HeaderLogoFallback').Visibility='Collapsed'
Assert-View ($w.FindName('UserDetailsExpander').IsExpanded -eq $false) 'Details are collapsed initially'
$tabs=$w.FindName('MainTabs');$grid=$w.FindName('ContentGrid');$exp=$w.FindName('UserDetailsExpander');$panel=$w.FindName('UserSimplePanel')
Assert-View ($tabs.Items.Count -eq 4) 'Keep all existing technical/admin tabs'
[void]$grid.Children.Remove($tabs);$exp.Content=$tabs;$exp.Visibility='Visible';$panel.Visibility='Visible'
$w.FindName('OptionsTab').Visibility='Collapsed';$w.FindName('RefreshPrtButton').Visibility='Collapsed';$w.FindName('RepairButton').Visibility='Collapsed';$w.FindName('UserCopyButton').Visibility='Visible'
Assert-View ($exp.Content -eq $tabs -and $exp.IsExpanded -eq $false) 'User tabs stay intact inside the collapsed details'
$userSimplePanel=$panel;$userDetailsExpander=$exp;$getSelectedMode={'User'}
foreach($handler in $gui.FindAll({param($n)$n -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Expression.Extent.Text -eq '$userDetailsExpander' -and $n.Member.Value -in @('Add_Expanded','Add_Collapsed')},$true)){Invoke-Expression $handler.Extent.Text}
$exp.IsExpanded=$true
Assert-View ($panel.Visibility -eq 'Collapsed') 'Expanded support details must release the compact panel space'
$exp.IsExpanded=$false
Assert-View ($panel.Visibility -eq 'Visible') 'Collapsing support details must restore the simple panel'
if($CaptureDirectory){[void](New-Item -ItemType Directory -Path $CaptureDirectory -Force)}
foreach($language in 'en','fr'){
 $t=Get-DrtUserViewStrings $language
 foreach($scenario in 'Healthy','Review','Unknown','Error'){
  $sample=$healthy.PSObject.Copy();$findings=@()
  if($scenario -eq 'Review'){$sample.MdmNonIntuneEnrollmentDetected=$true;$findings=@('[WARNING] fictional additional trace')}
  if($scenario -eq 'Unknown'){$sample.LocalMdmState='Unknown';$sample.DiagnosticCoverage='Partial'}
  if($scenario -eq 'Error'){$sample.Status='ERROR'}
  $v=Get-DrtUserViewModel $sample -Findings $findings -Language $language
  foreach($pair in @(@('UserResultTitle','Title'),@('UserResultDetail','Detail'),@('UserNextStep','Next'),@('UserAccountValue','Account'),@('UserDeviceValue','Device'),@('UserManagementValue','Management'))){$w.FindName($pair[0]).Text=$v.($pair[1])}
  foreach($pair in @(@('UserAccountLabel','Account'),@('UserDeviceLabel','Device'),@('UserManagementLabel','Management'),@('UserNoChangesText','NoChanges'),@('UserLocalEvidenceText','LocalEvidence'))){$w.FindName($pair[0]).Text=$t.($pair[1])}
  $w.FindName('RunButton').Content=$t.Check;$w.FindName('UserCopyButton').Content=$t.Copy;$exp.Header=$t.Details
  $w.FindName('UserCopyButton').IsEnabled=$true
  $w.FindName('SubtitleText').Text='DEMO - Sample data | '+$t.Subtitle;$w.FindName('DeviceNameText').Text='CTSO-LAP-0142';$w.FindName('StatusText').Text='DEMO - Sample data';$w.FindName('FooterText').Text='DEMO - Fictional presentation only. No device commands executed.'
  $w.FindName('UserResultBanner').Background=switch($v.Tone){'Good'{'#ECFDF3'}'Review'{'#FFFAEB'}'Error'{'#FEF3F2'}default{'#F8FAFC'}}
  $w.Width=1240;$w.Height=900;$w.Content.Measure([Windows.Size]::new(1240,900));$w.Content.Arrange([Windows.Rect]::new(0,0,1240,900));$w.Content.UpdateLayout()
  Assert-View ($panel.ActualHeight -gt 150 -and $w.FindName('UserResultTitle').ActualHeight -gt 10) ('Native WPF layout: '+$language+' '+$scenario)
  if($CaptureDirectory){$bmp=[Windows.Media.Imaging.RenderTargetBitmap]::new(1240,900,96,96,[Windows.Media.PixelFormats]::Pbgra32);$bg=[Windows.Media.DrawingVisual]::new();$dc=$bg.RenderOpen();$dc.DrawRectangle($w.Background,$null,[Windows.Rect]::new(0,0,1240,900));$dc.Close();$bmp.Render($bg);$bmp.Render($w.Content);$encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new();$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp));$stream=[IO.File]::Create((Join-Path $CaptureDirectory ('User-'+$language+'-'+$scenario+'.png')));try{$encoder.Save($stream)}finally{$stream.Dispose()}}
 }
}
$exp.Content=$null;[void]$grid.Children.Add($tabs);$panel.Visibility='Collapsed';$exp.Visibility='Collapsed'
Assert-View ($tabs.Parent -eq $grid) 'Admin technical layout can be restored without losing tabs'
$w.Close()
[pscustomobject]@{Pass=$true;Checks=$script:checks;NativeDeviceCommandsExecuted=$false;CaptureDirectory=$CaptureDirectory}|ConvertTo-Json

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBETkLDtU+0EaLX
# jBEW0IMWmM3eUl4EKS43v/F/QHHoGqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEmtTFO+l9J3U/uRKoO75oLmkVvLvDS22kY5diJb/xyNMA0GCSqG
# SIb3DQEBAQUABIIBgA6wKghRzmxV6zgbIPDF0C4xsm1Bhc9Mu/a9mOfyJzF1mVlb
# eBXaka9T3kYEXNZjemSnH+QxuunM695Lbp4WgyT8q5e7X2Sx3jFIGpN9cwjeWFuE
# Tw8nPEnAoKQ8T2PxmEfdLRl42VHc/k9PxAWGDmbo94LPYpiHbmUtWyvV3+HoJhHg
# zjn0gfSvC9EKYn+Wn8TYfS++d7VbTBEHzeRxjRdyByjaeXSIlC5YkZG6fiSZTRZA
# I9ekeK1NSrsEp0RRwCNV7jNJ1oYvboX9EsiwNunTH/ZsQPKYwth0hRs0hLI+H7si
# hzHyXxJbVw7Wk1REkAltUD/9X/QqJ7btFvXgtYzSH3+A+lDVuOj0/46sQan3Zhiy
# p3A/gu7fYStq+NgcZl+Vqirem92/gOBiuhDMT/byeQwRp9tTNc5gqz2c1gTO7Pc3
# iUY/fxI4xXcmuzFTX+bSrKd7yjsOUW88DXoeMyQRvXJa0e92MR1IsloAXnvKgSUI
# bD2LDoS69c9DMXjWNqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTIxOTQy
# MDNaMC8GCSqGSIb3DQEJBDEiBCCIxpdrVmKEw284jKvurudjJV0MhnOL/dPF3RXS
# 0y1NTjANBgkqhkiG9w0BAQEFAASCAgAUXr8vXxWJGICQgHh32BXz/XW5U6sekt/x
# IXRaxlf/TNO4ySKiojA6VJgPVyuXCeI6uqCVxG5L/ycAgyQRmwpPB70Dluk/S82w
# jhf7bbw69+NHHu53tXqrpDN0PbHCsiekoRNCXKUHUTqI4py8BVj1oaL6la6Flzka
# METiBYap9ctJ6WH81p3masRtHAmytDZu1x5ZxFsovVJw2xc0y61Grd9VeGM9OUrK
# mw6XQCFbSPdOAVkc1n+Bme06llGKCl5Sqm1BvtZ52U19pj/7V3MJm7XpsyB5t9RU
# h8EmfsT32aishDD9XBy1Ke0GzWvh6CbTUXT0z2q1uyRN1x/Cu6+Gw7pUPcVyiDUh
# ZIOEeRqVjPJmE3nRrNsvQtYmibzxKdqGjPxLKsOsTtHgvA/WS+fBoNhEC4df7Nwj
# YNPJE65rZQ8F4QIX4fDj9tDfXkU2NGPjMwImkWYZ0q7sek5gjSLuEpOT7plCb5T8
# Tg2kBYfGs3PhNIPDuctZdAWJZJHgcJZJE7pcrZ99zus3yq1rd6SrUJDTjMyYM+/F
# PsQjuOjdeCHMUZgoqClAJVevIFLd9foCTesQtgFWjzF9CR+GMZhfjinNxVS/Dv33
# azAWu/dsE1Bjc/ePO4biKLwGPS/i5/AZt7g5L7tgTCSF13tGDYHYeOBUQrPLR7U/
# FR3BwrvaSA==
# SIG # End signature block
