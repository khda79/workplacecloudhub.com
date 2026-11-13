#Requires -Version 5.1
<# .VERSION
1.0.3
#>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. (Join-Path $repo 'Release\WorkplaceCloudHub.ReleaseSafety.ps1')
$script:checks=0
function Assert-Policy($value,$message){if(-not $value){throw $message};$script:checks++}
function Must-Reject([scriptblock]$action,[string]$message){$rejected=$false;try{& $action|Out-Null}catch{$rejected=$true};Assert-Policy $rejected $message}
function Make-Zip($path,[hashtable]$entries){$z=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Create);try{foreach($name in $entries.Keys){$entry=$z.CreateEntry($name);$stream=$entry.Open();try{$bytes=$entries[$name];if($bytes -isnot [byte[]]){$bytes=[Text.Encoding]::UTF8.GetBytes([string]$bytes)};$stream.Write($bytes,0,$bytes.Length)}finally{$stream.Dispose()}}}finally{$z.Dispose()}}
$temp=Join-Path $env:TEMP ('WCH-ReleasePolicy-'+[guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($temp)
try{
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $safe=Join-Path $temp safe;[void][IO.Directory]::CreateDirectory($safe)
 [IO.File]::WriteAllText((Join-Path $safe 'New-DemoAssets.ps1'),'# executable local demo generator')
 [IO.File]::WriteAllText((Join-Path $safe 'app.ps1'),'# distributable application')
 Assert-WchNoDemoAssets -Path $safe;Assert-Policy $true 'Runtime demo code must remain allowed'
 $dirty=Join-Path $temp dirty;[void][IO.Directory]::CreateDirectory((Join-Path $dirty 'nested\dEmOaSsEtS'))
 [IO.File]::WriteAllText((Join-Path $dirty 'nested\dEmOaSsEtS\capture.png'),'synthetic')
 Must-Reject {Assert-WchNoDemoAssets -Path $dirty} 'Nested/case-insensitive DemoAssets folder accepted'
 $badZip=Join-Path $temp innocent.zip;Make-Zip $badZip @{'nested\DemoAssets\image.png'='synthetic'}
 Must-Reject {Assert-WchNoDemoAssets -Path $badZip} 'Backslash ZIP entry accepted'
 $nested=Join-Path $temp wrapper.zip;Make-Zip $nested @{'content/payload.zip'=[IO.File]::ReadAllBytes($badZip)}
 Must-Reject {Assert-WchNoDemoAssets -Path $nested} 'Hidden nested archive media accepted'
 $renamed=Join-Path $temp payload.dat;[IO.File]::WriteAllBytes($renamed,[IO.File]::ReadAllBytes($badZip))
 Must-Reject {Assert-WchNoDemoAssets -Path $renamed} 'Renamed ZIP bypassed inspection'
 $wrappedRenamed=Join-Path $temp disguised.zip;Make-Zip $wrappedRenamed @{'payload.dat'=[IO.File]::ReadAllBytes($badZip)}
 Must-Reject {Assert-WchNoDemoAssets -Path $wrappedRenamed} 'Nested renamed ZIP bypassed inspection'
 $named=Join-Path $temp 'Product-DemoAssets.zip';Make-Zip $named @{'otherwise-safe.txt'='safe'}
 Must-Reject {Assert-WchNoDemoAssets -Path $named} 'Media archive attachment name accepted'
 $nupkg=Join-Path $temp innocent.nupkg;Make-Zip $nupkg @{'content/DemoAssets/image.png'='synthetic'}
 Must-Reject {Assert-WchNoDemoAssets -Path $nupkg} 'NuGet media accepted'
 $corrupt=Join-Path $temp corrupt.zip;[IO.File]::WriteAllText($corrupt,'not a ZIP')
 Must-Reject {Assert-WchNoDemoAssets -Path $corrupt} 'Corrupt archive was silently accepted'
 $safeZip=Join-Path $temp safe.zip;Make-Zip $safeZip @{'Runtime/New-DemoAssets.ps1'='# local generator';'Runtime/app.ps1'='# application'}
 Assert-WchNoDemoAssets -Path $safeZip;Assert-Policy $true 'Valid runtime ZIP rejected'
 $deep=$safeZip
 for($level=0;$level -lt 10;$level++){$next=Join-Path $temp ('depth'+$level+'.zip');Make-Zip $next @{'payload.bin'=[IO.File]::ReadAllBytes($deep)};$deep=$next}
 Must-Reject {Assert-WchNoDemoAssets -Path $deep} 'Excessive nesting was not blocked'
 $paths=@('SmartM365\Devices\SmartEndpointMigrationToolkit\Release\SmartM365-EndpointMigration-Build-Release.ps1','SmartM365\Devices\DeviceRebootManager\Release\SmartM365-Build-DeviceRebootManagerRelease.ps1','SmartM365\Devices\DeviceRebootManager\PowerShellGallery\SmartM365-New-DeviceRebootManagerGalleryVmBundle.ps1','SmartM365\Devices\DeviceRebootManager\PowerShellGallery\SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1','SmartM365\Devices\EndpointDiagnosticsAnalyzer\PowerShellGallery\SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1','SmartM365\Devices\EndpointDiagnosticsAnalyzer\PowerShellGallery\SmartM365-Publish-EndpointDiagnosticsAnalyzerGalleryPackage.ps1','SmartM365\Devices\EndpointDiagnosticsAnalyzer\IntuneWin32\SmartM365-Build-EndpointDiagnosticsAnalyzerIntunePackage.ps1','SmartM365\Devices\Windows11UpgradeToolkit\IntuneWin32\Build-SmartM365Windows11IntunePackage.ps1','SmartM365\Devices\SmartEndpointMigrationToolkit\Campaigns\Windows11Upgrade\IntuneWin32\SmartM365-EndpointMigration-Build-Windows11IntunePackage.ps1','SmartM365\Intune\IntunePackager\SmartM365.IntunePackager.Packaging.psm1')
 $tokens=$null;$errors=$null;$canonical=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'Release\WorkplaceCloudHub.ReleaseSafety.ps1'),[ref]$tokens,[ref]$errors)
 foreach($path in $paths){
  $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo $path),[ref]$tokens,[ref]$errors);Assert-Policy (-not $errors) ('Syntax error '+$path)
  foreach($name in 'Test-WchDemoAssetPath','Assert-WchNoDemoAssets'){
   $expected=@($canonical.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$true)|Where-Object Name -eq $name)[0]
   $actual=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$true)|Where-Object Name -eq $name)
   Assert-Policy ($actual.Count -eq 1 -and ($actual[0].Extent.Text -replace '\r\n',"`n") -ceq ($expected.Extent.Text -replace '\r\n',"`n")) ('Standalone policy drift '+$path+' '+$name)
  }
  Assert-Policy (@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Assert-WchNoDemoAssets'},$true)).Count -gt 0) ('Guard not invoked '+$path)
  # Every embedded standalone guard executes the adversarial fixture independently.
  foreach($fn in $actual){Invoke-Expression $fn.Extent.Text}
  Must-Reject {Assert-WchNoDemoAssets -Path $nested} ('Embedded guard accepted contaminated archive '+$path)
  if($path -like '*DeviceRebootManager*' -and $path -match '(Release|VmBundle)'){
   $zipFunction=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$true)|Where-Object Name -eq New-ForwardSlashZip)
   if($zipFunction.Count){Invoke-Expression $zipFunction[0].Extent.Text;$blockedOutput=Join-Path $temp ([guid]::NewGuid().ToString('N')+'.zip');Must-Reject {New-ForwardSlashZip -SourceRoot $dirty -DestinationPath $blockedOutput} 'Actual ZIP writer accepted contaminated staging';Assert-Policy (-not(Test-Path $blockedOutput)) 'Rejected writer created an artifact'}
  }
 }
 $zipBuilder=Join-Path $repo 'Release\New-WorkplaceCloudHubReleaseZip.ps1';$out=Join-Path $temp built.zip
 $files=@((Join-Path $safe app.ps1),(Join-Path $safe New-DemoAssets.ps1))
 $skip=Join-Path $safe 'DemoAssets\capture.png';[void][IO.Directory]::CreateDirectory((Split-Path $skip -Parent));[IO.File]::WriteAllText($skip,'synthetic');$files+=$skip
 $result=& $zipBuilder -SourceRoot $safe -FilePath $files -OutputPath $out -Build
 Assert-Policy ($result.ExcludedDemoAssets.Count -eq 1 -and $result.Files.Count -eq 2) 'Generic builder dropped runtime code or shipped media'
 Assert-WchNoDemoAssets -Path $out
 $blockedOutput=Join-Path $temp not-created.zip
 Must-Reject {& $zipBuilder -SourceRoot $temp -FilePath $nested -OutputPath $blockedOutput -Build} 'Generic ZIP builder accepted a contaminated nested input'
 Assert-Policy (-not(Test-Path $blockedOutput)) 'Contaminated build created a publishable artifact'
 $notes=Join-Path $temp notes.md;[IO.File]::WriteAllText($notes,'# release')
 $publisher=Join-Path $repo 'Release\Publish-WorkplaceCloudHubRelease.ps1';$sha='a'*40
 $script:ghCalls=0;function gh {$script:ghCalls++;throw 'No GitHub call permitted in this offline test'}
 $preview=& $publisher -Repository example/tool -Tag v1.0.0 -TargetCommit $sha -AssetPath $out -NotesPath $notes
 Assert-Policy ($preview.Validated -and -not $preview.PublicationAttempted -and $script:ghCalls -eq 0) 'Preview contacted GitHub'
 Must-Reject {& $publisher -Repository example/tool -Tag v1.0.0 -TargetCommit $sha -AssetPath $nested -NotesPath $notes -Publish} 'Publisher allowed nested media before upload'
 [IO.File]::WriteAllText($notes,'Download Product-DemoAssets.zip')
 Must-Reject {& $publisher -Repository example/tool -Tag v1.0.0 -TargetCommit $sha -AssetPath $out -NotesPath $notes -Publish} 'Publisher accepted stale media attachment notes'
 Assert-Policy ($script:ghCalls -eq 0) 'Rejected publication contacted GitHub'
 if($PSVersionTable.PSVersion.Major -ge 7){
  Import-Module (Join-Path $repo 'SmartM365\Intune\IntunePackager\SmartM365.IntunePackager.Packaging.psm1') -Force
  $issues=@(Test-IpPackageBuildRequest -SourceFolder $dirty -SetupFile setup.exe -OutputFolder (Join-Path $temp out))
  Assert-Policy (($issues -join ' ') -match 'DemoAssets are local-only') 'Intune build preflight missed contaminated source'
  Must-Reject {Invoke-IpIntuneWinAppUtil -ToolPath missing.exe -SourceFolder $dirty -SetupFile setup.exe -OutputFolder (Join-Path $temp out)} 'Direct Intune engine did not reject before starting external tool'
 }
 [pscustomobject]@{Pass=$true;Checks=$script:checks;PowerShell=$PSVersionTable.PSVersion.ToString();GitOperations=0;Publications=0;EndpointActions=0}|ConvertTo-Json
}finally{if([IO.Path]::GetFullPath($temp).StartsWith([IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $temp -Recurse -Force}}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDN6hYnlDev3MMz
# Crdk6UuvlmCsjLIZ/ZGYFBc7Eb+rTqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMSys7gdZLm/RVx+2JX/5G8T1CNXS9OpxpI03MhiW4+LMA0GCSqG
# SIb3DQEBAQUABIIBgKMOI2YdRxzAthmWvzLi1yqpzL047MxkGGiTccgiZtrHyz3+
# kqWaffti7QJ7piUbI7WBHx/KtezS/equu79MZ7Je3kOWCBqxHptT7NPMVc+d5kDn
# j+v72GbBiDZPdjGEQ2fW2/VsKFftnm5TLHRz3ow/yJfAXjW1vHxd+qlFv8tZvCkH
# cQhbjyQcHKBpoh9ioVRUEubV2ISBIOX4MUA3wI0Ng27dYMWdIAp/C1MrCpaQ9W70
# kWB7ngl81fJ71G2B9+K0nfvhzZY00wJGeP7nCmBpeDEzX0ecKKiAlQlIwTelPxm0
# ygOtCSIvoF+Cpxibgv9b8wZM1WZj5OgbOY1ptSfG8hy6NllXuVibSHaLTlApy+HJ
# PdREkbqq2ZqhojJzoDXLgrgZUUTEEnVR1wdsMKovVQ1OUJOrvxOSsXOOmlaJuCdW
# QuTavBwPAZv3mHa+8dLqeiAvfeo8QTCpFUeVqgg+wh4cDKbf7+nE2rWWNk6hiCYk
# ZBb4G6HgBslsy4yF06GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMDE4
# MDZaMC8GCSqGSIb3DQEJBDEiBCDQb7DYtLfdOMU8NIOo/Yz+7PoIwZTOzMTFaIVR
# 7eAReDANBgkqhkiG9w0BAQEFAASCAgCKOoJl3FMY+5M/YK8YYqMq8noo0agtsdSS
# t6uhqs9aPGKAn0jMyB0h2xe+Dr8U07rp/PD0pWwN8fbRZrhTdDFGOyGRpMBBwnl2
# 7DCxXboPLIbQhfIUFxi918WCLpw95cZlEqnsuXmdscXzrPStMlcRjKLQ55J7Fg3M
# lJ1kkqf2PDnAVhaq7cALn3QJQVv8nj85WwB7usyHKvAGrCRIrYm48odsw5Fosnj2
# wZhGELoSyGvhGIM3/QIZ/c6TJ7ljyrAUp7cm04M2r/3s9jPtFziaK0cxasEf/Zns
# JCZpe+ZBx8LS4fAkPtercxNFXrN+FUcn5Cpp9wGhqAu/5ShdjnTLfzoGEiiP5cfw
# v+UtilTh+EAesKIpZ475jIqGFMoGpRT7mOZ7omP3RVMnx5jY5P9FlqR/Y9P0UZxP
# sXyi5P3cDbo8sR5dgIreJ6gWBWlB+krxZiOuXAIV+UpCmL4dMCqzBzNv1GJ6wbUO
# PSUHXa9WVevvyIBtPZMi5ufRlmHRNKGx9Jgp/JBDrqJ7Z5EmPPer9msFE6EJbFIo
# l8DqJc5CeX0ewFAacUQzoKr2AmX0DXnrJRrMKjzrGQYXheN9fsBjS6CvZ85dRD9a
# Q9dFp/NI6NJgeXmCGUZhmPoHPbsTzJ4xYGGFZk7ggl4kmrqYYqib9LdYuzE4jY0L
# a0x0eVIo4w==
# SIG # End signature block
