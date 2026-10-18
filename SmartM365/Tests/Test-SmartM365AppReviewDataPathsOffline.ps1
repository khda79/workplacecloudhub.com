#Requires -Version 5.1
<#
.SYNOPSIS
Offline regressions for enrollment backups, assessment provenance and Unicode launchers.
.VERSION
1.0.0
.NOTES
Registry, scheduled tasks and assessment transports are mocked. CMD probes only echo a temporary directory.
#>
[CmdletBinding()]
param([ValidateSet('All','EnrollmentBackup','ReadinessState','LotUnicode')][string]$Case='All')
$ErrorActionPreference='Stop'
$repoRoot=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$script:checks=0
function Assert-Review {param([bool]$Condition,[string]$Message)if(-not $Condition){throw $Message};$script:checks++}
function Import-ReviewFunction {
    param([string]$RelativePath,[string]$Name)
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot $RelativePath),[ref]$tokens,[ref]$errors)
    if($errors.Count){throw "Source parse failed: $RelativePath"}
    $definition=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name},$true)
    if(-not $definition){throw "Function not found: $Name"}
    Set-Item -Path ('Function:script:'+$Name) -Value $definition.Body.GetScriptBlock()
}
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-AppReview-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
try {
if($Case -in @('All','EnrollmentBackup')){ & {
    $script:ComputerName='SYNTHETIC';$script:RunId='Offline'
    function ConvertTo-CleanNativeOutput {param($Output);return [string]$Output}
    function Write-SmartM365DeviceRegistrationLog {param($Message)}
    function Write-RunLog {param($Message,$Level)}
    function Test-Path {
        param($LiteralPath,$Path,$PathType)
        $target=if($LiteralPath){$LiteralPath}else{$Path}
        if($target -like 'Registry::*'){return $target -match '^Registry::HKLM\\SOFTWARE\\Microsoft\\Enrollments\\(?:Status\\)?synthetic$'}
        return Microsoft.PowerShell.Management\Test-Path -LiteralPath $target -PathType $(if($PathType){$PathType}else{'Any'})
    }
    function reg.exe {
        $script:Exports++
        if($script:Exports -eq 2 -and $script:BackupFailure -eq 'Exit'){$global:LASTEXITCODE=1;return 'Synthetic export failure'}
        if(-not ($script:Exports -eq 2 -and $script:BackupFailure -eq 'Missing')){
            $text=if($script:Exports -eq 2 -and $script:BackupFailure -eq 'Empty'){''}else{'Windows Registry Editor Version 5.00'}
            [IO.File]::WriteAllText([string]$args[2],$text)
        }
        $global:LASTEXITCODE=0
    }
    function Remove-Item {param($LiteralPath,[switch]$Recurse,[switch]$Force,$ErrorAction);if($LiteralPath -notlike 'Registry::*'){throw 'Unexpected deletion'};$script:Removals++}
    function Get-ScheduledTask {param($TaskPath,$ErrorAction);$script:TaskOperations++;return @()}
    function Unregister-ScheduledTask {throw 'Unexpected scheduled task mutation'}
    function New-Object {
        param($TypeName,$ComObject)
        if($ComObject){$script:TaskOperations++;throw 'Synthetic COM boundary'}
        return Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName
    }
    $implementations=@(
        @('SmartM365/Devices/DeviceRegistrationTool/SmartM365-DeviceRegistration-Tool.ps1','Remove-DeviceRegistrationMdmEnrollment'),
        @('SmartM365/Devices/IntuneHybridJoinToolkit/Scripts/SmartM365-Invoke-IntuneHybridJoinRepair.ps1','Remove-NonIntuneMdmEnrollment')
    )
    foreach($implementation in $implementations){
        Import-ReviewFunction $implementation[0] $implementation[1]
        foreach($failure in @('Exit','Missing','Empty','None')){
            $script:BackupFailure=$failure;$script:Exports=0;$script:Removals=0;$script:TaskOperations=0
            $output=Join-Path $testRoot ($implementation[1]+'-'+$failure)
            $parameters=@{EnrollmentIds='synthetic';OutputDirPath=$output;RemovalLabel='Offline'}
            if($implementation[1] -eq 'Remove-NonIntuneMdmEnrollment'){$parameters.ComputerNameValue='SYNTHETIC';$parameters.RunIdValue='Offline'}
            $result=& $implementation[1] @parameters
            Assert-Review ($script:Exports -eq 2) ("Both existing registry keys were not exported: {0}; {1}; count={2}; detail={3}" -f $implementation[1],$failure,$script:Exports,$result.Detail)
            if($failure -eq 'None'){
                Assert-Review ($result.Success -and $script:Removals -eq 2 -and $script:TaskOperations -gt 0) 'Verified backups did not permit cleanup.'
            }else{
                Assert-Review (-not $result.Success -and $script:Removals -eq 0 -and $script:TaskOperations -eq 0) 'A failed backup allowed registry or task removal.'
                Assert-Review ($result.Detail -match 'blocked before any removal') 'Backup failure did not explain that cleanup was blocked.'
            }
        }
    }
}}
if($Case -in @('All','ReadinessState')){ & {
    $source='SmartM365/Exchange/MigrationReadiness/SmartM365-ExchangeMigrationReadiness-GUI.ps1'
    foreach($name in @('Clear-SemrAssessmentState','Import-SemrCsvToGui')){Import-ReviewFunction $source $name}
    $controls=@{}
    foreach($name in @('SummaryGrid','FindingsGrid','GlobalFindingsGrid','PermissionsGrid','BatchGrid')){$controls[$name]=[pscustomobject]@{ItemsSource=@('old')}}
    foreach($name in @('GoCountText','WarningCountText','NoGoCountText','UnknownCountText','CsvPathBox','CsvMetadataText','StatusText')){$controls[$name]=[pscustomobject]@{Text='old'}}
    foreach($name in @('OpenOutputButton','ComparePermissionsButton','RunButton')){$controls[$name]=[pscustomobject]@{IsEnabled=$true}}
    $controls.RunProgress=[pscustomobject]@{Value=100}
    $script:ProgressWindow=$null
    function Start-SemrProgressWindow {param($Title,$Stage,$Detail)}
    function Update-SemrProgressWindow {param($Stage,$Detail,$Current,$Total,[switch]$Indeterminate)}
    function Complete-SemrProgressWindow {param($Stage,$Summary,[switch]$Failed)}
    function Write-SemrActivity {param($Message,$Level)}
    function Show-SemrError {param($Title,$ErrorRecord)}
    function Import-SemrBatchCsv {param($Path);if($Path -eq 'invalid.csv'){throw 'Synthetic invalid CSV'};return [pscustomobject]@{Path=$Path;Rows=@([pscustomobject]@{Identity='fictional@example.invalid'});Encoding='UTF-8';Delimiter=';';IdentityColumn='Identity'}}
    foreach($path in @('batch-b.csv','invalid.csv')){
        $script:Assessment='old-assessment';$script:Export='old-export';$script:AssessmentBatchPath='batch-a.csv'
        $controls.CsvPathBox.Text=$path
        Import-SemrCsvToGui
        Assert-Review ($null -eq $script:Assessment -and $null -eq $script:Export -and $script:AssessmentBatchPath -eq '') 'Loading another batch retained an assessment or export.'
        Assert-Review (-not $controls.ComparePermissionsButton.IsEnabled -and -not $controls.OpenOutputButton.IsEnabled) 'Old assessment actions remained enabled.'
        Assert-Review (@('SummaryGrid','FindingsGrid','GlobalFindingsGrid','PermissionsGrid'|Where-Object {$null -ne $controls[$_].ItemsSource}).Count -eq 0) 'An old results grid remained populated.'
        Assert-Review ($controls.GoCountText.Text -eq '0' -and $controls.RunProgress.Value -eq 0) 'Old counters remained visible.'
        Assert-Review (($path -eq 'invalid.csv' -and $null -eq $script:Batch -and -not $controls.RunButton.IsEnabled) -or ($path -eq 'batch-b.csv' -and $script:Batch.Path -eq $path -and $controls.RunButton.IsEnabled)) 'Batch import state is inconsistent.'
    }
}}
if($Case -in @('All','LotUnicode')){ & {
    $unicodePath=Join-Path $testRoot 'équipe 東京'
    [void][IO.Directory]::CreateDirectory($unicodePath)
    foreach($relative in @('SmartM365/Devices/IntuneHybridJoinToolkit/Scripts/SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1','SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1')){
        foreach($name in @('Assert-CmdLiteralValue','ConvertTo-CmdArgument','ConvertTo-CmdWindowTitle','New-GuiLaunchCommandFile','Write-ToolkitConfig')){Import-ReviewFunction $relative $name}
        $launcher=New-GuiLaunchCommandFile -WorkingDirectory $unicodePath -Commands @('echo REVIEW-CWD=%CD%','cmd /d /c exit 7') -NamePrefix 'Offline-Review'
        $start=[Diagnostics.ProcessStartInfo]::new();$start.FileName=$env:ComSpec;$start.Arguments='/d /s /c ""'+$launcher+'""';$start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true;$start.StandardOutputEncoding=[Text.Encoding]::UTF8;$start.StandardErrorEncoding=[Text.Encoding]::UTF8
        $process=[Diagnostics.Process]::Start($start)
        try{$stdout=$process.StandardOutput.ReadToEnd();$stderr=$process.StandardError.ReadToEnd();$process.WaitForExit();Assert-Review ($process.ExitCode -eq 7) ('CMD exit code lost: '+$stderr);Assert-Review ($stdout.Contains('REVIEW-CWD='+$unicodePath)) ('Unicode working directory changed: '+$stdout)}finally{$process.Dispose()}
        $config=Join-Path $unicodePath 'toolkit.config'
        Write-ToolkitConfig -Path $config -Values @{W11UT_MEDIA=$unicodePath;EHJIR_MEDIA=$unicodePath}
        $bytes=[IO.File]::ReadAllBytes($config)
        Assert-Review ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191 -and ([IO.File]::ReadAllText($config)).Contains($unicodePath)) 'Config did not preserve Unicode for Windows PowerShell.'
        foreach($unsafe in @('bad%PATH%','bad!value',"bad`nvalue",'bad"value')){$caught=$false;try{ConvertTo-CmdArgument $unsafe|Out-Null}catch{$caught=$true};Assert-Review $caught 'Unsafe CMD literal was accepted.'}
    }
}}
Write-Output ('[{0}] PASS: {1}; {2} data/path regression checks; mocked operational boundaries.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Case,$script:checks)
}finally{
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $tempPrefix=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($tempPrefix,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notlike 'SmartM365-AppReview-*'){throw 'Unsafe test cleanup path'}
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKS9YX2cf7TzC0
# TYTgF1jTKXmtxR4aXoZzW7/XO3Gy4aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE/TB8Uh56D1G3pcTrsN6F5wZ+VFMy9W67tv/M58M8E5MA0GCSqG
# SIb3DQEBAQUABIIBgCJzF0Y3Xg7BF2Wkn5QENmrKCgemzuEamFkfz9+ssguXsZvZ
# QW8fdK6e1YedVzE16EgmLGjYc+mAzmny5V4t3mflX4i1E3AGU8XM4MLmah2JKwnY
# rjqwLTLXjIPF58v69i9rwqNYycN3TsdkWrKBO39YC5D5FfzuaymHTLBHpj2tyIpD
# BC8ONtQuzafjBtVSJ7a3jxmyUcE/150OKLbviTDRfzgKwyADAfgp3zj4I6BwsdJG
# FcfAyQoYCdaHHofwOtwsfeUkz/0dU3h+q2VhKHxBmybEkB66eeV181bL8myeZ5ke
# FASVGxFxG91fxu0NKhhRmWkQplVT4srBk4+R2Bk9uLQt5sEh+RkIrhsHTGHIsHmk
# YEeTWM1wCli4F9QbOw7p7FisPtNAtGyTIr0ey9WI3kxXBGoY/0QrT/vpRbolwxcV
# ET+cP7cx8AWhRkvT6WbUryn/37MQzhZHHyznQrPIxiK3nlfEGedRBvkIRZjYkg43
# c0u+J+J2UpdDk6jKS6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMjQw
# MThaMC8GCSqGSIb3DQEJBDEiBCC7nQljFf0H2o9SNqnXGka4ijXey05pur4ihgPR
# ai8tWjANBgkqhkiG9w0BAQEFAASCAgAhFLTuD/LKSBh9pBasPW7LSY751OY9Dt6Q
# 9aA8oiPSUEBag1kHk+AwLrPkSM7lZeWjMYZJz4bt0hSFfDn7XXaWwsQWq2C49BqT
# emwViL2NGCvjNiPa8TUg11VhG9TVnxTk2uKTPZLgtqqZo0R8l/oXpjuaJQVBFg4M
# q0tqOmDg3spTk/xz7AGmeneVrr5tq/znE5Mx3mJMFG+5PXuY1xuzfICPIeiOgjfx
# KNQBGAkXnVzeD1e0MHJtg2zq2K5PC5XLLzzw338CN4l64au0G+fSLTdHCe3wVqvd
# TXD/TKtrJqczXEccaR3K0yqVeQLH97jO+VJ7iOGrfPfOQN6DaDejUgYjGLjves3K
# J6SJU3CGwJa/ZA77tUHdIrFZd/p0AyhNYTwWZ+7OXfpHPrtP9ZKWsuJuxv7Nc1cl
# CBquHARrOCjx5zRZOlu4+YPjzF/xHvBA2YYK1byO9WXTM2wY/P9HFhUyO5ze5334
# 306le4KwgBIo1xei+8LqKOCNphK9oiBLquWht6hPzVQaDhqnzwaOXV5Fj+lvn2+F
# l7rFOxq3+MRAW5jnEJFMvBlQHBd00Wk6Y0XZL8Y7/Epnhj0bfEbuAlSm7z1A4dgX
# P3N64kFTUl/j5p1Kfxly59mWsJBBIkj0QGRduu0a/xkLfVfjxA7tcegNkpoahJAB
# JBHpFVLP9g==
# SIG # End signature block
