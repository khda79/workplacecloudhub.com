#Requires -Version 7.0
<#
.SYNOPSIS
Offline regressions for application process handling and destructive-action recovery.
.VERSION
1.0.0
.NOTES
Only synthetic child processes and mocked Graph operations are used. Run with -STA.
#>
[CmdletBinding()]
param([ValidateSet('All','NotificationProcess','RemediationReset')][string]$Case='All')
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
if($Case -in @('All','NotificationProcess')){ & {
    Add-Type -AssemblyName PresentationFramework
    $source='SmartM365/Exchange/ExchangeUserNotifications/SmartM365-ExchangeUserNotifications-GUI.ps1'
    foreach($name in @('Complete-Run','Invoke-CurrentRunStop','Stop-CurrentRunMonitor','Update-CurrentRunOutput','Invoke-CampaignRun')){Import-ReviewFunction $source $name}
    $script:CurrentProcess=$null;$script:CurrentRunMonitor=$null
    $script:ReviewLines=[Collections.Generic.List[string]]::new()
    $script:ReviewHost=(Get-Command pwsh -ErrorAction Stop).Source
    $script:ReviewChildCommand=''
    $script:ReviewRunning=$false
    function Get-SelectedCampaign {return [pscustomobject]@{Name='Synthetic test';ScriptPath=$PSCommandPath}}
    function Show-Status {param($Text,$Color)}
    function Show-ValidationResult {param($Campaign,[switch]$TemplatesOnly,[switch]$SkipTemplates);return $true}
    function Get-RunMode {return 'DryRun'}
    function Get-PowerShellPath {return $script:ReviewHost}
    function Get-CommandArgument {return @('-NoProfile','-NonInteractive','-Command',$script:ReviewChildCommand)}
    function Add-OutputLine {param($Text);$script:ReviewLines.Add([string]$Text)}
    function Show-RunState {param([bool]$Running);$script:ReviewRunning=$Running}
    function Wait-ReviewProcess {
        $deadline=[datetime]::UtcNow.AddSeconds(30)
        while($script:CurrentProcess -and [datetime]::UtcNow -lt $deadline){Update-CurrentRunOutput;Start-Sleep -Milliseconds 15}
        if($script:CurrentProcess){Invoke-CurrentRunStop;Stop-CurrentRunMonitor;throw 'Synthetic process did not complete.'}
    }
    try {
        $script:ReviewChildCommand='1..400 | ForEach-Object { [Console]::Out.WriteLine("OUT-$_"); [Console]::Error.WriteLine("ERR-$_") }; [Console]::Out.Write("FINAL-WITHOUT-NEWLINE"); exit 5'
        Invoke-CampaignRun
        Assert-Review $script:ReviewRunning 'The process did not enter the running state.'
        Wait-ReviewProcess
        Assert-Review (@($script:ReviewLines | Where-Object {$_ -match '^OUT-\d+$'}).Count -eq 400) 'Stdout lines were lost or duplicated.'
        Assert-Review (@($script:ReviewLines | Where-Object {$_ -match '^STDERR: ERR-\d+$'}).Count -eq 400) 'Stderr lines were lost or duplicated.'
        Assert-Review ($script:ReviewLines.Contains('FINAL-WITHOUT-NEWLINE')) 'The final unterminated line was lost.'
        Assert-Review ($script:ReviewLines.Contains('ExitCode=5') -and -not $script:ReviewRunning) 'The nonzero exit code or UI state was lost.'
        Assert-Review ($null -eq $script:CurrentRunMonitor) 'The monitor remained attached after completion.'
        $script:ReviewLines.Clear();$script:ReviewChildCommand="[Console]::WriteLine('SECOND-RUN'); exit 0"
        Invoke-CampaignRun;Wait-ReviewProcess
        Assert-Review ($script:ReviewLines.Contains('SECOND-RUN') -and $script:ReviewLines.Contains('ExitCode=0')) 'A subsequent run did not complete cleanly.'
        $script:ReviewChildCommand="[Console]::WriteLine('STOP-PROBE'); Start-Sleep -Seconds 60"
        Invoke-CampaignRun;Invoke-CurrentRunStop;Wait-ReviewProcess
        Assert-Review (-not $script:ReviewRunning -and $null -eq $script:CurrentRunMonitor) 'Cancellation did not clean up the process monitor.'
        $script:ReviewHost=Join-Path ([IO.Path]::GetTempPath()) ('missing-review-host-'+[guid]::NewGuid().ToString('N')+'.exe')
        Invoke-CampaignRun
        Assert-Review (-not $script:ReviewRunning -and $null -eq $script:CurrentProcess -and $null -eq $script:CurrentRunMonitor) 'A start failure left the GUI busy.'
    }
    finally {if($script:CurrentProcess){Invoke-CurrentRunStop;Stop-CurrentRunMonitor}}
}}
if($Case -in @('All','RemediationReset')){ & {
    Import-ReviewFunction 'SmartM365/Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1' 'Reset-CloudRemediationHistory'
    $script:GraphBaseUri='https://graph.invalid/beta'
    function Ensure-GuiGraphConnection {}
    function Get-CloudRemediationDetail {param($Id);return @{displayName='Synthetic remediation'}}
    function New-DeviceHealthScriptBodyFromCloudDetail {param($Detail,$DisplayNameOverride);return @{displayName=$DisplayNameOverride}}
    function Get-ObjectValue {param($InputObject,$Name);return $InputObject.$Name}
    function Get-CloudAssignments {param($Id);return @([pscustomobject]@{target=@{groupId='synthetic'}})}
    function Write-GuiLog {param($Message)}
    function Copy-CloudAssignments {param($SourceId,$TargetId,$Assignments);if($script:ReviewFailure -eq 'Assignments'){throw 'Synthetic assignment failure.'}}
    function Invoke-GraphJsonRequest {
        param($Method,$Uri,$Body)
        if($Method -ne 'POST' -or $Uri -notlike 'https://graph.invalid/*'){throw 'Unexpected transport operation.'}
        $script:ReviewObjects.new=$true
        return @{id='new'}
    }
    function Invoke-MgGraphRequest {
        param($Method,$Uri,$ErrorAction)
        if($Method -ne 'DELETE' -or $Uri -notlike 'https://graph.invalid/*'){throw 'Unexpected transport operation.'}
        if($Uri.EndsWith('/old')){
            if($script:ReviewFailure -eq 'BeforeDelete'){throw [TimeoutException]::new('Synthetic response failure before fake DELETE.')}
            [void]$script:ReviewObjects.Remove('old')
            if($script:ReviewFailure -eq 'AfterDelete'){throw [TimeoutException]::new('Synthetic response failure after fake DELETE.')}
        }
        elseif($Uri.EndsWith('/new')){[void]$script:ReviewObjects.Remove('new')}
        else{throw 'Unexpected object ID.'}
    }
    foreach($failure in @('None','Assignments','BeforeDelete','AfterDelete')){
        $script:ReviewFailure=$failure;$script:ReviewObjects=@{old=$true};$caught=$null;$result=$null
        try{$result=Reset-CloudRemediationHistory -CloudRemediation ([pscustomobject]@{Id='old';DisplayName='Synthetic remediation'})}catch{$caught=$_}
        switch($failure){
            'None' {Assert-Review ($null -eq $caught -and $result.Id -eq 'new' -and $script:ReviewObjects.ContainsKey('new') -and -not $script:ReviewObjects.ContainsKey('old')) 'Successful replacement did not retain the new object.'}
            'Assignments' {Assert-Review ($null -ne $caught -and $script:ReviewObjects.ContainsKey('old') -and -not $script:ReviewObjects.ContainsKey('new')) 'Pre-delete rollback did not preserve the original.'}
            'BeforeDelete' {Assert-Review ($null -ne $caught -and $script:ReviewObjects.ContainsKey('old') -and $script:ReviewObjects.ContainsKey('new')) 'Ambiguous original DELETE removed the replacement.'}
            'AfterDelete' {Assert-Review ($null -ne $caught -and -not $script:ReviewObjects.ContainsKey('old') -and $script:ReviewObjects.ContainsKey('new')) 'Applied original DELETE removed the replacement.'}
        }
        if($failure -in @('BeforeDelete','AfterDelete')){Assert-Review ($caught.Exception.Message -match 'Replacement new was retained') 'The operator did not receive reconciliation instructions.'}
    }
}}
Write-Output ('[{0}] PASS: {1}; {2} operational-handler regression checks; no real campaign or Graph mutation.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Case,$script:checks)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBpvCPrkiqGbf8m
# tNbTzpDz5AfL8zBaSO3sVfADPZLLN6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIwPK1SYnrwhr+85uADFBlIngPNm51kd0rXG6aZcb0scMA0GCSqG
# SIb3DQEBAQUABIIBgHtgcNewQBE0/kRhspSNcEi3NriRU4uo8grhQZztTXbsPJpj
# wrl4d/d1m46JPxh1utYS8f/lMDq6icW9cwCBw+uoSds6wLTYrk90n0PwcvLkbCUC
# WFTt5YU04RFfcMprVNLWdZcDp16v0FC90vi1CaSlvSME6WtU/k4l2HZPkCa8wt++
# Q+D6VE8i217CffclPPfil60B+QuptHYa/0dDjnpRv3nERU24iZtUaK76USJUfEWR
# NKeu3pc1QJcW57PDnwzU1S3dA3bW3dO2SVm6jD9qSRbY2Qqy/k3BYkD8Qrb0mwHn
# VTz6XeX4gJWfiWPsRTkMSQq9+q2dRjL0s+si0+0U6lUTY6xTrHTRxl/vIZk3HBt8
# CBwv/NA0y9K/x7YlmsfzBNpU9//7Q6J9QSNVZfruKpslvyEEi3MTtlRVIAuyEmVY
# 71tPV/pqsg3ty/xoJ7kdVvhvSFy/GE/Q38ptw1j/Qyegcsp5aS5u4fWDqlUJ4740
# DrCisE6BYu1p8tAZ8qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMjQw
# MTdaMC8GCSqGSIb3DQEJBDEiBCBGyqEylRbZ+WZmjYK35lIUZ32Gsne+GhP6Xd3I
# 9w8eBzANBgkqhkiG9w0BAQEFAASCAgCjewL1kY6EpGWYk+L+ZOfNqb3/PLvCuI9+
# 4pf3PhRn4wroEX8O2xzbY7BDf6O5NlcNwgRyGTZ60QTOvQTMFjEuz9VK4mQGrDTM
# ObCxPNzDXolz3qrKeSJH/BYBPksBMe24RsmtYmH53HRft5hpi00vkR6dhjivZClE
# Qd+tUZDkhmbWWjbNoVQ8qeP545MyMFWK9taRfVlzDx6uEjWLOSTwEVrseZR2N9RG
# 9eO6h3j2Elyl5ZKJAj+SVWD2IYRr+TUY55u04TQ5fYCInW+V1YT9bhnrbsLb+Rm9
# gXRMiBe29mdsx8Ni6LM6ts+NFKXTkyP7Z9fOv6b0MoqzvB6kNd9TqtAh+J8kRN7p
# LiDwpAIkl/ThqMJgDGy1bMJzH3K1A+cipSZsxjeROVD12PYGxmBprgWcve9bPnMd
# M02msjdSJiR5Rw4fk4djIQ3Y+13qivOTqomypOyS5LmKzxhHcozrZvM3ScavdefG
# BjzTsXC5nfQoCJdot1RTkvFNjHnCEHhN8U8LAjXcWrF9GSdlz8k77OnXxkxfCsfr
# BV3tmwhsgZxA8qO9L86zmBBdHnsp54cObLnJc9z0nJxnJ+jTPnjDqwwh30Jeqx5M
# /qHH1q6CDezG12LneLvvzTbj4DBuwIjQJ+CSPcvLep8FikMYjn6Mg2yt8uJAnWWb
# z4nJLUQnxw==
# SIG # End signature block
