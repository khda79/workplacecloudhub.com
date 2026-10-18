#Requires -Version 5.1
<#
.SYNOPSIS
Native report rendering with exclusively fictional inputs for SmartM365 demos.
.VERSION
1.0.0
#>

function script:Export-SmartM365DemoNativeReport {
    param([string]$Kind,[string]$SourcePath,[string]$Name,[AllowNull()]$Data)
    $s=$script:SmartM365Demo
    $path=Get-SmartM365DemoOutputPath $Name
    $entry=@{Analyzer='Export-SedaHtmlReport';Readiness='New-SemrHtmlReport';HybridJoin='New-CycleHtmlReport';Windows11='New-Windows11UpgradeCycleHtmlReport'}[$Kind]
    if(-not $entry){throw 'Unknown native demo report.'}
    $ast=Get-SmartM365DemoAst $SourcePath
    $definitions=@{}
    foreach($fn in $ast.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst]},$true)){$definitions[$fn.Name]=$fn}
    $queue=[Collections.Generic.Queue[string]]::new();$queue.Enqueue($entry)
    $selected=@{}
    while($queue.Count){
        $nameToLoad=$queue.Dequeue()
        if($selected.ContainsKey($nameToLoad)){continue}
        if(-not $definitions.ContainsKey($nameToLoad)){throw ('Native demo report entry is missing: '+$nameToLoad)}
        $fn=$definitions[$nameToLoad];$selected[$nameToLoad]=$fn
        foreach($call in $fn.FindAll({param($node)$node -is [Management.Automation.Language.CommandAst]},$true)){
            $called=$call.GetCommandName()
            # Logging is supplied below and never loads its operational implementation.
            if($called -and $called -ne 'Write-SedaLog' -and $definitions.ContainsKey($called) -and -not $selected.ContainsKey($called)){$queue.Enqueue($called)}
        }
    }
    foreach($fn in $selected.Values){
        $text=$fn.Extent.Text
        # The Hybrid report normally records the real operator. A demo must never do so.
        $text=$text.Replace('[System.Security.Principal.WindowsIdentity]::GetCurrent()','$null').Replace('[Environment]::UserName',"'demo.operator'").Replace('$env:COMPUTERNAME',"'CTSO-ADMIN-01'")
        . ([scriptblock]::Create($text))
    }
    function Write-SedaLog {param($Message,$Level) Write-SmartM365DemoActivity $Message}
    function New-HtmlLogLink {param($Path,$Label='Open') return '<a href="#demo-log">Sample log</a>'}
    $demoLogoPath=Join-Path $s.AssetRoot 'WorkplaceCloudHub-lockup-WPF.png'
    $demoLogoUri=if(Test-Path -LiteralPath $demoLogoPath){'data:image/png;base64,'+[Convert]::ToBase64String([IO.File]::ReadAllBytes($demoLogoPath))}else{''}
    function Get-IntuneHybridJoinBrandLogoDataUri {return $demoLogoUri}
    function Get-Windows11BrandLogoDataUri {return $demoLogoUri}
    # Supply only safe literal defaults to the renderer, never run its script body.
    if($ast.ParamBlock){foreach($parameter in $ast.ParamBlock.Parameters){
        $value=$false
        if($parameter.DefaultValue){try{$value=$parameter.DefaultValue.SafeGetValue()}catch{$value=''}}
        Set-Variable -Name $parameter.Name.VariablePath.UserPath -Value $value -Scope Local
    }}
    foreach($fn in $selected.Values){foreach($variable in $fn.FindAll({param($node)$node -is [Management.Automation.Language.VariableExpressionAst] -and $node.VariablePath.UserPath -like 'script:*'},$true)){
        Set-Variable -Name $variable.VariablePath.UserPath.Substring(7) -Value $null -Scope Script
    }}
    $script:AppName='Smart Endpoint Diagnostics Analyzer - DEMO'
    if($Kind -eq 'Analyzer'){[void](Export-SedaHtmlReport -Analysis $Data -Path $path)}
    if($Kind -eq 'Readiness'){[void](New-SemrHtmlReport -Assessment $Data -Path $path)}
    if($Kind -in @('HybridJoin','Windows11')){
        $LotRoot='Demo/LOT-CONTOSO';$LauncherVersion='DEMO';$RunOnce=$true;$DryRun=$true
        $script:LotRoot=$LotRoot;$script:LauncherVersion='DEMO';$script:LauncherLogPath='Demo/launcher.log'
        $script:PsExecSecurityEvidenceRows=@();$script:LauncherOptionRows=@([pscustomobject]@{Category='DEMO';Option='Execution';Value='Simulated only'})
        $script:CurrentCycleComputerListStats=$null;$script:CurrentCycleTotalComputers=12
        $script:UseEffectiveTechnicianRunGuardHistory=$false;$script:TechnicianRunGuardHistoryPath='Demo/run-guard.csv'
        if($Kind -eq 'HybridJoin'){$columns=@(Get-LauncherReportColumns)}else{
            $assignment=$ast.Find({param($node)$node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$reportColumns'},$true)
            if(-not $assignment){throw 'Windows 11 report schema missing.'}
            $reportColumns=@($assignment.Right.Expression.SafeGetValue())
            $columns=$reportColumns
        }
        $rows=@(for($i=0;$i -lt 12;$i++){
            $status=@('SUCCESS','SUCCESS','PENDING_REBOOT','ADMIN_SHARE_UNREACHABLE')[$i%4]
            New-SmartM365DemoRow -Columns $columns -Index $i -Values @{
                Computer=('CTSO-LAP-{0:0000}' -f ($i+1));ComputerName=('CTSO-LAP-{0:0000}' -f ($i+1));Cycle=1;CycleNumber=1
                Status=$status;EffectiveStatus=$status;RemoteStatus=$status;RemoteNextAction='Review sample result';NextAction='Review sample result'
                EffectiveNextAction='Review sample result';RemoteExitCode=0;PsExecExitCode=0;DryRun=$true
                InteractiveUserName=('user{0:00}' -f ($i+1));InteractiveUserDomain='CONTOSO';UserName=('user{0:00}@contoso.example' -f ($i+1))
                Timestamp='2026-10-18 09:30:00';LauncherVersion='DEMO';RemoteDetail='Fictional endpoint evidence.'
                LogPath='Demo/device.log';PsExecLogPath='Demo/psexec.log';RemoteLogsPath='Demo/CollectedLogs';RemoteCurrentRunLogsPath='Demo/CurrentRun'
            }
        })
        $script:AttemptRowCount=$rows.Count
        $script:AttemptStatusCounter=@{SUCCESS=6;PENDING_REBOOT=3;ADMIN_SHARE_UNREACHABLE=3}
        $script:AttemptNextActionCounter=@{'Review sample result'=12};$script:AttemptAdminShareFailureCounter=@{DNS_FAILED=3}
        if($Kind -eq 'HybridJoin'){New-CycleHtmlReport -Summary $rows -Path $path -CycleNumber 1 -GeneratedAt ([datetime]'2026-10-18T09:30:00')}
        else{New-Windows11UpgradeCycleHtmlReport -Summary $rows -Path $path -CycleNumber 1 -GeneratedAt ([datetime]'2026-10-18T09:30:00')}
    }
    Set-SmartM365DemoMarking $path
    $html=[IO.File]::ReadAllText($path)
    $html=$html -replace '(?i)(Generated:|Generated: </strong>)\s*\d{4}-\d{2}-\d{2} \d{2}:\d{2}','Generated: 2026-10-18 09:30'
    $html=$html.Replace('</body>','<details id="demo-log"><summary>Sample log</summary><pre>[2026-10-18 09:30:00] DEMO - Contoso operation simulated successfully.</pre></details></body>')
    [IO.File]::WriteAllText($path,$html,[Text.UTF8Encoding]::new($false))
    $s.Reports[$Kind]=$path
    return $path
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJSM+lQQmT3Lgx
# 0/WzZJ3SNw4/ovuy4mk1/FYTVHznl6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPafqDicFUAR3Ja/BrF/oT+yImLdS6SfUg5S0zqsiCy7MA0GCSqG
# SIb3DQEBAQUABIIBgD19WTaqSLXG73AcQxkvjn9vxMUndSqKcaz1zch2iaYe0x5m
# s0sMDsZrHlOQmWHcSQ43h3aYIEVmif+DSRmGu+fPz3wGDaf5aE5KyYgI1x0Vl6eD
# tJIIAu2BE1uaVoBPtwZlmnYfWtNHddAb8kd+YClK1mA8T6zbzFHI0dyq2tqcSqeb
# 4j0hho2jTPwOfQZ/AFdY1vbxdFZcuzYvjTARZlYWGaEvKy5eSWhv12UcuXln1pAz
# ADV3TOIbrnKBh/0u9dTBVfCnlpgQLUhtVtCZNQ9egd6q64qjYlH5uQHi94kXwbYk
# eaHcWMs9T2lfNPo8CB/X+hVFZzHUqCjQaq8EVe68/VnyejAtAHXaEWzjEXL3rrC0
# gFy+rjIW7rFjttZJI9aGIi7v0UsqCLG3dIKGwRSQNlpigYnNgJLChEolkNdohLW1
# +YrYVOpCkxOtl1/cq9ww4OwZUVgK4Rwy0UUVy5XYHxYUgMryCpsHeANuCOd4Rnwt
# /XPme3OVut9ZMN4X1aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMDQ2
# MzZaMC8GCSqGSIb3DQEJBDEiBCAi8kESEL1otgFg5fKTKPE26X6juiIqtRQPlVlj
# yKrYKjANBgkqhkiG9w0BAQEFAASCAgAazFi6SfgBQkLV4ICgwPs8JJdJFAX3s6Si
# 4JhPChZuFaWjFQOEcBxaHX57lVLBPm2i6938EZ+NeB+rbkzNtS3nrWkoEwb17Dr8
# u2MLWEtQoQHJkZIznzvy5TcH96pf+EDSSXw/hV8boc9Vo/k2kW4vdhFE9YM1oEy7
# 8cYQ+1l4bMcG4J9YdhD0ilRFnkNo/WDlT3XY1nWE4Da/fhKGSE05QgbMDxBOrtfn
# BSQ13dqfZlTV4FOTnbg5JcJti9hBvti6cstgWtxns+HQxzQZ6iEvMXVIMAKx5hEB
# y6PxIIpXYd3y1S/mkWV4kHkG8wb/nXMKzwaYns/AiBtjZyqa+qH1kWySYh8Ou635
# uEml1+qlvVuOWc6mvrVmLCjc8+w7KVl9+QMvmAIDjLK6/zIecp2k+p8gvULSGVFC
# MSWhD8gDecTZfslXDUiu2EeQ6NR7CAN+yMAa4K/TwS5F95xxiVpBMS0oblvf60AN
# kpLDm4AUpRhDHpb7WevADbTVAB4iFPvj2/bTz/xgYrXXcTpyEqvs0qVC3nxHbiGX
# qUtCbxg9ig/xsitpIBBJuk00PP2SRNLGdcnS0ds6Qscx/to7iEMfqPy3DSabyAc3
# lvKs5pn9FXaQRA5kuTE6H1Eih9P8QTnckiO0NE6yKVmD+mHk0x95Kqq8o4S8TnSy
# juuYkQwSmA==
# SIG # End signature block
