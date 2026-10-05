<#
.SYNOPSIS
    Offline checks for automatic ShareGate analysis scheduling and worker completion.
.VERSION
    1.0.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-GUI.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors | ForEach-Object Message)}
foreach($name in @('Get-DiagnosticReportKey','Start-AutomaticDiagnosticAnalysis','Complete-DiagnosticAnalysis')){
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    . ([scriptblock]::Create($node.Extent.Text))
}
$startNode=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Start-DiagnosticAnalysis'},$true)
function New-TestProcess {
    $process=[pscustomobject]@{HasExited=$false;ExitCode=0;Disposed=$false}
    $process | Add-Member ScriptMethod Dispose { $this.Disposed=$true }
    return $process
}
function Refresh-ActivityList { }
function Write-SmartM365GuiActivityEvent { param($Path,$Status,$ExitCode,$Detail) }
function Update-DiagnosticReportStatus {
    param([string]$Failure='')
    if($Failure){$script:Failures.Add($Failure)}
}
function Refresh-DiagnosticReportState {
    param([switch]$Force,[switch]$SkipAutoAnalysis)
    if(-not $SkipAutoAnalysis){throw 'Worker refresh did not suppress recursive automatic launch.'}
    if($Force -and $script:RestoreSummary){$script:DiagSummary=@{Lines=2}}
}
function Start-DiagnosticAnalysis {
    param([switch]$Automatic)
    if(-not $Automatic){throw 'Scheduler omitted automatic mode.'}
    $script:Launches.Add($script:CurrentMigration.Name)
    $script:DiagProcess=New-TestProcess
    $script:DiagProjectRoot=$script:CurrentMigration.Root
    $script:DiagActiveReportKey=Get-DiagnosticReportKey
}
$parent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
$root=[IO.Path]::GetFullPath((Join-Path $parent ('SmartM365-auto-diagnostics-'+[guid]::NewGuid().ToString('N'))))
if(-not $root.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe test directory.'}
try{
    [void](New-Item -ItemType Directory -Path $root)
    $reportPath=Join-Path $root 'latest.csv'; 'Status,Message' | Set-Content $reportPath
    $script:CurrentMigration=[pscustomobject]@{Name='A';Root=$root}
    $script:DiagLatestReport=$null; $script:DiagReportSignature='input-A'
    $script:DiagSummary=$null; $script:DiagProcess=$null
    $script:DiagAutoAttempts=@{}; $script:DiagActiveReportKey=''; $script:DiagProjectRoot=''
    $script:DiagOutputDirectory=$root; $script:DiagActivity=''; $script:RestoreSummary=$false
    $script:Launches=[Collections.Generic.List[string]]::new()
    $script:Failures=[Collections.Generic.List[string]]::new()
    $lblDiagProgress=[pscustomobject]@{Text=''}
    $btnDiagAnalyze=[pscustomobject]@{IsEnabled=$true}
    $script:DiagTimer=[pscustomobject]@{Started=0;Stopped=0}
    $script:DiagTimer | Add-Member ScriptMethod Start { $this.Started++ }
    $script:DiagTimer | Add-Member ScriptMethod Stop { $this.Stopped++ }
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count){throw 'Missing report launched an analysis.'}
    $script:DiagLatestReport=Get-Item $reportPath
    $script:DiagSummary=@{Lines=2}
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count){throw 'An existing analysis was automatically repeated.'}
    $script:DiagSummary=$null
    Start-AutomaticDiagnosticAnalysis
    $keyA=Get-DiagnosticReportKey
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count -ne 1){throw 'Missing analysis was not launched exactly once.'}
    $script:DiagProcess.HasExited=$true
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count -ne 1){throw 'An unconsumed worker receipt was overwritten.'}
    $script:DiagProcess.ExitCode=7
    Complete-DiagnosticAnalysis
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count -ne 1 -or -not $script:DiagAutoAttempts[$keyA] -or $lblDiagProgress.Text -notmatch 'retry'){
        throw 'Failed automatic analysis entered a retry loop or lost its error.'
    }
    # A changed input has a new identity and is eligible for automatic analysis.
    $script:DiagReportSignature='input-A-updated'
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count -ne 2){throw 'New report signature did not start analysis.'}
    $oldKey=Get-DiagnosticReportKey
    $script:CurrentMigration=[pscustomobject]@{Name='B';Root=(Join-Path $root 'B')}
    [void](New-Item -ItemType Directory -Path $script:CurrentMigration.Root)
    $reportB=Join-Path $script:CurrentMigration.Root 'latest.xlsx'; 'Synthetic input' | Set-Content $reportB
    $script:DiagLatestReport=Get-Item $reportB; $script:DiagReportSignature='input-B'
    $lblDiagProgress.Text='Waiting for B'
    $script:Failures.Clear()
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count -ne 2){throw 'Two automatic analyses ran at the same time.'}
    $script:DiagProcess.HasExited=$true; $script:DiagProcess.ExitCode=9
    Complete-DiagnosticAnalysis
    if($script:Launches.Count -ne 3 -or $script:Launches[2] -ne 'B' -or $script:Failures.Count -ne 0 -or
        -not $script:DiagAutoAttempts[$oldKey]){throw 'Migration change lost the queued analysis or displayed the previous migration error.'}
    $script:RestoreSummary=$true
    $script:DiagProcess.HasExited=$true
    Complete-DiagnosticAnalysis
    Start-AutomaticDiagnosticAnalysis
    if($script:Launches.Count -ne 3 -or -not $script:DiagSummary){throw 'Completed analysis was automatically repeated.'}

    # Exercise real worker argument construction, replacing the external process boundary.
    . ([scriptblock]::Create($startNode.Extent.Text))
    function New-SmartM365GuiActivity { param($ProjectRoot,$Migration,$Action) return (Join-Path $root 'activity.log') }
    function Start-Process {
        param($FilePath,$ArgumentList,$WorkingDirectory,$WindowStyle,[switch]$PassThru,$RedirectStandardOutput,$RedirectStandardError,$ErrorAction)
        $script:WorkerArguments=@($ArgumentList)
        if($WindowStyle -ne 'Hidden'){throw 'Automatic analysis must run in the background.'}
        return New-TestProcess
    }
    $script:ScriptRoot=$root; $script:DiagProcess=$null; $script:DiagSummary=$null
    $cmbDiagSession=[pscustomobject]@{SelectedItem='one-session'}
    Start-DiagnosticAnalysis -Automatic
    if($script:WorkerArguments -contains '-SessionId' -or $script:WorkerArguments -notcontains '-NonInteractive' -or
        $script:DiagActiveReportKey -ne (Get-DiagnosticReportKey)) {throw 'Automatic worker did not cover all sessions or lost its input identity.'}
    $script:DiagProcess=$null
    Start-DiagnosticAnalysis
    if($script:WorkerArguments -notcontains '-SessionId' -or $script:WorkerArguments -notcontains '"one-session"'){
        throw 'Manual session analysis lost the user selection.'
    }
    Write-Host 'PASS: automatic analysis, input changes, one-worker queue, failure suppression, manual retry and worker arguments.'
}finally{
    if(Test-Path -LiteralPath $root){Remove-Item -LiteralPath $root -Recurse -Force}
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAnww2VAyG0ZhRh
# sbTbwjj2M1SkAPSeW/WMgEV7WmL02qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDSVxbx3G4H2LS9V86R0lmK
# IzAzLisDJGf5Gnfm129OiTANBgkqhkiG9w0BAQEFAASCAYAu+m12lDm2QvMuvw/+
# bj+9++ws3Dwc4EJRERpChYOta32quiRTOYBY/85MYxGp6Avr5pJBORhKFOzXRqxW
# /5YzjEjmOomVQTsmcRgTYl9QTaLzzFP+lUgKODMmuzscVNkt8yoMzRwFLHxd6n4O
# dVE2roIlQsAYnm12LQsrXFZQqjCEu0PF/q88Etcy8OIQRHTwPiFj7vygWfGYcPeu
# 61uwnXqbvXPvLkGiU7uBB9X4EMNcaETxwRmoUKZhOYlmSw04rH31Tnbza1ONAWuv
# BWjAze97NxmToYrgHKP4dFCrmPqk5F/mJj3jR6crmwgVItPCWopu7dzXvPUqwu5G
# Vs+RtBH8Sk3ohiJUWeE+Q3CW3SmGHGc0Ct01uCQimWq4i0O5ZHGMXrhBG3FA4DHF
# xImwvVaXTGBJZ2LdDnx2SaDYesWoc1cjMIsAOs2pywO4czTYv4EI4fFUElQp0Tor
# 4xfWyEMWmxJoRNxQ8smGsEEf7TYZOkp742oHm8wD8VufOWI=
# SIG # End signature block
