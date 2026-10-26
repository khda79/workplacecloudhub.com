#Requires -Version 7.0
<#
.SYNOPSIS
Launches synthetic Windows PowerShell 5.1 and PowerShell 7 jobs through the real orchestrator job launcher
with the PowerShell 7 module path first, and checks that the 5.1 job still loads its own modules.

.DESCRIPTION
No collector, tenant or network is used. The orchestrator's Start-InventoryJob, Get-JobEngine and
Get-OrchestratorWindowsPowerShellModulePathCommand run unchanged from pwsh, with $env:PSModulePath starting with
the PowerShell 7 module folders (as in the resident orchestrator). Checks:
- a WindowsPowerShell job loads Microsoft.PowerShell.Security from Windows PowerShell and Get-Acl works;
- without the prefix the same job fails (the production failure is reproduced);
- the PSModulePath keeps the other entries (Documents\WindowsPowerShell\Modules, Program Files\WindowsPowerShell\Modules);
- a PowerShell 7 job is unchanged: no prefix, it still loads the PowerShell 7 module;
- the orchestrator process environment is not changed.

.VERSION
1.0.1
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Orchestrator parse failed.' }
$definitions = foreach ($name in @('Start-InventoryJob', 'Get-JobEngine', 'Get-PwshPath', 'Get-OrchestratorWindowsPowerShellModulePathCommand')) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $node) { throw "Missing orchestrator function: $name" }
    $node.Extent.Text
}

$script:Checks = 0
function Assert-ModulePath {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ModulePath-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root
$collectorPath = Join-Path $root 'SyntheticCollector.ps1'
[IO.File]::WriteAllText($collectorPath, @'
param([string]$Tenant)
try {
    $null = Get-Acl -LiteralPath $PSScriptRoot
    'ACL OK edition=' + $PSVersionTable.PSEdition + ' security=' + (Get-Module Microsoft.PowerShell.Security).ModuleBase
    'PATH=' + $env:PSModulePath
    exit 0
}
catch {
    'ACL FAILED edition=' + $PSVersionTable.PSEdition + ' ' + $_.Exception.Message
    exit 7
}
'@, [Text.UTF8Encoding]::new($false))

$windowsPowerShellModules = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\Modules'
$pwshModules = Join-Path $env:ProgramFiles 'PowerShell\7\Modules'
$savedModulePath = $env:PSModulePath
$pollutedModulePath = (@($pwshModules, (Join-Path $env:USERPROFILE 'Documents\PowerShell\Modules'), (Join-Path $env:ProgramFiles 'PowerShell\Modules'), (Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules'), (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'), $windowsPowerShellModules) -join ';')

function Invoke-SyntheticJob {
    param([string]$Name, [string]$Edition, [switch]$WithoutPrefix)
    $folder = Join-Path $root $Name
    $null = New-Item -ItemType Directory -Path $folder
    $module = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
    & $module {
        param($folder, $jobName, $withoutPrefix)
        $script:Settings = @{ JobLogFolderPath = $folder; SharedDataFolderPath = $folder }
        $script:RunningJobs = @{}
        $script:State = @{ Jobs = @{ $jobName = @{ Running = $null } } }
        $script:JobState = $script:State.Jobs[$jobName]
        $script:ScriptName = 'SyntheticOrchestrator'
        $script:Tenant = 'test'
        $script:Connect = $false
        function script:Get-JobState { param($JobName) $script:JobState }
        function script:Get-OrchestratorJobLaunchDefinition { param($Job) $null }
        function script:Resolve-OrchestratorJobPath { param($Path) $Path }
        function script:Invoke-OrchestratorAuthenticodeValidation { param($Job, $ScriptFullPath) [pscustomobject]@{ Allowed = $true } }
        function script:ConvertTo-StateTime { param($Value) ([datetime]$Value).ToString('o') }
        function script:Save-OrchestratorState {}
        function script:Write-OrchestratorLog { param($Message, $Level) }
        function script:Complete-JobRun { throw 'Synthetic launch failed before a child was started.' }
        if ($withoutPrefix) { function script:Get-OrchestratorWindowsPowerShellModulePathCommand { '' } }
    } $folder $Name $WithoutPrefix.IsPresent
    $job = [pscustomobject]@{ Name = $Name; ScriptPath = $collectorPath; Arguments = ''; PowerShellEdition = $Edition; TimeoutMinutes = 2 }
    & $module { param($job) Start-InventoryJob -Job $job -Occurrence (Get-Date) } $job
    $process = & $module { param($n) $script:RunningJobs[$n].Process } $Name
    if (-not $process.WaitForExit(60000)) { throw "Synthetic job $Name did not exit." }
    $record = & $module { param($n) $script:State.Jobs[$n].Running } $Name
    $result = Get-Content -LiteralPath $record.ResultPath -Raw | ConvertFrom-Json
    $log = Get-ChildItem -LiteralPath $folder -Filter '*.log' -File -Recurse | Select-Object -First 1
    return [pscustomobject]@{ ExitCode = $result.ExitCode; Log = [IO.File]::ReadAllText($log.FullName) }
}

try {
    $env:PSModulePath = $pollutedModulePath
    $windowsJob = Invoke-SyntheticJob -Name 'WindowsPowerShellJob' -Edition 'WindowsPowerShell'
    $controlJob = Invoke-SyntheticJob -Name 'WindowsPowerShellControl' -Edition 'WindowsPowerShell' -WithoutPrefix
    $coreJob = Invoke-SyntheticJob -Name 'PowerShell7Job' -Edition 'PowerShell7'
    Assert-ModulePath ($env:PSModulePath -ceq $pollutedModulePath) 'The orchestrator process keeps its own PSModulePath.'

    Assert-ModulePath ($windowsJob.ExitCode -eq 0 -and $windowsJob.Log -match ('ACL OK edition=Desktop security=' + [regex]::Escape((Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0')))) "A Windows PowerShell job loads its own security module and Get-Acl works: $($windowsJob.Log)"
    $windowsPath = ([regex]::Match($windowsJob.Log, 'PATH=(?<Path>[^\r\n]+)')).Groups['Path'].Value -split ';'
    Assert-ModulePath ($windowsPath[0] -ieq $windowsPowerShellModules) 'Windows PowerShell modules come first.'
    Assert-ModulePath (@($windowsPath | Where-Object { $_ -match '(?i)(^|\\)PowerShell\\(7[^\\]*\\)?Modules\\?$' }).Count -eq 0) "PowerShell 7 module folders are removed: $($windowsPath -join ';')"
    Assert-ModulePath (($windowsPath -contains (Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules')) -and ($windowsPath -contains (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'))) 'Windows PowerShell user and machine module folders are kept.'
    Assert-ModulePath ($windowsJob.Log -match 'Child command: \$smartM365ModulePath = ') 'The repair is visible at the start of the logged child command.'

    if (Test-Path -LiteralPath (Join-Path $pwshModules 'Microsoft.PowerShell.Security')) {
        Assert-ModulePath ($controlJob.ExitCode -eq 7 -and $controlJob.Log -match 'ACL FAILED edition=Desktop') "Without the repair the same job fails, as in production: $($controlJob.Log)"
    }

    Assert-ModulePath ($coreJob.ExitCode -eq 0 -and $coreJob.Log -match 'ACL OK edition=Core security=[^\r\n]*PowerShell\\7') "A PowerShell 7 job still loads its PowerShell 7 module: $($coreJob.Log)"
    Assert-ModulePath ($coreJob.Log -notmatch 'smartM365ModulePath') 'A PowerShell 7 job gets no module path prefix.'
}
finally {
    $env:PSModulePath = $savedModulePath
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
}

"PASS: $script:Checks orchestrator Windows PowerShell module path checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCByMT4uxjNd/wZg
# nxx8WYv43CdKJK8FcSfAhHc6+/JB/6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGBG3Sc68zvcUUeTe3zcI6aETOYs7jjnaCIoomwArbHnMA0GCSqG
# SIb3DQEBAQUABIIBgHkeoCRYP+XgT5tIgaUm5EWFKBhXQaCyM0DV5ESSDIdk0ptx
# ZOVsWwVzeGo4VlhFsGecteWlt3wGk6rCcMqm6iS3brOVrpgcAnCHJ6N8Zuaqbs2k
# x583qd+hvfbgmJEivaIp3oEgFVyM9she+m/DMrE5gUSPkIoggf0id4jVWaWqJiPl
# A2yYAYKjxd4gaObqdIwqbIlXDEGbBg89ciaDAAUrxU7X4Gw8T4PYR6RikTn4RM1X
# XeZ1KL+bhstFVDTq0VRkeoY3UFybsrERA7h6kx5St4BqCMw+Ts0+3QKYEGGVx/3z
# zUSvM9fry1s4pSTuH8O5u5IijhrgbDvHUB20qIcGCpb1hmhFwMRDrYx2Jo0Baf/n
# 7fRPGWtmZS71/Y9RAfFGW5t83BBpKa5oOW0aVxcZfy4/uZ4JN2rauPZZ4yEeo216
# ov6l1/a5Wxl/puRcd9Kesyn9PKKPgvQ5ZzaKAzwTzrWgiZNW4M2z92DhWNt5rBsO
# VVck0RGYAQAcyrip+aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYwOTU2
# MTRaMC8GCSqGSIb3DQEJBDEiBCBlp+zg272ZWg0MHHdKh671uRxRHD4SZ3sjs4if
# sCrZuTANBgkqhkiG9w0BAQEFAASCAgB/V54TicfoHLuw0BXQEP5R2xXGn9RyKfX2
# aKMnfzyzKhoiuaVMztUncNpvkRxTgx3StZkt4ZE9ICtV54PrfsUNtZjZBrLdONMs
# waYeihxZhUPeUK2wkJQZYw9zIf8Tk9UQqhspnRtUWKmcEX19sX5v72my0aQxzyww
# POe+NuHaO7U68eFf3SdkYDvVOIG+awc7YOG227gQ+SfXH6rAdrw5ic8j2LgrpA/f
# dF0QJ8U2N16ZzPtFI9+wz/62euHPMkMoMkbA7WYZAJqgpEIWnj8Ox2prY4LRTs80
# 8dXP7g/jU9gDSAMMgMlUloREPAOjpXm+LFbVArfEynpcGnqe69cVc+MnaTzAnYiJ
# UK0+aKvj1zdiHai5tQrjuV8fqKxB694uFHyBjrIaKH9uvTjB7fWvyhSnURWKVFUZ
# x6xF524sa7SgEovIAhLH/KMShQlvrrHZPMf5SMttWwvfQ1WjTUERZ95jJX5RZ4Bi
# KtU00xtfnr2mrtgWimqocShxyCqyujLQg86dFUSO22tLlg8Ei0wedrACr1kweCea
# yP/g5ZulNshV+iO76//GKv1cb9FD4qCxPUECietJ5BJ8z6cdHI3epjRJ/24p/4Vo
# lyD14qV4liP7IgVZcpTXtuNOWdR9Ah+YeVEMeemtnQBgxXxzowXtKokdaTgR4Zt8
# XkgTVEWFLg==
# SIG # End signature block
