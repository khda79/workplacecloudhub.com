<#PSScriptInfo
.VERSION 1.0.2
.GUID 7b0b7843-4544-4a5e-9c5d-8f22b499fbf7
.AUTHOR https://github.com/khda79/workplacecloudhub.com
.COMPANYNAME WorkplaceCloudHub
.COPYRIGHT (c) WorkplaceCloudHub. All rights reserved.
.TAGS SmartM365 Orchestrator Restart ScheduledTask
.LICENSEURI https://github.com/khda79/workplacecloudhub.com
.PROJECTURI https://github.com/khda79/workplacecloudhub.com
#>
<#
.SYNOPSIS
Stops the SmartM365 Inventory Orchestrator cleanly and restarts its scheduled task.

.DESCRIPTION
This helper is intended for deployment/restart operations after updating the
repository. It first delegates the clean stop to SmartM365-Inventory-Orchestrator.ps1
with -Stop so the resident instance can save state and release its lock. It then
starts the existing Windows scheduled task so the orchestrator runs under the
registered service account and Task Scheduler settings.

.PARAMETER Tenant
Tenant key used by the orchestrator task. Defaults to prod.

.PARAMETER TaskPath
Windows Task Scheduler folder. Defaults to \WCH\.

.PARAMETER TaskName
Windows Task Scheduler task name. Defaults to "SmartM365 Inventory Orchestrator - <Tenant>".

.PARAMETER StopTimeoutSeconds
Maximum time to wait for the clean stop request.

.PARAMETER RestartDelaySeconds
Delay between the stop phase and task start.

.PARAMETER StartWaitSeconds
Maximum time to wait for the task to enter Running state after Start-ScheduledTask.
#>
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Tenant = 'prod',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TaskPath = '\WCH\',

    [Parameter()]
    [string]$TaskName = '',

    [Parameter()]
    [ValidateRange(1, 3600)]
    [int]$StopTimeoutSeconds = 180,

    [Parameter()]
    [ValidateRange(0, 600)]
    [int]$RestartDelaySeconds = 5,

    [Parameter()]
    [ValidateRange(1, 600)]
    [int]$StartWaitSeconds = 60
)

$ErrorActionPreference = 'Stop'

$smartM365Root = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$tenantContextPath = Join-Path -Path $smartM365Root -ChildPath 'Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) {
    throw "SmartM365 tenant context not found: $tenantContextPath"
}
. $tenantContextPath
$script:RestartStartedAt = Get-Date
$global:SmartM365ScriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$global:SmartM365ExecutionStartTime = $script:RestartStartedAt
$restartLogFolder = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath 'SmartM365\Logs\Orchestrator'
if (-not (Test-Path -LiteralPath $restartLogFolder)) { New-Item -ItemType Directory -Path $restartLogFolder -Force | Out-Null }
$global:LogTextFile = Join-Path -Path $restartLogFolder -ChildPath ('{0}_{1}.log' -f $global:SmartM365ScriptName, $script:RestartStartedAt.ToString('yyyyMMdd_HHmmss'))
$script:RestartCompletionStatus = 'Auto'

Write-SmartM365StartupBanner

function Write-RestartLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line -ForegroundColor $color
    Add-Content -LiteralPath $global:LogTextFile -Value $line -Encoding UTF8
}

function Get-CurrentPowerShellExecutable {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $processPath = (Get-Process -Id $PID -ErrorAction Stop).Path
        if (-not [string]::IsNullOrWhiteSpace($processPath) -and (Test-Path -LiteralPath $processPath -PathType Leaf)) {
            return $processPath
        }
    }

    $programFilesPwsh = Join-Path -Path $env:ProgramFiles -ChildPath 'PowerShell\7\pwsh.exe'
    if (Test-Path -LiteralPath $programFilesPwsh -PathType Leaf) {
        return $programFilesPwsh
    }

    $command = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
        return $command.Source
    }

    throw 'PowerShell 7 (pwsh.exe) was not found.'
}

function Test-ScheduledTaskRunning {
    param(
        [Parameter(Mandatory = $true)]
        $Task
    )

    if ([string]$Task.State -eq 'Running') {
        return $true
    }

    try {
        return ([int]$Task.State -eq 4)
    }
    catch {
        return $false
    }
}

function Get-OrchestratorTask {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        return Get-ScheduledTask -TaskPath $Path -TaskName $Name -ErrorAction Stop
    }
    catch {
        throw ("Scheduled task not found. TaskPath={0}; TaskName={1}; {2}" -f $Path, $Name, $_.Exception.Message)
    }
trap {
    $script:RestartCompletionStatus = 'Failed'
    try { Write-RestartLog -Level ERROR -Message $_.Exception.Message } catch {}
    Write-SmartM365CompletionBanner -Status 'Failed' -ScriptName $global:SmartM365ScriptName -StartedAt $script:RestartStartedAt -ErrorCount 1 -LogPath $global:LogTextFile
    exit 1
}

}

if ([string]::IsNullOrWhiteSpace($TaskName)) {
    $TaskName = "SmartM365 Inventory Orchestrator - $Tenant"
}

if (-not $TaskPath.EndsWith('\')) {
    $TaskPath = "$TaskPath\"
}

$orchestratorScript = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Inventory-Orchestrator.ps1'
if (-not (Test-Path -LiteralPath $orchestratorScript -PathType Leaf)) {
    throw "Orchestrator script not found: $orchestratorScript"
}

$pwsh = Get-CurrentPowerShellExecutable

Write-RestartLog -Message ("Restart requested for tenant {0} on {1}." -f $Tenant, $env:COMPUTERNAME)
Write-RestartLog -Message ("TaskPath={0}; TaskName={1}" -f $TaskPath, $TaskName)
Write-RestartLog -Message ("Stop command: `"{0}`" -NoProfile -ExecutionPolicy Bypass -File `"{1}`" -Tenant {2} -Stop -StopTimeoutSeconds {3}" -f $pwsh, $orchestratorScript, $Tenant, $StopTimeoutSeconds)

& $pwsh -NoProfile -ExecutionPolicy Bypass -File $orchestratorScript -Tenant $Tenant -Stop -StopTimeoutSeconds $StopTimeoutSeconds
$stopExitCode = $LASTEXITCODE
if ($stopExitCode -ne 0) {
    throw "Clean stop failed with exit code $stopExitCode. The scheduled task was not restarted."
}

if ($RestartDelaySeconds -gt 0) {
    Write-RestartLog -Message ("Waiting {0} second(s) before restart." -f $RestartDelaySeconds)
    Start-Sleep -Seconds $RestartDelaySeconds
}

$task = Get-OrchestratorTask -Path $TaskPath -Name $TaskName
if (Test-ScheduledTaskRunning -Task $task) {
    throw ("Scheduled task is still Running after clean stop. TaskPath={0}; TaskName={1}; State={2}. Restart aborted." -f $TaskPath, $TaskName, $task.State)
}

Write-RestartLog -Message ("Starting scheduled task: {0}{1}" -f $TaskPath, $TaskName)
Start-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop

$deadline = (Get-Date).AddSeconds($StartWaitSeconds)
do {
    Start-Sleep -Seconds 2
    $task = Get-OrchestratorTask -Path $TaskPath -Name $TaskName
    if (Test-ScheduledTaskRunning -Task $task) {
        Write-RestartLog -Level SUCCESS -Message ("Scheduled task restarted and is Running. TaskPath={0}; TaskName={1}" -f $TaskPath, $TaskName)
        $script:RestartCompletionStatus = 'Auto'
        Write-SmartM365CompletionBanner -Status $script:RestartCompletionStatus -ScriptName $global:SmartM365ScriptName -StartedAt $script:RestartStartedAt -LogPath $global:LogTextFile
        exit 0
    }
} while ((Get-Date) -lt $deadline)

$taskInfo = $null
try {
    $taskInfo = Get-ScheduledTaskInfo -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
}
catch {
    Write-RestartLog -Level WARN -Message ("Unable to read scheduled task info: {0}" -f $_.Exception.Message)
}

if ($null -ne $taskInfo) {
    throw ("Scheduled task did not enter Running state within {0} second(s). State={1}; LastTaskResult={2}; LastRunTime={3}" -f $StartWaitSeconds, $task.State, $taskInfo.LastTaskResult, $taskInfo.LastRunTime)
}

throw ("Scheduled task did not enter Running state within {0} second(s). State={1}" -f $StartWaitSeconds, $task.State)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCHCviZR1aYUfWH
# uZIyKXguqMG+1qQUixHpyEWZnVztTaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIGVgjILy6v6ZbTJlafKXFrSb2FLNnZ6227xzKfpccWwjMA0GCSqG
# SIb3DQEBAQUABIIBgEVqplvPCOk9xzl3Vg71/v9wWb7o39Hw0ksXy1nFLswzQmTC
# zGTt0sxFoGMJiru6KnPE0nUSspDVZQc+DeGWnuGpQtapwFVyXwrkfzwutc7PeJuj
# 7+TkvZHEMOqGK3bsHki/WbMq4MoAxfu9r1PZmAzja9asywejKgpG+d1O49Gby9er
# DKaxaf9yqG+/8JYi2tKVJPQ+wCmLVJX3WgGDbXM+Wt6mLKulqV881ryfNr0/NIw8
# GdnrlZZrLemNX2KRMknEL0MKT+hZL0ZhG5nGILSuujG+NmU7OctTo2NIr8Z4Vrw6
# bgzPCLDyno9KHFaEmVllPoZvAmSNt2KUiECNGPc6GnkWkmfAIVJlqqL0bU7dcuLC
# byDn54dDc7FUt7FzPvZ2RPHGsQcX9VVwbfsYi8lcohb2KZ0FDimD5Y151E88Yf8O
# yxebrsSdu5d+UlCfgALbu6S/C2EJApKprJ6swJISG3eBeglG2Fqg9APCz5BTS4hk
# V4WweytObJjV6zV556GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQwMTEy
# NThaMC8GCSqGSIb3DQEJBDEiBCDbrZLBaFQGxsti9/IN1B7CbMZ9jtzOnwP/1GRu
# DRIxNzANBgkqhkiG9w0BAQEFAASCAgCkrp7ueRbZZ/f05Pe+VplRxSUu4hKgPSpi
# 0oVisfIjZrcx6cwM71H3dPZlTJdyuwxmAyRbSMUVblRNMyUujVpu6W/fdCdUHV2p
# x1RnIwhfwH06DpLX3AF2zT761806/Nubmj+P7kd7txh7aQgbbQq+Tu7op9gH6yQk
# PAO2qqcYu7xTWcPvMaCvGX+2nsNA8zl6rSqIJflQnyf7PhgVjSTSyvEZkAIqfX6d
# SvTSkQIa8OG/F3Kftb0LXGzdSMp3b3w52ikSp338sAFXMhzJkL4MskynAMJ22hLu
# nBu9F69qTKxOVox0J6cFyRSwonx0hKaNlJscuxNU5ZLb2lv6TRHKJI9I+e4ZzyDq
# /kOneXxxcQTfPiF6nb/sOODAi32Di6jaE4j96312mqHU/HjHuJq4DZbkbiqUKAa8
# CGKG/Hn6KE1yq5M+uqkGR3Bc6PjxaiVMEENlNInEg+QQUBL24YnEWvcrvx+n7tPk
# Dfjt/3qfDPXzprz1HzXid5e0tda+WD1fyjCZaIwzdyw3875O6maZ/eZWnbmcUwSe
# jV7DiRrYFbmy3ZPeLfmlLqgBob4pIiLoWFyaSrmkuV8cGejV5S/CJQlzE5hhk2Vr
# InHdG0YQtHjINKUxyjoYp3ov5HhhNSRJ5kWoUmbd2i44LaIDYOpHVKuDW6Bo/06B
# 8lqTaKt9tw==
# SIG # End signature block
