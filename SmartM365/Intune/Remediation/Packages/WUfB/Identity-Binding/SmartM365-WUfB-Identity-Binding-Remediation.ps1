<#
    Name: SmartM365-WUfB-Identity-Binding-Remediation.ps1
    Version: 1.1
    Description: Repairs Windows Update for Business identity binding by resetting computed PolicyState and refreshing MDM/WU policy state safely.

    Intended use:
    - Microsoft Intune remediation script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell
    - No PRT refresh
    - No forced reboot

    Logs:
    - C:\ProgramData\SmartM365\IntuneRemediation\Logs\Remediate-WUfB-Identity-Binding\

    Exit codes:
    0 = Remediation script completed
    1 = Technical script error
#>

[CmdletBinding()]
param(
    [switch]$ExportPolicyStateReg
)

$ErrorActionPreference = "Stop"

# Relaunch in 64-bit PowerShell if Intune starts the script in 32-bit PowerShell
if ($env:PROCESSOR_ARCHITEW6432) {
    & "$env:WINDIR\Sysnative\WindowsPowerShell\v1.0\powershell.exe" `
        -ExecutionPolicy Bypass `
        -NoProfile `
        -File "$PSCommandPath" `
        -ExportPolicyStateReg:$ExportPolicyStateReg

    exit $LASTEXITCODE
}

$RemediationName = "Remediate-WUfB-Identity-Binding"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$RemediationName"

if (-not (Test-Path -Path $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

$LogPath = Join-Path -Path $LogRoot -ChildPath "$RemediationName.log"

$PolicyStatePath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState"
$MdmUpdatePolicyPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"
$UsoClientPath = Join-Path -Path $env:SystemRoot -ChildPath "System32\UsoClient.exe"

$ErrorFound = $false
$RemediationErrors = New-Object System.Collections.Generic.List[string]

function Format-CompactText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [int]$MaxLength = 180
    )

    $compactText = ($Text -replace "\s+", " ").Trim()

    if ($compactText.Length -gt $MaxLength) {
        return ($compactText.Substring(0, $MaxLength) + "...")
    }

    return $compactText
}

function Write-IntuneResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Status,

        [hashtable]$Data = @{}
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("Status=$Status")

    foreach ($key in ($Data.Keys | Sort-Object)) {
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 240
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not (Test-Path -Path $LogRoot)) {
        New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    }

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Add-RemediationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log "ERROR: $Message"
    $script:ErrorFound = $true
    $script:RemediationErrors.Add($Message)
}

function Invoke-ServiceRestartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=${Name}"
            return
        }

        if ($service.Status -ne "Stopped") {
            Write-SmartM365Log "ServiceStopRequested=${Name}"
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        }
        else {
            Write-SmartM365Log "ServiceAlreadyStopped=${Name}"
        }

        Start-Sleep -Seconds 2

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStartupTypeChanged=${Name} StartupType=Manual"
        }

        Write-SmartM365Log "ServiceStartRequested=${Name}"
        Start-Service -Name $Name -ErrorAction SilentlyContinue
    }
    catch {
        Add-RemediationError "Failed to restart service ${Name}: $($_.Exception.Message)"
    }
}

function Invoke-EnterpriseMgmtPush {
    try {
        $enterpriseMgmtTasks = @(
            Get-ScheduledTask -TaskPath "\Microsoft\Windows\EnterpriseMgmt\" -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskName -eq "PushLaunch" }
        )

        if ($null -eq $enterpriseMgmtTasks -or $enterpriseMgmtTasks.Count -eq 0) {
            Write-SmartM365Log "EnterpriseMgmtPushLaunch=NotFound"
            return
        }

        foreach ($task in $enterpriseMgmtTasks) {
            try {
                Start-ScheduledTask -InputObject $task -ErrorAction Stop
                Write-SmartM365Log "EnterpriseMgmtPushLaunch=Triggered TaskPath=$($task.TaskPath) TaskName=$($task.TaskName)"
            }
            catch {
                Write-SmartM365Log "EnterpriseMgmtPushLaunch=Failed TaskName=$($task.TaskName) Message=$($_.Exception.Message)"
            }
        }
    }
    catch {
        Add-RemediationError "Failed to enumerate EnterpriseMgmt tasks: $($_.Exception.Message)"
    }
}

function Get-RegistryPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -Path $Path)) {
        return $null
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue

    if ($null -eq $item) {
        return $null
    }

    if ($item.PSObject.Properties.Name -contains $Name) {
        return $item.$Name
    }

    return $null
}

function Get-IsWUfBConfigured {
    $value = Get-RegistryPropertyValue -Path $PolicyStatePath -Name "IsWUfBConfigured"

    if ($null -eq $value) {
        return $null
    }

    try {
        return [int]$value
    }
    catch {
        return $null
    }
}

function Export-PolicyState {
    try {
        if (-not (Test-Path -Path $PolicyStatePath)) {
            Write-SmartM365Log "PolicyStateExport=Skipped Reason=PolicyStateNotFound"
            return
        }

        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $exportPath = Join-Path -Path $LogRoot -ChildPath "WUfB-PolicyState-$timestamp.reg"
        $regExe = Join-Path -Path $env:SystemRoot -ChildPath "System32\reg.exe"

        & $regExe export "HKLM\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState" "$exportPath" /y | Out-Null

        if (Test-Path -LiteralPath $exportPath -PathType Leaf) {
            Write-SmartM365Log "PolicyStateExport=Completed Path=$exportPath"
        }
        else {
            Write-SmartM365Log "PolicyStateExport=Failed Path=$exportPath"
        }
    }
    catch {
        Write-SmartM365Log "PolicyStateExport=Failed Message=$($_.Exception.Message)"
    }
}

function Invoke-UsoClientSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    try {
        if (-not (Test-Path -LiteralPath $UsoClientPath -PathType Leaf)) {
            Write-SmartM365Log "UsoClient=${Action} Status=UsoClientNotFound"
            return
        }

        Start-Process -FilePath $UsoClientPath -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-SmartM365Log "UsoClient=${Action} Status=Triggered"
    }
    catch {
        Write-SmartM365Log "UsoClient=${Action} Status=Failed Message=$($_.Exception.Message)"
    }
}

try {
    Write-SmartM365Log "===== WUfB identity binding remediation started ====="

    $mdmPresentBefore = Test-Path -Path $MdmUpdatePolicyPath
    $wuConfiguredBefore = Get-IsWUfBConfigured

    Write-SmartM365Log "Before_MDMWUfBPolicyPresent=$mdmPresentBefore"
    Write-SmartM365Log "Before_IsWUfBConfigured=$wuConfiguredBefore"

    # Stop/restart Windows Update services around PolicyState reset
    foreach ($serviceName in @("UsoSvc", "wuauserv")) {
        Invoke-ServiceRestartSafe -Name $serviceName
    }

    if ($ExportPolicyStateReg) {
        Export-PolicyState
    }

    if (Test-Path -Path $PolicyStatePath) {
        try {
            Write-SmartM365Log "PolicyStateReset=Start Path=$PolicyStatePath"
            Remove-Item -Path $PolicyStatePath -Recurse -Force -ErrorAction Stop
            Write-SmartM365Log "PolicyStateReset=Completed"
        }
        catch {
            Add-RemediationError "Failed to remove PolicyState: $($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "PolicyStateReset=Skipped Reason=PolicyStateNotFound"
    }

    # Restart related services again after reset
    foreach ($serviceName in @("wuauserv", "UsoSvc")) {
        Invoke-ServiceRestartSafe -Name $serviceName
    }

    # Trigger MDM and Windows Update policy refresh
    Invoke-EnterpriseMgmtPush

    Start-Sleep -Seconds 20

    Invoke-UsoClientSafe -Action "RefreshSettings"
    Invoke-UsoClientSafe -Action "StartScan"

    Start-Sleep -Seconds 15

    $mdmPresentAfter = Test-Path -Path $MdmUpdatePolicyPath
    $wuConfiguredAfter = Get-IsWUfBConfigured

    Write-SmartM365Log "After_MDMWUfBPolicyPresent=$mdmPresentAfter"
    Write-SmartM365Log "After_IsWUfBConfigured=$wuConfiguredAfter"

    if ($mdmPresentAfter -and $wuConfiguredAfter -eq 0) {
        Write-SmartM365Log "Status=CompletedButDriftStillPresent"
        Write-SmartM365Log "Result=MDM WUfB policy is present but IsWUfBConfigured is still 0"
        Write-IntuneResult -Status "CompletedButDriftStillPresent" -Data @{
            BeforeIsWUfBConfigured = $wuConfiguredBefore
            AfterIsWUfBConfigured = $wuConfiguredAfter
            MDMWUfBPolicyPresent = $mdmPresentAfter
            LogPath = $LogPath
        }
        exit 0
    }

    if ($mdmPresentAfter -and $wuConfiguredAfter -eq 1) {
        Write-SmartM365Log "Status=CompletedHealthy"
        Write-SmartM365Log "Result=WUfB identity binding is healthy after remediation"
        Write-IntuneResult -Status "CompletedHealthy" -Data @{
            BeforeIsWUfBConfigured = $wuConfiguredBefore
            AfterIsWUfBConfigured = $wuConfiguredAfter
            MDMWUfBPolicyPresent = $mdmPresentAfter
            LogPath = $LogPath
        }
        exit 0
    }

    if (-not $mdmPresentAfter) {
        Write-SmartM365Log "Status=CompletedNotApplicable"
        Write-SmartM365Log "Result=No MDM WUfB policy detected after remediation"
        Write-IntuneResult -Status "CompletedNotApplicable" -Data @{
            BeforeIsWUfBConfigured = $wuConfiguredBefore
            AfterIsWUfBConfigured = $wuConfiguredAfter
            MDMWUfBPolicyPresent = $mdmPresentAfter
            LogPath = $LogPath
        }
        exit 0
    }

    Write-SmartM365Log "Status=CompletedInconclusive"
    Write-SmartM365Log "Result=Sanity check inconclusive"
    Write-IntuneResult -Status "CompletedInconclusive" -Data @{
        BeforeIsWUfBConfigured = $wuConfiguredBefore
        AfterIsWUfBConfigured = if ($null -eq $wuConfiguredAfter) { "Unknown" } else { $wuConfiguredAfter }
        MDMWUfBPolicyPresent = $mdmPresentAfter
        LogPath = $LogPath
    }
    exit 0
}
catch {
    try {
        Add-RemediationError $_.Exception.Message
    }
    catch {
        Write-IntuneResult -Status "ErrorDuringErrorHandling" -Data @{
            Message = $_.Exception.Message
        }
        exit 1
    }

    $sampleErrors = @($RemediationErrors | Select-Object -First 3)
    Write-IntuneResult -Status "Error" -Data @{
        ErrorCount = $RemediationErrors.Count
        LogPath = $LogPath
        Message = $_.Exception.Message
        Samples = ($sampleErrors -join " | ")
    }
    exit 1
}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD+JaJCglu0nGGH
# feOxyGNmDTSDh7pbfGQqF7+u78/t6KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJ/Rtkl2RA2BqrtAAL8z6x4nJ1Ifb17gsWtVV9ZvVT8nMA0GCSqG
# SIb3DQEBAQUABIIBgHSG3cvxYd1YHfbRgA6Bj1J8uXVvaqrgP320ySOdgbRGXPmW
# sG+CgtO8mk6DlLKTilD7DpRyagA8ANm8HB1WUux6wU3oAPCk2ArAYS2h5xydXB2W
# LcbrW3AcAtb9LptRMNSZ6GAn/AwvXvFUV/3CD/TDJ6xzSjM+k/J3Q1YPINzL77r7
# g8UP/yJFI5ZShXxlwrSbaa6ySNg8in1QXHeQDME8yYfyAI3FQGYo3umMnAAKbLPw
# saakgEj1SmoA3Ku5Vxb5exTc+UVQHy7Ox7Vmhjec02c0In2rsYNnJkI0IdhilJe7
# k4Wc/E9K1h2B41yoYDD8/Cj34cUn+yn8JEKiTdvMnZNuz4t/rtwlb1GE8ZPEyxps
# vBqtjsXf7lpEYJIXU6QFTCXIEx3UunzVe9T8B1iwiT+wPac3fVoXrPD6VBm6a4nP
# 0jcAbOb4PyfytaPJK4WeIrq0Sj0QlpEOgnWVyVZBKcr2yIzsI97fILB+W8iqcdh/
# T3JTsbDajYvF1/j+VKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MDZaMC8GCSqGSIb3DQEJBDEiBCA9X5Y9KosuNCHzYgXiE2OU6aq94aMPjQxVS7sk
# uw6HTzANBgkqhkiG9w0BAQEFAASCAgDHbvz5u8XKBUNo0bBnIqz7SQcM9QTkjQSp
# UDBPmCDqcLmc6z0NgcduVSCl5f+/lqOz2eZ9yg2Iy7ad0t2Q9u52A5bnAufBOQhA
# ywV/bRtfDi8cg4+hMMhUCAa7RfVQy17twzsTUhl8qrHEGr7UrWx+91YteQDOYXi8
# ioCVQo1awzslE4oI4nxgxkn8EainttfYbVi5MAzZxz8RJQriaUheXYaJrLbDQmLa
# bLRYNPuRcxHXK1DPU2WHjf9yqv8sXEVR9GxsT/oMurds1X1c1AJOptJ0nF0GOvoK
# maiqpAERhDACki0yxdsNY9iIM9l8olj93GeqbJoAVfDXYrQLezi99azVTrhREMj2
# A3rL1LFsnFQJusSXk+9U4aGF7K+GgkIxOEo0XsE4s2KIzROjqu+4pMhmX6BQ9Ee4
# B3Gh6ga+757Kt3epmt+VaRorRFkepkz6Z2LFCQpR14vzgDQlKCxeMjge1xtrRAhJ
# +V8/atr6WJxZEGZqF/PekugZ9rnk62RYEV8tDBh/2ZzCxnBnHqdhcf+CpOfz0B/u
# 3IAGZxJQv9aoejfGbVo8u0wGnZiGj8j1R3mhxJKwFig/1qBQfgLSYnGJcxnC3gSc
# mun8ZuKCevGIPFe5rByKIFE2K4fPT8BimI+DvKoMqPIdBofAEAGRvhMcowbbiShb
# Rf5kfjT/yw==
# SIG # End signature block
