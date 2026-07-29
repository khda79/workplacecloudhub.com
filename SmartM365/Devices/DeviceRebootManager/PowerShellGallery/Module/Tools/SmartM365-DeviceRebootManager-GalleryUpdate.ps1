#Requires -Version 5.1

<#
.SYNOPSIS
Updates and deploys SmartM365 Device Reboot Manager from PowerShell Gallery.

.DESCRIPTION
This helper is packaged with the PowerShell Gallery module and copied to the
stable ProgramData installation. It updates the module, validates the pinned
WorkplaceCloudHub signer, deploys the runtime, preserves local configuration,
and optionally registers the automatic update task.
#>

[CmdletBinding()]
param(
    [string]$ModuleName = 'SmartM365.DeviceRebootManager',
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$ConfigSourcePath = '',
    [switch]$ForceConfig,
    [switch]$SkipScheduledTask,
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [ValidateRange(15, 10080)]
    [int]$RepeatIntervalMinutes = 240,
    [bool]$EnableAutomaticUpdate = $false,
    [string]$UpdateTaskPath = '\SmartM365\',
    [string]$UpdateTaskName = 'Device Reboot Manager Update',
    [ValidateRange(1, 168)]
    [int]$UpdateIntervalHours = 24,
    [bool]$IncludePrerelease = $false,
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [switch]$SkipPackageUpdate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-UpdateLog {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')

    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Output $line

    try {
        $logRoot = Join-Path -Path $InstallPath -ChildPath 'Logs'
        if (-not (Test-Path -LiteralPath $logRoot -PathType Container)) {
            New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
        }
        Add-Content -LiteralPath (Join-Path $logRoot 'PowerShellGalleryUpdate.log') -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning ("Unable to write update log. {0}" -f $_.Exception.Message)
    }
}

function Get-NormalizedThumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Get-LatestModuleInfo {
    param([Parameter(Mandatory = $true)][string]$Name)

    $module = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if (-not $module) {
        throw "Installed module not found: $Name"
    }
    return $module
}

function Invoke-GalleryPackageUpdate {
    param([Parameter(Mandatory = $true)][string]$Name)

    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $powerShellGetInstall = $null
    if (Get-Command Get-InstalledModule -ErrorAction SilentlyContinue) {
        $powerShellGetInstall = Get-InstalledModule -Name $Name -ErrorAction SilentlyContinue
    }

    if ($powerShellGetInstall -and (Get-Command Update-Module -ErrorAction SilentlyContinue)) {
        Write-UpdateLog "Updating $Name with PowerShellGet."
        $updateModuleParameters = @{
            Name        = $Name
            Force       = $true
            ErrorAction = 'Stop'
        }
        if ($IncludePrerelease) {
            $updateModuleParameters.AllowPrerelease = $true
        }
        Update-Module @updateModuleParameters
        return
    }

    if (Get-Command Update-PSResource -ErrorAction SilentlyContinue) {
        Write-UpdateLog "Updating $Name with PSResourceGet."
        $updateResourceParameters = @{
            Name            = $Name
            Repository      = 'PSGallery'
            Scope           = 'AllUsers'
            TrustRepository = $true
            Quiet           = $true
            ErrorAction     = 'Stop'
        }
        if ($IncludePrerelease) {
            $updateResourceParameters.Prerelease = $true
        }
        Update-PSResource @updateResourceParameters
        return
    }

    throw 'No supported PowerShell Gallery update command is available.'
}

function Assert-ModuleSignature {
    param(
        [Parameter(Mandatory = $true)][string]$ModuleBase,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $expected = Get-NormalizedThumbprint $Thumbprint
    $powerShellFiles = @(Get-ChildItem -LiteralPath $ModuleBase -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
    if ($powerShellFiles.Count -eq 0) {
        throw "No PowerShell files found in package: $ModuleBase"
    }

    foreach ($file in $powerShellFiles) {
        $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName
        $actual = if ($signature.SignerCertificate) {
            Get-NormalizedThumbprint $signature.SignerCertificate.Thumbprint
        }
        else {
            ''
        }

        if ($signature.Status -ne 'Valid' -or $actual -ne $expected) {
            throw ("Package signature validation failed: file={0}; status={1}; signer={2}" -f $file.FullName,$signature.Status,$actual)
        }
    }
}

function Register-GalleryUpdateTask {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$ScheduledTaskPath,
        [Parameter(Mandatory = $true)][string]$ScheduledTaskName,
        [Parameter(Mandatory = $true)][int]$IntervalHours
    )

    $powerShellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) {
        $powerShellPath = 'powershell.exe'
    }

    $arguments = @(
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        ('"{0}"' -f $ScriptPath)
        '-ModuleName'
        ('"{0}"' -f $ModuleName)
        '-InstallPath'
        ('"{0}"' -f $InstallPath)
        '-TaskPath'
        ('"{0}"' -f $TaskPath)
        '-TaskName'
        ('"{0}"' -f $TaskName)
        '-RepeatIntervalMinutes'
        $RepeatIntervalMinutes
        '-EnableAutomaticUpdate:$true'
        '-UpdateTaskPath'
        ('"{0}"' -f $ScheduledTaskPath)
        '-UpdateTaskName'
        ('"{0}"' -f $ScheduledTaskName)
        '-UpdateIntervalHours'
        $IntervalHours
        ('-IncludePrerelease:${0}' -f $IncludePrerelease.ToString().ToLowerInvariant())
        '-ExpectedSignerThumbprint'
        ('"{0}"' -f $ExpectedSignerThumbprint)
    )

    $action = New-ScheduledTaskAction -Execute $powerShellPath -Argument ($arguments -join ' ')
    $now = Get-Date
    $firstRun = $now.Date.AddHours(3)
    if ($firstRun -le $now) {
        $firstRun = $firstRun.AddDays(1)
    }
    $trigger = New-ScheduledTaskTrigger `
        -Once `
        -At $firstRun `
        -RepetitionInterval (New-TimeSpan -Hours $IntervalHours) `
        -RepetitionDuration (New-TimeSpan -Days 3650)
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1)

    Register-ScheduledTask `
        -TaskPath $ScheduledTaskPath `
        -TaskName $ScheduledTaskName `
        -Action $action `
        -Trigger $trigger `
        -Principal $principal `
        -Settings $settings `
        -Description 'Updates SmartM365 Device Reboot Manager from PowerShell Gallery.' `
        -Force | Out-Null
}

if (-not (Test-IsAdministrator)) {
    throw 'Device Reboot Manager Gallery deployment must run elevated.'
}

try {
    Write-UpdateLog "Starting Device Reboot Manager Gallery deployment. Module=$ModuleName; SkipPackageUpdate=$SkipPackageUpdate"

    if (-not $SkipPackageUpdate) {
        Invoke-GalleryPackageUpdate -Name $ModuleName
    }

    $moduleInfo = Get-LatestModuleInfo -Name $ModuleName
    $moduleBase = $moduleInfo.ModuleBase
    Assert-ModuleSignature -ModuleBase $moduleBase -Thumbprint $ExpectedSignerThumbprint

    $latestToolPath = Join-Path -Path $moduleBase -ChildPath 'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    if (-not $SkipPackageUpdate -and $latestToolPath -ne $PSCommandPath) {
        & $latestToolPath `
            -ModuleName $ModuleName `
            -InstallPath $InstallPath `
            -ConfigSourcePath $ConfigSourcePath `
            -ForceConfig:$ForceConfig `
            -SkipScheduledTask:$SkipScheduledTask `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -RepeatIntervalMinutes $RepeatIntervalMinutes `
            -EnableAutomaticUpdate:$EnableAutomaticUpdate `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName `
            -UpdateIntervalHours $UpdateIntervalHours `
            -IncludePrerelease:$IncludePrerelease `
            -ExpectedSignerThumbprint $ExpectedSignerThumbprint `
            -SkipPackageUpdate
        return
    }

    $installerPath = Join-Path -Path $moduleBase -ChildPath 'Runtime\Deploy\SmartM365-DeviceRebootManager-Install.ps1'
    if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
        throw "Runtime installer not found in package: $installerPath"
    }

    $installParameters = @{
        InstallPath           = $InstallPath
        ConfigSourcePath      = $ConfigSourcePath
        ForceConfig           = $ForceConfig
        SkipScheduledTask     = $SkipScheduledTask
        TaskPath              = $TaskPath
        TaskName              = $TaskName
        RepeatIntervalMinutes = $RepeatIntervalMinutes
        PackageSource         = 'PowerShellGallery'
        UpdateTaskPath        = $UpdateTaskPath
        UpdateTaskName        = $UpdateTaskName
    }
    & $installerPath @installParameters

    $installedUpdateScript = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    Copy-Item -LiteralPath $latestToolPath -Destination $installedUpdateScript -Force

    if ($EnableAutomaticUpdate) {
        Register-GalleryUpdateTask `
            -ScriptPath $installedUpdateScript `
            -ScheduledTaskPath $UpdateTaskPath `
            -ScheduledTaskName $UpdateTaskName `
            -IntervalHours $UpdateIntervalHours
        Write-UpdateLog ("Automatic update task registered: {0}{1}" -f $UpdateTaskPath,$UpdateTaskName)
    }
    else {
        $existingUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        if ($existingUpdateTask) {
            Unregister-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -Confirm:$false
        }
        Write-UpdateLog 'Automatic update task is disabled.'
    }

    $metadataPath = Join-Path $InstallPath 'SmartM365-DeviceRebootManager.installation.json'
    $metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
    $metadata | Add-Member -NotePropertyName ModuleName -NotePropertyValue $ModuleName -Force
    $metadata | Add-Member -NotePropertyName SignerThumbprint `
        -NotePropertyValue (Get-NormalizedThumbprint $ExpectedSignerThumbprint) -Force
    $metadata | Add-Member -NotePropertyName AutomaticUpdateEnabled `
        -NotePropertyValue ([bool]$EnableAutomaticUpdate) -Force
    $metadata | Add-Member -NotePropertyName UpdateIntervalHours `
        -NotePropertyValue $UpdateIntervalHours -Force
    $metadata | ConvertTo-Json -Depth 4 |
        Set-Content -LiteralPath $metadataPath -Encoding UTF8

    Write-UpdateLog ("Device Reboot Manager deployed successfully. PackageVersion={0}" -f $metadata.PackageVersion)
}
catch {
    Write-UpdateLog -Level ERROR -Message $_.Exception.Message
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCqRVJVefXuqISM
# ecSsajV7Co8AQqaiPu1cKUqB4UU2MqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOF5EE1inR67fseaXIjPqYm8mpfj+oLyC4z2uauD2pQWMA0GCSqG
# SIb3DQEBAQUABIIBgEwBd7wHYc7ogJ0Ni6ugd+aS5EFvknuS+J4Ymm9Ei1lhUTMU
# 0ZnyfXF8yHiNZ7gR72x4vaRvmXDhPW+DpCY9KNBlz/K4sQzhuOrQGKnjQRo2LqQA
# exh1Zdpg4kAgGB7o5CRAcBpuwFlQ8u589o0h0J9SQrAxn6tNmpAOk5oWhvCpb0Fj
# FlwTzDlK6QR/8lPU7g1kDH1kOBW5gqtWti/JNBP3/o5nRg1VM4Wbh1SSS7Sk89qJ
# rNvrKuMDJCmX90DDjOB8r5buQjv9rcwWS0x1C6Ulo3BQq9u9gzZGIvLaKnFZ0aQ9
# f9auQBqdsfJyO4MZvD46W8e3WnfS2azFS3ts2h3l8/3AbnFIjgQ4wt7DWiul874A
# zmo6IxUJfR12dOzc0z+7+FGS9Bf4cZZOtyuhmnsv9p3pysYcY3r+0jEwc5W3lbA5
# ns4nsDnt7E/7p8dFOOskmH+4JlIqT1ZANVahMsOMTASa/44KINAFimgqNE1jqcJd
# N40jQgo5y7hsw7tfqaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjkwNzU0
# MThaMC8GCSqGSIb3DQEJBDEiBCCA8IHvMF8VwQRTdVE1A/hnX50bdwOtlzSmzqZr
# YZHwIDANBgkqhkiG9w0BAQEFAASCAgB6rCvP/2Ysf3ILDeIdoUndDlSNhXMRgjAd
# ZnJXqS2pBYEe4J2T/OS5u0WCOnEzISfxtwgeaYr0VaxDDlTqYaMBX4xlO3JIA1WZ
# cR3gSn+O4uioLiU3xO+AyMcgHRWfjOYIBPwm4YUelPa3REFR7wV2XEpgBn+xPwjh
# Pncc2YMdrSLxQgTcmxMJ81ub12mZcIazIq/1RoqBeX9T2O4yqrskqacWmFex47s9
# dsLUmpVlJNEf1EtW4v4Xp4zR2SffVLKdiPbl2/bD7fMSvxDQ1TyHLeD0lgkePTVe
# 8/DLKNrhfMPtzg+EtNDd9QRtfoa+s7skH6CXzYsC9yViMR3YdRi7etRM3idN0bEI
# YE0GXHYLerrQUbXPN9DFvKuTH4XlrkZgGkoiJb1MnxItJjOL0/x27yfdHFKo5Vwz
# y1oR3HiC2ocs5LSfT6zxgsJRJKR9JB5aDy4KE2QkVRWiXrmqFW9ykjXjHVoElfTh
# nktEONkIWG3NlA/RzgCHccfEr69iryRiMrnnmQZSkQUYb++7YN+ADzAVPYI+YSgb
# lSKdQp35BDm+l1vIKicQC20OX3+4NTJ/So+kK2fWcNAyuj68LNaeHur9oFJ0W0G1
# USmU8VNuatVD5K3ntCrpofPlljf2yL8jR8TQC+WNBn7nRkZy7YHLyVbjIElx35vO
# q5HJk1ToXA==
# SIG # End signature block
