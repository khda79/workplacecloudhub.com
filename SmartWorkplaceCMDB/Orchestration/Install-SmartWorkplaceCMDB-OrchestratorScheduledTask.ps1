#Requires -Version 5.1

<#
.SYNOPSIS
Installs or removes the SmartWorkplaceCMDB orchestrator scheduled task.

.DESCRIPTION
Creates a daily Windows scheduled task that runs the autonomous
SmartWorkplaceCMDB orchestrator with PowerShell 7 and the explicit -Collect
switch. The task runs under a dedicated Windows account so its CurrentUser
certificate store, Active Directory permissions, network access, and
SharePoint permissions remain available.

The script is preview-only unless -Execute is supplied. During installation,
the service-account password is requested through Get-Credential and is never
accepted as a command-line parameter or written to a file.

.PARAMETER Tenant
Tenant profile passed to the orchestrator. Defaults to prod.

.PARAMETER ServiceAccount
Dedicated Windows account used to run the scheduled task. SYSTEM and
LocalSystem are refused.

.PARAMETER DailyAt
Daily start time in 24-hour HH:mm format. Defaults to 02:00.

.PARAMETER TaskName
Optional scheduled-task name. Defaults to
SmartWorkplaceCMDB Orchestrator - <Tenant>.

.PARAMETER StartNow
Starts the task after successful installation. This immediately launches a
live full collection and may publish CSV files to SharePoint.

.PARAMETER Uninstall
Removes the scheduled task instead of installing it.

.PARAMETER Execute
Applies the planned installation, update, or removal. Without this switch, the
script only returns the proposed task definition.

.PARAMETER Interactive
Runs the guided workflow used by the centralized CMD launcher.

.EXAMPLE
.\Install-SmartWorkplaceCMDB-OrchestratorScheduledTask.ps1 `
  -Tenant prod -ServiceAccount 'CONTOSO\svc-cmdb'

.EXAMPLE
.\Install-SmartWorkplaceCMDB-OrchestratorScheduledTask.ps1 `
  -Tenant prod -ServiceAccount 'CONTOSO\svc-cmdb' -DailyAt '02:00' -Execute

.EXAMPLE
.\Install-SmartWorkplaceCMDB-OrchestratorScheduledTask.ps1 `
  -Tenant prod -Uninstall -Execute

.VERSION
0.1.0
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Alias('ProfileKey')]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._-]*$')]
    [string]$Tenant = 'prod',

    [string]$ServiceAccount = '',

    [ValidatePattern('^(?:[01]\d|2[0-3]):[0-5]\d$')]
    [string]$DailyAt = '02:00',

    [string]$TaskName = '',

    [switch]$StartNow,

    [switch]$Uninstall,

    [switch]$Execute,

    [switch]$Interactive
)

$ScriptVersion = '0.1.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$taskFolderName = 'WCH'
$taskPath = '\WCH\'
$installerPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDirectory
$orchestratorPath = Join-Path $scriptDirectory (
    'SmartWorkplaceCMDB-Orchestrator.ps1'
)
$globalConfigPath = Join-Path $projectRoot (
    'Config\SmartWorkplaceCMDB.global.local.json'
)
$tenantConfigPath = Join-Path $projectRoot (
    "Config\Tenants\$Tenant.local.json"
)

function Write-SmartWorkplaceCMDBInstallerMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Information (
        '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    ) -InformationAction Continue
}

function Test-SmartWorkplaceCMDBAdministrator {
    [CmdletBinding()]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $identity
    $principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

function Get-SmartWorkplaceCMDBPowerShell7Path {
    [CmdletBinding()]
    param()

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
        $candidates.Add(
            (Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe')
        )
    }
    $programFilesX86 = [Environment]::GetEnvironmentVariable(
        'ProgramFiles(x86)'
    )
    if (-not [string]::IsNullOrWhiteSpace($programFilesX86)) {
        $candidates.Add(
            (Join-Path $programFilesX86 'PowerShell\7\pwsh.exe')
        )
    }

    $pwshCommand = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($null -ne $pwshCommand) {
        $candidates.Add($pwshCommand.Source)
    }

    foreach ($candidate in @($candidates)) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            continue
        }
        $resolvedCandidate = (Resolve-Path -LiteralPath $candidate).Path
        $majorVersion = & $resolvedCandidate -NoProfile -NonInteractive `
            -Command '$PSVersionTable.PSVersion.Major' 2>$null
        if ($LASTEXITCODE -eq 0 -and [int]$majorVersion -ge 7) {
            return $resolvedCandidate
        }
    }

    throw 'PowerShell 7 (pwsh.exe) was not found.'
}

function ConvertTo-SmartWorkplaceCMDBQuotedArgument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    if ($Value -match '[\r\n"]') {
        throw 'Command arguments must not contain quotes or line breaks.'
    }
    '"{0}"' -f $Value
}

function Resolve-SmartWorkplaceCMDBAccountSid {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccountName
    )

    try {
        $account = New-Object Security.Principal.NTAccount $AccountName
        $account.Translate(
            [Security.Principal.SecurityIdentifier]
        ).Value
    }
    catch {
        throw "Service account '$AccountName' could not be resolved. $($_.Exception.Message)"
    }
}

function Assert-SmartWorkplaceCMDBServiceAccount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccountName,

        [switch]$ResolveSid
    )

    if ([string]::IsNullOrWhiteSpace($AccountName)) {
        throw 'ServiceAccount is required when installing the task.'
    }
    if ($AccountName -notmatch (
            '^(?:[A-Za-z0-9._-]+\\[A-Za-z0-9._$+@-]+|' +
            '[A-Za-z0-9._+@-]+@[A-Za-z0-9.-]+)$'
        )) {
        throw 'ServiceAccount must use DOMAIN\user or user@domain format.'
    }

    $normalized = $AccountName.Trim().ToLowerInvariant()
    $forbiddenAccounts = @(
        'system',
        'localsystem',
        'nt authority\system',
        's-1-5-18'
    )
    if ($normalized -in $forbiddenAccounts) {
        throw 'SYSTEM and LocalSystem are refused. Use the dedicated account that owns the collection certificate and required access.'
    }

    if ($ResolveSid) {
        $sid = Resolve-SmartWorkplaceCMDBAccountSid -AccountName $AccountName
        if ($sid -eq 'S-1-5-18') {
            throw 'SYSTEM and LocalSystem are refused.'
        }
        return $sid
    }
}

function Read-SmartWorkplaceCMDBInstallerYesNo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [bool]$Default = $false
    )

    $defaultText = if ($Default) { 'Y' } else { 'N' }
    while ($true) {
        $answer = Read-Host "$Prompt [y/n, default=$defaultText]"
        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $Default
        }
        switch ($answer.Trim().ToLowerInvariant()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
            default {
                Write-SmartWorkplaceCMDBInstallerMessage (
                    'Enter y or n.'
                )
            }
        }
    }
}

function Initialize-SmartWorkplaceCMDBTaskFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderName
    )

    $taskService = $null
    $rootFolder = $null
    $targetFolder = $null
    try {
        $taskService = New-Object -ComObject 'Schedule.Service'
        $taskService.Connect()
        $normalizedPath = '\{0}' -f $FolderName.Trim('\')
        try {
            $targetFolder = $taskService.GetFolder($normalizedPath)
        }
        catch {
            $rootFolder = $taskService.GetFolder('\')
            $targetFolder = $rootFolder.CreateFolder(
                $FolderName.Trim('\')
            )
        }
    }
    finally {
        foreach ($comObject in @(
                $targetFolder,
                $rootFolder,
                $taskService
            )) {
            if ($null -ne $comObject -and
                [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                    $comObject
                )
            }
        }
    }
}

function Start-SmartWorkplaceCMDBElevatedInstaller {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PowerShellPath
    )

    foreach ($value in @($Tenant, $ServiceAccount, $DailyAt, $TaskName)) {
        if (-not [string]::IsNullOrWhiteSpace($value) -and
            $value -match '[\r\n"]') {
            throw 'Installer parameters must not contain quotes or line breaks.'
        }
    }

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        (ConvertTo-SmartWorkplaceCMDBQuotedArgument `
            -Value $installerPath),
        '-Tenant',
        (ConvertTo-SmartWorkplaceCMDBQuotedArgument -Value $Tenant),
        '-DailyAt',
        (ConvertTo-SmartWorkplaceCMDBQuotedArgument -Value $DailyAt),
        '-Execute'
    )
    if (-not [string]::IsNullOrWhiteSpace($ServiceAccount)) {
        $arguments += @(
            '-ServiceAccount',
            (ConvertTo-SmartWorkplaceCMDBQuotedArgument `
                -Value $ServiceAccount)
        )
    }
    if (-not [string]::IsNullOrWhiteSpace($TaskName)) {
        $arguments += @(
            '-TaskName',
            (ConvertTo-SmartWorkplaceCMDBQuotedArgument -Value $TaskName)
        )
    }
    if ($StartNow) {
        $arguments += '-StartNow'
    }
    if ($Uninstall) {
        $arguments += '-Uninstall'
    }

    Write-SmartWorkplaceCMDBInstallerMessage (
        'Administrator rights are required. Requesting UAC elevation.'
    )
    $process = Start-Process `
        -FilePath $PowerShellPath `
        -ArgumentList ($arguments -join ' ') `
        -WorkingDirectory $projectRoot `
        -Verb RunAs `
        -Wait `
        -PassThru
    return $process.ExitCode
}

$interactiveMode = $Interactive -or $PSBoundParameters.Count -eq 0

if ($interactiveMode) {
    Write-SmartWorkplaceCMDBInstallerMessage (
        'SmartWorkplaceCMDB orchestrator scheduled-task installer'
    )
    $action = Read-Host (
        'Action: 1=install/update, 2=uninstall [default=1]'
    )
    if ([string]::IsNullOrWhiteSpace($action)) {
        $action = '1'
    }
    if ($action -notin @('1', '2')) {
        throw 'Action must be 1 or 2.'
    }
    $Uninstall = $action -eq '2'

    $tenantAnswer = Read-Host "Tenant profile [default=$Tenant]"
    if (-not [string]::IsNullOrWhiteSpace($tenantAnswer)) {
        if ($tenantAnswer -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
            throw 'Tenant profile contains invalid characters.'
        }
        $Tenant = $tenantAnswer
    }

    if (-not $Uninstall) {
        while ($true) {
            $accountAnswer = Read-Host (
                'Dedicated service account (DOMAIN\user or user@domain)'
            )
            try {
                Assert-SmartWorkplaceCMDBServiceAccount `
                    -AccountName $accountAnswer
                $ServiceAccount = $accountAnswer
                break
            }
            catch {
                Write-SmartWorkplaceCMDBInstallerMessage $_.Exception.Message
            }
        }

        $timeAnswer = Read-Host "Daily start time HH:mm [default=$DailyAt]"
        if (-not [string]::IsNullOrWhiteSpace($timeAnswer)) {
            if ($timeAnswer -notmatch '^(?:[01]\d|2[0-3]):[0-5]\d$') {
                throw 'Daily start time must use HH:mm format.'
            }
            $DailyAt = $timeAnswer
        }
        $StartNow = Read-SmartWorkplaceCMDBInstallerYesNo `
            -Prompt (
                'Start immediately after installation? This launches a live collection'
            ) `
            -Default $false
    }

    $Execute = Read-SmartWorkplaceCMDBInstallerYesNo `
        -Prompt 'Apply this scheduled-task change?' `
        -Default $false
}

$tenantConfigPath = Join-Path $projectRoot (
    "Config\Tenants\$Tenant.local.json"
)

if ([string]::IsNullOrWhiteSpace($TaskName)) {
    $TaskName = "SmartWorkplaceCMDB Orchestrator - $Tenant"
}
if ($TaskName.IndexOfAny([char[]]'\/:*?"<>|[]') -ge 0) {
    throw 'TaskName contains an invalid scheduled-task character.'
}
if ($Uninstall -and $StartNow) {
    throw 'StartNow cannot be combined with Uninstall.'
}
if (-not $Uninstall) {
    Assert-SmartWorkplaceCMDBServiceAccount `
        -AccountName $ServiceAccount
}

$powerShell7Path = Get-SmartWorkplaceCMDBPowerShell7Path
$actionArguments = @(
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    (ConvertTo-SmartWorkplaceCMDBQuotedArgument -Value $orchestratorPath),
    '-Tenant',
    (ConvertTo-SmartWorkplaceCMDBQuotedArgument -Value $Tenant),
    '-Collect'
) -join ' '

$plan = [pscustomobject]@{
    Status                 = if ($Execute) { 'PendingExecution' } else {
        'Preview'
    }
    Operation              = if ($Uninstall) { 'Uninstall' } else {
        'InstallOrUpdate'
    }
    TaskPath               = $taskPath
    TaskName               = $TaskName
    Tenant                 = $Tenant
    ServiceAccount         = $ServiceAccount
    DailyAt                = $DailyAt
    StartNow               = [bool]$StartNow
    PowerShellPath         = $powerShell7Path
    OrchestratorPath       = $orchestratorPath
    ActionArguments        = $actionArguments
    GlobalConfigPath       = $globalConfigPath
    TenantConfigPath       = $tenantConfigPath
    Execute                = [bool]$Execute
    ScriptVersion          = $ScriptVersion
}

if (-not $Execute -or $WhatIfPreference) {
    $plan
    if (-not $Execute) {
        Write-SmartWorkplaceCMDBInstallerMessage (
            'Preview only. Add -Execute to apply the change.'
        )
    }
    return
}

if (-not (Test-SmartWorkplaceCMDBAdministrator)) {
    $exitCode = Start-SmartWorkplaceCMDBElevatedInstaller `
        -PowerShellPath $powerShell7Path
    if ($exitCode -ne 0) {
        throw "Elevated installer failed with exit code $exitCode."
    }
    return
}

if ($Uninstall) {
    $existingTask = Get-ScheduledTask `
        -TaskPath $taskPath `
        -TaskName $TaskName `
        -ErrorAction SilentlyContinue
    if ($null -eq $existingTask) {
        Write-SmartWorkplaceCMDBInstallerMessage (
            "Scheduled task not found: $taskPath$TaskName"
        )
        return
    }
    if ($PSCmdlet.ShouldProcess(
            "$taskPath$TaskName",
            'Unregister scheduled task'
        )) {
        Unregister-ScheduledTask `
            -TaskPath $taskPath `
            -TaskName $TaskName `
            -Confirm:$false
        [pscustomobject]@{
            Status        = 'Uninstalled'
            TaskPath      = $taskPath
            TaskName      = $TaskName
            ScriptVersion = $ScriptVersion
        }
    }
    return
}

$requiredFiles = @(
    $orchestratorPath,
    $globalConfigPath,
    $tenantConfigPath
)
foreach ($requiredFile in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "Required runtime file not found: '$requiredFile'."
    }
}

$serviceAccountSid = Assert-SmartWorkplaceCMDBServiceAccount `
    -AccountName $ServiceAccount `
    -ResolveSid

$dailyTime = [datetime]::ParseExact(
    $DailyAt,
    'HH:mm',
    [Globalization.CultureInfo]::InvariantCulture
)
$action = New-ScheduledTaskAction `
    -Execute $powerShell7Path `
    -Argument $actionArguments `
    -WorkingDirectory $projectRoot
$trigger = New-ScheduledTaskTrigger -Daily -At $dailyTime
$settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 12)

$target = "$taskPath$TaskName"
if (-not $PSCmdlet.ShouldProcess(
        $target,
        "Register daily task as $ServiceAccount"
    )) {
    return
}

$credential = $null
$passwordPointer = [IntPtr]::Zero
$plainTextPassword = $null
try {
    $credential = Get-Credential `
        -UserName $ServiceAccount `
        -Message (
            'Enter the password for the SmartWorkplaceCMDB scheduled-task account.'
        )
    if ($null -eq $credential) {
        throw 'Credential entry was cancelled.'
    }
    $credentialSid = Resolve-SmartWorkplaceCMDBAccountSid `
        -AccountName $credential.UserName
    if ($credentialSid -ne $serviceAccountSid) {
        throw "The entered credential does not belong to '$ServiceAccount'."
    }

    $passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR(
        $credential.Password
    )
    $plainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
        $passwordPointer
    )

    Initialize-SmartWorkplaceCMDBTaskFolder `
        -FolderName $taskFolderName
    Register-ScheduledTask `
        -TaskPath $taskPath `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Description (
            "Runs the full SmartWorkplaceCMDB collection for tenant '$Tenant'."
        ) `
        -User $ServiceAccount `
        -Password $plainTextPassword `
        -RunLevel Highest `
        -Force | Out-Null

    if ($StartNow) {
        Start-ScheduledTask -TaskPath $taskPath -TaskName $TaskName
    }

    [pscustomobject]@{
        Status          = 'Installed'
        TaskPath        = $taskPath
        TaskName        = $TaskName
        Tenant          = $Tenant
        ServiceAccount  = $ServiceAccount
        DailyAt         = $DailyAt
        StartNow        = [bool]$StartNow
        PowerShellPath  = $powerShell7Path
        OrchestratorPath = $orchestratorPath
        ActionArguments = $actionArguments
        ScriptVersion   = $ScriptVersion
    }
}
finally {
    $plainTextPassword = $null
    $credential = $null
    if ($passwordPointer -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDCTc/kXw+SZmGU
# YLQ+xd563FCHXkGDPksLMwrddbxRHKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBNMXk4zHop/XsI2zT1BZdt8Qz2QxcCzRet+zJm9OVrVMA0GCSqG
# SIb3DQEBAQUABIIBgBuWKPk2lsElDw3Tww98s/uSTFiaq2IuZVPJK2JXd6RcTdN/
# wkskO9XV4Be1mAeP/S2SKChEZXGu1oNRo4de0JQCnvGeFrmDTHEFX7m8uFYulTI6
# 1EPBB4Z17rjxY8BRfxtdU4eBtirveQq3/eUJ+E5zIe7XuQPvla3yM7TYaXAQy9zG
# BR44R4cEJW7r0V5EMXKP2tqpTrspV18FYFdI3RK9E2og88sAxLoVu0AHjQGseO5z
# 81raAJROAqaUP2D/MO5NeE/88woTBi2x3O9WS0cc8CQPqc5Cj+z28g51BZMQ8Qhs
# lBwlBx2spkwE4ab2b0P4aODkRhyYagVscju4nmSAL0w5JxXY0G2Kf1GwD+uSepzi
# 8wCOZ/ypq5WNyqje0aLP8a35KGqRCOBQ99d2siieF4saHlGLkK7mfwJ9kkw0kzhb
# XR3lVbQtKHL/GD8wFPoLosNxGaiKVT72OP/G4A2MpsgviY6Vth+/y8f4IKVfzt9/
# 9DMhpZJWfeQih7k3DKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAwMDAz
# MTBaMC8GCSqGSIb3DQEJBDEiBCDH1iy7O2X0RUZ6pxMAcm+vKLwta1/7OupdMOZP
# Kp3xYDANBgkqhkiG9w0BAQEFAASCAgBBrg8/Z9PpS90oDtcHGPCxhKKGa7YumOvL
# hx5k1iJE/YKRPVBPE/mPBi437NDquNhJch+BLqtcEQ8W5DVndIqLMs8InwpnXKZ7
# xtjl64LSoqVbMiJiILmmQ9fzdUuzw0syG/lHdLWGIyHWT3MZM0KhpTDn8Phj0Luh
# HID/bu5okmX9SesnBveZGRr8grxsJ99PUk0iIu91V/53Rwb7QkUIuaexHyG5/Ezr
# 3kd5syS1RTUq6u13ByWary5By0diG02PXA0HWaV6SsbpmREiUKbeJhfEYk347iqC
# BhQXjA/EAZGmZkkHs4yWmfpsNjz9xjdxuiLFw1lv0qnF/zkdVOQAS2dLS3GF81OF
# B2tJB6ZMFwgn2yHw7sj7tz4nvKyDE0vU7j1+WYeS7XMP/Sm+4AVLH34SrCDE3CqR
# CcCwo02EvWItbUoIDVUVMaPDOwKz59k7edHubqQRCrxYRi39QsgxfIvJfXkZkG80
# cMcxIYcFiWWbThiJXg+PT6J4pkOfqb5dR1VrVW5N1AWZqCB73Ta//fzCSO2aJVN1
# /1Eww0xa8IPhfYE4m94TTsx+R+pXBfclzdFSw0vqNtijLh1AIK5zHexo06T44IO5
# JiV/ML1/nx5dJVBBcnlUehdwxWI/obZgh3/XA1aAemqQDN7SJHc/MGEmrKQ0Z5QC
# pcEL0HmyQQ==
# SIG # End signature block
