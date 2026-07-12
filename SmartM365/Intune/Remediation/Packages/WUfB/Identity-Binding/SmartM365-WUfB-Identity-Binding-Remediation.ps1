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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCALMuvpRjBLDrjt
# 5CcQbnaqF728Lm9SF2tOPl3lIcWJJ6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCBsN/70A4ziH309iLLIuEQfPsk6Z1wVsd530+HVB/tWTzANBgkqhkiG9w0B
# AQEFAASCAYB606qYutJ+DkiFEMMkFmo/z6gq6Wvr9s1pjUYCjF82O55DatXPrHgs
# guzQEdIbekE6HEi+oAGso5NyQbxVyPb+qolHFfEATnaDZ/mjCSdCeBeNhDkJVx8m
# UNxHKXOUq5F3MRvjFBzn3yoDLD1OldGL6MAOuRgeFyrZ946uqbkeR3LONyJ72m+W
# EGjVdxPLYoIwoEQ6go5xcw2yYEPudKxec973bGlf+kGLhBDD3nVGo5risFBSAe+7
# 5d3PSuo6TjBzJLFpDjdaAIgLFo0cv5wOlJlsfqiR7S72ZfRw8Zcah3L4OJw35IVY
# nc6beah9RHjHVgUc+jNgmPb/qsS/z0ntyjbx4CZCA9OLB0v2wklaIolb17O8ddWz
# IOSfzaTkESWb2BzHA20mtSN7Q6cEYgXTdj7vHQNnrynEpjUqfrml6ufQN/sexh4O
# Ndj7XKM4CyhECQLgr7eCB7nUYReKhdlHhGvgUfUgzIveyqNAfdziJPPUWOdFKzmi
# pK7Nmfssnbc=
# SIG # End signature block
