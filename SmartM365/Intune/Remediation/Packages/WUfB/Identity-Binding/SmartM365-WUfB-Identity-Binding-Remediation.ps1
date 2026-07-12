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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCALMuvpRjBLDrjt
# 5CcQbnaqF728Lm9SF2tOPl3lIcWJJ6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBsN/70A4ziH309iLLI
# uEQfPsk6Z1wVsd530+HVB/tWTzANBgkqhkiG9w0BAQEFAASCAYCY4KUjVZfAoLKB
# ikwAmG93TcckcuHhYgEA2Es9CABG99CVqAZMa5Dg4jLJ9MmHzEo1Vm/an5XTzQ+5
# 7kPlDLQbjOaahWw1ELAJzspgy5qVhyHjRmPGVzBKniMYhio70GGyrjXMCNGKmVFt
# 83p36Ts8dGivPYQCq5lxq9SnxPVDM81hJmD1yqPzvPYNXchbvmlwzYj4kY4li0Xp
# RzKrsqw6pc98rGWwBmWrTURw+rrvHCgRkdv7A8PFk2cNn6B3tkSgHZN2k/XuR4Ac
# YznvDnQooaD/PRPaFjfpsR6LTaePlfVVekKQQT4ntv9XK8E8cjC6K1GNh71BpgQD
# s/6oW8sj5x91sulKmVRK05zZUpS+qJVRBB/3yRMjmVzxx5b+4ypLX2rqnQxciYM9
# z0OxrvFyytRrn7p/w+iRTXo2Of2DkeLip4GPAW5mW/MLslKAGlEy375Pjcxmp+1t
# JQW5fXuN/7kgGa8/LpMPT5rzaOQ6EY5ZMCjtTEu2fy+qrQiRC9c=
# SIG # End signature block
