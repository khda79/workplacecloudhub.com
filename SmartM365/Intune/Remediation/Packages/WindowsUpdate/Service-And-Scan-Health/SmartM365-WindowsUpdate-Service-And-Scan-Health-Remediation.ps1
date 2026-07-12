<#
    Name: SmartM365-WindowsUpdate-Service-And-Scan-Health-Remediation.ps1
    Version: 1.0
    Description: Consolidated remediation for Windows Update service health, settings refresh, and scan trigger.
#>

[CmdletBinding()]
param(
    [int]$PostScanWaitSeconds = 30
)

$ErrorActionPreference = "Stop"
$Scenario = "Service-And-Scan-Health"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"
$ErrorFound = $false

function Write-SmartM365Log {
    param([Parameter(Mandatory = $true)][string]$Message)

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Add-RemediationError {
    param([Parameter(Mandatory = $true)][string]$Message)

    Write-SmartM365Log "ERROR: $Message"
    $script:ErrorFound = $true
}

function Repair-WindowsUpdateService {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$StartupType
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Add-RemediationError "ServiceMissing=$Name"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStartupTypeChanged=$Name StartupType=$StartupType"
        }

        if ($service.Status -in @("StopPending", "PausePending")) {
            Write-SmartM365Log "ServicePending=$Name Status=$($service.Status)"
            Start-Sleep -Seconds 10
            $service.Refresh()
        }

        if ($service.Status -eq "Running") {
            Restart-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceRestartRequested=$Name"
        }
        else {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStartRequested=$Name"
        }
    }
    catch {
        Add-RemediationError "ServiceRepairFailed=$Name Message=$($_.Exception.Message)"
    }
}

function Invoke-UsoClientAction {
    param([Parameter(Mandatory = $true)][string]$Action)

    try {
        $usoClient = Join-Path -Path $env:SystemRoot -ChildPath "System32\UsoClient.exe"

        if (-not (Test-Path -LiteralPath $usoClient -PathType Leaf)) {
            Add-RemediationError "UsoClientMissing"
            return
        }

        Start-Process -FilePath $usoClient -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-SmartM365Log "UsoClient=$Action Status=Triggered"
    }
    catch {
        Add-RemediationError "UsoClientFailed=$Action Message=$($_.Exception.Message)"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $servicePlan = @(
        @{ Name = "bits"; StartupType = "Manual" },
        @{ Name = "wuauserv"; StartupType = "Manual" },
        @{ Name = "dosvc"; StartupType = "Manual" },
        @{ Name = "cryptsvc"; StartupType = "Manual" },
        @{ Name = "UsoSvc"; StartupType = "Automatic" }
    )

    foreach ($serviceItem in $servicePlan) {
        Repair-WindowsUpdateService -Name $serviceItem.Name -StartupType $serviceItem.StartupType
    }

    Invoke-UsoClientAction -Action "RefreshSettings"
    Invoke-UsoClientAction -Action "StartScan"

    if ($PostScanWaitSeconds -gt 0) {
        Start-Sleep -Seconds $PostScanWaitSeconds
    }

    $lastEvent = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 1 -ErrorAction SilentlyContinue
    if ($lastEvent) {
        Write-SmartM365Log ("LastWUEvent={0:s} Id={1}" -f $lastEvent.TimeCreated, $lastEvent.Id)
    }

    if ($ErrorFound) {
        Write-SmartM365Log "Status=CompletedWithErrors"
        exit 1
    }

    Write-SmartM365Log "Status=Completed"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Add-RemediationError $_.Exception.Message } catch { Write-Output "ErrorDuringErrorHandling=$($_.Exception.Message)" }
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC4cFTR73Gj+M+h
# hw+rh1NHNIi6XOAMQZk+H4FwBTuaYaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCC1SFfR6w1AgA9C9EmENOgbS1P/uA0Hqy01Vu69vPWIYTANBgkqhkiG9w0B
# AQEFAASCAYAHFxbvexeFU7PjgE37qGwYDUtR6m+pfxkQo7RfD0Va15giyO+mLPJe
# LWr4SNPCUHr2LMLRDLWAu4i/MTcyjD0CXkjXwvyXVRyCCw41kjR/Qs22F8422y5n
# V6oej3nFITn1LtAtGdwPu63T9dOXmjhlE9Xw4zPn4HAcCr4Ca1+zFRkgd6H12LqX
# iZ0PUxpevprR7Aq7lLuPZrnB0FxhmB0oQvzPjl6bj1NLy65lNuWyAnrO7cbfbOFd
# 0h1x72DV80R+/8nsTiDX4sFSQkd5Tjk+tkAtftVlSIdFxTgJ88IkhEZCaLPA/0M+
# IaeLw7E1VVVdF6fYrOmnJItAs2GSGI3ekpMWHQK6l1Pf0cY+w7RM4ym4xJkHpxHZ
# MZcFV9vTdVHi3T8JsS9/q2qpxY+zzgz5C3z4MQsdTDKhL2ArEiEjo44cEZqT9Rvh
# bz+haUqOYzgwW2GuemPB9o1duleJg8d2GrIvr3xh5cf8GMmTSSUe5W28rs1FmCSa
# oKAZ1vKePiQ=
# SIG # End signature block
