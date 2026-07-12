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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC4cFTR73Gj+M+h
# hw+rh1NHNIi6XOAMQZk+H4FwBTuaYaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC1SFfR6w1AgA9C9EmE
# NOgbS1P/uA0Hqy01Vu69vPWIYTANBgkqhkiG9w0BAQEFAASCAYAb28cHAmIj9N0w
# mKkZuUrvLRqdkT1z8Nf6ITqeQLtlUVZiTseYPpAJMiCkCa87j+8gagQzP0vVra66
# YFNoBJoXzYsXxfLuaNv+m7lBuadwvYiC0z0yH1uUy6BoanW36DQCmaoUHyR2L47D
# FddTW/EYgo8cSQCB7Cyzlsx/ygHBac56YHYi7FTgkLCQGhg6IbZcQq5MXvrK8FDr
# KyjsrXTq+9U0VL0wH//LloaUkTtGaajpnYUuQrNpjyEOJEgKKIW1mYmcDYGa1T8m
# r2O7XNlM6V1jqGaIZUq47UBQf1Fw5KITWwfRPXhr2URQELzi41WVptdCwtncWOgg
# 8qJTIumNkVrdyu8GK5TuLDBPNYrjlO3wEheVWetk/g92Xbs/Z0i3JFA9Dd1e3DUS
# DM03dWD7IX2tksnJwc0UFghB+aWJ9fQh3trTCW4Y+ISrgXj80OgGoJTmQDYHhilh
# JRKHwmlKZnMoZFX9riNo0fkJz90/7kY6om21kEmrkIFQnf8VC2c=
# SIG # End signature block
