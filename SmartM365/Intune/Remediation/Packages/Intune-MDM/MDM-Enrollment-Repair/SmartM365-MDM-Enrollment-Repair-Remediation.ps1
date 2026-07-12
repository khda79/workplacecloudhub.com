# Name: SmartM365-MDM-Enrollment-Repair-Remediation.ps1
# Version: 1.0
# Description: Repairs local device registration and MDM enrollment signals without deleting the current identity.

$ErrorActionPreference = "Stop"

$Scenario = "MDM-Enrollment-Repair"
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Invoke-ProcessSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string]$Arguments,

        [int]$TimeoutSeconds = 120
    )

    if (-not (Test-Path -LiteralPath $FilePath)) {
        Write-SmartM365Log "ProcessSkipped FilePath=$FilePath Reason=NotFound"
        return
    }

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden -ErrorAction Stop
        $completed = $process.WaitForExit($TimeoutSeconds * 1000)

        if (-not $completed) {
            Write-SmartM365Log "ProcessTimeout FilePath=$FilePath Arguments=$Arguments"
            return
        }

        Write-SmartM365Log "ProcessCompleted FilePath=$FilePath Arguments=$Arguments ExitCode=$($process.ExitCode)"
    }
    catch {
        Write-SmartM365Log "ProcessFailed FilePath=$FilePath Arguments=$Arguments Message=$($_.Exception.Message)"
    }
}

function Invoke-TaskStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Task
    )

    try {
        if ($Task.State -eq "Disabled") {
            Enable-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath -ErrorAction SilentlyContinue | Out-Null
            Write-SmartM365Log "TaskEnabled=$($Task.TaskPath)$($Task.TaskName)"
        }

        Start-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath -ErrorAction Stop
        Write-SmartM365Log "TaskStarted=$($Task.TaskPath)$($Task.TaskName)"
    }
    catch {
        Write-SmartM365Log "TaskStartFailed=$($Task.TaskPath)$($Task.TaskName) Message=$($_.Exception.Message)"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\gpupdate.exe") -Arguments "/target:computer /force" -TimeoutSeconds 180
    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\DeviceEnroller.exe") -Arguments "/c /AutoEnrollMDM" -TimeoutSeconds 180
    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\dsregcmd.exe") -Arguments "/join" -TimeoutSeconds 180
    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\dsregcmd.exe") -Arguments "/refreshprt" -TimeoutSeconds 60

    Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object {
            $_.TaskPath -like "\Microsoft\Windows\EnterpriseMgmt\*" -and
            ($_.TaskName -like "*PushLaunch*" -or $_.TaskName -like "*Schedule*")
        } |
        ForEach-Object { Invoke-TaskStartSafe -Task $_ }

    $imeService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue

    if ($imeService) {
        Restart-Service -Name "IntuneManagementExtension" -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceRestartRequested=IntuneManagementExtension"
    }
    else {
        Write-SmartM365Log "ServiceNotFound=IntuneManagementExtension"
    }

    Write-SmartM365Log "RemediationCompleted"
    exit 0
}
catch {
    Write-SmartM365Log "RemediationFailed Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASkKiZzH+28V3/
# c8JOr8G6op8WAP66qtbBwOMOlHDPDqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCClClxeeibbzBfXne+ERfJnrI96nEbrE6/bVf2SjM/HZzANBgkqhkiG9w0B
# AQEFAASCAYCHyrtsUckEq/yQXlMSXuUg0+ji3+Inx6y42qyr7JPQ3NmPM6T8UAuv
# CuBc8be7RA89BWvFpZyUU1pFKFiuItqOZu9ozltDpRv7mK8DHWDRCSKrQmmvqS+S
# FWkFp8BUXU2xo7kHG2bk02JZoIejjjwjZ3yBE57eByx2ljNajIzvDUXLIraFcxfK
# bwGfdAYaeInEN76nClRZNkvvn2GRKDrpw0ppG3esh47DL347sfMPbrgs1fUIv2Vg
# +v3GQTtH40FMmzWYF1dLv0GZ3+aS1lc6szMgDdkKI3YjNCf8LbAZuNGB4W2kuS7Z
# chNX941XoroB4bpJewtja1u9cqB64/YBU2qCECVhQtEPnysTKFXSoXJswNfa1ys1
# qAl2pd75R92n25sTWE1LhPJScKhgqkVTGXyYAYhxHP0ytQm3ocf8od5zQkqAezkM
# 5sw51zmwFDRq2XxAqwm7o87xQW+A9iCrknWs7eymoevE7RD364QWJAX0IFHFWZ/f
# oWMHyhRAurs=
# SIG # End signature block
