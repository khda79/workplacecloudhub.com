# Name: SmartM365-WindowsUpdate-Reset-Remediation.ps1
# Version: 1.0
# Description: Resets Windows Update components and removes legacy WSUS policies that block WUfB or Autopatch.

$ErrorActionPreference = "Stop"

$ScriptName = "Remediate-WindowsUpdate-Reset"
$Scenario = "WindowsUpdate-Reset"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$ScriptName.log"
$WuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$WuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($service -and $service.Status -ne "Stopped") {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStopRequested=$Name"
    }
}

function Invoke-ServiceStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($service) {
        try {
            $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

            if ($serviceCim -and $serviceCim.StartMode -eq "Disabled") {
                Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
                Write-SmartM365Log "ServiceStartupTypeChanged=$Name StartupType=Manual"
            }
        }
        catch {
            Write-SmartM365Log "ServiceStartupTypeCheckFailed=$Name Message=$($_.Exception.Message)"
        }

        Start-Service -Name $Name -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStartRequested=$Name"
    }
}

function Invoke-RegistryValueRemovalSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -Path $Path)) {
        return
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue

    if ($item -and $item.PSObject.Properties.Name -contains $Name) {
        Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "RegistryValueRemoved=$Path\$Name"
    }
}

function Rename-FolderSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-SmartM365Log "FolderNotFound=$Path"
        return
    }

    $suffix = Get-Date -Format "yyyyMMddHHmmss"
    $destination = "{0}.SmartM365.old.{1}" -f $Path, $suffix

    try {
        Rename-Item -LiteralPath $Path -NewName (Split-Path -Leaf $destination) -Force -ErrorAction Stop
        Write-SmartM365Log "FolderRenamed=$Path Destination=$destination"
    }
    catch {
        Write-SmartM365Log "FolderRenameFailed=$Path Message=$($_.Exception.Message)"
    }
}

function Invoke-UsoClientSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    $usoClient = Join-Path -Path $env:windir -ChildPath "System32\UsoClient.exe"

    if (Test-Path -LiteralPath $usoClient) {
        Start-Process -FilePath $usoClient -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-SmartM365Log "UsoClient=$Action Status=Triggered"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate", "DoNotConnectToWindowsUpdateInternetLocations", "DisableWindowsUpdateAccess", "SetDisableUXWUAccess")) {
        Invoke-RegistryValueRemovalSafe -Path $WuPolicyPath -Name $name
    }

    foreach ($name in @("UseWUServer", "NoAutoUpdate", "AUOptions")) {
        Invoke-RegistryValueRemovalSafe -Path $WuAuPolicyPath -Name $name
    }

    foreach ($service in @("bits", "wuauserv", "dosvc", "cryptsvc")) {
        Invoke-ServiceStopSafe -Name $service
    }

    Rename-FolderSafe -Path (Join-Path -Path $env:windir -ChildPath "SoftwareDistribution")
    Rename-FolderSafe -Path (Join-Path -Path $env:windir -ChildPath "System32\catroot2")

    foreach ($service in @("cryptsvc", "dosvc", "wuauserv", "bits")) {
        Invoke-ServiceStartSafe -Name $service
    }

    Invoke-UsoClientSafe -Action "RefreshSettings"
    Invoke-UsoClientSafe -Action "StartScan"

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBLvXiLdYZWDJmf
# mTh+xdRFvPDlkXKf+8n3gAVHmpeTfqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDmHqC5x0HfNdac6bth6tbxyox1pbf3SeFsGjrXc3PjOzANBgkqhkiG9w0B
# AQEFAASCAYCrWSvu0pOplopDlO+4atwAbbr6fkonadnSbld3qsTyYZ+8zMK3jhbA
# CIBEdoRYOjyD/L4Gl5KSNKZ629bMN6G9/Ny5N6FJpIhiXfklLR51YWs8eyLFoTMZ
# FlEeNoq82k6JQIk77ynyVMxrFUvM4LOCbTxd2CtWbKTVFNb4yY1+DbpDq/Fvm61U
# mitQZvbqmzisupslJb/qX88Ukqog9lyXAMFGvIG+9LLZjcW1OwCBXlZnz7zCm4af
# WHGzEmIWZROZIsIL9mroqH0hKsfiB9BwM0gYaK9NZu3x7Vxg9OB2jPj05st6lozz
# 4Nof8Cx3kJDrltfclP4OKliJKrIVBdcdbS2ta3JNKjM18Xzlkt2wXiagGkEKJHeq
# LLRaMYhpTADZA14GBGpN3d5MDCZ0aYAEdlX8Uv7Zs4z9wJm4mtMvdmzR9yIWfWwO
# qw+VFhNTojTIlAvs4rODdgKkwy5x8lek+GuM5M3ctBev/5HtGpC1UBppBKwLdJX5
# 11X5Ww1MtHY=
# SIG # End signature block
