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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBLvXiLdYZWDJmf
# mTh+xdRFvPDlkXKf+8n3gAVHmpeTfqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDmHqC5x0HfNdac6bth
# 6tbxyox1pbf3SeFsGjrXc3PjOzANBgkqhkiG9w0BAQEFAASCAYA/qNdtRt79XeZV
# OiFbQznlSAixIOQDugonymH1bX1sHC84gK5+R5FxCoMNZaNhgsfqy9IwicFnkJ2B
# xqMvbFC0Zv5RhvoAETPOxcC/vdWjJFQXOzfFLB51StDOR7OKJY/4ryTQ8tBGddzr
# n66cIauhfORGmGbgbLX90xq/30aw+CscxYndlP44ksisr3WFVdwVy8wwYnPDmxvi
# 3ypMOZsVTt7LUROLWRM4KgQgAWzZxYfZm7IBkWDJBI2J2qoKmiBGMLEcI0jqrNyj
# yYGP7sFALuW2J+9FmWunPKvVd+L17D/xyYEdsON1J9+J/fdBH4NOSpdI8WPR04ht
# 6gV3GD1PeMEefWSCc0Ag8d0srw2FkeOO/rzG/pNMv3hPOzw1kjvYDCUZcQdP9P+e
# +noDDZ/hfrwhGHOfRZe6T0tMYiMsrwpD1e6LAxU7ePOeVJgmMA+L9IoiX4CyQ8IB
# 5QyZxYH2LKcWwXI8ny2nCnghHolwuP9ALfyX4ChvN5tszIMsSRU=
# SIG # End signature block
