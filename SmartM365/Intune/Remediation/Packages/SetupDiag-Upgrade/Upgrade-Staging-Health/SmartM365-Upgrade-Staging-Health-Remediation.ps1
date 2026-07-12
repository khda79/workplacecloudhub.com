# Name: SmartM365-Upgrade-Staging-Health-Remediation.ps1
# Version: 1.0
# Description: Removes stale Windows upgrade staging folders only when no recent setup activity is detected.

$ErrorActionPreference = "Stop"

$Scenario = "Upgrade-Staging-Health"
$RecentSetupActivityHours = 6
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
$UpgradePaths = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS')
$SetupIndicators = @(
    "C:\Windows\Panther\setupact.log",
    "C:\Windows\Panther\setuperr.log",
    'C:\$WINDOWS.~BT\Sources\Panther\setupact.log',
    'C:\$WINDOWS.~BT\Sources\Panther\setuperr.log'
)

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStopRequested=$Name"
    }
}

function Invoke-ServiceStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Start-Service -Name $Name -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStartRequested=$Name"
    }
}

function Invoke-UsoClient {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    $uso = Join-Path $env:SystemRoot "System32\UsoClient.exe"

    if (Test-Path -LiteralPath $uso) {
        Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-SmartM365Log "UsoClient=$Action Status=Triggered"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $recentSetupActivity = $false

    foreach ($indicator in $SetupIndicators) {
        if (Test-Path -LiteralPath $indicator) {
            $lastWrite = (Get-Item -LiteralPath $indicator -ErrorAction SilentlyContinue).LastWriteTime

            if ($lastWrite -and ((Get-Date) - $lastWrite).TotalHours -le $RecentSetupActivityHours) {
                $recentSetupActivity = $true
            }
        }
    }

    if ($recentSetupActivity) {
        Write-SmartM365Log "RecentSetupActivityDetected=True CleanupSkipped=True"
        exit 0
    }

    foreach ($service in @("bits", "wuauserv", "dosvc")) {
        Invoke-ServiceStopSafe -Name $service
    }

    foreach ($path in $UpgradePaths) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "UpgradeFolderRemoved=$path"
        }
        else {
            Write-SmartM365Log "UpgradeFolderNotFound=$path"
        }
    }

    foreach ($service in @("dosvc", "wuauserv", "bits")) {
        Invoke-ServiceStartSafe -Name $service
    }

    Invoke-UsoClient -Action "RefreshSettings"
    Invoke-UsoClient -Action "StartScan"
    Invoke-UsoClient -Action "StartDownload"

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAjTCa6vyBNDIVi
# 8pDClkrJ/+9jAlYLXvnPSWE6hHCPKKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBkwfSkH3eZc/pdqWX1
# ehFbCCEPBRGL/27vQHuS7vVCszANBgkqhkiG9w0BAQEFAASCAYBmPv7RLW5Bq9st
# xpLi4GTcW73luGEZ2TtdLGjdYRkfAVSqY26AR+7DLQ/7fVoBVsF6YnfgjwyRI+DQ
# ++KHK0/CrXrrIsRtUnyeNRC1duiCRD17VMAk0BJNAnartN86aumPOh5W8R071FAR
# HUyPbGg2+p1POE8t+qPydsLy/ER6Gj/BVWIJ9hZoBTjVSUwqiaCUCLH/mcvXAloS
# zf+bYeEqgB5k4kJy+wFy9/mix4CPWfpRFGMcdySMMpzadTiPCxq3PBD0MEk9FyX3
# d76V4d1jeRqw17ArxZahBSYP/yLtXlj/Vj08y5oZ0tW1OtGA/Ue4CTyLFTeI3zrW
# sr43WOiKftJQyTvP+CtB4/NAq5uG8rKk1O5hKz3wmTGIF2Dn5jXfmPw4copnB47Z
# HZik/AO5u8zNUey4T9cNqWUAwwUxn9SRgAlME6Kn0SVAOgsfNACD40hIFdtBFgME
# 6xR7YS3Mx/A5R123ZxpZ5ShjbfKwaV+h5gyB6+oLg2b1XAIde38=
# SIG # End signature block
