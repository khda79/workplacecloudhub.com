<#
.SYNOPSIS
    Version: 1.0
    Repairs Intune Management Extension health for IntuneManagementExtension-Health.
.DESCRIPTION
    Ensures the IME service is configured to start automatically, recreates the expected log directory, restarts the IME service, starts EnterpriseMgmt tasks, and triggers a Windows Update scan to refresh policy-driven activity.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'IntuneManagementExtension-Health'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"

function Write-SmartM365Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }
function Invoke-UsoClient { param([string]$Action) $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'; if (Test-Path -LiteralPath $uso) { Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue; Write-SmartM365Log "UsoClient $Action triggered." } }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    New-Item -Path 'C:\ProgramData\SmartM365\IntuneRemediation\Logs' -ItemType Directory -Force | Out-Null
    Write-SmartM365Log 'Remediation started.'

    $service = Get-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
    Set-Service -Name 'IntuneManagementExtension' -StartupType Automatic -ErrorAction SilentlyContinue
    if ($service.Status -eq 'Running') {
        Restart-Service -Name 'IntuneManagementExtension' -Force -ErrorAction Stop
        Write-SmartM365Log 'Intune Management Extension service restarted.'
    }
    else {
        Start-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
        Write-SmartM365Log 'Intune Management Extension service started.'
    }

    Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.State -ne 'Disabled' -and ($_.TaskName -like '*PushLaunch*' -or $_.TaskName -like '*Schedule*') } |
        ForEach-Object {
            try { Start-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction Stop; Write-SmartM365Log "Started scheduled task: $($_.TaskPath)$($_.TaskName)" } catch { Write-SmartM365Log "Could not start scheduled task $($_.TaskName): $($_.Exception.Message)" }
        }

    Invoke-UsoClient -Action 'RefreshSettings'
    Invoke-UsoClient -Action 'StartScan'

    Write-SmartM365Log 'Remediation completed.'
    exit 0
}
catch {
    Write-SmartM365Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}


# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC9Znw9MBANPaDl
# zYHhFKa1BX0/VZeoCYvVzPUKw8ZCjaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAmpktjTFLRa+xaeU9Q
# U6tzv96R88TCY3HTvfm7VgadbzANBgkqhkiG9w0BAQEFAASCAYCSzKw9hQXPxbiF
# mM2MI792NVs/WX0zPbk5CkF9AjMmgf57kAso5a7MkaLVX0/eZFrI0YM2/onaz4h5
# Sl6YWY/oODBmKl/TdwjtNLr72katVuNyxG/meI9DNqIF+WNsv71/FHfgE5ub5bkQ
# 9ODRktWrVJb9LO7/9tBjrxUiP9tGEJpJTpLbuG0YaUY5KzMJpozkSGulOkTlcgR6
# UwYPDmNOnuUP+hVuvsCCns5F64wxvJ1Pz6PIgU8bM2J32nkfYxurwVgcIRgaQDBg
# +zf0+Lcj7Jj+3UlG/qaDYxklioiI5NOsAHzlyRpd/ogq0A9z3tjkaUgJWLw0CkF4
# ShSwiW0p8KrpWi7efJVDrwMT8HdkRbkRjY8CnX83KJJ3dzvXXeReTNO6MfWIMryh
# NDAKOvJ+Al0uniB2q93RrSp22wRf+cSeHLlci7iml5okAExI0ShGxTjRdwwAve0a
# md/UT91V4ikCk9x7X12Ags/dWUD3fYDb8pi1ID+lEJzyvEfNEJI=
# SIG # End signature block
