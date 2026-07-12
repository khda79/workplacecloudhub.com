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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASkKiZzH+28V3/
# c8JOr8G6op8WAP66qtbBwOMOlHDPDqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCClClxeeibbzBfXne+E
# RfJnrI96nEbrE6/bVf2SjM/HZzANBgkqhkiG9w0BAQEFAASCAYBLrtyFmKGQZWYT
# Uh2VDdOwZknQOMVfyX4FVBIAXSdtuiOZmaocI8FzFy9tNxf2JqZBfhALFm00b4tF
# /ondU98q8OVdvNWhDl2sTztdM4yOlpdCafiaZQPv5oqAhDgAF06t9XbKdzfMymrr
# nAG77a3tuyoOkV4b1iX+25Ld3IKQNubFAo08RlBbwXu2/9ffCUvHClBvnJMDiNCa
# XZIzBuqKMVkVKBbrVdDTDhlKG+9FODr9KzmBplosnG8bsXkGAS1XFNuqY1Y2Dhks
# eZGy/pM1YHvM78B+9qfecBjwIkJyGqahbo1V9oPQa3FIGhfECxmfN6VuSTP/z5ia
# cU3Y1uoXfZIc2pg88rx/ctnEiq1FVHAnSdGXD9/FJgmPVDeUpJHmX5DGhUB+T2Wc
# 13YiyrC/o3s2Ojh/v4DPAivcMhGq6Jk1y+yAcs2YJ0qlyb6NZlkptSWtiiqaxdHk
# te/+afNPeqEc0snIMDi5dfJHxr588EzctAJZPQKMogfUUGLYzJQ=
# SIG # End signature block
