# Name: SmartM365-GroupPolicy-Stale-Remediation.ps1
# Version: 1.0
$ErrorActionPreference = "Stop"
$Scenario = "GroupPolicy-Stale"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"
$groupPolicyStatePath = "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}"

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $process = Start-Process -FilePath (Join-Path $env:SystemRoot "System32\gpupdate.exe") -ArgumentList "/force" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
    Write-SmartM365Log "GpupdateExitCode=$($process.ExitCode)"

    if ($process.ExitCode -ne 0) {
        Write-SmartM365Log "Status=GpupdateFailed"
        exit 1
    }

    $state = Get-ItemProperty -Path $groupPolicyStatePath -ErrorAction Stop
    $fileTime = ([Int64]$state.startTimeHi -shl 32) -bor [UInt32]$state.startTimeLo
    $lastGPUpdateDate = [datetime]::FromFileTime($fileTime)
    $lastGPUpdateHours = (New-TimeSpan -Start $lastGPUpdateDate -End (Get-Date)).TotalHours

    if ($lastGPUpdateHours -le 24) {
        Write-SmartM365Log ("Status=Completed LastGPUpdateHours={0:N1}" -f $lastGPUpdateHours)
        exit 0
    }

    Write-SmartM365Log ("Status=CompletedButStillStale LastGPUpdateHours={0:N1}" -f $lastGPUpdateHours)
    exit 1
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Write-SmartM365Log "Status=Error Message=$($_.Exception.Message)" } catch { Write-Output "LogWriteFailed=$($_.Exception.Message)" }
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBExWjLUcV8oUTp
# 4o1sMpbjNHQmlnsfeM0JOkqF41B6EKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBzV+FFmPnDiucByqxj
# FFNsdulGTE3w9cHFuGMb2ol8LDANBgkqhkiG9w0BAQEFAASCAYCBvsh8iyGPEb7i
# gXBEwgXnAmqxoiKAyX1lV+uCj6hF1dunBJ8J4Mq9XmSFaZNfBsenMn+gSRGbHY40
# HT9blE8vCHc7EVZ2bI8t0Hcl/7nRfxsNGJ4WjLHN/2InsIeDmEFOlhTUQVNpMQHx
# h2V7lwcLVQuHT0ScvG0UTubHnQldm0IuU/2wOBPnPr7lwvSuce+3TGXqSBlnqmul
# mpfqj6zJipgU0DokB1SQD1ksGQ6lJ+aEJ5g44vuujMMcWy9tcFRrGVgjmKICA8xo
# djw7nxAmxbQGR9XEtDcLDpsW/i6LY/YC7vmfELP/Q6g/1Mwf/xs++CSdCRUQYP28
# edJvGgj3DX6NYA6e2kK4z7HhOG88QGY+ayCxuz5yzEwFOwWDQhrgnLxNWgCluafA
# Lk/qANV9XJwBUDjCmhug+xqH+2tSBEK48PZRxLpoqsqL33tuAV4c9jDTw6lVgkMh
# e+e0EbVcjLlliPQrcOVF1qug0fcUTFlmoYuWhIroyX6CQ3seKWs=
# SIG # End signature block
