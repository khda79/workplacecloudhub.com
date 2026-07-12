# Name: SmartM365-RepairWindowsUpdateandDISMissues-Remediation.ps1
# Version: 1.0

# Remediation Script for Windows Update and DISM
# Purpose: Repair Windows Update and DISM issues blocking Windows 11 upgrade

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Repair-DISM'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'RepairWindowsUpdateandDISMissues.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
$hadError = $false

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [Repair-DISM] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

# Step 1: Restart Windows Update service
Write-SmartM365Log "Restarting Windows Update service..."
Try {
    Restart-Service -Name wuauserv -Force -ErrorAction Stop
    Write-SmartM365Log "Windows Update service restarted successfully."
} Catch {
    Write-SmartM365Log "Failed to restart Windows Update service: $($_.Exception.Message)"
    $hadError = $true
}

# Step 2: Run DISM to repair system image
Write-SmartM365Log "Running DISM /Online /Cleanup-Image /RestoreHealth..."
Try {
    $dismResult = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "DISM completed with exit code: $($dismResult.ExitCode)"
    if ($dismResult.ExitCode -notin @(0, 3010)) {
        $hadError = $true
    }
} Catch {
    Write-SmartM365Log "DISM failed: $($_.Exception.Message)"
    $hadError = $true
}

# Step 3: Run SFC to repair system files
Write-SmartM365Log "Running SFC /scannow..."
Try {
    $sfcResult = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "SFC completed with exit code: $($sfcResult.ExitCode)"
    if ($sfcResult.ExitCode -notin @(0, 1)) {
        $hadError = $true
    }
} Catch {
    Write-SmartM365Log "SFC failed: $($_.Exception.Message)"
    $hadError = $true
}

# Step 4: Clear Windows Update cache
Write-SmartM365Log "Clearing Windows Update cache..."
Try {
    net stop wuauserv
    Remove-Item -Path "$env:windir\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
    net start wuauserv
    Write-SmartM365Log "Windows Update cache cleared."
} Catch {
    Write-SmartM365Log "Failed to clear Windows Update cache: $($_.Exception.Message)"
    $hadError = $true
}

# Step 5: Log completion
if ($hadError) {
    Write-SmartM365Log "Status=CompletedWithErrors"
    exit 1
}

Write-SmartM365Log "Status=Completed"
exit 0

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOjr3yD9hpnwA/
# ilYSlvg7z1RYXt7EEltSbeo65fYodaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD1jQ9qTiqld5BxhvM0
# AScsAKLMeiT8cl1UqwxAs97FQzANBgkqhkiG9w0BAQEFAASCAYBO5PtlOeN8EIRC
# M5yW+6fLX2yKWDuxApYldouDmBZpE7LAw+IgPDw/X3PImB/ryrkSI43z58E7jDl9
# i+B6yS6aicqGGeWPk0J9G4TxmYEuUdG3kd3d8GKbHl5aJP5pfgiMLf2ao3y7Isv7
# NaIGNFfJbQkOt/GLJtUXLPiG7VB2hNVlvcSa+QBzZhZMyMj3U4UEDiGyUhHIwZB6
# eyk2P0/E8YmctBjNu0WsWS8zJNqOGx3It9G50wAs2V9ZPHEQjprR6sMpkdj2wGk3
# YQcewyRF1fyFMb/y498Nmb3w2F8QHCIG3wQ6z+SqYmc04mKrGDXnz/gEi+LE9hy1
# DVTSF84ypzi0NJ56jrnwdiE6VBsQBc0GTc9JhmwqdXtNUv4/etJnZn7JAFeJVcVk
# Lc+dPjuYdy/ooW2dlkB3Tm259KkU/dKADyEn+Nj3sFn/XiI9F+AJNxT/h/1n4LR3
# Y/SU7HUDr/J4m9JAFp+EkazQfcHN1kJ/+RxDnXdP1DrA82dOcwY=
# SIG # End signature block
