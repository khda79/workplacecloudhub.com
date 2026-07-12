# Name: SmartM365-GroupPolicy-Stale-Detection.ps1
# Version: 1.0
$ErrorActionPreference = "Stop"
$groupPolicyStatePath = "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}"
$maxAgeDays = 7

try {
    $state = Get-ItemProperty -Path $groupPolicyStatePath -ErrorAction Stop

    if ($null -eq $state.startTimeHi -or $null -eq $state.startTimeLo) {
        Write-Output "Status=GroupPolicyTimestampMissing"
        exit 1
    }

    $fileTime = ([Int64]$state.startTimeHi -shl 32) -bor [UInt32]$state.startTimeLo
    $lastGPUpdateDate = [datetime]::FromFileTime($fileTime)
    $lastGPUpdateDays = (New-TimeSpan -Start $lastGPUpdateDate -End (Get-Date)).TotalDays

    if ($lastGPUpdateDays -gt $maxAgeDays) {
        Write-Output ("Status=GroupPolicyStale AgeDays={0:N1} MaxAgeDays={1}" -f $lastGPUpdateDays, $maxAgeDays)
        exit 1
    }

    Write-Output ("Status=GroupPolicyFresh AgeDays={0:N1}" -f $lastGPUpdateDays)
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAHw4Zhyp2hPBck
# cxSXUzSV4hbf+oPPeJF2V+lG/UrmIqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAWBrBUAMS0dxhLHxgU
# BSUw5Dku8N4Z7r4of3fORPZDgTANBgkqhkiG9w0BAQEFAASCAYAaiZ87A0VkCvNf
# 96UYWrENlmbDTuUZMH8SlvWfAT9pQ6pixSJBFF8bD6b1x76bqBBlUXw+HdfyQKkH
# 5q0j/rTozYP2tzgguRUX/mrJAus6tpNt2hxvb3A6gvC3+gIXgivnzRyGe4JtN20F
# gDEYblCdf0T2FtnA9PuTGtJGz7iDGEk7v1MJtA0q3k6sCpkO5mMq3FADdugqYG2j
# MITZkmAlwzjA9O2NWf66PjtRaVKqestqXeSLDQUxFHpmv5lo7ZAwBrNA5U6Vd3Ke
# aJb/HXGuHmdgOZaFWZPxSF40ms6i08O31xP7/033PENsBqn/vs+dpWrxA/0bXt6z
# ipuTc31ED4kJS/+b71QMmYmQQl6IMmUpf0NRAIQq1uPcn2ijNTj/R/tz8akwGeW5
# UvL4PpoR3ZwB5mwAUATUl9hIuraQCR1ZZaiRZGlW7sYcWFq7Mj0tv99peK8pW65z
# qV2CCVfSYe3YnEAcoNcFQcCnz6BG+ht7krksoEJqHQ06SP1oRYg=
# SIG # End signature block
