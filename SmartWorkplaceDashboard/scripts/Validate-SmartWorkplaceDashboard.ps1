[CmdletBinding()]
param(
    [string]$DashboardRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$CheckSourceFiles
)

$ErrorActionPreference = 'Stop'

function Assert-JsonFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing file: $Path" }
    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json | Out-Null
    Write-Host "OK JSON $Path"
}

function Assert-FileContainsNoLegacyTerm {
    param([Parameter(Mandatory)][string]$Path)
    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $legacyPattern = ([string][char]101) + ([string][char]109) + ([string][char]101) + ([string][char]114) + ([string][char]105) + ([string][char]116)
    if ($content -match (('(?i)' + $legacyPattern))) { throw "Forbidden legacy term found in $Path" }
}

$pbipRoot = Join-Path $DashboardRoot 'pbip'
$reportDir = Join-Path $pbipRoot 'SmartWorkplaceDashboard.Report'
$modelDir = Join-Path $pbipRoot 'SmartWorkplaceDashboard.SemanticModel'

Assert-JsonFile -Path (Join-Path $pbipRoot 'SmartWorkplaceDashboard.pbip')
Assert-JsonFile -Path (Join-Path $reportDir 'definition.pbir')
Assert-JsonFile -Path (Join-Path $modelDir 'definition.pbism')
Assert-JsonFile -Path (Join-Path $modelDir 'model.bim')

$model = Get-Content -LiteralPath (Join-Path $modelDir 'model.bim') -Raw | ConvertFrom-Json
if (-not $model.model.tables -or $model.model.tables.Count -lt 10) { throw 'model.bim does not contain the expected initial tables.' }
if (-not $model.model.expressions -or $model.model.expressions.Count -lt 5) { throw 'model.bim does not contain the expected Power Query expressions.' }
if (-not (@($model.model.tables.name) -contains 'FactSourceFreshness')) { throw 'FactSourceFreshness table is missing.' }
if (-not (@($model.model.tables.name) -contains 'Measures')) { throw 'Measures table is missing.' }
Write-Host ("OK model tables={0} expressions={1}" -f $model.model.tables.Count, $model.model.expressions.Count)

Get-ChildItem -LiteralPath $DashboardRoot -Recurse -File | Where-Object { $_.Extension -in '.md','.pq','.json','.bim','.pbip','.pbir','.pbism','.ps1' } | ForEach-Object {
    Assert-FileContainsNoLegacyTerm -Path $_.FullName
}
Write-Host 'OK forbidden term scan'

if ($CheckSourceFiles) {
    $cmdbPath = ($model.model.expressions | Where-Object name -eq 'CMDBPowerBIPath').expression.Trim('"')
    $cmdbFullPath = if ([System.IO.Path]::IsPathRooted($cmdbPath)) { $cmdbPath } else { Join-Path $DashboardRoot $cmdbPath }
    $requiredCmdb = @('DimTenant.csv','DimDate.csv','DimUser.csv','DimDevice.csv','DimLicenseSku.csv','FactUserLicense.csv','FactDeviceCompliance.csv','FactMailbox.csv','FactDataQuality.csv')
    foreach ($fileName in $requiredCmdb) {
        $candidate = Join-Path $cmdbFullPath $fileName
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw "Missing CMDB source file: $candidate" }
    }
    Write-Host "OK CMDB source files $cmdbFullPath"
}

Write-Host 'Smart Workplace Dashboard validation completed.'

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAi0I6Pmp5I+NRA
# SyWmPkE7lNWgyIUZ9OUa1L78Vp1+x6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDASMbx0TPSyNHGCpSd
# DliVDtu0/IN9za9e0WwFsB2OeDANBgkqhkiG9w0BAQEFAASCAYBOUfPsa5YC1T/k
# jkA23fiBH+Hu6TQ+3riNMCVKDx1KVGrMfMdY1BGFKTTka7pVDcMbt9aU0XyajcDk
# 568fjYLHXdzUNHBvcIl43KAe6bTvuv5+LgIb14QFFinc/3+dmO5vTUixHknF0VGm
# 2M36C6c1hCZiK7E1A87B2zUncXfCbv0ggsHOm4Y8w4iaicpmtX2YTdCMo6uqs4YU
# opYX8S8TFdhZCEQRpf1yHeRKiITsnedo0vCDcWNPTo3Ed5nZQXiOxSHwUO33SVYi
# 0t6M6k3Qern+n48vIYjh9Q4/hXbjggk4xPIO/xsO7z+W6ePocXdg1C4LEbUSji6K
# tvINB/Ml60iqLryYtWAsT1AUSobeurKSb4FGTaFFnzq2F8PeAHDYWaFcNDwtEfgK
# u31bclTmP95lilQs8JenCymDlyLXbzzEPrSWgfXTvMBduRQ7MMb+AqSZgQ6rxwxh
# NubRGKTFp18lbHhMrBZdLQViEYMTy+2dKzWkZgznTtOoVmVvpLI=
# SIG # End signature block
