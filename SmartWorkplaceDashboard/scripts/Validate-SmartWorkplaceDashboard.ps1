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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAi0I6Pmp5I+NRA
# SyWmPkE7lNWgyIUZ9OUa1L78Vp1+x6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDASMbx0TPSyNHGCpSdDliVDtu0/IN9za9e0WwFsB2OeDANBgkqhkiG9w0B
# AQEFAASCAYBi9p14QhPkgc2QouH587o3NamiwrQIRhtFqcKiiMe9VkzB55TxYciW
# URM+SGfxibaRyKWOyWSomIT9VsRT6QnGx6IR5xlnEwJaHEUe5nF6gfxARES++jPB
# 3ZeWN2Fj5HJDDTFZnxd0vdFU9vQErGNluiYP2tJpwzsNH24uh7PAayPBk/OK5CWz
# xj6lCCvJJ/S6F/2m3LFOcvLLCqyH7ruEcDDGuuhAKLcEhV1TaybXDYJVxBs20/SX
# 1LNJ5hkyZsafWC2M3UOuuYqTzwcF32+S0tRrfbLVjNV/Yl6JNf4ywlIlGvwJdpBw
# LrPHh7iQ7eKXZXzLi9MPTMwQtAoFnG30nihOe5vaL//FY2nf4YB/LfjWzuAtGRix
# /e9BkwXL7ipneQBOnBHaZZaAbmfkNx5Zy5MjMgD1ZdDozIN4Z3iA8v7p9nriXxbk
# FxOZwPabebGKjrHdOCPKHhln4PpqNMFy8c0Gjva9QX9/Rm/+QvhBU6wxn+XJS2vQ
# pd44yyClr2c=
# SIG # End signature block
