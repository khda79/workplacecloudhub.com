<#
.SYNOPSIS
Offline check that atomic writes stay under MAX_PATH (260) for deep targets, as on Windows PowerShell 5.1 /
.NET Framework without long path support (launcher cache copies of the repository): every file the JSON
transport, the Core and the Windows PowerShell 5.1 module create while writing a target whose full path
is 245 characters long must itself stay below 260 characters.
.VERSION
1.0.0
.NOTES
Runs in the current host and, from PowerShell 7, again in Windows PowerShell 5.1 with the 5.1 module.
Long paths may be enabled on the test machine, so file creation is observed rather than expected to fail.
Synthetic TEMP files only; no tenant, network or synchronized data access.
#>
[CmdletBinding()]
param([switch]$InnerWindowsPowerShell)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$coreFolder = Join-Path $root 'Modules/SmartM365.Core'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

Import-Module (Join-Path $coreFolder 'SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.8' -Global -Force -ErrorAction Stop
$isWindowsPowerShell = $PSVersionTable.PSEdition -ne 'Core'
if ($isWindowsPowerShell) {
    Import-Module (Join-Path $coreFolder 'Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psd1') -MinimumVersion '1.0.63' -Global -Force -WarningAction SilentlyContinue -ErrorAction Stop
}
else {
    Import-Module (Join-Path $coreFolder 'SmartM365.Core.psd1') -MinimumVersion '1.0.102' -Global -Force -WarningAction SilentlyContinue -ErrorAction Stop
}

# A folder deep enough for a 245-character target named like the Exchange infrastructure local configuration.
$leaf = 'SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.local.json.txt'
$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-maxpath-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$folder = $base
while (($folder.Length + 1 + $leaf.Length) -lt 230) { $folder = Join-Path $folder 'SmartInventory' }
$folder = Join-Path $folder ('p' * (245 - $folder.Length - 2 - $leaf.Length))
$target = Join-Path $folder $leaf
[void][IO.Directory]::CreateDirectory($folder)
Check ($target.Length -eq 245) "Synthetic target length is $($target.Length), not 245."

$sourceId = 'SmartM365ShortTemporaryPaths-' + [guid]::NewGuid().ToString('N')
$watcher = New-Object IO.FileSystemWatcher $folder
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true
$null = Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier $sourceId
try {
    # Helper: same folder, shorter than the target whenever the target name is longer than the short name.
    $temporary = Get-SmartM365SiblingTemporaryPath -Path $target -Extension '.pending'
    Check ([IO.Path]::GetDirectoryName($temporary) -eq $folder -and $temporary.Length -lt $target.Length -and $temporary -like '*.pending') "Sibling temporary path is not short or not beside the target: $temporary"

    # JSON transport: create then replace (temporary and previous copies).
    $json = [Text.Encoding]::UTF8.GetBytes('{"Name":"synthetic","Version":1}')
    Write-SmartM365JsonBytesAtomically -Path $target -Bytes $json -Validate { param($document) $null = $document } -ExpectedSHA256 'ABSENT' | Out-Null
    $json2 = [Text.Encoding]::UTF8.GetBytes('{"Name":"synthetic","Version":2}')
    $hash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    Write-SmartM365JsonBytesAtomically -Path $target -Bytes $json2 -Validate { param($document) $null = $document } -ExpectedSHA256 $hash | Out-Null
    Check ([IO.File]::ReadAllText($target) -like '*"Version":2*') 'The JSON transport did not replace the deep target.'

    # Core (PowerShell 7) or the Windows PowerShell 5.1 module: atomic copy and text write next to a deep target.
    $copyTarget = Join-Path $folder ($leaf -replace '\.local\.json\.txt$', '_Copy.csv')
    $source = Join-Path $base 'source.csv'
    [IO.File]::WriteAllText($source, "Id`r`n1`r`n")
    Copy-SmartM365FileAtomically -SourcePath $source -DestinationPath $copyTarget | Out-Null
    Write-SmartM365TextAtomically -Path $copyTarget -Content "Id`r`n2`r`n"
    Check ([IO.File]::ReadAllText($copyTarget) -like '*2*') 'The atomic text write did not replace the deep target.'

    Start-Sleep -Milliseconds 800
    $created = @(Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.SourceEventArgs.FullPath })
    $tooLong = @($created | Where-Object { $_.Length -ge 260 })
    Check ($created.Count -ge 4) "Too few file creations observed ($($created.Count))."
    Check ($tooLong.Count -eq 0) ("Files created at or beyond MAX_PATH: " + (($tooLong | ForEach-Object { "$($_.Length) $([IO.Path]::GetFileName($_))" }) -join '; '))

    # Launcher case: the cache is rebuilt from the share on every run, so a script local configuration missing
    # on the share is created again from its template, then must not be rewritten while no key is missing.
    Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | Remove-Event
    . (Join-Path $root 'Config/SmartM365-TenantContext.ps1') | Out-Null
    Remove-Item -LiteralPath $target -Force
    $templatePath = $target + '.template'
    [IO.File]::WriteAllText($templatePath, '{"Alpha":"__USE_GLOBAL__","HealthCheckerTimeoutMinutesPerServer":30,"ExcludedServers":[],"Nested":{"Value":1}}')
    $legacyPath = $target -replace '\.txt$', ''
    $config = Read-SmartM365JsonConfig -Path $legacyPath -Required 6>$null
    Check ((Test-Path -LiteralPath $target) -and $config['HealthCheckerTimeoutMinutesPerServer'] -eq 30) 'The deep local configuration was not created from its template.'
    $createdHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    $null = Read-SmartM365JsonConfig -Path $legacyPath -Required 6>$null
    Check ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq $createdHash) 'A complete local configuration was rewritten.'
    # Same keys, another order and other values (an edit made with ConvertTo-Json): still no rewrite.
    $edited = [Text.Encoding]::UTF8.GetBytes('{"Nested":{"Value":2},"ExcludedServers":["Server1"],"HealthCheckerTimeoutMinutesPerServer":45,"Alpha":"x"}')
    [IO.File]::WriteAllBytes($target, $edited)
    $editedHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    $config = Read-SmartM365JsonConfig -Path $legacyPath -Required 6>$null
    Check ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq $editedHash -and $config['HealthCheckerTimeoutMinutesPerServer'] -eq 45) 'A reordered local configuration with every template key was rewritten.'
    Start-Sleep -Milliseconds 800
    $created = @(Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.SourceEventArgs.FullPath })
    $tooLong = @($created | Where-Object { $_.Length -ge 260 })
    Check ($tooLong.Count -eq 0) ("Configuration files created at or beyond MAX_PATH: " + (($tooLong | ForEach-Object { "$($_.Length) $([IO.Path]::GetFileName($_))" }) -join '; '))
}
finally {
    Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue
    Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue | Remove-Event
    $watcher.Dispose()
    if (Test-Path -LiteralPath $base) { cmd /c rd /s /q "$base" | Out-Null }
}

# Static: no temporary name built from the full target name plus a GUID in the Core files.
$files = @('SmartM365.JsonTransport.psm1', 'SmartM365.Core.psm1', 'SmartM365-CmdbReceipt.ps1', 'SmartM365-DailyCaptureRetention.ps1', 'Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1') | ForEach-Object { Join-Path $coreFolder $_ }
$files += Join-Path $root 'Config/SmartM365-TenantContext.ps1'
foreach ($file in $files) {
    $long = @(Select-String -LiteralPath $file -Pattern "\[guid\]::NewGuid\(\)[^\r\n]*\.(tmp|pending|previous|partial|download)\b" | Where-Object { $_.Line -notmatch 'GetTempPath|~sm' })
    Check ($long.Count -eq 0) ("Long temporary name in {0}: line {1}" -f [IO.Path]::GetFileName($file), (($long | ForEach-Object { $_.LineNumber }) -join ','))
}

if (-not $isWindowsPowerShell -and -not $InnerWindowsPowerShell) {
    # The production case: Windows PowerShell 5.1 with the 5.1 module.
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $windowsPowerShell) {
        $output = & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -InnerWindowsPowerShell 2>&1 | Out-String
        Check ($LASTEXITCODE -eq 0 -and $output -match 'PASS:') ("Windows PowerShell 5.1 run failed: " + $output.Trim())
        Write-Output $output.Trim()
    }
}
"PASS: $script:checks short temporary path checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAEJIbZCLFTNjTl
# JJu3fZxne6rO6xgNSHssv3WJY6LguKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIKmAiX5UnQs2nscaHRw8gSt8cK/nCeVrg0wpDINvqXY6MA0GCSqG
# SIb3DQEBAQUABIIBgAjDOiMC859ybi0BapF7pVSmYrtSXgtPMuvK+vvk9CEosm2v
# dJJClcVyPRjUavZr6Pt4O22qr0jqcgse3nk2RxIhX3l/58WRK5kghFytkRjbb32Q
# HzuGvYHOXJBJNo/QXoL4Gqnt5j4bFf0j4Csa8WtbFqYM9luAGZJCNzMojYFaWW6E
# p1g9BWBMd4ZUIbWLQMEMiUgCJn+GcB97Uva9+jd+NMx7//5TttqlgCylcj0HLB97
# VdqVaeZuRlVgRt8jOeCs/VNP4kOtn7/YcnlziG7fkmQ+fMEXnk0yTMKpnwEgRh1H
# hLCbAgIGR6zBHkQPg05M5KTSiVghFvbCxC5C2woQhOIqNK3Jb9g7G6VVxcDYH63g
# gXoR925x+jZx9/QVcMPAQwlhyE037nJkez7r0xUxTnL06EdnSo/zIKEgrSXkbGOA
# OLdaCU4r4FcpBXikE7HmJMiOBlyYDBYZHp23Kd8rAMxe43D7/bA5SmNw+SbWZQ8B
# Fr1DTbgxz/R27lzr36GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNzE2
# MjRaMC8GCSqGSIb3DQEJBDEiBCDlw8eaBLnHMt7MFGGDWE+5Xuimp0syF1KT9Soi
# hkIAODANBgkqhkiG9w0BAQEFAASCAgBtA6KC2tWBXZBZg0WVf+gcNGWWb8SQ7lZg
# 6NiceAnrs6yTL+x7/Ew+YDUr97K2Kikjmxsab4bwQ1DnruJ13in+Y2zsryF0PRMY
# O537uELMgIXNSueXf+IU6TPqh9sIHyYsmrZFvxGO1zHvxlh61lyKR7V7giqSiZe3
# 7Ef0DQmUHX0MPmzmQ7f3vpHZWzIPk9S0nhPoYXRTILUXNsGvp8x+Xr48yyTIfRQb
# iaFTRrLaTetWBKtDYyzJSuQXk4dcp0bw/9tx1opSJ3qXwVQsl1pKUz3cSeXvpwlD
# uop69SlV9KbBT5PEWzeK7JNiBn+o3zevZSR/z2lfZ+s1vIuh2maBGMVIABsxjn+o
# NxwuwWnAY3jeT2Hn/x2EJpda6ljNqqdPa/9cPyefO4dDR32TYYHHerbgzwEPm7Fa
# 1a9WOliSHRml/3rMd30gRJZ3mdquL6Yy4Guix642FbLiA6IKhlM7iBnkPiNaFBjp
# T/+x1cvgrWhTYSlC/9yd4JKWaZ+Fz0DxX8pow0ESkdMzvnS2GByqZcTsiGhgT5mj
# 6V5qJUNwy3dw0MI6sWQB24I9vBRlh+roWj7wlESeu5KVTxBmFo1CuWBdHSrcsQl+
# SEmi9NKRPBPAwW8qTKx72kQeqEPEmbrll+PTkOolyedBWDRnI5YaVFg8qw1Hh0yo
# jygV0NUAwg==
# SIG # End signature block
