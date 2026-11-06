<#
.SYNOPSIS
    Offline tests of the setup media manifest value used to compare two Windows 11 setup media.
.DESCRIPTION
    Extracts the real Get-SetupMediaIntegrityManifestHash function of the Windows11Upgrade
    endpoint script and runs it on synthetic manifests. Checks that two manifests describing the
    same files give the same value whatever their row order, encoding, line ending, path case
    or separator, and that any change of path, length or hash, or an added file, gives another
    value. Also checks the fallback for a manifest that cannot be read as rows, and that the
    fingerprint comparison still uses the manifest value.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$w11Path = Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($w11Path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ('Windows 11 upgrade script does not parse: {0}' -f $parseErrors[0].Message) }
$functionText = @{}
foreach ($name in 'Get-SetupMediaIntegrityManifestHash','Get-SetupMediaIntegrityManifestPath','Get-SmartFileSha256','Open-SmartFileReadStream','Test-SetupFingerprintMatch') {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Function not found in the Windows 11 upgrade script: $name" }
    $functionText[$name] = $definition.Extent.Text
    . ([scriptblock]::Create($definition.Extent.Text))
}
function ConvertTo-LongLiteralPath { param([string]$Path) return $Path }
$script:SetupMediaManifestFileName = 'SmartM365-SetupMediaManifest.sha256.csv'

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SemtMediaFingerprint-{0}' -f [guid]::NewGuid().ToString('N'))
$rows = @(
    [pscustomobject]@{ RelativePath = 'setup.exe'; Length = '102400'; SHA256 = 'AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000AAAA0000' },
    [pscustomobject]@{ RelativePath = 'sources\install.wim'; Length = '5000000000'; SHA256 = 'BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111BBBB1111' },
    [pscustomobject]@{ RelativePath = 'sources\lang.ini'; Length = '120'; SHA256 = 'CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222CCCC2222' }
)

function New-ManifestMedia {
    param([string]$Name, [object[]]$Rows, [string]$Encoding = 'UTF8', [string]$NewLine = "`r`n", [switch]$Raw, [string]$RawText)
    $media = Join-Path $tempRoot $Name
    New-Item -ItemType Directory -Path $media -Force | Out-Null
    $path = Join-Path $media 'SmartM365-SetupMediaManifest.sha256.csv'
    if ($Raw) { [System.IO.File]::WriteAllText($path, $RawText, (New-Object System.Text.UTF8Encoding($false))); return $media }
    $lines = @($Rows | ConvertTo-Csv -NoTypeInformation)
    $text = ($lines -join $NewLine) + $NewLine
    $encodingObject = switch ($Encoding) {
        'UTF8' { New-Object System.Text.UTF8Encoding($true) }
        'UTF8NoBom' { New-Object System.Text.UTF8Encoding($false) }
        'ASCII' { [System.Text.Encoding]::ASCII }
        'Unicode' { [System.Text.Encoding]::Unicode }
    }
    [System.IO.File]::WriteAllText($path, $text, $encodingObject)
    return $media
}

try {
    $reference = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'reference' -Rows $rows)
    Assert-Test 'manifest value is a content value' ($reference -match '^content:[0-9A-F]{64}$')

    $reordered = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'reordered' -Rows @($rows[2],$rows[0],$rows[1]))
    Assert-Test 'other row order: same value' ($reordered -eq $reference)

    $asciiLf = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'ascii-lf' -Rows $rows -Encoding ASCII -NewLine "`n")
    Assert-Test 'ASCII with LF line endings: same value' ($asciiLf -eq $reference)

    $unicode = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'unicode' -Rows $rows -Encoding Unicode)
    Assert-Test 'UTF-16 encoding: same value' ($unicode -eq $reference)

    $caseRows = @($rows | ForEach-Object { [pscustomobject]@{ RelativePath = ('\' + ([string]$_.RelativePath).ToLowerInvariant().Replace('\','/')); Length = (' {0} ' -f $_.Length); SHA256 = ([string]$_.SHA256).ToLowerInvariant() } })
    $caseValue = Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'case' -Rows $caseRows -Encoding UTF8NoBom)
    Assert-Test 'path case, separator, leading slash, spaces and hash case: same value' ($caseValue -eq $reference)

    $changedHash = @($rows | ForEach-Object { $_.PSObject.Copy() }); $changedHash[1].SHA256 = 'DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333DDDD3333'
    Assert-Test 'changed file hash: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'hash' -Rows $changedHash)) -ne $reference)

    $changedLength = @($rows | ForEach-Object { $_.PSObject.Copy() }); $changedLength[0].Length = '102401'
    Assert-Test 'changed file length: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'length' -Rows $changedLength)) -ne $reference)

    $changedPath = @($rows | ForEach-Object { $_.PSObject.Copy() }); $changedPath[2].RelativePath = 'sources\lang2.ini'
    Assert-Test 'changed file path: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'path' -Rows $changedPath)) -ne $reference)

    $extraRows = @($rows) + @([pscustomobject]@{ RelativePath = 'sources\extra.dll'; Length = '1'; SHA256 = 'EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444EEEE4444' })
    Assert-Test 'added file: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'extra' -Rows $extraRows)) -ne $reference)

    $missingRows = @($rows[0],$rows[1])
    Assert-Test 'missing file: other value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot (New-ManifestMedia -Name 'missing' -Rows $missingRows)) -ne $reference)

    $emptyMedia = Join-Path $tempRoot 'no-manifest'; New-Item -ItemType Directory -Path $emptyMedia -Force | Out-Null
    Assert-Test 'no manifest: empty value' ((Get-SetupMediaIntegrityManifestHash -MediaRoot $emptyMedia) -eq '')

    $badMedia = New-ManifestMedia -Name 'bad' -Raw -RawText "Path,Size`r`nsetup.exe,1`r`n"
    $badValue = Get-SetupMediaIntegrityManifestHash -MediaRoot $badMedia
    $badFileHash = (Get-FileHash -LiteralPath (Join-Path $badMedia 'SmartM365-SetupMediaManifest.sha256.csv') -Algorithm SHA256).Hash
    Assert-Test 'manifest without the expected columns: file hash fallback' ($badValue -eq ('file:' + $badFileHash))

    # The fingerprint comparison still includes the manifest value.
    $left = [pscustomobject]@{ MediaId = 'Win11'; ExpectedLanguage = 'fr-FR'; SetupExeLength = 1; SetupExeHash = 'A'; InstallImageName = 'install.wim'; InstallImageLength = 2; LangIniHash = 'B'; IntegrityManifestHash = $reference }
    $right = $left.PSObject.Copy(); $right.IntegrityManifestHash = $reordered
    Assert-Test 'fingerprints of reordered manifests match' (Test-SetupFingerprintMatch -Left $left -Right $right)
    $right.IntegrityManifestHash = 'content:' + ('0' * 64)
    Assert-Test 'fingerprints with another manifest content do not match' (-not (Test-SetupFingerprintMatch -Left $left -Right $right))
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $passed,$failed)
if ($failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBruev9gluCTwP8
# T9Rm8miZL1O0RpecGQhB86AajCsGeKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBviiZui5LG3Hwe2obpV11M0Lg9lo3BPTDbJY0wV4A7qMA0GCSqG
# SIb3DQEBAQUABIIBgCWFqyJs2vfrNfd4woGB760K1PvAdVO0mfSCohTFpBjxrkxJ
# TjlQQeV7y5zQfO11hGObYw+C4u1r0x+BNrv49n4azb3UQ30wgfuzKZBzCxYa2sfZ
# Odm14jUwbmW8HWsGe1UnWJ53jKBdzlwclcaLhp4qAqD7fnJV4NPRyPCtf1gg6hWb
# zz57Gmws/wplK3ZdaMmHGHuSzdwjxw4e9rXZjx0VlSo2YAY0rjOA6qc55UO6ieu1
# r4eAKO251EFME3P/lp7Qt7M6/PxJpD1PPOIRK5Vp4eaYlnSXpq4p2SvztKKaIyHI
# +CahzlJQULqh9444yjzn3CzVy5Kn4fU5N2l9eEBmrWr0MGkaQpwBsTbjxV08jG6F
# 1yFVFce77w/x/9aU5DbCqTynCf9YyNO9KuvzdLxaNPlkvYrRRkTCA4wZzzo9P/RF
# GJpF1ar0r+XDEJouWjgzQWQ/rB3m1UGQLVxy+zRD5qglM/T4czTCoHHHpfKq7Xnc
# K+IK29cMPGdG0vAVxaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYyMTQ3
# NDVaMC8GCSqGSIb3DQEJBDEiBCCnlvP/bLUgNo8M5o8Uv2qmuG4ZlTgTpZYXmCM3
# vFiu1TANBgkqhkiG9w0BAQEFAASCAgCZvxERnaVTDq1omNAwEOjRYHUs7PsoAkrk
# BFW+hGUTBRXYl58DrIQxSpf8OXY+I0UK9H/7hci/CcQMu3WB6UdrwzO7frmAORNy
# 6Ga2yoBpRV++InQFXqTdh/V0rgEKbHDvIieC/HhDvF4aRdpT9WS468uD75HserVg
# 4aN/XjsKn1eaHEniu83uz6BwUHBwCDdj+h4COuln0e26W1QTwpIA+stYQJSUGfI3
# FF6JXD1rDoxJ296v/Szc4PncvaDgPMQGo65WLoJ2Qu0eDLOw//1cTOq0N7VXTUmq
# 0jK7F2XkWLeiVXf+z1EBPyC8jHW/Mw297z6hrD7dgE18itdb1B9J4XbbuCVN6VZQ
# Aym/q7oCImSKNz0LqI4ClHvzVHGxmd0SZbgiErE5iCQbdAQRSXlJRgJ7yeix9hbr
# PYi5SVnfKG4ScPeDx6p/pCXEtQdUdcXGIIQBYuJ336ktm3Al4Vvy9/TbUaXoMAdU
# HEZ0X8HWGJ2UrLo6nXAC0Fohk4yKg8SIjQu4vpJPuxyk/5hzpuR48txR/pQwgE29
# zFad1P9HauRt5RiSS3GvzMYukwsyVRynRl1z2piu3oV6Pvt6fAs0y6l8mp1S/fRW
# F29qRCzdkNOY1qiX2u6GkbfADYQKnJ6RUZjt3f1Tg2TEYUrQqxwTzG/VMKldzJC1
# NNli3gzbQA==
# SIG # End signature block
