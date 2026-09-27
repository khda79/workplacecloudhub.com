[CmdletBinding()]
param([string]$FixtureRoot = (Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonTransport-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psm1') -Force
$root = [IO.Path]::GetFullPath($FixtureRoot)
if (Test-Path -LiteralPath $root) { throw 'Use a new empty fixture path. Existing data is never accepted.' }
New-Item -ItemType Directory -Path $root | Out-Null
$script:passed = 0
$validate = { param($document) if ($document.Owner -ne 'Synthetic' -or $document.Id -ne 42) { throw 'Invalid synthetic identity.' } }
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Reject([scriptblock]$Action, [string]$Message) { $failed = $false; try { & $Action | Out-Null } catch { $failed = $true }; Check $failed $Message }
function Fixture([string]$Name) {
    $folder = Join-Path $root $Name
    New-Item -ItemType Directory -Path $folder | Out-Null
    $old = Join-Path $folder 'state.json'
    # Preserve CRLF, whitespace, BOM and non-ASCII bytes, not only parsed values.
    [IO.File]::WriteAllText($old, "{`r`n  `"Owner`": `"Synthetic`", `"Id`": 42, `"Text`": `"caf$([char]0x00e9)`"`r`n}`r`n", [Text.UTF8Encoding]::new($true))
    return $folder
}
try {
    $one = Fixture 'old'
    $hash = (Get-FileHash (Join-Path $one 'state.json')).Hash
    Check ((Read-SmartM365JsonDocument (Join-Path $one 'state.json') -Validate $validate).Document.Id -eq 42) 'Legacy read failed.'
    $result = Move-SmartM365OwnedJsonFile -Root $one -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate
    Check ($result.Status -eq 'Completed' -and $result.SHA256 -eq $hash) 'Migration changed hash.'
    Check (-not (Test-Path (Join-Path $one 'state.json'))) 'Legacy remains after rename.'
    Check ((Read-SmartM365JsonDocument (Join-Path $one 'state.json') -Validate $validate).SHA256 -eq $hash) 'Preferred read failed.'
    Check ((Move-SmartM365OwnedJsonFile -Root $one -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate).Status -eq 'Completed') 'Replay failed.'
    $both = Fixture 'identical'
    Copy-Item (Join-Path $both 'state.json') (Join-Path $both 'state.json.txt')
    Check ((Read-SmartM365JsonDocument (Join-Path $both 'state.json')).Coexistence -eq 'Identical') 'Identical pair not detected.'
    Check ((Move-SmartM365OwnedJsonFile -Root $both -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate).Status -eq 'PendingLegacy') 'Unconfirmed duplicate removed.'
    Check (Test-Path (Join-Path $both 'state.json')) 'Pending legacy lost.'
    Move-SmartM365OwnedJsonFile -Root $both -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate -RemoveIdenticalLegacy | Out-Null
    Check (-not (Test-Path (Join-Path $both 'state.json'))) 'Confirmed identical pair not deduplicated.'
    $conflict = Fixture 'conflict'
    [IO.File]::WriteAllText((Join-Path $conflict 'state.json.txt'), '{"Owner":"Synthetic","Id":43}')
    Reject { Read-SmartM365JsonDocument (Join-Path $conflict 'state.json') } 'Conflicting payload accepted.'
    Reject { Move-SmartM365OwnedJsonFile -Root $conflict -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate } 'Conflicting pair migrated.'
    Check ((Get-FileHash (Join-Path $conflict 'state.json')).Hash -eq $hash) 'Conflict damaged legacy.'
    foreach ($bad in @('', '{', 'null', '{"Owner":"Wrong","Id":42}')) {
        $invalid = Fixture ('invalid-' + [guid]::NewGuid().ToString('N'))
        [IO.File]::WriteAllText((Join-Path $invalid 'state.json.txt'), $bad)
        Reject { Read-SmartM365JsonDocument (Join-Path $invalid 'state.json') -Validate $validate } 'Invalid preferred fell back to legacy.'
    }
    $directory = Fixture 'directory'
    New-Item -ItemType Directory (Join-Path $directory 'state.json.txt') | Out-Null
    Reject { Get-SmartM365JsonReadPath (Join-Path $directory 'state.json') } 'Directory at preferred name enabled fallback.'
    Check ($null -eq (Get-SmartM365JsonReadPath (Join-Path $root 'absent.json') -Optional)) 'Missing optional file should be absent.'
    Reject { Read-SmartM365JsonDocument (Join-Path $root 'absent.json') } 'Missing required file accepted.'
    Reject { Move-SmartM365OwnedJsonFile -Root $one -RelativePath '../outside.json' -Owner 'Synthetic' -Validate $validate } 'Owner traversal accepted.'
    Reject { Move-SmartM365OwnedJsonFile -Root '\\synthetic.invalid\never-access' -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate } 'Unqualified UNC accepted.'
    $locked = Fixture 'locked'
    $stream = [IO.File]::Open((Join-Path $locked 'state.json.transport.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    try {
        Reject { Move-SmartM365OwnedJsonFile -Root $locked -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate -LockTimeoutSeconds 0 } 'Lock bypassed.'
        Check (Test-Path (Join-Path $locked 'state.json')) 'Locked source changed.'
    } finally { $stream.Dispose() }
    Check ((Move-SmartM365OwnedJsonFile -Root $locked -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate).Status -eq 'Completed') 'Lock release did not permit retry.'
    # Simulate a process interruption after atomic rename, before the completion event.
    $resumed = Fixture 'interrupted'
    [IO.File]::WriteAllText((Join-Path $resumed 'state.json.migration.log'), '{"Owner":"Synthetic","Phase":"Prepared"}' + [Environment]::NewLine + '{"Phase":"Comp')
    [IO.File]::Move((Join-Path $resumed 'state.json'), (Join-Path $resumed 'state.json.txt'))
    Check ((Move-SmartM365OwnedJsonFile -Root $resumed -RelativePath 'state.json' -Owner 'Synthetic' -Validate $validate).SHA256 -eq $hash) 'Interrupted rename not recoverable.'
    Check ((Get-Content (Join-Path $resumed 'state.json.migration.log') -Tail 1 | ConvertFrom-Json).Phase -eq 'Completed') 'Recovery journal incomplete.'
    $journalHash=(Get-FileHash (Join-Path $resumed 'state.json.migration.log')).Hash
    Move-SmartM365OwnedJsonFile -Root $resumed -RelativePath state.json -Owner Synthetic -Validate $validate | Out-Null
    Check ((Get-FileHash (Join-Path $resumed 'state.json.migration.log')).Hash -eq $journalHash) 'Idempotent replay grew a completed journal.'
    $writePath = Join-Path $one 'state.json.txt'
    $nextBytes = [Text.UTF8Encoding]::new($false).GetBytes('{"Owner":"Synthetic","Id":42,"Next":true}')
    $write = Write-SmartM365JsonBytesAtomically -Path $writePath -Bytes $nextBytes -Validate $validate
    Check ((Get-FileHash $writePath).Hash -eq $write.SHA256) 'Atomic write hash mismatch.'
    Check ((Read-SmartM365JsonDocument $writePath).Document.Next -eq $true) 'Atomic write payload missing.'
    Reject { Write-SmartM365JsonBytesAtomically -Path $writePath -Bytes ([Text.Encoding]::UTF8.GetBytes('{')) -Validate $validate } 'Invalid replacement accepted.'
    Check ((Get-FileHash $writePath).Hash -eq $write.SHA256) 'Invalid replacement destroyed last valid state.'
    Reject { Write-SmartM365JsonBytesAtomically -Path $writePath -Bytes $nextBytes -Validate $validate -ExpectedSHA256 $hash } 'Stale writer overwrote newer state.'
    Reject { Write-SmartM365JsonBytesAtomically -Path $writePath -Bytes $nextBytes -Validate $validate -ExpectedSHA256 'ABSENT' } 'Create-only writer overwrote existing state.'
    $arrayPath = Join-Path $root 'assignments.json.txt'
    Write-SmartM365JsonBytesAtomically -Path $arrayPath -Bytes ([Text.Encoding]::UTF8.GetBytes('[]')) -Validate { param($value) if (@($value).Count -ne 0) { throw 'Expected empty assignments.' } } | Out-Null
    Check ((Get-Content $arrayPath -Raw) -eq '[]') 'Valid empty array damaged.'
    [pscustomobject]@{ Passed = $script:passed; FixtureRoot = $root; Evidence = 'Synthetic local filesystem only; no collector, tenant, SharePoint or real UNC access.' }
} catch {
    Write-Warning "Fixtures retained for diagnosis: $root"
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCABTVqrBGXvqtPr
# DbH4hGM3HZ6B8hY3w6QEwAyYuCQ1RKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL9fhLVimYiPGOCKrz0GQ5i/8iOrKCrDABZMNmgjzFv9MA0GCSqG
# SIb3DQEBAQUABIIBgFnTr2plW6j5wS0obCIYPZPcTJJUCW/PfnIaGnMaM29RusVY
# dtFhGE0E9rconYQSPll6VfWWgL2pb3EqfWrAb3LdKgOGwnEswsUFVjbKxvwxQaZL
# Vj2RMCvup/ZKtmU42UcjqjfM+IMN7rqWqE41miyXqsJpg8WGxWcKB2TKbNpCaq2I
# hZGzhDlq24A0OTmjC84b7ymWaOeXOvvpAXQe4Rky9PBgBQ3GgwMhGMGKoWZXKPs7
# 1PZCFSwN8rOwmeZ9rcoF4rzUMJOzvowFjxTPzAcJlRimN148aKOws3w5twWmB3FE
# d86VxuA8BfIWV471RKLGhDqmZQ4h6vv6/+wReH7zHdXR1EH+1dL+2BgTEVkhuE24
# p2MzxsYsmLxV9AyUF7BxpAgu1wqa5fBt35QlVTsflGL/ErZnWin0FuMziaUT1OJ4
# OF2oXrqsml9rb0kTq+fu7x+x3n9mZfH1D95N6nfQX+McbK4v3njPkjCwUAraTQTI
# bZlfb0H35iOwtsxkPaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU3
# NTdaMC8GCSqGSIb3DQEJBDEiBCBNmykBtaP4I9b/H5O00kdRXIonsFcG7BhKdOH9
# vL/FJDANBgkqhkiG9w0BAQEFAASCAgCG58UATjJM/u9+leeD0LHOZo3NFkH1KWAW
# t21ePxUCtE258vF6Mixu5yXFCcXZ7XmFSc7XNPEoNReos3IesQKnTa3hU4d2/jh/
# kxAUaVzi0lbOhSJE98Fp0XNhZmKVhNx1ObJxcdv1TX0yhbfwu0/w4cZOWTqNHofA
# vgP6Wbe34qWl1qctqRf3SaQELdtOTLP8Qgoc1yYn8X0VhUwsobHhol+bNgqD5/Rk
# SG43XuxPBrnoh4pv3PU2qDjVRNkjDDMj1uumB1cQwGOpy9617x+BzXjYJItb0U+Q
# kSmWaEd1mhUKilktaaGb8405YDhLCz2fPk1kguAK2MklQJydXyjrZ9+Q1GZ08Yp3
# Udsv6tLQtSkejg7a8Ovb+NMSzO8IidnIKdB0/IJSRqdWZtggxPsZt7CshsDYxj39
# LA1uyaXE8p+YyaOM+18LHZcSj1Cp/dehKTXRrPLwltdJCdnO6GnEnOPMkvKX6gkt
# vbfKYH4XJHERZfnD4MAcQfKmPw3Tauhz/ShMisuZpPcDq9CyfGa6ktiemw6qQCuY
# N1MIlkDjtMZHMfcurmJDNSaN/fmEya5eJrlyPNU0vzSzJG3bYATi3STegIZ5728k
# Sv3eYVUs+LQMtd5w5AoUbpLS6WYqokd4SM6yF8Btj+CHgWhYDvwTgI38Q+zrlWjM
# cJ0YXdfcPg==
# SIG # End signature block
