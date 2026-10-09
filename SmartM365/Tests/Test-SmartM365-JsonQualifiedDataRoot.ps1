[CmdletBinding()]
param([string]$FixtureRoot = (Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonQualified-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1'
Import-Module $modulePath -MinimumVersion '1.0.6' -Force
$module = Get-Module SmartM365.JsonTransport
$root = [IO.Path]::GetFullPath($FixtureRoot)
if (Test-Path -LiteralPath $root) { throw 'Use a new fixture path; existing data is never accepted.' }
New-Item -ItemType Directory -Path $root | Out-Null
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Reject([scriptblock]$Action, [string]$Message) {
    $failed = $false
    try { & $Action | Out-Null } catch { $failed = $true }
    Check $failed $Message
}
function Guard([string]$Path, [object[]]$Qualifications) {
    & $module { param($p,$q) Assert-SmartM365JsonQualifiedDataPath -Path $p -Qualifications $q } $Path $Qualifications
}
$target = Join-Path $root 'physical-data'
$alias = Join-Path $root 'DATA'
$outside = Join-Path $root 'outside'
New-Item -ItemType Directory -Path $target,$outside | Out-Null
New-Item -ItemType Junction -Path $alias -Target $target | Out-Null
$qualification = @([pscustomobject]@{Root=$alias;Target=$target})
$destination = Join-Path $alias 'state.json.txt'
$validate = { param($value) if ($value.Owner -ne 'Synthetic' -or $value.Id -ne 42) { throw 'Invalid synthetic payload.' } }
$bytes = [Text.Encoding]::UTF8.GetBytes('{"Owner":"Synthetic","Id":42}')
try {
    for ($index=0; $index -lt 16; $index++) {
        $tag = [uint32]2415919130 + [uint32]($index * 4096)
        Check (& $module { param($v) Test-SmartM365JsonCloudTag $v } $tag) 'Documented cloud tag rejected.'
    }
    foreach ($tag in @([uint32]2684354563,[uint32]2684354572,[uint32]2147483681,[uint32]2415919132,[uint32]0)) {
        Check (-not (& $module { param($v) Test-SmartM365JsonCloudTag $v } $tag)) 'Non-cloud tag accepted.'
    }
    Check ((& $module { param($p) (Get-SmartM365JsonPathNode $p).Tag } $alias) -eq [uint32]2684354563) 'Native junction tag inspection failed.'
    Guard $destination $qualification
    Check $true 'Approved alias rejected.'
    Guard (Join-Path $target 'state.json.txt') $qualification
    Check $true 'Approved physical path rejected.'
    Guard (Join-Path $target 'state.json.txt') @([pscustomobject]@{Root=$target;Target=$target})
    Check $true 'Identical alias/target qualification rejected.'
    Reject { Guard $destination @() } 'Unqualified alias accepted.'
    Reject { Guard $destination @([pscustomobject]@{Root=$alias;Target=$outside}) } 'Changed junction target accepted.'
    Reject { Guard $destination @($qualification[0],$qualification[0]) } 'Overlapping qualification accepted.'
    $nested = Join-Path $target 'nested'
    New-Item -ItemType Junction -Path $nested -Target $outside | Out-Null
    Reject { Guard (Join-Path $alias 'nested/blocked.json.txt') $qualification } 'Nested redirect accepted through alias.'
    Reject { Guard (Join-Path $target 'nested/blocked.json.txt') $qualification } 'Nested redirect accepted through physical path.'
    $prefixAlias = Join-Path $root 'DATA-other'
    New-Item -ItemType Junction -Path $prefixAlias -Target $outside | Out-Null
    Reject { Guard (Join-Path $prefixAlias 'blocked.json.txt') $qualification } 'Prefix sibling bypassed root boundary.'
    foreach ($badRoot in @('C:\','relative\DATA','C:\DATA\..\other','C:\DATA*','C:\DATA\file:stream','C:\DATA\bad.','C:\DATA\{{Token}}','\\server\share\DATA')) {
        Reject { & $module { param($p) ConvertTo-SmartM365JsonQualifiedRoot $p } $badRoot } 'Unsafe root syntax accepted.'
    }
    # Inject only metadata tags for cloud cases; real OneDrive data is never written.
    & $module {
        param($p)
        $script:OriginalPathNode = ${function:Get-SmartM365JsonPathNode}
        $script:SyntheticCloudRoot = $p
        $script:SyntheticCloudTag = [uint32]2415919130
        function script:Get-SmartM365JsonPathNode {
            param([string]$Path)
            $node = & $script:OriginalPathNode $Path
            if ($Path -eq $script:SyntheticCloudRoot) {
                $node.Attributes = $node.Attributes -bor [uint32]1024
                $node.Tag = $script:SyntheticCloudTag
            }
            $node
        }
    } $target
    for ($index=0; $index -lt 16; $index++) {
        & $module { param($v) $script:SyntheticCloudTag = $v } ([uint32]2415919130 + [uint32]($index * 4096))
        Guard $destination $qualification
        Check $true 'Cloud-marked physical root rejected.'
    }
    Guard (Join-Path $target 'state.json.txt') @([pscustomobject]@{Root=$target;Target=$target})
    Check $true 'Cloud-marked identical alias/target rejected.'
    Reject { Guard (Join-Path $target 'state.json.txt') @() } 'Unqualified cloud root accepted.'
    & $module { $script:SyntheticCloudTag = [uint32]2415919132 }
    Reject { Guard $destination $qualification } 'Unknown physical-root tag accepted.'
    & $module { Set-Item -Path Function:script:Get-SmartM365JsonPathNode -Value $script:OriginalPathNode }
    # Override only this isolated process's module policy; never write deployed policy/data.
    & $module {
        param($q)
        $script:SyntheticQualifiedRoots = $q
        function script:Get-SmartM365JsonTransportPolicy {
            @{ Mode='JsonText'; QualifiedUncRoots=@(); QualifiedSharePointDrives=@(); QualifiedLocalDataRoots=$script:SyntheticQualifiedRoots }
        }
    } $qualification
    $written = Write-SmartM365JsonBytesAtomically -Path $destination -Bytes $bytes -Validate $validate
    Check ((Get-FileHash -LiteralPath $destination).Hash -eq $written.SHA256) 'Approved junction write hash mismatch.'
    $next = [Text.Encoding]::UTF8.GetBytes('{"Owner":"Synthetic","Id":42,"Next":true}')
    $replaced = Write-SmartM365JsonBytesAtomically -Path $destination -Bytes $next -Validate $validate -ExpectedSHA256 $written.SHA256
    Check ((Read-SmartM365JsonDocument $destination).Document.Next -eq $true) 'Atomic replacement through junction failed.'
    Reject { Write-SmartM365JsonBytesAtomically -Path $destination -Bytes $bytes -Validate $validate -ExpectedSHA256 $written.SHA256 } 'Stale writer accepted.'
    Check ((Get-FileHash -LiteralPath $destination).Hash -eq $replaced.SHA256) 'Stale writer damaged last good state.'
    Reject { Write-SmartM365JsonBytesAtomically -Path $destination -Bytes $bytes -Validate $validate -ExpectedSHA256 'ABSENT' } 'Create-only writer overwrote state.'
    Reject { Write-SmartM365JsonBytesAtomically -Path $destination -Bytes ([Text.Encoding]::UTF8.GetBytes('{')) -Validate $validate } 'Invalid JSON accepted.'
    Check ((Get-FileHash -LiteralPath $destination).Hash -eq $replaced.SHA256) 'Invalid JSON damaged last good state.'
    $changingPolicyValidator = {
        & $module { param($a,$t) $script:SyntheticQualifiedRoots = @([pscustomobject]@{Root=$a;Target=$t}) } $alias $outside
    }
    try {
        Reject { Write-SmartM365JsonBytesAtomically -Path $destination -Bytes $bytes -Validate $changingPolicyValidator } 'Changed qualification after validation accepted.'
        Check ((Get-FileHash -LiteralPath $destination).Hash -eq $replaced.SHA256) 'Post-lock qualification rejection damaged state.'
    } finally {
        & $module { param($q) $script:SyntheticQualifiedRoots = $q } $qualification
    }
    $blocked = Join-Path $alias 'nested/blocked.json.txt'
    Reject { Write-SmartM365JsonBytesAtomically -Path $blocked -Bytes $bytes -Validate $validate } 'Writer followed nested redirect.'
    Check (-not (Test-Path -LiteralPath (Join-Path $outside 'blocked.json.txt'))) 'Rejected writer created an outside target.'
    $lockPath = (Get-SmartM365JsonNames $destination).Lock
    $lock = [IO.File]::Open($lockPath,'OpenOrCreate','ReadWrite','None')
    try { Reject { Write-SmartM365JsonBytesAtomically -Path $destination -Bytes $bytes -Validate $validate -LockTimeoutSeconds 0 } 'Exclusive lock bypassed.' }
    finally { $lock.Dispose() }
    Reject { Complete-SmartM365JsonConsumption -Path $destination -Owner Synthetic -ExpectedSHA256 $replaced.SHA256 } 'Destructive consumption protection relaxed.'
    Reject { Move-SmartM365OwnedJsonFile -Root $alias -RelativePath state.json -Owner Synthetic -Validate $validate } 'Legacy migration protection relaxed.'
    Reject { Write-SmartM365JsonBytesAtomically -Path $destination -Bytes $bytes -Validate $validate -Owner Synthetic -RetireLegacyAfterPublication } 'Legacy retirement protection relaxed.'
    Check ((Get-FileHash -LiteralPath $destination).Hash -eq $replaced.SHA256) 'Rejected deletion/retirement damaged state.'
    [pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Synthetic local filesystem only; no prepared publication, Power BI refresh, collector, tenant or SharePoint access.'}
} catch {
    Write-Warning "Synthetic fixtures retained: $root"
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCP71V5iUEtWLZ+
# GOV+mFGdTZSzxBtlG2NyoL7HRS4nhaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAnNsljPXZyJ2WgIqoQVQZY9upI0G3BwIv53DDTEwSVQMA0GCSqG
# SIb3DQEBAQUABIIBgKZA1ll2qqlnL5TWi5Nfagnmzl6gjZVW54sNRGhQsL1oT/7L
# n+SOjxn8ZqaPm/nk8arhuROkZFcu+6lvlE+fjQ/C1M6xsoU7AEI4+MFrxP4Qgg/I
# A99J9DBVGYgzAJsIYBJlwmrT819jKiRzxsh2Ur6s/11dMijgLJwmvurzIMjm6DJr
# m0SKgAJnbHT5G6VCQc84LVYGoZoEvFKSjueeTQ3SlXGzg5OD0v6xvdfW4Ej0Zoii
# FgunOSVFK2Ij9qzn5A0P894QIKVwZ7ljRm+jP3l1IVUwNlN9KI+c5jKTv8gf8sfG
# jlE0AIjyOQhgicJMiPsxSPW5gSbZfaKO7M3q/tnNO78ztDXq8qz1gShQnkK93SlD
# 4XJgH3mL9hWm5QniPP/TPK5hqXAgYbfBvCAJ+Pi3k61IH4ncDnnuM97MAiougaPG
# zBVeBrhM0INB529kPZx3GsSFvePYseywaQSkotiBpP9fTUFUlmg4+EInDguFCnE2
# jJ2oL/mtZ5ZA27xkY6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkxNjM0
# MzBaMC8GCSqGSIb3DQEJBDEiBCCI2ZkI5/6tfPvhjJwg2EUi7b//wYyqokb+EvU+
# m/gY2zANBgkqhkiG9w0BAQEFAASCAgBFMaDWpqEYQm584Q0IJ2PR++eInIuBv553
# J0gXyJzgj9YG4YBG2pdvwn2ou6OGtWE0k/xqV+irDw7WlolbM7ZKb+q6q8O/z9SJ
# /oDjrYA9/rv8J5YZaJlDKlonWgz8jSAAVAEZxOqyacWpfdT4697VqNq0TCgk96kg
# bt+xcDYzKW1IISbF+08e2VYErsDC/Jv2esdYok0PhbQkldA/L56Fp88gl8iwygai
# sQaJHOrLb99clLHLdM8hrSGiDrho9naPMMtbPFoJGTCAT9fDRvljIp9qT4gYMCdw
# 15RJ5GgTV+0GKklYGfDCDOo1oMw+5zXS31M+YbD9BhEsVv2KX/jnqY0GKyQ6qvQT
# XlJdX+zodyFUXkVq9ObZVI2jOlTJxMejxzgYrfddt/+Duf6nvXJpxtbAZPVfq+04
# l06HLt0+DrgIZL6VIZJsBFu32aPkZNq4Gp+TyqKewiBjVqU+vle8S8Lyunrf81v0
# DtqspSRe1JHAm9RwDW4zgvh1vfusxXlavSdrfdbYWEOmdvdD4/D/i4B8kg+rOclg
# nN+qGlEvZWJTCH39LV7Pv65U6sEO1eKPc7eUthCuL8U4Gv/oXd1aZTRVw9BtkQ0E
# FQ5ITAlNaJn1yX9Da1hmP2fsGzBvJBGim+77Q2HLdGzSJvZwGF21tWwfVbwW7lH5
# qM51XRqvyg==
# SIG # End signature block
