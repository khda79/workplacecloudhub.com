<#
.SYNOPSIS
    Wave computer list helpers.

.DESCRIPTION
    Parses Computers.txt (quotes trimmed, comments ignored, case-insensitive de-duplication),
    computes short keys and moves completed computers to archive lists under an exclusive
    file lock shared by every process and session of the technician computer.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

function ConvertTo-SemtComputerName {
    [CmdletBinding()]
    param([AllowNull()][string]$Line)

    $name = ([string]$Line).Trim().Trim([char]34).Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) { return '' }
    return $name
}

function Get-SemtComputerKey {
    [CmdletBinding()]
    param([AllowNull()][string]$ComputerName)

    $name = ConvertTo-SemtComputerName -Line $ComputerName
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    return ($name.Split('.')[0]).ToUpperInvariant()
}

function Get-SemtComputerList {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Computer list not found: $Path" }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $name = ConvertTo-SemtComputerName -Line $line
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($seen.Add($name)) { $result.Add($name) }
    }
    return @($result.ToArray())
}

function Get-SemtComputerListStats {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Computer list not found: $Path" }
    $names = @(Get-Content -LiteralPath $Path -ErrorAction Stop | ForEach-Object { ConvertTo-SemtComputerName -Line $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $groups = @($names | Group-Object { $_.ToUpperInvariant() } | Where-Object { $_.Count -gt 1 } | Sort-Object -Property @{ Expression = 'Count'; Descending = $true },Name)
    $duplicateLines = 0
    foreach ($group in $groups) { $duplicateLines += ([int]$group.Count - 1) }
    return [pscustomobject]@{
        RawLines = [int]$names.Count
        Unique = [int](@($names | ForEach-Object { $_.ToUpperInvariant() } | Select-Object -Unique).Count)
        DuplicateGroups = [int]$groups.Count
        DuplicateLines = [int]$duplicateLines
        DuplicateSamples = [string](($groups | Select-Object -First 10 | ForEach-Object { $_.Group[0] }) -join ', ')
    }
}

function Test-SemtComputerListPresent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 20)][int]$Attempts = 5,
        [ValidateRange(0, 5000)][int]$DelayMilliseconds = 200
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        if (Test-Path -LiteralPath $Path -PathType Leaf) { return $true }
        if ($attempt -lt $Attempts -and $DelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $DelayMilliseconds }
    }
    return $false
}

function Get-SemtComputerListLockPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ComputerListPath)

    $key = (Get-SemtStringSha256 -Text ([System.IO.Path]::GetFullPath($ComputerListPath).ToUpperInvariant())).Substring(0, 32)
    return (Join-Path -Path (Get-SemtLauncherStateRoot) -ChildPath ("ComputerListLocks\{0}.lock" -f $key))
}

function Invoke-SemtWithComputerListLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )

    return (Invoke-SemtWithFileLock -LockPath (Get-SemtComputerListLockPath -ComputerListPath $ComputerListPath) -TimeoutSeconds 60 -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList)
}

function Write-SemtTextLinesAtomic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Lines
    )

    $temp = '{0}.tmp.{1}.txt' -f $Path,([guid]::NewGuid().ToString('N'))
    try {
        [System.IO.File]::WriteAllLines($temp, [string[]]$Lines, (New-Object System.Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $temp -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $temp -PathType Leaf) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
    }
}

function Remove-SemtComputerListDuplicates {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    return (Invoke-SemtWithComputerListLock -ComputerListPath $Path -ArgumentList @($Path) -ScriptBlock {
        param($ListPath)
        $lines = @(Get-Content -LiteralPath $ListPath -ErrorAction Stop)
        $seen = @{}
        $output = New-Object System.Collections.Generic.List[string]
        $duplicates = 0
        $samples = New-Object System.Collections.Generic.List[string]
        foreach ($line in $lines) {
            $name = ConvertTo-SemtComputerName -Line $line
            if ([string]::IsNullOrWhiteSpace($name)) { $output.Add([string]$line); continue }
            $key = $name.ToUpperInvariant()
            if ($seen.ContainsKey($key)) {
                $duplicates++
                if ($samples.Count -lt 10 -and -not $samples.Contains($seen[$key])) { $samples.Add($seen[$key]) }
                continue
            }
            $seen[$key] = $name
            $output.Add([string]$line)
        }
        if ($duplicates -eq 0) { return [pscustomobject]@{ Changed = $false; DuplicateLines = 0; DuplicateSamples = ''; BackupPath = '' } }
        $backup = '{0}.dedup_{1}.bak' -f $ListPath,(Get-Date -Format 'yyyyMMdd-HHmmss')
        Copy-Item -LiteralPath $ListPath -Destination $backup -Force -ErrorAction Stop
        Write-SemtTextLinesAtomic -Path $ListPath -Lines @($output.ToArray())
        return [pscustomobject]@{ Changed = $true; DuplicateLines = $duplicates; DuplicateSamples = ($samples -join ', '); BackupPath = $backup }
    })
}

function Move-SemtComputersFromList {
    <#
    Moves the named computers from Computers.txt to an archive list in the same Wave folder.
    Both files are rewritten atomically under the shared computer-list lock.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ComputerNames,
        [Parameter(Mandatory = $true)][string]$ArchiveFileName,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $moveKeys = @{}
    foreach ($computer in @($ComputerNames)) {
        $key = Get-SemtComputerKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $moveKeys.ContainsKey($key)) { $moveKeys[$key] = $true }
    }
    if ($moveKeys.Count -eq 0) { return [pscustomobject]@{ Moved = 0; ArchivePath = ''; Detail = "No $Label computer to move." } }

    return (Invoke-SemtWithComputerListLock -ComputerListPath $ComputerListPath -ArgumentList @($ComputerListPath,$moveKeys,$ArchiveFileName,$Label) -ScriptBlock {
        param($ListPath, $Keys, $ArchiveName, $MoveLabel)
        if (-not (Test-Path -LiteralPath $ListPath -PathType Leaf)) { return [pscustomobject]@{ Moved = 0; ArchivePath = ''; Detail = 'Computers.txt is missing.' } }
        $remaining = New-Object System.Collections.Generic.List[string]
        $moved = New-Object System.Collections.Generic.List[string]
        foreach ($line in @(Get-Content -LiteralPath $ListPath -ErrorAction Stop)) {
            $name = ConvertTo-SemtComputerName -Line $line
            if (-not [string]::IsNullOrWhiteSpace($name) -and $Keys.ContainsKey((Get-SemtComputerKey -ComputerName $name))) { $moved.Add($name); continue }
            $remaining.Add([string]$line)
        }
        if ($moved.Count -eq 0) { return [pscustomobject]@{ Moved = 0; ArchivePath = ''; Detail = "$MoveLabel computers were no longer present in Computers.txt." } }
        $archivePath = Join-Path -Path (Split-Path -Path $ListPath -Parent) -ChildPath $ArchiveName
        $archiveLines = New-Object System.Collections.Generic.List[string]
        $existing = @{}
        if (Test-Path -LiteralPath $archivePath -PathType Leaf) {
            foreach ($line in @(Get-Content -LiteralPath $archivePath -ErrorAction Stop)) {
                $archiveLines.Add([string]$line)
                $existingKey = Get-SemtComputerKey -ComputerName $line
                if (-not [string]::IsNullOrWhiteSpace($existingKey)) { $existing[$existingKey] = $true }
            }
        }
        foreach ($name in $moved) {
            $key = Get-SemtComputerKey -ComputerName $name
            if (-not $existing.ContainsKey($key)) { $archiveLines.Add($name); $existing[$key] = $true }
        }
        Write-SemtTextLinesAtomic -Path $archivePath -Lines @($archiveLines.ToArray())
        Write-SemtTextLinesAtomic -Path $ListPath -Lines @($remaining.ToArray())
        return [pscustomobject]@{ Moved = $moved.Count; ArchivePath = $archivePath; Detail = ('Moved {0} {1} computer(s) to {2}.' -f $moved.Count,$MoveLabel,$ArchiveName) }
    })
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDyMwp43XxWxlJR
# M9iSm4mBS0LIu44L8wP5i+MxVX4DDaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDW9HIPC2SPgAyVinO9pqA1GzivzgqoE5wEamfi0HJfWMA0GCSqG
# SIb3DQEBAQUABIIBgD2QquskMT5rKYTKbc45kJc/EUSj0EaAbCmr65TN7vb8PDV9
# 3yqyRyuHc2A3TrykE1mm+hymoy1XG8eOTRVl8ZL1Iv4GFJqjN1nySE1zcXUR8sgy
# oBfNnxfgPrBnzFgTQ7EVHHw5jsxwrCJi4hssasjcDnKNEJEeMrZhv7ffWPtoZfF0
# mF2kS2wq7WdLvOqi2th9tjF5Ifz61wsEMKtYf7qlAWyM90Car1fF102Ij4FQdFrr
# jGT7Dw35+85kqTra3iwYvbPAu64KKzPNJp+sanQKnPLJzPueEN4xZ32fYGOQnW5i
# 0nzHZlAaP4Vb6Otk7VZnkDj2QICxYy6xptq1Jk1a8xeA0BXSC7eu7CaIDrlIfhia
# kDbSD84DbR4wxAbv4ouxJGwziXeHQhHsLOTCJosLsrUYnglipDe+ydmwlMvQPHN3
# 78xeoofMLSDSy0U8Efx1950W+Pu4jgaep8Sq3Pcq7eRWPR5B5qfaytdAs4+n2JSE
# 1LkxdRL4W63oMm+jfKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTFaMC8GCSqGSIb3DQEJBDEiBCCgoFnTitLy99zKpyKQNx/SQ3p4qpV64RISgZRL
# RdULQDANBgkqhkiG9w0BAQEFAASCAgA4wGb/ZXiK5jLOi57SOcX8lkvJprRzI4gl
# VLSZSMFJ8Jw7RtXiBImSv4M04pxpU/c5NF1ud9HWjGCukbOPXafQU1ybPOqNwoZv
# cXiqbjwQ791124Md/u7FOCnAx1DewU51Of4pTlCcMnK5qFj5RsRFcFLguFzbBweA
# q2id7pvd6eUWJ4tOb8OQOJTLdh9XZkEv5UXOyluOFY6b9xqThyI73Fg0eWOrPd17
# f3TPLLST0M7AIa5EPAD+125xBex0/LUdZK67WJzYWnPdA9L37gXUv1aQuJ5VWg8+
# 93oeROlSjyhqTGhNCNWcPRqAGjjVGJ0pStSG2jjt5Re0cwv2KUINT8N/37c6D8eR
# OibPpYD5JSqfWXdQtgMYPjT2g5Nn4++mddurz+LjWLnxdh1J+LYMjbPtXswKl8/t
# 8kyCvSUz+G2wPZhCFmEgm+XoKXNSG+aXUpA8KHkyboragMt3OsIKErOnIepxAWxU
# 8g09U7u/0nAnz8QoRaWmKW7CfJBhELlx5fYnwTNVZ2sR+2Eip6PZ2/w7lQ0C9ebQ
# kB3xA1+TjBgBzaju6KGT125QppjOr7SYFDDRPz4eGHL6rFAX9vlEsg8okSBZDhs9
# vW5wgC+yqX0RJKhu4giq2eIgVH1uP1gczTFDcvCD8bB9q90OaE8BTIiafEpug3Wg
# 0H0o1HWQ0g==
# SIG # End signature block
