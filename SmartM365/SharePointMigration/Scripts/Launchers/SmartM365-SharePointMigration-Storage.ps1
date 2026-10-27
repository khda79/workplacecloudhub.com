<#
.SYNOPSIS
    Resolves the SharePointMigration data roots (Migrations, ArchivedMigrations, shared logs,
    Entra users cache).
.DESCRIPTION
    Config/storage.json.txt may point the data roots to a shared folder. Empty values keep
    the historical layout inside the toolkit: Migrations, ArchivedMigrations and
    Migrations\logs, _Local\EntraUsersCache. The runtime file is created from, and completed with, its template.
    Windows PowerShell 5.1 compatible; no dependency on other toolkit helpers.
.VERSION
    1.0.1
#>

function Initialize-SmartM365MigrationStorageConfig {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)

    $configRoot = Join-Path -Path $ProjectRoot -ChildPath 'Config'
    $template = Join-Path -Path $configRoot -ChildPath 'storage.json.template'
    $runtime = Join-Path -Path $configRoot -ChildPath 'storage.json.txt'
    if (-not (Test-Path -LiteralPath $template -PathType Leaf)) { return $runtime }
    try {
        if (-not (Test-Path -LiteralPath $runtime -PathType Leaf)) {
            Copy-Item -LiteralPath $template -Destination $runtime -ErrorAction Stop
            return $runtime
        }
        $templateData = Get-Content -LiteralPath $template -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $runtimeData = Get-Content -LiteralPath $runtime -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $added = @()
        foreach ($property in $templateData.PSObject.Properties) {
            if (-not $runtimeData.PSObject.Properties[$property.Name]) {
                $runtimeData | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
                $added += $property.Name
            }
        }
        if ($added.Count) {
            $json = $runtimeData | ConvertTo-Json -Depth 4
            [System.IO.File]::WriteAllText($runtime, $json + "`r`n", [System.Text.UTF8Encoding]::new($false))
            Write-Verbose ("storage.json.txt: added keys {0}" -f ($added -join ', '))
        }
    }
    catch {
        # A read-only toolkit copy must still run; an unreadable runtime file fails in the resolver.
        Write-Verbose ("storage.json.txt self-heal skipped: {0}" -f $_.Exception.Message)
    }
    return $runtime
}

function Resolve-SmartM365MigrationStoragePath {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory = $true)][string]$Default
    )

    $text = if ($null -eq $Value) { '' } else { $Value.Trim() }
    if (-not $text) { $text = $Default }
    $text = [Environment]::ExpandEnvironmentVariables($text)
    if (-not [System.IO.Path]::IsPathRooted($text)) { $text = Join-Path -Path $ProjectRoot -ChildPath $text }
    $full = [System.IO.Path]::GetFullPath($text)
    if ($full.Length -gt 3 -and $full.EndsWith('\')) { $full = $full.TrimEnd('\') }
    return $full
}

function Get-SmartM365MigrationStorage {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)

    $root = [System.IO.Path]::GetFullPath($ProjectRoot).TrimEnd('\')
    $configPath = Initialize-SmartM365MigrationStorageConfig -ProjectRoot $root
    $stamp = ''
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        $stamp = (Get-Item -LiteralPath $configPath).LastWriteTimeUtc.Ticks.ToString()
    }
    $cacheKey = $root + '|' + $stamp
    $cache = Get-Variable -Name SmartM365MigrationStorageCache -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($cache -and $cache.Key -eq $cacheKey) { return $cache.Value }

    $data = $null
    if ($stamp) {
        try { $data = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
        catch { throw "Invalid SharePointMigration storage configuration '$configPath': $($_.Exception.Message)" }
    }
    $values = @{}
    foreach ($name in @('MigrationsRoot', 'ArchivedMigrationsRoot', 'LogsRoot', 'EntraUsersCacheRoot')) {
        $values[$name] = ''
        if ($data -and $data.PSObject.Properties[$name] -and $null -ne $data.$name) { $values[$name] = [string]$data.$name }
    }
    $migrationsRoot = Resolve-SmartM365MigrationStoragePath -ProjectRoot $root -Value $values.MigrationsRoot -Default 'Migrations'
    $archivedRoot = Resolve-SmartM365MigrationStoragePath -ProjectRoot $root -Value $values.ArchivedMigrationsRoot -Default 'ArchivedMigrations'
    $logsRoot = Resolve-SmartM365MigrationStoragePath -ProjectRoot $root -Value $values.LogsRoot -Default (Join-Path $migrationsRoot 'logs')
    $entraRoot = Resolve-SmartM365MigrationStoragePath -ProjectRoot $root -Value $values.EntraUsersCacheRoot -Default '_Local\EntraUsersCache'
    $storage = [pscustomobject]@{
        ProjectRoot            = $root
        ConfigPath             = $configPath
        MigrationsRoot         = $migrationsRoot
        ArchivedMigrationsRoot = $archivedRoot
        LogsRoot               = $logsRoot
        EntraUsersCacheRoot    = $entraRoot
        IsDefault              = -not ($values.MigrationsRoot.Trim() -or $values.ArchivedMigrationsRoot.Trim() -or $values.LogsRoot.Trim() -or $values.EntraUsersCacheRoot.Trim())
    }
    $script:SmartM365MigrationStorageCache = @{ Key = $cacheKey; Value = $storage }
    return $storage
}

function Get-SmartM365MigrationsRoot {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)
    return (Get-SmartM365MigrationStorage -ProjectRoot $ProjectRoot).MigrationsRoot
}

function Get-SmartM365MigrationLogsRoot {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot, [string]$ChildPath = '')
    $logsRoot = (Get-SmartM365MigrationStorage -ProjectRoot $ProjectRoot).LogsRoot
    if ($ChildPath) { return (Join-Path -Path $logsRoot -ChildPath $ChildPath) }
    return $logsRoot
}

function Get-SmartM365DirectoryFingerprint {
    param([Parameter(Mandatory = $true)][string]$Path)
    $files = @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction Stop)
    $bytes = [long]0
    foreach ($file in $files) { $bytes += $file.Length }
    [pscustomobject]@{ Files = $files.Count; Bytes = $bytes }
}

# Directory.Move only works inside one volume; shared roots can sit on different shares.
function Move-SmartM365MigrationDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )
    if (Test-Path -LiteralPath $Destination) { throw "Destination already exists: $Destination" }
    $sameVolume = [string]::Equals([System.IO.Path]::GetPathRoot($Source), [System.IO.Path]::GetPathRoot($Destination),
        [System.StringComparison]::OrdinalIgnoreCase)
    if ($sameVolume) {
        try { [System.IO.Directory]::Move($Source, $Destination); return }
        catch [System.IO.IOException] {
            # Mapped drives or DFS links can share a root text and still be distinct volumes.
            if (Test-Path -LiteralPath $Destination) { throw }
        }
    }
    $expected = Get-SmartM365DirectoryFingerprint -Path $Source
    $staging = Join-Path -Path (Split-Path -Path $Destination -Parent) -ChildPath ('.moving-{0}-{1}' -f
        (Split-Path -Path $Destination -Leaf), [guid]::NewGuid().ToString('N'))
    try {
        Copy-Item -LiteralPath $Source -Destination $staging -Recurse -Force -ErrorAction Stop
        $copied = Get-SmartM365DirectoryFingerprint -Path $staging
        if ($copied.Files -ne $expected.Files -or $copied.Bytes -ne $expected.Bytes) {
            throw ("Copy verification failed: {0} files / {1} bytes expected, {2} / {3} copied." -f
                $expected.Files, $expected.Bytes, $copied.Files, $copied.Bytes)
        }
        Rename-Item -LiteralPath $staging -NewName (Split-Path -Path $Destination -Leaf) -ErrorAction Stop
    }
    catch {
        if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
        throw
    }
    try { Remove-Item -LiteralPath $Source -Recurse -Force -ErrorAction Stop }
    catch { throw "The folder was copied to '$Destination' but the source could not be removed: $($_.Exception.Message)" }
}

# Per-migration .cmd launchers read this pointer when Migrations lives outside the toolkit.
function Set-SmartM365MigrationToolkitPointer {
    param(
        [Parameter(Mandatory = $true)][string]$MigrationsRoot,
        [Parameter(Mandatory = $true)][string]$ToolkitRoot
    )
    if ($ToolkitRoot -match '[^\x20-\x7E]') { throw "The toolkit path must be ASCII for cmd launchers: $ToolkitRoot" }
    $path = Join-Path -Path $MigrationsRoot -ChildPath 'toolkit.root.txt'
    [System.IO.File]::WriteAllText($path, $ToolkitRoot.TrimEnd('\') + "`r`n", [System.Text.Encoding]::ASCII)
    return $path
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD2lnqoF6AE9Z4C
# rE3ZDZ2oT9zZrxISf+Cgp7tu/gWMKaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJ7eJjAFkSjASx989FQJ3oP1MCIA5eTazg0OhKxhqSNPMA0GCSqG
# SIb3DQEBAQUABIIBgBiUpnN0FLIa0pJD0vcR/RNvkFXGHGaJuNd3eY7mMLsPMrHs
# 3ht0VTTVJDqVDyCPrrVOg9dVmEqUZJZEWy2rLO5wdPHuvvHkAw0h9ArspQEhvNxq
# mfwKrLFbZB5OBsrWd3mzRn/tI+RIPjLmNnyH4uC6uwQVQlu3D/PGvGyrc0ghlG5Q
# gWpykGLEpjBmjegV7QwVg8UPYvzgIYa9kq01FcYO8+B9nBSY0BjhL6T7jKrmdtE9
# n/s4bLzMe1OL8SInW7tZam0ATHpk5rqm5ILUrrEW7gwFObxb8BuVHzNTL1aKp2jV
# aDi7DrhaJFLkMHX55kpN0DjM64hd3MAL4hSaEt8qbtrLJkSBq7uLUN5dIFt1ZDNM
# O2ZaHMvI8kpFsoCOdrVRok3vDr7fAjJmm/s4dEV/7/nqJp8/hFnIpm0+GgqMAgwu
# QBefxsy6Qp34UvnhUs1Lv6bJStkklUQnzT5jHp+EjkYCuq4TxMnoVPN3d38UIpmp
# 4RSOjB0ooDQR43xwoqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzZaMC8GCSqGSIb3DQEJBDEiBCBrxclQsdUpAPM7xo3Pa9uxze0BXLsNTRn/SNOL
# EpLboTANBgkqhkiG9w0BAQEFAASCAgBfNPfpDtoGGVGHklx0yTwtmAk1a7i4ue15
# A0xlHOQ5xFEf9gUh+6yDXRAACNfkAJwrt1uWgb3MRb0kgsuxzFsK0vLXWaMyPppp
# A5hdWHCXgPL1v6gyOlOiWelqUcsRb2zsDgASjwGDSZzGFBElx2tpozY0wGBqav4N
# 2ld5Q0sdzb2f3uX0c12nMPeh2G3JNbt4wNFixxvO7IDMNP6HcupWrJE0qlJY1CZl
# M7hfkS7HaeQpdcJ0etDeZwIHK44+GjYXa9wUEaImWCKls01IwgOXOI5ueEDAHHlJ
# swYg+aGIn580dt3CVzFncN2JHVmQ35UVQ2fFfnZB1EXqGc07TkUcnumHvRLiXlKh
# lK4srV+2YqRprpqWUZJxeECPLkGL08ZQOzVnnmA5x7J8g3L2mVQ9cQVcWQBrQTm5
# T0w6aGYuUkdVdgPky5Imi2QBWSaVaojRNDCxcl8V4Kdz5g/rn0c/yoY3Jlw7YQmZ
# 0wQHwl3kkDivKmglEQllfTo9Qcir2mCMC8+xpDGfZYCCCmbZv3ut2UGSiqvK7oU6
# kt/sena3elwdgBhY/RAoZBEdrHhokJVtRWlBAzt4q9xyjPWI71xDgxbwkhmHeBXa
# 3E8YY5wILKZUVy4BYmOBDbcbHEEtaFpHYPVepKdKXYgY/L9ilA+aejEJIfGdDs0k
# 70xLe913VA==
# SIG # End signature block
