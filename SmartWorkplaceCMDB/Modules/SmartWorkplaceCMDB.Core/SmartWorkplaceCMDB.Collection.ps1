# Version: 1.0.0
$script:SmartWorkplaceCMDBCollectionVersion = '1.0.0'

function Read-SmartWorkplaceCMDBCollectionFixture {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $document = Read-SmartWorkplaceCMDBJsonFile -Path $Path
    if ($document -is [array]) { return $document }
    if ($null -ne $document -and $null -ne $document.PSObject.Properties['value'] -and $document.value -is [array]) {
        return $document.value
    }
    throw "Offline collection JSON must be an array or contain a value array: $Path"
}

function Resolve-SmartWorkplaceCMDBCollectionPaths {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Paths, [switch]$Fixture, [int]$MaxItems,
        [switch]$ExplicitDataRoot, [switch]$NoWrite)
    $kind = if ($Fixture) { 'Fixture' } else { 'Live' }
    if ($MaxItems -gt 0) { $kind += '-Bounded' }
    $isTest = $Fixture -or $MaxItems -gt 0
    $root = $Paths.DataRootPath
    $markerPath = Join-Path $root '.collection-root.json'
    $marker = if (Test-Path -LiteralPath $markerPath) { Read-SmartWorkplaceCMDBJsonFile $markerPath } else { $null }
    $matches = $null -ne $marker
    if ($matches) {
        foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
            if ([string]$marker.$name -ne [string]$Paths.$name) { $matches = $false }
        }
        if ([string]$marker.Kind -ne $kind) { $matches = $false }
    }
    if (-not $isTest -and $null -ne $marker -and -not $matches) {
        throw 'Collection root identity or mode mismatch. Choose a dedicated live data root.'
    }
    if ($isTest) {
        $occupied = (Test-Path -LiteralPath $root) -and @(Get-ChildItem -LiteralPath $root -Force | Select-Object -First 1).Count -gt 0
        if (-not $matches -and (-not $ExplicitDataRoot -or $occupied)) {
            $base = if ($ExplicitDataRoot) { $root } else { Join-Path $Paths.ProjectRootPath 'Data' }
            $root = Join-Path $base ('TestRuns\{0}_{1}' -f $kind, [guid]::NewGuid().ToString('N'))
        }
        $Paths = Resolve-SmartWorkplaceCMDBTenantPath -Tenant $Paths.ProfileKey -OrganizationKey $Paths.OrganizationKey -EnvironmentKey $Paths.EnvironmentKey -TenantKey $Paths.TenantKey -TenantId $Paths.TenantId -DataRootPath $root
    }
    if (-not $NoWrite) {
        $document = [ordered]@{Version=1; Kind=$kind}
        foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) { $document[$name]=$Paths.$name }
        Write-SmartWorkplaceCMDBJsonAtomically -InputObject $document -Path (Join-Path $root '.collection-root.json')
    }
    return $Paths
}

function Start-SmartWorkplaceCMDBSourceCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Paths, [Parameter(Mandatory)][string[]]$RawPath,
        [switch]$Fixture, [int]$MaxItems, [switch]$Scoped, [switch]$NoWrite)
    $coverage = if ($MaxItems -gt 0) { 'Bounded' } elseif ($Fixture) { 'Fixture' } elseif ($Scoped) { 'Scoped' } else { 'Complete' }
    $run = [pscustomobject]@{Paths=$Paths; RawPath=$RawPath; Fixture=[bool]$Fixture; Coverage=$coverage; MaxItems=$MaxItems; RunId=[guid]::NewGuid().ToString('N'); StartedUtc=[datetime]::UtcNow.ToString('o'); NoWrite=[bool]$NoWrite; Locks=(New-Object 'System.Collections.Generic.List[object]')}
    if ($NoWrite) { return $run }
    try {
        foreach ($path in $RawPath) {
            New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
            $run.Locks.Add([IO.File]::Open(($path + '.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None))
        }
        foreach ($path in $RawPath) {
            Write-SmartWorkplaceCMDBSourceState -Run $run -Path $path -Status 'InProgress'
        }
    } catch { foreach ($handle in $run.Locks) { $handle.Dispose() }; throw }
    return $run
}

function Write-SmartWorkplaceCMDBSourceState {
    param($Run, [string]$Path, [string]$Status, [string[]]$NotCollectedPath = @())
    $state = [ordered]@{Version=1; SourceName=[IO.Path]::GetFileName($Path); RunId=$Run.RunId; Status=$Status; Coverage=$Run.Coverage; Mode=$(if ($Run.Fixture) {'Fixture'} else {'Live'}); MaxItems=$Run.MaxItems; StartedUtc=$Run.StartedUtc; CompletedUtc=''; RowCount=$null; SHA256=''}
    foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) { $state[$name]=$Run.Paths.$name }
    if ($Status -eq 'Completed') {
        $state.CompletedUtc=[datetime]::UtcNow.ToString('o')
        $state.RowCount=(Import-Csv -LiteralPath $Path | Measure-Object).Count
        $state.SHA256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
        if ($Path -in $NotCollectedPath) { $state.Coverage='NotCollected' }
    }
    Write-SmartWorkplaceCMDBJsonAtomically -InputObject $state -Path ($Path + '.status.json')
}

function Complete-SmartWorkplaceCMDBSourceCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Run, [switch]$Failed, [string[]]$NotCollectedPath=@())
    if ($Run.NoWrite) { return }
    try {
        foreach ($path in $Run.RawPath) {
            Write-SmartWorkplaceCMDBSourceState -Run $Run -Path $path -Status $(if ($Failed) {'Failed'} else {'Completed'}) -NotCollectedPath $NotCollectedPath
        }
    } finally { foreach ($handle in $Run.Locks) { $handle.Dispose() } }
}

function Import-SmartWorkplaceCMDBSourceCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LiteralPath, [Parameter(Mandatory)]$Paths)
    $statePath = $LiteralPath + '.status.json'
    if (Test-Path -LiteralPath $statePath) {
        $state = Read-SmartWorkplaceCMDBJsonFile $statePath
        foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
            if ([string]$state.$name -ne [string]$Paths.$name) { throw 'Source evidence identity mismatch.' }
        }
        if ($state.Status -ne 'Completed' -or
            $state.SHA256 -ne (Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).Hash) {
            throw "Unusable source snapshot: '$LiteralPath'. Recollect before normalization."
        }
    }
    Import-Csv -LiteralPath $LiteralPath -ErrorAction Stop
}

function Assert-SmartWorkplaceCMDBCollectionPage {
    [CmdletBinding()]
    param([AllowNull()]$Response)
    $value = $null
    if ($Response -is [System.Collections.IDictionary]) {
        if ($Response.Contains('value')) { $value = $Response['value'] }
    } elseif ($null -ne $Response -and $null -ne $Response.PSObject.Properties['value']) {
        $value = $Response.value
    }
    if ($null -eq $value -or $value -is [string] -or $value -isnot [System.Collections.IList]) {
        throw 'Invalid collection page: expected a value array. An unavailable response is not an empty snapshot.'
    }
}

function Get-SmartWorkplaceCMDBSourceHealth {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Paths, [datetimeoffset]$ReferenceDateTime=[datetimeoffset]::UtcNow,
        [int]$WarningHours=48, [int]$CriticalHours=168)
    $contract = Get-SmartWorkplaceCMDBTableContract (Join-Path $Paths.ProjectRootPath 'Schema/SmartWorkplaceCMDB.raw.tables.json')
    foreach ($table in $contract.tables) {
        $path=Join-Path $Paths.LatestOutputRootPath (Join-Path $table.area $table.name)
        $statePath=$path+'.status.json'
        $health='Unknown'; $coverage='Unknown'; $completed=''; $severity='Warning'; $rowCount=$null
        if (Test-Path -LiteralPath $statePath) {
            try {
                $state=Read-SmartWorkplaceCMDBJsonFile $statePath
                foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
                    if ([string]$state.$name -ne [string]$Paths.$name) { throw 'Source evidence identity mismatch.' }
                }
                $health=[string]$state.Status; $coverage=[string]$state.Coverage; $rowCount=$state.RowCount
                # PS7 can materialize JSON dates as DateTime; a plain string cast
                # loses their offset and fractional seconds. Round-trip typed
                # dates while retaining the string path used by Windows PS5.1.
                $completed = if ($state.CompletedUtc -is [datetime] -or $state.CompletedUtc -is [datetimeoffset]) {
                    $state.CompletedUtc.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
                } else { [string]$state.CompletedUtc }
                if ($health -eq 'Completed') {
                    $snapshotMatches=(Test-Path -LiteralPath $path) -and $state.SHA256 -eq (Get-FileHash -LiteralPath $path).Hash
                    $health=$coverage
                    $date=[datetimeoffset]::MinValue
                    if (-not $snapshotMatches) { $health='SnapshotMismatch'; $severity='Critical' }
                    elseif (-not [datetimeoffset]::TryParse($completed,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$date)) { $health='InvalidDate' }
                    elseif ($date -gt $ReferenceDateTime.AddMinutes(5)) { $health='FutureDate' }
                    elseif (($ReferenceDateTime-$date).TotalHours -gt $WarningHours) {
                        $health='Stale'; $severity=if (($ReferenceDateTime-$date).TotalHours -gt $CriticalHours) {'Critical'} else {'Warning'}
                    }
                    elseif ($coverage -eq 'Complete') { $severity='Information' }
                }
            } catch { $health='InvalidEvidence'; $severity='Critical' }
        }
        [pscustomobject]@{SourceName=$table.name; Status=$health; Coverage=$coverage; CompletedUtc=$completed; RowCount=$rowCount; Severity=$severity; HasEvidence=(Test-Path -LiteralPath $statePath)}
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATEeTZAGqDj1O3
# 1Ow9SnwRefn5EwHeZCirRp05Pg8dYqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIiRz0ntU3gnTWohXwH1Pt07FpDxctGLXX8O+QNMBdQsMA0GCSqG
# SIb3DQEBAQUABIIBgDKoDliAiCIaCK6Yk4544Xivj7hQW2pwK4bGnWKjcoirYq8B
# MLZiULzu0PMCKZdBGiO4MuJJ1qdtfUPo8bGGT/BBLe5ED6IREUOA1p0Y0lnbs2nO
# 18S/hvKucz3otrNkqzpok49Ru2y1Sj/2ujry1j2moLuOzqkuGDlv/VStZMqHZcv7
# 8Dd2imdQk7kL84TUBIUxTJDmJoZUwJWN3JpH+bT1i8OREglXk7QmjQ9Yoh8DysDI
# YZQ+NCu2r9ux779anAibF5HC4DWyA0m3Fxt/q/0/uuC6mVARwU6bJExd5RYfIfU8
# yMTpzjUZPNQpvAkwNDLEvrAAJ1boydMkwYitZ4KCYfxS3dfsIbVfXUd95+epoC5U
# h+TwBKSrSAs1vRaGJja1JKpJ201mvR2d+VevqceZ7O4si++gsmlqGOjZWUsrJx9S
# SH43pPwne/q2ZFvvKY5PRSx+lTFK00cj45Pk44TY1OEXOrLfJQ0aVIn9UFayiGnV
# A8ImZjuKbiSbIHdMKqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MTVaMC8GCSqGSIb3DQEJBDEiBCAywDKxJqjl6S/dE6h2bhyC6wa5FmJBtXl8HZtJ
# Nxi6jzANBgkqhkiG9w0BAQEFAASCAgBnvH13447UVS3FD5c5iXeV5/YXAEHJ+jmA
# T89D8dtUZkT//uW5ctR63n5Pnj1TPLixT9WI8Yz7FZ2LFLDTWz0FRAl2vsGb+RXr
# 9jFo/BCc3VE/PtdTlOtucdijsKPq7YF6lsy50Eh7/UkFrNvPVxGbQD25bCvxGJL5
# IPDd4ykpwW/5nVf0vfG7V2BaEicCbN6+ScEoGBXkfRAuo5fQxPGKVhFKKTxSYAtC
# BNxEkdgTSfmydgMAEEYMBRlCrkYbeZDVkqx3wXi2NblV35a0O1FhXJybf88ad+ZY
# J/Dlv8o6kkQ9h992RttC1wuGNU958g1Y9WHNYAgrh4AJ3AmKYHJj+kGaoN3/SKS2
# XGWQy7mjRJYFztW7+Wbn0z06D0bDIHTL9Om3nOz/uPiIZX81xBM/ztLvyJXEp1Yw
# 5o8lyAPpDgrGZbDCcK7WIJJ3AMA1JMRT251KtkQd/xB+55co5HPnVTA6Jz9vPKQ+
# A8rMkeVqnCp9TWoxmSyf3W8HNOffubViSOt8wCwjBUC3cGwwdspUA03lHAqxpij+
# uoJ9YzLrefxV5MfdG2pWm7KY5oiRlwbw7Ip1S+UtCpvk6fx/emjRLZsToZTBqC3Q
# Ei4FNRKAEOu9rq4WDrv4hr/NJyKZisCYLzL0r/v+aDUyKMjYww68pveUyPx/xwfw
# Fu10BxDpoQ==
# SIG # End signature block
