<#
.SYNOPSIS
    Shared SharePoint upload policy; compatible with Windows PowerShell 5.1.
.VERSION
    1.0.0
#>
function Get-SmartM365SharePointUploadCategory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowNull()]$DataAllRoot,
        [AllowNull()]$DataLastRoot,
        [AllowNull()]$DataPowerBIRoot,
        [AllowNull()]$LogAllRoot
    )
    $candidate = [IO.Path]::GetFullPath($Path).TrimEnd([char[]]@(92,47))
    $best = $null
    foreach ($entry in @(@('DataAll',$DataAllRoot), @('DataLast',$DataLastRoot), @('DataPowerBI',$DataPowerBIRoot), @('LogAll',$LogAllRoot))) {
        if ([string]::IsNullOrWhiteSpace([string]$entry[1])) { continue }
        $root = [IO.Path]::GetFullPath([string]$entry[1]).TrimEnd([char[]]@(92,47))
        $prefix = $root + [IO.Path]::DirectorySeparatorChar
        if (($candidate.Equals($root,[StringComparison]::OrdinalIgnoreCase) -or $candidate.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) -and
            ($null -eq $best -or $root.Length -gt $best.Length)) {
            $best = @{Category=$entry[0];Length=$root.Length}
        }
    }
    if ($null -ne $best) { return $best.Category }
    return $null
}

function Get-SmartM365KnownUploadArtifactCategory {
    param([Parameter(Mandatory)][string]$Path, [AllowNull()]$LogPaths, [AllowNull()]$ReceiptPaths)
    $fullPath = [IO.Path]::GetFullPath($Path)
    foreach ($entry in @(@('DataLast', $ReceiptPaths), @('LogAll', $LogPaths))) {
        foreach ($candidate in @($entry[1])) {
            if ([string]::IsNullOrWhiteSpace([string]$candidate)) { continue }
            try { $candidatePath = [IO.Path]::GetFullPath([string]$candidate) } catch { continue }
            if ($fullPath.Equals($candidatePath, [StringComparison]::OrdinalIgnoreCase)) { return $entry[0] }
        }
    }
    return $null
}

function Resolve-SmartM365SharePointUploadPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('DataAll','DataLast','DataPowerBI','LogAll')][string]$Category,
        [AllowNull()]$LocalConfig,
        [AllowNull()]$GlobalConfig,
        [Parameter(Mandatory)][bool]$LegacyEnabled
    )

    $name = 'EnableSharePointUpload' + $Category
    foreach ($entry in @(@('Local', $LocalConfig), @('Global', $GlobalConfig))) {
        $source = $entry[0]
        $config = $entry[1]
        if ($null -eq $config) { continue }
        $value = $null
        if ($config -is [Collections.IDictionary]) {
            if ($config.Contains($name)) { $value = $config[$name] }
        } else {
            $property = $config.PSObject.Properties[$name]
            if ($null -ne $property) { $value = $property.Value }
        }
        if ($null -eq $value) { continue }
        if ($value -is [string] -and ($value.Trim() -in @('', '__USE_GLOBAL__', 'USE_GLOBAL'))) { continue }
        if ($value -isnot [bool]) { throw "$name must be a JSON Boolean or an inheritance marker ($source configuration)." }
        return [pscustomobject]@{ Category=$Category; Enabled=$value; Source=$source; Setting=$name }
    }
    return [pscustomobject]@{ Category=$Category; Enabled=$LegacyEnabled; Source='Legacy'; Setting='EnableSharePointUpload' }
}

function Resolve-SmartM365PreparedUploadPolicy {
    [CmdletBinding()]
    param(
        [AllowNull()]$LocalConfig, [AllowNull()]$GlobalConfig,
        [Parameter(Mandatory)][bool]$LegacyBatchEnabled,
        [switch]$Offline, [switch]$ValidateOnly, [switch]$WorkforceDiagnostic, [switch]$TransferOnly
    )
    if ($TransferOnly -and ($Offline -or $ValidateOnly -or $WorkforceDiagnostic)) {
        throw 'TransferOnly cannot be combined with offline, validation or diagnostic mode.'
    }
    $batch = Resolve-SmartM365SharePointUploadPolicy -Category DataPowerBI -LocalConfig $LocalConfig -GlobalConfig $GlobalConfig -LegacyEnabled $LegacyBatchEnabled
    $logs = Resolve-SmartM365SharePointUploadPolicy -Category LogAll -LocalConfig $LocalConfig -GlobalConfig $GlobalConfig -LegacyEnabled $true
    return [pscustomobject]@{
        BatchEnabled = [bool]($TransferOnly -or (-not ($Offline -or $ValidateOnly -or $WorkforceDiagnostic) -and $batch.Enabled))
        LogEnabled = [bool](-not $Offline -and $logs.Enabled)
        BatchSource = $batch.Source
        LogSource = $logs.Source
    }
}

function Resolve-SmartM365CollectorUploadPolicies {
    [CmdletBinding()]
    param(
        [AllowNull()]$LocalConfig, [AllowNull()]$GlobalConfig,
        [Parameter(Mandatory)][bool]$LegacyEnabled,
        [Parameter(Mandatory)][bool]$ExternalActionsAllowed
    )
    $result = @{}
    foreach ($category in @('DataAll','DataLast','LogAll')) {
        $policy = Resolve-SmartM365SharePointUploadPolicy -Category $category -LocalConfig $LocalConfig -GlobalConfig $GlobalConfig -LegacyEnabled $LegacyEnabled
        $result[$category] = [pscustomobject]@{
            Enabled = [bool]($ExternalActionsAllowed -and $policy.Enabled)
            Source = if ($ExternalActionsAllowed) { $policy.Source } else { 'RunGuard' }
        }
    }
    return $result
}

function Test-SmartM365CollectorUploadEnabled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('DataAll','DataLast','LogAll')][string]$Category,
        [AllowNull()]$Policies,
        [Parameter(Mandatory)][bool]$RunEnabled
    )
    if (-not $RunEnabled) { return $false }
    if ($null -eq $Policies) { return $true }
    if ($Policies -isnot [Collections.IDictionary] -or -not $Policies.Contains($Category) -or
        $null -eq $Policies[$Category] -or $Policies[$Category].Enabled -isnot [bool]) {
        throw "Invalid collector upload policy for $Category."
    }
    return $Policies[$Category].Enabled
}

function Assert-SmartM365CollectorUploadRoots {
    param([string]$DataAllRoot, [string]$DataLastRoot, [string]$LogAllRoot, [Parameter(Mandatory)]$Policies)
    $roots=@{DataAll=$DataAllRoot;DataLast=$DataLastRoot;LogAll=$LogAllRoot}
    $categories=@('DataAll','DataLast','LogAll')
    for($i=0;$i -lt $categories.Count;$i++) {
        for($j=$i+1;$j -lt $categories.Count;$j++) {
            $left=$categories[$i];$right=$categories[$j]
            if ([string]::IsNullOrWhiteSpace([string]$roots[$left]) -or [string]::IsNullOrWhiteSpace([string]$roots[$right])) { continue }
            $leftPath=[IO.Path]::GetFullPath([string]$roots[$left]).TrimEnd([char[]]@(92,47))
            $rightPath=[IO.Path]::GetFullPath([string]$roots[$right]).TrimEnd([char[]]@(92,47))
            if ($leftPath.Equals($rightPath,[StringComparison]::OrdinalIgnoreCase) -and $Policies[$left].Enabled -ne $Policies[$right].Enabled) {
                throw "Cannot independently control $left and $right uploads: their configured local roots are identical."
            }
        }
    }
}

function Initialize-SmartM365CollectorUploadContext {
    [CmdletBinding()]
    param(
        [AllowNull()]$LocalConfig, [AllowNull()]$GlobalConfig,
        [Parameter(Mandatory)][string]$DataAllRoot,
        [AllowEmptyString()][string]$DataLastRoot,
        [Parameter(Mandatory)][string]$LogAllRoot,
        [Parameter(Mandatory)][bool]$ExternalActionsAllowed
    )
    $state=Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if (-not $state) { throw 'Collector upload context requires an active inventory lifecycle.' }
    $legacy=[bool]$global:EnableSharePointUpload
    $policies=Resolve-SmartM365CollectorUploadPolicies -LocalConfig $LocalConfig -GlobalConfig $GlobalConfig -LegacyEnabled $legacy -ExternalActionsAllowed $ExternalActionsAllowed
    Assert-SmartM365CollectorUploadRoots -DataAllRoot $DataAllRoot -DataLastRoot $DataLastRoot -LogAllRoot $LogAllRoot -Policies $policies
    $state.ManagedCollectorLegacyUpload=@{Value=$legacy}
    $global:SmartM365CollectorUploadRoots=[pscustomobject]@{DataAllRoot=$DataAllRoot;DataLastRoot=$DataLastRoot;LogAllRoot=$LogAllRoot}
    $global:SmartM365CollectorUploadPolicies=$policies
    $global:EnableSharePointUpload=@($policies.Values | Where-Object Enabled).Count -gt 0
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDj5JNKKqhPKZsG
# OqPf7i3lXA8o7S2uoeHT8B+ngRaVCaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGzrUss3X5QWQkv2hd3020qJu28tqFYGSG/QnR9cLyzhMA0GCSqG
# SIb3DQEBAQUABIIBgGShFO0mVs0zgQGfV4OKbGq7xlkwDfzOkyjT/t/e46dveM46
# lEPiXe9EFyzmahsikcr4HvfFqBo7PPiXmeMmwDfpGB3IYYpDZEZT8hEHLMKtLhps
# P5Ycxxf5asDXX/mwobDZ8hVd0aJwalqL5MBVbyLF5VdruWtvqRhGlDZXIVjm0WVQ
# GUZe0U7k7hg7liXkuoCJVm/n9BVihUSNUkALWeavZG9C1JmXQjG5Am4HAMsXNRAG
# cMV5tRcz7qwG6//vblIGzGW1RA9oF4S5dpBdQavIsS0lMF6gbAgX5KB8GvuChj55
# d4cBrVw0Cu33vm9AzDmSK+HKgwxQRkXxPrYmdhnePa6w+mK9JBnAfJdjdSPY1Oxt
# ktm+ue97EImhgtIV2YNke/X6dQ0muHUG4LMoAnHoYiIKU09aMwDUfdo2gYB1dcCt
# PXkMquuyu00g3xbTsErBEDj9FnKWofzyxh+a61xcSVJwRZiZpKLhsYkZmUosxSpG
# SISENwGpJINL2zPcG6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExMzQz
# MjNaMC8GCSqGSIb3DQEJBDEiBCBJVSjjsynoLce3DaJpc+/sNo2ZmcmNgtvsURYU
# N4luJTANBgkqhkiG9w0BAQEFAASCAgB04+nE8eTuMyqcOVHchxQwM7pWZ7v+kz/g
# g45TMdFtMUplXwaucQv39alfE0le9JWoEXTdc0pb4KA18ShnuY9h5H+Hz4b3AIX2
# cklSJKOaer+/BKAKBU042mitab+e2Dhy22xTCav5R9jSUGpCuCtWw60/dKIbeQIu
# qUil68bXnZivg/2TQA7Ihjc4d6KxCSbR/HfFBTcAlLHNpK9684c75CWBDf9X6aip
# e6ePlDnnnEz6hWxXWylTF03p0oZF03C+pBGdn95x4XUBZQexlc8qCEt6bWBpRmrx
# ZX3RhaikfGj3VU+1g4eOoY6sDbzTZh6rz73pjYWulP9ERdwBifC09CIp+7p1ApE7
# 7g9XWkzYSQxQFB6fdP8fo9U61Nbm7asDUBcZxW4NfXNkZgqC9fibNm80gZKyWQa1
# H5TtvBoKyVGiRurGifEzH94VZmLXF2YeE+BbtBYw4vKVuWXglq65qU1Yz1SiwNUN
# ddImIR7vCS1ujYhR5OVGBfqRI8XMsMU3sdiZqK/OYr8nUYm2I+wV8OF0B4jVkwcN
# QDCohHiafQV6Ch0wPAldTpgMIZcrYSF2lhDF8lWW0iXU5D92jruCRXJFLxVUQmoR
# Ae7K3cxJyAinPkNXsYOesxKRsSef7j8rSw4hidkS0IcLY1N9LytCLXy+KWAP+REb
# gYpA5H6WHA==
# SIG # End signature block
