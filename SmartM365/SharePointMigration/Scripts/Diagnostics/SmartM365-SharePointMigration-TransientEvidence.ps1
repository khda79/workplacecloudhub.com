<#
.SYNOPSIS
    Validates the private analysis and five-item ShareGate pilot evidence.
.VERSION
    1.0.1
#>

function Get-SmartM365DiagnosticsChild {
    param([Parameter(Mandatory)][string]$Parent, [Parameter(Mandatory)][string]$Path)
    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    if (-not $resolved.StartsWith($Parent.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside the project's ShareGate Diagnostics directory: $resolved"
    }
    return $resolved
}

function Get-SmartM365PilotManifest {
    param([Parameter(Mandatory)][string]$PilotDirectory)
    $files = @(Get-ChildItem -LiteralPath $PilotDirectory -Recurse -File -ErrorAction Stop | Sort-Object {
        $_.FullName.Substring($PilotDirectory.Length + 1).Replace('\', '/').ToLowerInvariant()
    })
    if ($files.Count -lt 9) { throw 'Pilot evidence is incomplete.' }
    $entries = @($files | ForEach-Object {
        $relative = $_.FullName.Substring($PilotDirectory.Length + 1).Replace('\', '/').ToLowerInvariant()
        [pscustomobject]@{ RelativePath=$relative; SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
    })
    $canonical = (($entries | ForEach-Object { $_.RelativePath + "`t" + $_.SHA256 }) -join "`n") + "`n"
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($canonical)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $digest = ([BitConverter]::ToString($hash.ComputeHash($bytes))).Replace('-', '') }
    finally { $hash.Dispose() }
    return [pscustomobject]@{ SHA256=$digest; Files=$entries }
}

function Get-SmartM365TransientEvidence {
    param(
        [Parameter(Mandatory)][string]$ProjectRoot,
        [Parameter(Mandatory)][string]$AnalysisDirectory,
        [Parameter(Mandatory)][string]$PilotDirectory,
        [Parameter(Mandatory)][string]$SessionId
    )
    $project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    $diagnostics = Join-Path $project 'ShareGate\Diagnostics'
    $analysis = Get-SmartM365DiagnosticsChild -Parent $diagnostics -Path $AnalysisDirectory
    $pilot = Get-SmartM365DiagnosticsChild -Parent $diagnostics -Path $PilotDirectory
    if ($analysis -eq $pilot) { throw 'Analysis and pilot directories must differ.' }
    $classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
    $pilotResultsPath = Join-Path $pilot 'Pilot-Results.csv'
    $pilotPlanPath = Join-Path $pilot 'Pilot-Plan.csv'
    $pilotLogPath = Join-Path $pilot 'Pilot.log'
    foreach ($required in @($classifiedPath,$pilotResultsPath,$pilotPlanPath,$pilotLogPath)) {
        if (-not (Test-Path -LiteralPath $required -PathType Leaf)) { throw "Required evidence is missing: $required" }
    }
    $analysisHash = (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash
    $logText = Get-Content -LiteralPath $pilotLogPath -Raw -ErrorAction Stop
    if ($logText -notmatch ('(?i)Session=' + [regex]::Escape($SessionId) + '(?:;|\s)') -or
        $logText -notmatch ('(?i)AnalysisSHA256=' + [regex]::Escape($analysisHash) + '(?:;|\s)')) {
        throw 'Pilot log does not match the selected analysis hash and session.'
    }
    $access = @{}
    $outOfBatch = [System.Collections.Generic.List[object]]::new()
    foreach ($row in @(Import-Csv -LiteralPath $classifiedPath -Encoding UTF8)) {
        if ($row.SessionId -ne $SessionId -or $row.RuleId -ne 'SG-ACCESS-SOURCE' -or
            $row.State -ne 'To fix' -or $row.AccessSide -ne 'Source') { continue }
        if ([string]::IsNullOrWhiteSpace([string]$row.SourceItemId)) {
            $outOfBatch.Add([pscustomobject]@{
                Disposition='Hors lot - à traiter à part'; Reason='Source ID absent';
                RowId=$row.RowId; Timestamp=$row.Timestamp; ObjectType=$row.ObjectType; ItemName=$row.ItemName;
                SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; SourcePath=$row.'Raw: Source path';
                DestinationUrl=$row.DestinationUrl; DestinationList=$row.DestinationList;
                DestinationPath=$row.'Raw: Destination path'
            })
            continue
        }
        if (-not $row.ItemKey -or
            -not $row.SourceUrl -or -not $row.SourceList -or -not $row.DestinationUrl -or -not $row.DestinationList) { continue }
        $id = 0
        if (-not [int]::TryParse([string]$row.SourceItemId, [ref]$id) -or $id -le 0) { continue }
        if ($access.ContainsKey($row.ItemKey)) {
            $existing = $access[$row.ItemKey]
            foreach ($name in @('SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList')) {
                if ([string]$existing.$name -ne [string]$row.$name) { throw "Conflicting analysis rows for item $($row.ItemKey)." }
            }
            if ($existing.ObjectType -ne 'File' -and $row.ObjectType -eq 'File') { $access[$row.ItemKey] = $row }
        }
        else { $access[$row.ItemKey] = $row }
    }
    if ($access.Count -lt 5) { throw 'Analysis contains fewer than five eligible source access items.' }
    $pilotRows = @(Import-Csv -LiteralPath $pilotResultsPath -Encoding UTF8)
    $planRows = @(Import-Csv -LiteralPath $pilotPlanPath -Encoding UTF8)
    if ($pilotRows.Count -ne 5 -or $planRows.Count -ne 5) { throw 'Pilot plan and results must each contain exactly five items.' }
    $seen = @{}
    $review = @()
    for ($index = 0; $index -lt 5; $index++) {
        $pilotRow = $pilotRows[$index]
        $planRow = $planRows[$index]
        if ($pilotRow.SessionId -ne $SessionId -or -not $pilotRow.ItemKey -or $seen.ContainsKey($pilotRow.ItemKey) -or
            -not $access.ContainsKey($pilotRow.ItemKey) -or $planRow.ItemKey -ne $pilotRow.ItemKey) {
            throw "Pilot item $($index + 1) does not match the selected analysis and plan."
        }
        $seen[$pilotRow.ItemKey] = $true
        $original = $access[$pilotRow.ItemKey]
        foreach ($name in @('SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList')) {
            if ([string]$pilotRow.$name -ne [string]$original.$name -or
                [string]$planRow.$name -ne [string]$original.$name) {
                throw "Pilot item $($index + 1) endpoint or source ID differs from the analysis."
            }
        }
        $reportName = 'Pilot-{0:D2}.csv' -f ($index + 1)
        if ((Split-Path -Leaf $pilotRow.ReportPath) -ne $reportName) { throw "Unexpected export name for pilot item $($index + 1)." }
        $reportPath = Join-Path (Join-Path $pilot 'Reports') $reportName
        if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) { throw "Pilot export is missing: $reportPath" }
        $reportRows = @(Import-Csv -LiteralPath $reportPath -Encoding UTF8)
        $itemRows = @($reportRows | Where-Object { $_.'Source ID' -eq [string]$pilotRow.SourceItemId })
        if (-not $itemRows.Count) { throw "Pilot export has no item row for source ID $($pilotRow.SourceItemId)." }
        $states = @($itemRows | ForEach-Object { [string]$_.Status } | Where-Object { $_ } | Sort-Object -Unique)
        $result = if ($states.Count -eq 1) { $states[0] } else { 'Mixed' }
        $destinationPaths = @($itemRows | ForEach-Object { [string]$_.'Destination path' } | Where-Object { $_ } | Sort-Object -Unique)
        $sourcePaths = @($itemRows | ForEach-Object { [string]$_.'Source path' } | Where-Object { $_ } | Sort-Object -Unique)
        $importStates = @($reportRows | ForEach-Object { [string]$_.'Microsoft 365 Import: Status' } | Where-Object { $_ } | Sort-Object -Unique)
        $destinationPath = if ($destinationPaths.Count -eq 1) { $destinationPaths[0] } else { [string]$original.'Raw: Destination path' }
        $destinationPathEvidence = if ($destinationPaths.Count -eq 1) { 'Pilot export' }
                                   elseif ($destinationPath) { 'Original migration report' }
                                   else { 'Unavailable' }
        $sourcePath = if ($sourcePaths.Count -eq 1) { $sourcePaths[0] } else { [string]$original.'Raw: Source path' }
        $review += [pscustomobject]@{
            Role=$pilotRow.Role; SessionId=$SessionId; ItemKey=$pilotRow.ItemKey;
            SourceUrl=$pilotRow.SourceUrl; SourceList=$pilotRow.SourceList; SourceItemId=[int]$pilotRow.SourceItemId;
            SourcePath=$sourcePath; DestinationUrl=$pilotRow.DestinationUrl; DestinationList=$pilotRow.DestinationList;
            DestinationPath=$destinationPath; DestinationPathEvidence=$destinationPathEvidence;
            CopySessionId=$pilotRow.CopySessionId; Result=$result;
            ImportStatus=($importStates -join '; '); ReportPath=$reportPath; ItemRows=$itemRows.Count
        }
    }
    $manifest = Get-SmartM365PilotManifest -PilotDirectory $pilot
    return [pscustomobject]@{
        ProjectRoot=$project; DiagnosticsRoot=$diagnostics; AnalysisDirectory=$analysis; PilotDirectory=$pilot;
        ClassifiedPath=$classifiedPath; AnalysisSHA256=$analysisHash; PilotManifestSHA256=$manifest.SHA256;
        PilotManifestFiles=$manifest.Files; AccessItems=@($access.Values); PilotItems=$review;
        OutOfBatchRows=$outOfBatch.ToArray();
        PilotLogPath=$pilotLogPath; PilotResultsPath=$pilotResultsPath
    }
}

function Export-SmartM365TransientCsv {
    param([Parameter(Mandatory)][string]$Path, [object[]]$Rows, [Parameter(Mandatory)][string[]]$Columns)
    $temporary = Join-Path (Split-Path -Parent $Path) ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        if (@($Rows).Count) { $Rows | Select-Object -Property $Columns | Export-Csv -LiteralPath $temporary -NoTypeInformation -Encoding UTF8 }
        else { Set-Content -LiteralPath $temporary -Value (($Columns | ForEach-Object { '"' + $_.Replace('"','""') + '"' }) -join ',') -Encoding UTF8 }
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}

function Export-SmartM365TransientJson {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Value)
    $temporary = Join-Path (Split-Path -Parent $Path) ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $Value | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $temporary -Encoding UTF8
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCDbHfu57nhTm8Y
# t7lIq+bLdm4f7XDu3xcKGnEalYQmkaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBtZeERCpoY/vu6DH4Mk+acH+oDzEsu/ssQtz7v1tQClMA0GCSqG
# SIb3DQEBAQUABIIBgAPa9XKcozwZnU5Jv0832gJOAKc6/BGC6k+bz+8Uo8lteJN0
# QugniEKdWlTngDHDERKumnFjwesyBMPUxw1gT5CkXEqHrrT090w9cApa5nUR565Z
# eQyhzGRm7ake0/FCyaGPiIuwJOKEr0lI12QtMWMXXul9Do0ffW/NCk52/08/uA/D
# NXk0Gt+YmSs/5GcfB2suatg62EaN93yxWdQ5qBkW9RvFfw74RwZIKUaPQzdxroy+
# 8ORX1hJnBpDeUQGJD/+X1O/jCjIncBZdn5UIjvtOzHhneemJQmchH1oJzYZsu0Iu
# pjC19VVqPqBSJdch8FfhU7NMtxGsVAzKwDEHFf+X4LdeuihCdK306oxQHD7ihwFD
# S55e0s+7rkBl88TZXmZtaBpCQqJMe9RkKnEQCSVWzOd/RTHEiqzsaekLPR0w3ASu
# 3SYI5U5R8Rj/syFIKwd23J2jk1Qk/+XHaW+HVwJ1Vq4rMOVgm6Nal6jk2ROKsiNa
# 9NPibl+9pWvYzdlHPqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxMzMy
# MjFaMC8GCSqGSIb3DQEJBDEiBCDw4TJuOl4iFODWu1ttoE9CWXZfnwNdtM9UwaRm
# eSRNATANBgkqhkiG9w0BAQEFAASCAgCYjxc++itgkePiZPqUVm7mgmthmmVgppUM
# x59Xy3fFdKiJI+qfuf5a+MOquEdqgLtY5lbIJWVM20T4PMPV5IcJmUyaybrxKz6E
# 7HxhP/X3CNaDWaOweq0hTao25rYqHSWij4DUeqJ9kKNdnj0QAxa95cwUC8xNDyTP
# FFlB+F17WmT9wWOG1gJylueQrL1myWjjhxbTL4qrvqRJHMHd10Lh7MaRWPbcN0X6
# L3lPaprTxGIN5gMKKxCr+CgoIbgC1Cl/DUgcO8D6+Bt2IEuRTXqDPpHAzwQzE+Ts
# X415aVMBi2h3mHuPm5asm7xr5xzG20qF0nrORYolhYKhc4tV5HaaacbP7i4qxcOQ
# cJkRg4tn/lBDNjyR69rms7E3VHPFQXcZo3lFHIqEeNktOgRG8kvNKvq9dDj9/XRs
# F70VBTAlFomF6tLAJZIZS0bU+Ig0WjyqVQFOUS/jNuzak8gtTG6m9wDm2TWzlYav
# bk/zbAloi8dhloZyxB6N1I4rsj/FG2NRqp28ADUjL1xy9rCYODSB4VJp3pq2qZDc
# Ga6nVIMN0JSpD7iGResq8o3xczI0xx0ZC/e1JOtkpKqmvll7HFjhi6+rORRneqQW
# lZ34u6OlqJfA3fVe0wcE/UoVlvundDA87NtiedGAyJ3eM/5m+83twIyyQtKtu1G2
# ukMM+s6IJQ==
# SIG # End signature block
