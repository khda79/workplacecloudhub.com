<#
.SYNOPSIS
    Validates ShareGate destination-folder proof for the four copied pilot items.
.VERSION
    1.0.0
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest

function Get-SmartM365PlacementEvidenceHash {
    param([Parameter(Mandatory)][string[]]$Paths)
    $componentHashes = @($Paths | ForEach-Object {
        if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) { throw "Placement evidence is missing: $_" }
        (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash
    })
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($componentHashes -join ':'))))).Replace('-', '')
    }
    finally { $sha.Dispose() }
}

function Assert-SmartM365PlacementReport {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$SourceItemId,
        [Parameter(Mandatory)][string]$CopySessionId,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][hashtable]$Aliases
    )
    $rows = @(Import-Csv -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)
    if (-not $rows.Count -or
        -not (Test-SmartM365ShareGateReportField -Row $rows[0] -Field 'SourceItemId' -Aliases $Aliases) -or
        -not (Test-SmartM365ShareGateReportField -Row $rows[0] -Field 'DestinationPath' -Aliases $Aliases)) {
        throw "Placement report lacks required item or destination-path columns: $Path"
    }
    $itemRows = @($rows | Where-Object {
        (Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $Aliases) -eq [string]$SourceItemId
    })
    if ($itemRows.Count -ne 1 -or
        (Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'Status' -Aliases $Aliases) -ne 'Success' -or
        (Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'DestinationPath' -Aliases $Aliases) -ne $DestinationPath -or
        [string]$itemRows[0].'Session ID' -ne $CopySessionId -or
        (Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'Errors' -Aliases $Aliases) -or
        (Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'Warnings' -Aliases $Aliases)) {
        throw "Placement report does not prove exactly one successful item at $DestinationPath : $Path"
    }
    $otherItemRows = @($rows | Where-Object {
        $id = Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $Aliases
        $id -and $id -ne [string]$SourceItemId
    })
    $importStates = @($rows | ForEach-Object { [string]$_.'Microsoft 365 Import: Status' } | Where-Object { $_ } | Sort-Object -Unique)
    if ($otherItemRows.Count -or $importStates.Count -ne 1 -or $importStates[0] -ne 'Finished' -or
        @($rows | Where-Object {
            (Get-SmartM365ShareGateReportValue -Row $_ -Field 'Errors' -Aliases $Aliases) -or
            (Get-SmartM365ShareGateReportValue -Row $_ -Field 'Warnings' -Aliases $Aliases)
        }).Count) {
        throw "Placement report has another item, an unfinished import, or an error: $Path"
    }
}

function Assert-SmartM365PlacementItemRoute {
    param(
        [Parameter(Mandatory)]$AnalysisRow,
        [Parameter(Mandatory)][string]$SourceFilePath,
        [Parameter(Mandatory)][string]$DestinationFilePath,
        [Parameter(Mandatory)][string]$DestinationItemUrl
    )
    $route = Resolve-SmartM365ShareGateDestinationPath -Row $AnalysisRow
    if ($route.SourceFilePath -ne $SourceFilePath -or $route.DestinationFilePath -ne $DestinationFilePath -or
        -not $route.DestinationFolder) {
        throw "Placement proof differs from the classified route for item $($AnalysisRow.ItemKey)."
    }
    $targetUri = [uri]$DestinationItemUrl
    $expectedUri = [uri]$AnalysisRow.DestinationUrl
    $actualPath = [uri]::UnescapeDataString($targetUri.AbsolutePath).Replace('\','/')
    if ($targetUri.Scheme -ne 'https' -or $targetUri.Host -ne $expectedUri.Host -or
        -not $actualPath.EndsWith('/' + $route.DestinationFilePath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Verified ShareGate destination URL does not end at the classified file path for item $($AnalysisRow.ItemKey)."
    }
    return $route
}

function Get-SmartM365ShareGatePlacementEvidence {
    param(
        [Parameter(Mandatory)][string]$DiagnosticsRoot,
        [Parameter(Mandatory)][string]$QualificationDirectory,
        [Parameter(Mandatory)][string]$PathCorrectionDirectory,
        [Parameter(Mandatory)][string]$SessionId,
        [Parameter(Mandatory)][string]$AnalysisSHA256,
        [Parameter(Mandatory)][object[]]$AccessItems,
        [Parameter(Mandatory)][object[]]$PilotSuccessItems,
        [Parameter(Mandatory)][hashtable]$Aliases
    )
    $qualification = Get-SmartM365DiagnosticsChild -Parent $DiagnosticsRoot -Path $QualificationDirectory
    $correction = Get-SmartM365DiagnosticsChild -Parent $DiagnosticsRoot -Path $PathCorrectionDirectory
    if ($qualification -eq $correction -or
        (Split-Path -Leaf $qualification) -notlike 'PathQualification-*' -or
        (Split-Path -Leaf $correction) -notlike 'PathCorrection-*') {
        throw 'Qualification and correction must use distinct matching Diagnostics folders.'
    }
    $qResultPath = Join-Path $qualification 'PathQualification-Result.json.txt'
    $qReportPath = Join-Path $qualification 'ShareGate-Report.csv'
    $qLogPath = Join-Path $qualification 'PathQualification.log'
    $qualificationHash = Get-SmartM365PlacementEvidenceHash -Paths @($qResultPath,$qReportPath,$qLogPath)
    $qualified = Get-Content -LiteralPath $qResultPath -Raw -ErrorAction Stop | ConvertFrom-Json
    if ($qualified.SessionId -ne $SessionId -or $qualified.AnalysisSHA256 -ne $AnalysisSHA256 -or
        $qualified.Qualification -ne 'Passed' -or $qualified.ShareGateResult -ne 'Success' -or
        $qualified.Error -or $qualified.RootBefore -ne 'Absent' -or $qualified.RootAfter -ne 'Absent' -or
        -not $qualified.SourceRead -or -not $qualified.DestinationBefore -or
        $qualified.DestinationBefore -ne $qualified.DestinationAfter -or
        $qualified.CopySessionId -notmatch '^\d{6}-\d+$') {
        throw 'One-file qualification is incomplete or does not match the analysis and session.'
    }
    $qPilot = @($PilotSuccessItems | Where-Object {
        $_.SourceItemId -eq [int]$qualified.SourceItemId -and
        $_.SourceUrl -eq $qualified.SourceUrl -and $_.SourceList -eq $qualified.SourceList -and
        $_.DestinationUrl -eq $qualified.DestinationUrl -and $_.DestinationList -eq $qualified.DestinationList
    })
    if ($qPilot.Count -ne 1) { throw 'Qualified item is not exactly one of the successful five-item pilot entries.' }
    $accessByKey = @{}
    foreach ($row in $AccessItems) { $accessByKey[$row.ItemKey] = $row }
    $null = Assert-SmartM365PlacementItemRoute -AnalysisRow $accessByKey[$qPilot[0].ItemKey] `
        -SourceFilePath $qualified.SourceFilePath -DestinationFilePath $qualified.DestinationFilePath `
        -DestinationItemUrl $qualified.DestinationAfter
    Assert-SmartM365PlacementReport -Path $qReportPath -SourceItemId ([int]$qualified.SourceItemId) `
        -CopySessionId ([string]$qualified.CopySessionId) -DestinationPath ([string]$qualified.DestinationFilePath) -Aliases $Aliases
    $qLog = Get-Content -LiteralPath $qLogPath -Raw -ErrorAction Stop
    $versionMatch = [regex]::Match($qLog, 'ShareGate module version=(?<version>\d+(?:\.\d+){2,3});')
    if (-not $versionMatch.Success -or $qLog -notmatch 'Qualification=Passed; ShareGate=Success;') {
        throw 'Qualification log lacks a passed result or ShareGate version.'
    }
    $shareGateVersion = $versionMatch.Groups['version'].Value

    $cSummaryPath = Join-Path $correction 'PathCorrection-Summary.json.txt'
    $cResultsPath = Join-Path $correction 'PathCorrection-Results.csv'
    $cLogPath = Join-Path $correction 'PathCorrection.log'
    $reportDir = Join-Path $correction 'Reports'
    $reportFiles = @(Get-ChildItem -LiteralPath $reportDir -File -Filter 'Item-*.csv' -ErrorAction Stop | Sort-Object Name)
    if ($reportFiles.Count -ne 3) { throw 'Correction proof requires exactly three per-item exports.' }
    $correctionFiles = @($cSummaryPath,$cResultsPath,$cLogPath) + @($reportFiles | ForEach-Object FullName)
    $correctionHash = Get-SmartM365PlacementEvidenceHash -Paths $correctionFiles
    $summary = Get-Content -LiteralPath $cSummaryPath -Raw -ErrorAction Stop | ConvertFrom-Json
    if ($summary.SessionId -ne $SessionId -or $summary.AnalysisSHA256 -ne $AnalysisSHA256 -or
        $summary.QualificationSHA256 -ne $qualificationHash -or $summary.RunStatus -ne 'Completed' -or
        [int]$summary.Planned -ne 3 -or [int]$summary.Success -ne 3 -or
        [int]$summary.Skipped -ne 0 -or [int]$summary.Error -ne 0 -or
        [int]$summary.Unreported -ne 0 -or [int]$summary.NotAttempted -ne 0 -or
        $summary.PlanSHA256 -notmatch '^[0-9A-Fa-f]{64}$') {
        throw 'Three-item correction summary is incomplete or does not match the reviewed evidence.'
    }
    $cLog = Get-Content -LiteralPath $cLogPath -Raw -ErrorAction Stop
    if ($cLog -notmatch ('(?i)ShareGate=' + [regex]::Escape($shareGateVersion) + ';') -or
        $cLog -notmatch ('(?i)PlanSHA256=' + [regex]::Escape([string]$summary.PlanSHA256) + '(?:;|\s)') -or
        $cLog -notmatch 'Completed all three items;') {
        throw 'Correction log does not confirm the same ShareGate version and completed plan.'
    }
    $correctionRows = @(Import-Csv -LiteralPath $cResultsPath -Encoding UTF8)
    if ($correctionRows.Count -ne 3) { throw 'Correction results must contain exactly three items.' }
    $seen = @{}
    $seen[$qPilot[0].ItemKey] = $true
    for ($index = 0; $index -lt 3; $index++) {
        $result = $correctionRows[$index]
        $reportName = 'Item-{0:D2}.csv' -f ($index + 1)
        $pilot = @($PilotSuccessItems | Where-Object ItemKey -EQ $result.ItemKey)
        if ($pilot.Count -ne 1 -or $seen.ContainsKey($result.ItemKey) -or
            $pilot[0].SourceItemId -ne [int]$result.SourceItemId -or
            $result.Status -ne 'Success' -or $result.Error -or
            -not $result.RootBefore -or $result.RootBefore -ne $result.RootAfter -or
            $result.CopySessionId -notmatch '^\d{6}-\d+$' -or
            (Split-Path -Leaf $result.ReportPath) -ne $reportName -or
            $reportFiles[$index].Name -ne $reportName) {
            throw "Correction result $($index + 1) does not match a distinct successful pilot item or report."
        }
        $seen[$result.ItemKey] = $true
        $route = Assert-SmartM365PlacementItemRoute -AnalysisRow $accessByKey[$result.ItemKey] `
            -SourceFilePath $result.SourceFilePath -DestinationFilePath $result.DestinationFilePath `
            -DestinationItemUrl $result.DestinationItemUrl
        if ($result.DestinationFolder -ne $route.DestinationFolder) {
            throw "Correction folder differs from the classified path for item $($result.ItemKey)."
        }
        Assert-SmartM365PlacementReport -Path $reportFiles[$index].FullName -SourceItemId ([int]$result.SourceItemId) `
            -CopySessionId ([string]$result.CopySessionId) -DestinationPath $route.DestinationFilePath -Aliases $Aliases
    }
    if ($seen.Count -ne 4 -or $PilotSuccessItems.Count -ne 4 -or
        @($PilotSuccessItems | Where-Object { -not $seen.ContainsKey($_.ItemKey) }).Count) {
        throw 'The qualified and corrected items do not cover all four successful pilot items.'
    }
    return [pscustomobject]@{
        QualificationDirectory=$qualification; PathCorrectionDirectory=$correction;
        QualificationSHA256=$qualificationHash; PathCorrectionSHA256=$correctionHash;
        ShareGateVersion=$shareGateVersion; QualifiedItemKey=$qPilot[0].ItemKey;
        CorrectedItemKeys=@($correctionRows | ForEach-Object ItemKey); ProvenCount=4;
        QualificationFiles=@($qResultPath,$qReportPath,$qLogPath);
        CorrectionFiles=$correctionFiles
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBX68FsFj2p3y2P
# e2otm7qooerPJp8XN3BxvT5auM5SxKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKdyDUNYzQ+yBKjKYLAv+kXggEqwz3mUXVOYrUTEeJYcMA0GCSqG
# SIb3DQEBAQUABIIBgAARJlEIukpjv5fI2u+8LxS6helDlpA5RSn76czKdfCPD3AD
# BMDfphlt/SRs6u9q0zXOuIwPRKN2+VsQHa9/2kLvoPD2X4xy+cuXCg3K5zH3aSbU
# IGo/VZ+OnJrZn5Fs6/rLvamU9h8TZ2RJw967khE+2Tq27bKsJQqE18tBmSmT6LhP
# tigQlB6av/W3w5glcwnlQorOYYi3GCsadfuLzGIteSQ0Ki+xLQslIZ3ph88bHho2
# RG1RBJ2nPE9zmOj6gv6/TxJIUyboXSjoN8uvNuIVCeO3sBp0uiK4yztgUdOTauT1
# VGocRQJU/IZ+8v2uFlN5EBtTrjkxQCF/bw8aJE4R2YFDRgz1P0kA8nJ+aAIf8+H5
# vRtWlHkiSDU5JfWQG6lVJ3CmtYpzng5po0wAoVpXpHVHIZqwiWJxPOFAztyoUMxN
# E1e1ts1bMsyKeDpLzGojLMgfMByeZDD0hzOJ9f7yCxv29a3xRU9UY2MqKB1DheDh
# yMSztvkqMilNDv6QnqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxODIy
# NDZaMC8GCSqGSIb3DQEJBDEiBCAky4Zi8jnzVBLbSauAy2I/M7fJoJ9poHDvUb0R
# 7BnilzANBgkqhkiG9w0BAQEFAASCAgCXuRWfbSPCT+MOncvEB4wVeELvZRemMOKM
# Zfm3ygP/DUeve4fz6y6Xf3DIdrIsf54iFPInjrqv9/ovMunKjJNPu3J3r5msfRG3
# ODMMnS6b36aImfOKJdplQv/nJ7MI/84x+jcVbM3NyauDKy++onVGAfdMlpNNDPZk
# C5gFeJtZt8iCQ1Kevl2y7BA61nyKie1tO8MuMJZa6mz8/fZ8V9YUvaNmCXh0U6Mx
# 1LEupBMHs0HKgXriZHolk3I3ISDazvfMTx6bv4OYkqNiaLU3p/bWqkbK4CnpOfmC
# 5DfW5hSiDFe02bh9uEgIyoaBvhSIJNtazd4fv4stEQMsXVnBLpUDefhLpmI1qEro
# qjvog8RRJJ3m5wqgKhdGOGUTCqWi4VpRU/lOLnbYZRroUfh2RkNOMo49B2EuhyX/
# OZV5XUN8/73k69EalTv7Goq//dIcuq+Yf06HTit9JKM2sVq2w5RgD2YTxIS8Oi3u
# vxhtUZPCNijLly0hTR69WoAI4BbNiNJBaoruSS58vedD/dz0AEfFXxnvM4hjQfEq
# kic0amTbnezqxUVtSdhfjLLya6LKPUf6C5KDtXHqnCBlhaPp1ZzFPY1xifZLBdVs
# 9Jn/VhPx9a12i8i2NFm7IaDmYno9/lpXgl6nlizJ0puFSZeOLLGVJAohpvujJqqW
# FTh7X2RcPQ==
# SIG # End signature block
