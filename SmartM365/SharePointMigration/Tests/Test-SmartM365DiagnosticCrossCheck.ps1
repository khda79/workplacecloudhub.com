<#
.SYNOPSIS
    Offline evidence cross-check cases with synthetic ShareGate, file and permission reports.
.VERSION
    1.0.1
#>
#Requires -Version 7.4
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Summary.ps1')
. (Join-Path $PSScriptRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-CrossCheck.ps1')

$root = Join-Path $PSScriptRoot ('.cross-check-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }
try {
    $project = Join-Path $root 'Synthetic'
    $fileDir = Join-Path $project 'comparisons\files\Synthetic-20261004-120000'
    $permDir = Join-Path $project 'comparisons\permissions\Synthetic-20261004-120000'
    $scanDir = Join-Path $project 'scans'
    [void](New-Item -ItemType Directory -Path $fileDir,$permDir,$scanDir -Force)
    $sourceFile = Join-Path $scanDir 'source-files.csv'
    $targetFile = Join-Path $scanDir 'target-files.csv'
    $sourcePerm = Join-Path $scanDir 'source-permissions.csv'
    $targetPerm = Join-Path $scanDir 'target-permissions.csv'
    foreach ($path in @($sourceFile,$targetFile,$sourcePerm,$targetPerm)) { 'header' | Set-Content -LiteralPath $path }
    $fileHtml = Join-Path $fileDir 'files-summary.html'
    $permHtml = Join-Path $permDir 'permissions-summary.html'
    '<html></html>' | Set-Content -LiteralPath $fileHtml
    '<html></html>' | Set-Content -LiteralPath $permHtml
    [pscustomobject]@{ SourceCsv=$sourceFile; TargetCsv=$targetFile; MatchedKeys=8; SourceUniqueKeys=10; TargetUniqueKeys=9;
        MissingInTarget=2; ExtraInTarget=1; DifferentSize=0; TargetOlderThanSource=0;
        ScanEvidenceStatus='Verified'; HtmlSummary=$fileHtml } |
        Export-Csv -LiteralPath (Join-Path $fileDir 'Summary.csv') -Delimiter ';' -NoTypeInformation
    [pscustomobject]@{ SourceWebUrl='https://legacy.example/sites/a/'; TargetWebUrl='https://target.example/sites/a/'; SourceLibraryTitle='Documents'; TargetLibraryTitle='Documents'; LibraryKey='library-1';
        MissingInTarget=2; ExtraInTarget=1 } |
        Export-Csv -LiteralPath (Join-Path $fileDir 'LibrarySummary.csv') -Delimiter ';' -NoTypeInformation
    [pscustomobject]@{ SourceCsv=$sourcePerm; TargetCsv=$targetPerm; MatchedPermissions=7; SourceUniqueKeys=10; TargetUniqueKeys=9;
        MissingInSPO=1; ExtraInSPO=1; TargetHasMorePermissions=0; TargetHasLessPermissions=1; PermissionLevelDifferent=0;
        ScanEvidenceStatus='Verified'; HtmlSummary=$permHtml } |
        Export-Csv -LiteralPath (Join-Path $permDir 'Summary.csv') -Delimiter ';' -NoTypeInformation
    [pscustomobject]@{ SourceWebUrl='https://legacy.example/sites/a'; TargetWebUrl=''; ComparisonObjectPath='/sites/a/documents'; ListTitle='Documents'; MissingInSPO=1;
        TargetHasMorePermissions=0; TargetHasLessPermissions=1; PermissionLevelDifferent=0 } |
        Export-Csv -LiteralPath (Join-Path $permDir 'PermissionSummary.csv') -Delimiter ';' -NoTypeInformation
    $migration = [pscustomobject]@{ Name='Synthetic'; Root=$project;
        Config=@{ Target=@{ SiteUrl='https://target.example/sites/a' }; Output=@{ FileComparisons='comparisons\files'; PermissionComparisons='comparisons\permissions' } } }
    $status = [pscustomobject]@{ SourceFileCsv=Get-Item $sourceFile; TargetFileCsv=Get-Item $targetFile;
        SourcePermCsv=Get-Item $sourcePerm; TargetPermCsv=Get-Item $targetPerm }
    $diagnostic = @{ GeneratedAtUtc='2026-10-04T12:01:00Z'; DistinctItems=10; UnkeyedRows=1;
        IssueItemState=@{ 'To fix'=1 }; IssueLineState=@{ 'To fix'=2 }; ReportPath='diagnostics.html' }
    $rows = @(
        [pscustomobject]@{ Status='Warning'; State='To fix'; SourceUrl='https://source.example/sites/a'; DestinationUrl='https://target.example/sites/a'; SourceList='Documents';
            SourceListId='list-1'; ItemKey='site|list|10' },
        [pscustomobject]@{ Status='Error'; State='To fix'; SourceUrl='https://source.example/sites/a/'; DestinationUrl='https://target.example/sites/a/'; SourceList='Documents';
            SourceListId='list-1'; ItemKey='site|list|10' },
        [pscustomobject]@{ Status='Warning'; State='Accepted'; SourceUrl='https://source.example/sites/a'; DestinationUrl='https://target.example/sites/a'; SourceList='Documents';
            SourceListId='list-1'; ItemKey='site|list|11' }
    )
    Set-StrictMode -Version Latest
    $result = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status `
        -DiagnosticSummary $diagnostic -DiagnosticRows $rows -DiagnosticVerified $true
    if (@($result.Evidence).Count -ne 3 -or @($result.Scopes).Count -ne 1) { throw 'Expected three evidence rows and one matched scope.' }
    $renamed = @()
    foreach ($entry in @(
        @{ Directory=$fileDir; Name='Summary.csv' },
        @{ Directory=$fileDir; Name='LibrarySummary.csv' },
        @{ Directory=$permDir; Name='Summary.csv' },
        @{ Directory=$permDir; Name='PermissionSummary.csv' }
    )) {
        $old = Join-Path $entry.Directory $entry.Name
        $new = Join-Path $entry.Directory ("Synthetic-{0}" -f $entry.Name)
        Move-Item -LiteralPath $old -Destination $new
        $renamed += [pscustomobject]@{ Old=$old; New=$new }
    }
    try {
        $prefixedResult = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status `
            -DiagnosticSummary $diagnostic -DiagnosticRows $rows -DiagnosticVerified $true
        if (@($prefixedResult.Scopes).Count -ne 1 -or
            (Split-Path -Path $prefixedResult.FilesSummary -Leaf) -ne 'Synthetic-Summary.csv' -or
            (Split-Path -Path $prefixedResult.PermissionsSummary -Leaf) -ne 'Synthetic-Summary.csv') {
            throw 'Prefixed comparison details were not read.'
        }
    }
    finally { foreach ($entry in $renamed) { Move-Item -LiteralPath $entry.New -Destination $entry.Old } }
    $withoutAnalysis = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status -DiagnosticRows @()
    if (@($withoutAnalysis.Scopes).Count -ne 1 -or $withoutAnalysis.Scopes[0].ShareGateToFix -ne '—' -or
        $withoutAnalysis.Scopes[0].Assessment -ne 'ShareGate analysis missing') {
        throw 'Missing ShareGate analysis was presented as zero issues.'
    }
    $noIssues = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status -DiagnosticSummary $diagnostic -DiagnosticRows @() -DiagnosticVerified $true
    if ($noIssues.Scopes[0].ShareGateToFix -ne '0 items / 0 lines') { throw 'An available analysis without matching issues lost its measured zero.' }
    $scope = $result.Scopes[0]
    if ($scope.ShareGateToFix -ne '1 items / 2 lines' -or $scope.FilesMissing -ne 2 -or $scope.PermsMissing -ne 1 -or
        $scope.Assessment -ne 'Both report issues') { throw 'Scope reconciliation is incorrect.' }
    if ($result.Evidence[1].Rate -notmatch '80' -or $result.Evidence[2].Rate -notmatch '70') { throw 'Independent comparison rates are incorrect.' }
    if (-not $result.FilesReport -or -not $result.PermissionsReport) { throw 'Comparison HTML links are missing.' }
    if ($result.Scopes[0].Site.TrimEnd('/') -ne 'https://target.example/sites/a') { throw 'Target site was not used to reconcile source aliases.' }
    $fr = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
    $en = [Globalization.CultureInfo]::GetCultureInfo('en-US')
    $originalCulture = [Globalization.CultureInfo]::CurrentCulture
    $expectedDate = ''
    try {
        foreach ($culture in @($fr,$en)) {
            [Globalization.CultureInfo]::CurrentCulture = $culture
            $diagnostic.GeneratedAtUtc = [datetime]::Parse('2026-10-04T12:01:00Z', [Globalization.CultureInfo]::InvariantCulture)
            $dated = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status `
                -DiagnosticSummary $diagnostic -DiagnosticRows $rows -DiagnosticVerified $true
            if ($dated.Evidence[0].Date -notmatch '^2026-10-04 \d{2}:01$') { throw 'ShareGate date changed day or month with system culture.' }
            if ($expectedDate -and $dated.Evidence[0].Date -ne $expectedDate) { throw 'ShareGate date changed with system culture.' }
            $expectedDate = $dated.Evidence[0].Date
        }
    }
    finally {
        [Globalization.CultureInfo]::CurrentCulture = $originalCulture
        $diagnostic.GeneratedAtUtc = '2026-10-04T12:01:00Z'
    }

    $workerInput = [pscustomobject]@{ Migration=$migration; Status=$status; Summary=$diagnostic; Rows=$rows; Verified=$true }
    $summaryScript = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Summary.ps1'
    $crossCheckScript = Join-Path $PSScriptRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-CrossCheck.ps1'
    $job = Start-ThreadJob -ScriptBlock {
        param($InputContext,$SummaryScript,$CrossCheckScript)
        . $SummaryScript
        . $CrossCheckScript
        Get-SmartM365DiagnosticCrossCheck -Migration $InputContext.Migration -Status $InputContext.Status `
            -DiagnosticSummary $InputContext.Summary -DiagnosticRows @($InputContext.Rows) `
            -DiagnosticVerified ([bool]$InputContext.Verified)
    } -ArgumentList $workerInput,$summaryScript,$crossCheckScript
    try {
        $background = @(Receive-Job -Job $job -Wait -ErrorAction Stop)[0]
        if (-not $background -or @($background.Scopes).Count -ne 1 -or $background.Scopes[0].Assessment -ne 'Both report issues') {
            throw 'Background cross-check returned a different result.'
        }
    }
    finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }

    $newerTarget = Join-Path $scanDir 'newer-target-files.csv'
    'header' | Set-Content -LiteralPath $newerTarget
    $status.TargetFileCsv = Get-Item $newerTarget
    $outdated = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status `
        -DiagnosticSummary $diagnostic -DiagnosticRows $rows -DiagnosticVerified $true
    if ($outdated.Evidence[1].State -ne 'Newer scans available' -or $outdated.Scopes[0].Assessment -ne 'Both report issues; review evidence') {
        throw 'A comparison on older scans was treated as current.'
    }

    $status.TargetFileCsv = Get-Item $targetFile
    $ambiguousRows = @($rows + [pscustomobject]@{ Status='Warning'; State='To fix'; SourceUrl='https://source.example/sites/a'; DestinationUrl='https://target.example/sites/a';
        SourceList='Documents'; SourceListId='list-2'; ItemKey='site|other-list|12' })
    $ambiguous = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status `
        -DiagnosticSummary $diagnostic -DiagnosticRows $ambiguousRows -DiagnosticVerified $true
    if ($ambiguous.Ambiguous -ne 1 -or $ambiguous.Scopes[0].Assessment -ne 'Ambiguous scope') { throw 'Duplicate list titles were silently reconciled.' }

    Remove-Item -LiteralPath (Join-Path $permDir 'Summary.csv') -Force
    $missing = Get-SmartM365DiagnosticCrossCheck -Migration $migration -Status $status `
        -DiagnosticSummary $diagnostic -DiagnosticRows $rows -DiagnosticVerified $true
    if ($missing.Evidence[2].State -ne 'Comparison missing') { throw 'Missing permission comparison was not explained.' }
    Write-Output 'Diagnostic cross-check offline test passed.'
}
finally {
    if ((Test-Path -LiteralPath $root) -and $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBUV0GcgIgOVlqL
# DsadazAqpP82n291chmWWim9iaryYKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJybl4zOaCpAZeqc3FNgNRr2GJVf5ZYUPHSKHEaO/Fr0MA0GCSqG
# SIb3DQEBAQUABIIBgAUPjj+xc48SCbiU8YpvX5nwJDqOd2s3E9i7bkPs7wdkCEAO
# VIbPcXIu5c9BYcvLSIpI2EHocD7SF1JEVM/VAczBuwy3Gl8ZGAZOD9LVwwZVOerl
# XyFykNfaSUH5nTwMHQdg939SshLxDPuWKwreIQ2Q2SG1R7Gjy1ZmeIYD6hkOn64o
# vA61UMSJqth9gb4s5sUWKuTXf232BZFDuqpSgJo7bHLcETtxdd4IiHgUKPNQ08Y/
# U+KcDI4V0fWV9qbXMFdzXlqap23Pc8pzHrqVGEUsLPTV3962RnGp7gsXfvTJnrQk
# kx/obUMZYlbFoXUSpLwUyzyTmMJl/wJmgQFSWCicfhRwN8D7AKWAjag8hV9dEjqU
# NmLWPKTLW4kngK9I+8ySHYHpdoilhUjQmsOTmfizOrGzgturpHXhUBtzL9Ti5KEs
# 6Tvky6W14WWsz/bez+MJ39IrrOWLDubEMBZYowGFvbe9t6hHWU3fWmBW8Aztw1HV
# 9FJY6E2h4KVqZogDFaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTEwOTM1
# NDhaMC8GCSqGSIb3DQEJBDEiBCArTG9O0zDsH48MOWuB+ONuza6/GLDftPwoeLU8
# d6EhFjANBgkqhkiG9w0BAQEFAASCAgAdfOeQMYTRBth7lsxyg9OpRDAJBzQlRFOH
# 5bRw0ysAgXq59fAzKQWREnckIgcO13PghIi0si518hUldScunjFIMuFoOmTf0zRq
# 2rsuuIbDzhuHA4A9m9xE7W6vbwiU3B+vg2bd1tB4V1XfJOIOWYRvQDNZ3xwOy4Ck
# MWVwPyIkGkD4QKTWlhY4DMJfs6ynAqTvTAAxfPe1QFIxKm51GhNgvEBRQcLN3inj
# SvbFCStTQ+4owb9nHPqBOLXdX052A/DaySDikKcnsokV2aJv3/5DqeT+9vcEvh0c
# 3PwjnZkkie6/BRSnSzcqZgrVnbW+A0PNi5RX+OJjRpWgNY1zLoSpbFfL9or6dY1F
# SB8Zt/89RP04lss9YhUVgNEck2otumTRmPC0UjhbQsz4iu7gzp0qbjgbUsO2+tZd
# EmlRofOL8EnpniwL9TyQ9Y5jRlIXcC6T+64ZnJTMx7+fH6nRIQihxll4yfQBZZiP
# n79FJdruttOE1SPAAa4J+/i0HCGrEXESgM3BpT7WeET+NIIR4YudGCg5wSR29ktu
# xOsWrHhUGif4W9TQYU8m5TWS+8iqfZJ+qqtEZG5K59LGsGWrXzE4VBhJP2HuyyFN
# LlPl3CF0xxYw2vlEpuhUDi6ndRfgGrX8j1j1wH82l7GfYZGfBMVzEsag/Qjd5xrk
# smWaRYZSEQ==
# SIG # End signature block
