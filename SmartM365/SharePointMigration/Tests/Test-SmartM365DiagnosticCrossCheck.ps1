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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBw7ajJCbhntqiP
# m6rv030cd6vkgDhTqForZjSB74uY0KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD3FVJ/3WLCoUYSLxgsVEYa
# JBiXSrhHXEQt5EvF7+NcpzANBgkqhkiG9w0BAQEFAASCAYAXbDIM2FpDjOj4zUtz
# 8ahtM+hBDOft3MttbHm32prTiGxboYcnnrCFoCNJc09v+p+9wxsS7Cu3bgUAGQRu
# EqIUGfVSH6+TIRsQ/mTXf+oSkYekYP1Ht1aCGhlHuZHMEwRtVjLnu+jG3Yx2KcPU
# JmGH6jaoN9q1AKSJizRRXbARg67YX6QjiIrBFhzzfeGd2SQSICWkPTzwuewtK2Dg
# P0d7yURqOJWBQqTKh+B7oJ4XNdFxqyu0aOdbDq7rRM6jsRMyk7vWPb9cZBVpmULf
# uoNlcWsfbCpVd5aMQmz4bIlq75bD5zXcd6lDv1naPyC+k0UlMjoqdm0Mo4IZTkuB
# uS+i1g68USaYutqnLZdqvjr7hJ0J3TThCP2fr5H68wtqzoPmSar/WnDg9DDtNSO8
# GrbTk3/dFGeYhjc9afSEmRTQd0e4YKfTLMvmsHjbCU4N3ZG27GoCzVC+aiA03pHE
# oXWSrsCm2myS6pQbTou9T7PDhFzM6/z+p91AfgUJIdjEvWM=
# SIG # End signature block
