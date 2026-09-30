<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB SharePoint publication contract tests.

.VERSION
0.1.5
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.1.5'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:Passed = 0
$script:Failed = 0

function Invoke-SmartWorkplaceCMDBSharePointTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Test
    )

    try {
        & $Test
        $script:Passed++
        Write-Information "PASS $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Information "FAIL $Name - $($_.Exception.Message)" `
            -InformationAction Continue
    }
}

function Assert-SmartWorkplaceCMDBSharePointTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBSharePointThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Body,
        [Parameter(Mandatory)][string]$ExpectedText
    )

    try {
        & $Body
    }
    catch {
        if ($_.Exception.Message -notlike "*$ExpectedText*") {
            throw "Unexpected exception: $($_.Exception.Message)"
        }
        return
    }
    throw "Expected exception containing '$ExpectedText'."
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $projectRoot `
    'Modules\SmartWorkplaceCMDB.SharePoint\SmartWorkplaceCMDB.SharePoint.psd1'
$orchestratorPath = Join-Path $projectRoot `
    'Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1'
$globalTemplatePath = Join-Path $projectRoot `
    'Config\SmartWorkplaceCMDB.global.local.json.template'
$tenantTemplatePath = Join-Path $projectRoot `
    'Config\Tenants\tenant.local.json.template'
Import-Module $modulePath -Force
$sharePointModule = Get-Module -Name SmartWorkplaceCMDB.SharePoint

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
    'SmartWorkplaceCMDB-SharePoint-Tests-' +
    [guid]::NewGuid().ToString('N')
)
$dataAll = Join-Path $tempRoot 'DATA-ALL'
$dataLast = Join-Path $tempRoot 'DATA-LAST'
$logAll = Join-Path $tempRoot 'LOG-ALL'

try {
    foreach ($path in @(
            (Join-Path $dataAll 'Entra\Users\2026\07'),
            (Join-Path $dataLast 'CMDB'),
            (Join-Path $logAll 'Orchestration'),
            (Join-Path $logAll 'CollectionSummary\2026\09'),
            (Join-Path $logAll 'Jobs\SmartWorkplaceCMDB-EntraUsers-Collect')
        )) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
    $history = Join-Path $dataAll `
        'Entra\Users\2026\07\Entra_Users_20260719.csv'
    $latest = Join-Path $dataLast 'CMDB\CMDB_Users.csv'
    $log = Join-Path $logAll `
        'Orchestration\SmartWorkplaceCMDB-Orchestrator_20260719.csv'
    $jobLog = Join-Path $logAll `
        'Jobs\SmartWorkplaceCMDB-EntraUsers-Collect\EntraUsers_20260719.log'
    $jobTranscript = Join-Path $logAll `
        'Jobs\SmartWorkplaceCMDB-EntraUsers-Collect\EntraUsers_20260719.transcript.txt'
    $sourceStatus = $latest + '.status.json.txt'
    $reportManifest = Join-Path $dataLast 'CMDB\report-data.manifest.json.txt'
    $summaryHtml = Join-Path $logAll `
        'CollectionSummary\2026\09\SmartWorkplaceCMDB_CollectionSummary_20260914.html'
    foreach ($path in @($history, $latest, $log, $jobLog, $jobTranscript, $sourceStatus, $reportManifest, $summaryHtml)) {
        'TenantKey,Value' | Set-Content -LiteralPath $path -Encoding UTF8
    }
    $arbitraryText = Join-Path $logAll 'Orchestration\state.txt'
    'not a diagnostic transcript' | Set-Content `
        -LiteralPath $arbitraryText -Encoding UTF8

    Invoke-SmartWorkplaceCMDBSharePointTest `
        'Preserve DATA-ALL relative structure' {
        $relative = Get-SmartWorkplaceCMDBSharePointRelativePath `
            -LocalFilePath $history `
            -DataAllRootPath $dataAll `
            -LatestOutputRootPath $dataLast `
            -LogRootPath $logAll
        Assert-SmartWorkplaceCMDBSharePointTrue `
            ($relative -eq
                'DATA-ALL/Entra/Users/2026/07/Entra_Users_20260719.csv') `
            'DATA-ALL SharePoint relative path is invalid.'
    }

    Invoke-SmartWorkplaceCMDBSharePointTest `
        'Preserve DATA-LAST and LOG-ALL relative structures' {
        $latestRelative = Get-SmartWorkplaceCMDBSharePointRelativePath `
            $latest $dataAll $dataLast $logAll
        $logRelative = Get-SmartWorkplaceCMDBSharePointRelativePath `
            $log $dataAll $dataLast $logAll
        Assert-SmartWorkplaceCMDBSharePointTrue `
            ($latestRelative -eq 'DATA-LAST/CMDB/CMDB_Users.csv' -and
                $logRelative -eq
                'LOG-ALL/Orchestration/SmartWorkplaceCMDB-Orchestrator_20260719.csv') `
            'Latest or log SharePoint relative path is invalid.'
    }

    Invoke-SmartWorkplaceCMDBSharePointTest `
        'Preserve LOG-ALL job log and transcript relative structures' {
        $jobLogRelative = Get-SmartWorkplaceCMDBSharePointRelativePath `
            $jobLog $dataAll $dataLast $logAll
        $jobTranscriptRelative = Get-SmartWorkplaceCMDBSharePointRelativePath `
            $jobTranscript $dataAll $dataLast $logAll
        Assert-SmartWorkplaceCMDBSharePointTrue `
            ($jobLogRelative -eq
                'LOG-ALL/Jobs/SmartWorkplaceCMDB-EntraUsers-Collect/EntraUsers_20260719.log' -and
                $jobTranscriptRelative -eq
                'LOG-ALL/Jobs/SmartWorkplaceCMDB-EntraUsers-Collect/EntraUsers_20260719.transcript.txt') `
            'Job log or transcript SharePoint relative path is invalid.'
    }

    Invoke-SmartWorkplaceCMDBSharePointTest `
        'Accept CSV HTML logs transcripts and source status but reject arbitrary text' {
        $csvContentType = & $sharePointModule {
            param($Path)
            Get-SmartWorkplaceCMDBSharePointContentType `
                -FileInfo (Get-Item -LiteralPath $Path)
        } $history
        $logContentType = & $sharePointModule {
            param($Path)
            Get-SmartWorkplaceCMDBSharePointContentType `
                -FileInfo (Get-Item -LiteralPath $Path)
        } $jobLog
        $htmlContentType = & $sharePointModule {
            param($Path)
            Get-SmartWorkplaceCMDBSharePointContentType `
                -FileInfo (Get-Item -LiteralPath $Path)
        } $summaryHtml
        $transcriptContentType = & $sharePointModule {
            param($Path)
            Get-SmartWorkplaceCMDBSharePointContentType `
                -FileInfo (Get-Item -LiteralPath $Path)
        } $jobTranscript
        $statusContentType = & $sharePointModule {
            param($Path)
            Get-SmartWorkplaceCMDBSharePointContentType `
                -FileInfo (Get-Item -LiteralPath $Path)
        } $sourceStatus
        $manifestContentType = & $sharePointModule {
            param($Path)
            Get-SmartWorkplaceCMDBSharePointContentType `
                -FileInfo (Get-Item -LiteralPath $Path)
        } $reportManifest
        Assert-SmartWorkplaceCMDBSharePointThrow {
            & $sharePointModule {
                param($Path)
                Get-SmartWorkplaceCMDBSharePointContentType `
                    -FileInfo (Get-Item -LiteralPath $Path)
            } $arbitraryText
        } 'Only CSV, HTML, LOG, .transcript.txt, .status.json.txt, and .manifest.json.txt files'
        Assert-SmartWorkplaceCMDBSharePointTrue `
            ($csvContentType -eq 'text/csv' -and
                $htmlContentType -eq 'text/html; charset=utf-8' -and
                $logContentType -eq 'text/plain; charset=utf-8' -and
                $transcriptContentType -eq 'text/plain; charset=utf-8' -and
                $statusContentType -eq 'application/json; charset=utf-8' -and
                $manifestContentType -eq 'application/json; charset=utf-8') `
            'SharePoint content types are invalid.'
    }

    Invoke-SmartWorkplaceCMDBSharePointTest `
        'Reject files outside CMDB data roots' {
        $outside = Join-Path $tempRoot 'outside.csv'
        'Value' | Set-Content -LiteralPath $outside -Encoding UTF8
        Assert-SmartWorkplaceCMDBSharePointThrow {
            Get-SmartWorkplaceCMDBSharePointRelativePath `
                $outside $dataAll $dataLast $logAll | Out-Null
        } 'outside the configured'
    }

    Invoke-SmartWorkplaceCMDBSharePointTest `
        'Keep SharePoint disabled in templates with CMDB target' {
        $global = Get-Content -Raw -LiteralPath $globalTemplatePath |
            ConvertFrom-Json
        $tenant = Get-Content -Raw -LiteralPath $tenantTemplatePath |
            ConvertFrom-Json
        Assert-SmartWorkplaceCMDBSharePointTrue `
            (-not $global.SharePoint.Enabled -and
                -not $tenant.SharePoint.Enabled -and
                $global.SharePoint.TargetFolderPath -eq 'SMART-CMDB/DATA' -and
                $tenant.SharePoint.TargetFolderPath -eq 'SMART-CMDB/DATA' -and
                $global.Notifications.MailTimeoutSeconds -eq 120 -and
                $tenant.Notifications.MailTimeoutSeconds -eq 120) `
            'SharePoint template safety or target is invalid.'
    }

    Invoke-SmartWorkplaceCMDBSharePointTest `
        'Publish only successful unbounded and unscoped live orchestrations' {
        $content = Get-Content -Raw -LiteralPath $orchestratorPath
        $moduleContent = Get-Content -Raw -LiteralPath `
            (Join-Path $projectRoot `
                'Modules\SmartWorkplaceCMDB.SharePoint\SmartWorkplaceCMDB.SharePoint.psm1')
        Assert-SmartWorkplaceCMDBSharePointTrue `
            ($content.Contains("`$mode -in @('Collect', 'Finalize')") -and
                $content -match '\$MaxItems -eq 0' -and
                $content -match '\$DisableSharePointUpload' -and
                $content -match '-not \$activeDirectoryScoped' -and
                $content -match 'Publish-SmartWorkplaceCMDBSharePointFile' -and
                $content -match 'SharePoint summary and log synchronization completed' -and
                $content -match '\$finalSharePointSnapshot' -and
                $content -match '\$sharePointPublishedSnapshot' -and
                $content -match 'SharePoint file publication failed' -and
                $content -match 'failedUpload\.LocalFilePath' -and
                $content -match 'failedUpload\.Error' -and
                $content -match "\.Extension -ieq '\.log'" -and
                $content -match "\.Extension -ieq '\.html'" -and
                $content -match "\.transcript\.txt'" -and
                $content -match "\.status\.json\.txt'" -and
                $content -match "\.manifest\.json\.txt'" -and
                $moduleContent -match "\.Extension -ieq '\.csv'" -and
                $moduleContent -match "\.Extension -ieq '\.log'" -and
                $moduleContent -match "\.Extension -ieq '\.html'" -and
                $moduleContent -match "\.transcript\.txt'" -and
                $moduleContent -match "\.status\.json\.txt'" -and
                $moduleContent -match "\.manifest\.json\.txt'" -and
                $moduleContent -match "text/plain; charset=utf-8" -and
                $moduleContent -match 'Only CSV, HTML, LOG, \.transcript\.txt, \.status\.json\.txt, and \.manifest\.json\.txt files') `
            'Orchestrator SharePoint publication guard is incomplete.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolved.StartsWith(
            $tempBase,
            [StringComparison]::OrdinalIgnoreCase
        ) -and (Test-Path $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

Write-Information (
    'SmartWorkplaceCMDB SharePoint tests completed. Version={0}; Passed={1}; Failed={2}' -f
    $ScriptVersion,
    $script:Passed,
    $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfilKHZd5uQMfM
# gu0rYZKKNr7O4kRq/66IZqu4+aOGGqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCCZTW/rlzN3OZuDbWt5vm0
# ufPj0ovfmPly/jXB2Nt3FjANBgkqhkiG9w0BAQEFAASCAYCqretTu+iNzn0iDPjk
# cZdvHcg12+4ZHW8zCWHF4kA6Uj8u7GC6YoaJuZMU9szxG5K/oi7QjQL8BrYtxBb2
# twrkVY9eHiI2NA/hKSZeu+jTR63gy5xycpZ/v0zbIleucpgs5DRsHd4dN5JH778U
# 1b2VrVO/j0NHZH2u8XrKpqxbsM2HiQv/uaHsIwDRcW8UJv2DHrGi/3PEf79qutym
# 3mzIU6/7feDVIg8Ldsk4jlGrItnt3qGx8nlKb/OBHN0Sq1rBC5Vt9EbE8S0Vhl4J
# /UcwBDg4ao9veHimVkrRef3H/LUoaoBaM8hgfkIKPgtIkTi5XztzlUlr343XV+V9
# yeqluSV09g2I8ZIveYj7SUBd8xXoo/X6i+DlDcn6T+wwoj+o2OTDnZvV2lmT6mRU
# NSyGioXcNU3YHeqt20myQksTL7VGt4yvf/JoSWmISSOQB4MOvIakCcK+IgzgZtS8
# 2AfuOzJ1ZSnm7pVcsTTpUX0kARhcRIjOkZmxR/PFx5TuBvc=
# SIG # End signature block
