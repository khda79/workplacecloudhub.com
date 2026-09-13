<#
.SYNOPSIS
Runs offline tests for the full-collection summary and delta email renderer.

.VERSION
1.2.2
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:Passed = 0
$script:Failed = 0

function Invoke-SummaryTest {
    param([string]$Name, [scriptblock]$Test)
    try { & $Test; $script:Passed++; Write-Information "PASS $Name" -InformationAction Continue }
    catch { $script:Failed++; Write-Information "FAIL $Name - $($_.Exception.Message)" -InformationAction Continue }
}
function Assert-SummaryTrue {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$summaryPath = Join-Path $projectRoot 'Reports\SmartWorkplaceCMDB-CollectionSummary.ps1'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase ('SmartWorkplaceCMDB-Summary-Tests-' + [guid]::NewGuid().ToString('N'))
$latestRoot = Join-Path $tempRoot 'DATA-LAST'
$cmdbRoot = Join-Path $latestRoot 'CMDB'
$dataAllRoot = Join-Path $tempRoot 'DATA-ALL'
$logRoot = Join-Path $tempRoot 'LOG-ALL'

function Write-SummaryFixture {
    param([int]$DeviceCount, [int]$UserCount, [int]$MailboxCount,
        [int]$F1Assignments, [int]$F1Capacity,
        [int]$F3Assignments, [int]$F3Capacity,
        [int]$E3Assignments, [int]$E3Capacity,
        [int]$E5Assignments, [int]$E5Capacity,
        [int]$CopilotAssignments, [int]$CopilotCapacity)
    New-Item -ItemType Directory -Path $cmdbRoot -Force | Out-Null
    @(1..$DeviceCount | ForEach-Object {[pscustomobject]@{CmdbDeviceId="device-$_"}}) |
        Export-Csv (Join-Path $cmdbRoot 'CMDB_Devices.csv') -NoTypeInformation -Encoding UTF8
    @(1..$UserCount | ForEach-Object {[pscustomobject]@{CmdbUserId="user-$_"}}) |
        Export-Csv (Join-Path $cmdbRoot 'CMDB_Users.csv') -NoTypeInformation -Encoding UTF8
    @(1..$MailboxCount | ForEach-Object {[pscustomobject]@{CmdbMailboxId="mailbox-$_"}}) |
        Export-Csv (Join-Path $cmdbRoot 'CMDB_Mailboxes.csv') -NoTypeInformation -Encoding UTF8
    @(
        [pscustomobject]@{CmdbLicenseId='license-f1';SkuPartNumber='M365_F1';EnabledUnits=$F1Capacity},
        [pscustomobject]@{CmdbLicenseId='license-f3';SkuPartNumber='SPE_F1';EnabledUnits=$F3Capacity},
        [pscustomobject]@{CmdbLicenseId='license-e3';SkuPartNumber='SPE_E3';EnabledUnits=$E3Capacity},
        [pscustomobject]@{CmdbLicenseId='license-e5';SkuPartNumber='SPE_E5';EnabledUnits=$E5Capacity},
        [pscustomobject]@{CmdbLicenseId='license-copilot';SkuPartNumber='Microsoft_365_Copilot';EnabledUnits=$CopilotCapacity},
        [pscustomobject]@{CmdbLicenseId='license-other';SkuPartNumber='VISIOCLIENT';EnabledUnits=100}
    ) | Export-Csv (Join-Path $cmdbRoot 'CMDB_Licenses.csv') -NoTypeInformation -Encoding UTF8
    $relationships = New-Object System.Collections.Generic.List[object]
    if ($F1Assignments -gt 0) {
        foreach ($index in 1..$F1Assignments) {
            $relationships.Add([pscustomobject]@{RelationshipType='AssignedLicense';FromEntityType='User';FromEntityId="user-$index";ToEntityType='License';ToEntityId='license-f1'})
        }
    }
    foreach ($license in @(
            @{Count=$F3Assignments;Id='license-f3'},
            @{Count=$E3Assignments;Id='license-e3'},
            @{Count=$E5Assignments;Id='license-e5'})) {
        if ($license['Count'] -gt 0) {
            foreach ($index in 1..$license['Count']) {
                $relationships.Add([pscustomobject]@{RelationshipType='AssignedLicense';FromEntityType='User';FromEntityId="user-$index";ToEntityType='License';ToEntityId=$license['Id']})
            }
        }
    }
    if ($CopilotAssignments -gt 0) {
        foreach ($index in 1..$CopilotAssignments) {
            $relationships.Add([pscustomobject]@{RelationshipType='AssignedLicense';FromEntityType='User';FromEntityId="user-$index";ToEntityType='License';ToEntityId='license-copilot'})
        }
    }
    $relationships.Add([pscustomobject]@{RelationshipType='PrimaryUser';FromEntityType='User';FromEntityId='user-1';ToEntityType='Device';ToEntityId='device-1'})
    $relationships | Export-Csv (Join-Path $cmdbRoot 'CMDB_Relationships.csv') -NoTypeInformation -Encoding UTF8
}

$parameters = @{
    Tenant='audit';OrganizationKey='contoso';EnvironmentKey='prod';TenantKey='contoso-prod'
    TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';DataRootPath=$tempRoot
    DataAllRootPath=$dataAllRoot;LatestOutputRootPath=$latestRoot;LogRootPath=$logRoot
    NoConfigWrite=$true
}

try {
    Invoke-SummaryTest 'Capture private J-30 and J-7 aggregate baselines' {
        Write-SummaryFixture 1 1 1 1 10 0 8 1 10 0 12 0 5
        $first = & $summaryPath @parameters -RunId 'day30' -CaptureBaselineOnly `
            -SnapshotDateTime ([datetimeoffset]::UtcNow.AddDays(-31))
        Write-SummaryFixture 2 2 1 1 10 1 8 1 10 1 12 1 5
        $second = & $summaryPath @parameters -RunId 'day7' -CaptureBaselineOnly `
            -SnapshotDateTime ([datetimeoffset]::UtcNow.AddDays(-8))
        Assert-SummaryTrue ($first.Status -eq 'Captured' -and $second.Status -eq 'Captured') `
            'Historical summary baselines were not captured.'
    }

    Invoke-SummaryTest 'Render previous, J-7, and J-30 deltas without sending mail' {
        Write-SummaryFixture 4 3 2 2 10 2 8 3 10 1 12 2 5
        $result = & $summaryPath @parameters -RunId 'current' -RunStatus 'Completed' `
            -SnapshotDateTime ([datetimeoffset]::UtcNow) -PreviewOnly
        Assert-SummaryTrue ($result.Status -eq 'Previewed') 'Preview mode did not complete.'
        Assert-SummaryTrue ($result.Snapshot.Devices -eq 4 -and $result.Snapshot.Users -eq 3 -and $result.Snapshot.Mailboxes -eq 2) `
            'Core population counts are incorrect.'
        Assert-SummaryTrue ($result.Snapshot.M365F1Assigned -eq 2 -and
            $result.Snapshot.M365F3Assigned -eq 2 -and
            $result.Snapshot.M365E3Assigned -eq 3 -and
            $result.Snapshot.M365E5Assigned -eq 1 -and
            $result.Snapshot.M365CopilotAssigned -eq 2) `
            'License-family assignments are incorrect.'
        Assert-SummaryTrue ($result.BodyHtml -match 'Since previous' -and $result.BodyHtml -match 'Since J-7' -and $result.BodyHtml -match 'Since J-30') `
            'The required comparison columns are missing.'
        Assert-SummaryTrue ($result.Previous.RunId -match 'day7' -and $result.Day7.RunId -match 'day7' -and $result.Day30.RunId -match 'day30') `
            'The previous, J-7, or J-30 baseline was selected incorrectly.'
        Assert-SummaryTrue ($result.BodyHtml -notmatch 'device-1|user-1|mailbox-1') `
            'The aggregate email exposed an entity identifier.'
    }

    Invoke-SummaryTest 'Use the immediately previous snapshot when the dataset is unchanged' {
        $result = & $summaryPath @parameters -RunId 'unchanged' -RunStatus 'Completed' `
            -SnapshotDateTime ([datetimeoffset]::UtcNow.AddMinutes(1)) -PreviewOnly
        Assert-SummaryTrue ($result.Previous.RunId -eq 'current') `
            'An unchanged dataset skipped the immediately previous collection.'
        Assert-SummaryTrue ($result.BodyHtml -match '<td class="number">0</td>') `
            'An unchanged collection did not render zero deltas.'
    }

    Invoke-SummaryTest 'Render a failure alert without creating a business snapshot' {
        $historyBefore = @(Get-ChildItem -LiteralPath (Join-Path $dataAllRoot 'CollectionSummary') -Filter '*.csv' -File -Recurse).Count
        $result = & $summaryPath @parameters -RunId 'failed-run' -RunStatus 'Failed' `
            -SnapshotDateTime ([datetimeoffset]::UtcNow.AddMinutes(2)) `
            -OperationalError 'Synthetic collector failure' `
            -FailedStep 'Entra users collection' `
            -FailureLogPath 'C:\private\orchestrator.log' `
            -FailureTranscriptPath 'C:\private\collector.transcript.txt' `
            -PreviewOnly
        $historyAfter = @(Get-ChildItem -LiteralPath (Join-Path $dataAllRoot 'CollectionSummary') -Filter '*.csv' -File -Recurse).Count
        Assert-SummaryTrue ($result.Status -eq 'Previewed' -and
            $result.Subject -match 'collection failed' -and
            $result.BodyHtml -match 'Synthetic collector failure' -and
            $result.BodyHtml -match 'Entra users collection' -and
            $historyBefore -eq $historyAfter) `
            'The operational failure alert was not isolated from business snapshots.'
    }

    Invoke-SummaryTest 'Use the proven Graph SDK mail transport without touching snapshots' {
        $content = Get-Content -Raw -LiteralPath $summaryPath
        $tokens = $null
        $parseErrors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile(
            $summaryPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $mailBodyFunction = $ast.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq 'New-SmartWorkplaceCMDBSummaryGraphMailBody'
            }, $true)
        $mailBodyScript = $mailBodyFunction.Body.GetScriptBlock()
        $singleRecipientBody = & $mailBodyScript `
            -Subject 'Synthetic subject' `
            -BodyHtml '<p>Synthetic body</p>' `
            -ToRecipients @(@{
                    emailAddress=@{address='recipient@example.invalid'}
                })
        $singleRecipientPayload = $singleRecipientBody | ConvertFrom-Json
        $historyBefore = @(Get-ChildItem -LiteralPath `
                (Join-Path $dataAllRoot 'CollectionSummary') `
                -Filter '*.csv' -File -Recurse).Count
        $result = & $summaryPath @parameters -SendMailTestOnly -ValidateOnly
        $historyAfter = @(Get-ChildItem -LiteralPath `
                (Join-Path $dataAllRoot 'CollectionSummary') `
                -Filter '*.csv' -File -Recurse).Count
        Assert-SummaryTrue ($result.Status -eq 'Validated' -and
            $result.ScriptVersion -eq '1.2.2' -and
            $result.Subject -match 'mail transport test' -and
            $historyBefore -eq $historyAfter -and
            $parseErrors.Count -eq 0 -and
            $null -ne $mailBodyFunction -and
            $singleRecipientPayload.message.toRecipients -is [object[]] -and
            @($singleRecipientPayload.message.toRecipients).Count -eq 1 -and
            $singleRecipientPayload.message.toRecipients[0].emailAddress.address -eq
                'recipient@example.invalid' -and
            $singleRecipientPayload.saveToSentItems -eq $false -and
            $content -match 'Connect-MgGraph' -and
            $content -match 'Invoke-MgGraphRequest' -and
            $content -match 'ConvertTo-Json -Depth 12' -and
            $content -match 'Graph mail request failed') `
            'The isolated Graph SDK mail transport contract is incomplete.'
    }
}
finally {
    if ((Test-Path -LiteralPath $tempRoot) -and
        $tempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $tempRoot) -like 'SmartWorkplaceCMDB-Summary-Tests-*') {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}

Write-Information ('SmartWorkplaceCMDB collection summary tests completed. Version=1.2.2; Passed={0}; Failed={1}' -f $script:Passed,$script:Failed) -InformationAction Continue
if ($script:Failed -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBL8gaGbeAkDDdY
# gPLCy3kcdB3kjEmM0KA+keXbdlafcaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM0crwLdqAkyrS3vX1UKZ24QHA8+WZ03WmCUM4EoHF3pMA0GCSqG
# SIb3DQEBAQUABIIBgCHB5t/YGWhqGq0l58BNDUxHvKKYFHepSr4NHeoVAnsR/u2h
# fDTLhPUmS2mB21ju46+u7Kf7u1XIsL9qghxiDjlebyc6NzKKuCMlgiUpG3N+GnVI
# L8O+0IiJxw8tSHJfgMqRxZm0rsGHlYAwi4Agjyvnrb5lhI1dK1lqcIeBlOnw5bu7
# jOsi1prspv6RPund1K0URExblD8rwdiwhM0FBb5NytfS6cRw63a11VDGNhi8CIh1
# dOMOAQ6NnrcTjoCTPAntvp7tKpJKeBUC0B8UcIWtZgT6ppNCCfpCgGZ6pNhludh0
# 60/HEanhKnW2TYG/Slhm3RVLi8zf/ovkIuknnZyKUTC6RHLwf0g7ixbkjMPWNUHq
# riDZ+y9LNBytevZqKwI1kdzNkDCPq3JAosd7shvntwLCKNOAVi0yLqZdGTn4u2lo
# ADztGGK2bDjlPopdEVgBT55oL/7aNzA1gzVahUQ+jrZRhX5RgINXdna05RGy2rLL
# U2881HNoUp6rBbhdtqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTMwOTIz
# MzRaMC8GCSqGSIb3DQEJBDEiBCB7gxzqWkPH0UqWQJ6GTrp3A7v7ECHFWnl/1IN3
# dRyskDANBgkqhkiG9w0BAQEFAASCAgBAvK1R/qIBrw4ZZgbDWDxEkL+vZGPHak31
# /1D5xmVkmtvS2ITXjOSjuOTuk+YpaXaBiiGG1bGRdLX9ibJM1TsMuwB/VRLeb4s3
# 2JGCJrSi7S179eKq++dr+X45EMiJyBuXMd8BcJNWx/gy68kWm6xxObVv+w6QEyj3
# jhrxyCjiMhOvorGnknPPJ+9G+5oHcTl6mkEnKX1RBHUZxJNPxAnCQgM5nGCJd/vH
# ThRHWstBD5IjMfUrWKIFtfnfja0uJ94yk+ttfYHwqrIgRM/xSNYwDBbVYWDx3nih
# Gv4xK5Kt1xJRcgWwQx7AqYve/4pdJC9O7lprpDQ4Bsuw43WtUtg06XBicWJ3iW3V
# PQnZPQQGMgA528hr22FGxGoZWzNcOPbF65jLYFFuD/pVkFB9FucBVfDCdEzGqFCZ
# jLsLhAbcc4I2d5gDwRnCiO/c4lipIiwqdN4ZidHXwSL8JkrEu+iE5VgU++vjddzy
# x/qjiY66B3unU+Nx9j43lLub3HpPGxx1NCzAOBU2Nik7gB0e14+1PUzAfc8YkOEE
# kmKfaW+3sBJEVctXmYGw/oMT0+9x7XLnKTgz3+fe6Mgf4a71Qcenz7KPYv52HIdt
# XJ4OA+hoFAvrU2JRb1fXE83by1WGYeo4W1DVrEThDTiFdFZAvBMr8XH4XctfBl7I
# TyLyNEjKOg==
# SIG # End signature block
