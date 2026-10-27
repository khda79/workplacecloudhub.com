# Version 1.1.0
#Requires -Version 7.0
[CmdletBinding()]
param()
# Synthetic license report qualification checks: the Licences report is consumed only when the current
# licensing receipt records it. No collector, tenant, cloud or Power BI access.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedArtifacts.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('LicenseQualificationTest-'+[guid]::NewGuid().ToString('N'))
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Json([string]$Path,$Value){[IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))}
$spec=(Get-Content (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw|ConvertFrom-Json)
$license=$spec.licenseReport
$tenant='synthetic'
$snapshotId=('b'*32)
$runId='11111111-2222-3333-4444-555555555555'

function New-Capture {
    param([string]$Name,[switch]$WithDecisions,[switch]$NoReceipt,[string]$ReceiptStatus='Completed',[string]$SnapshotRunId=$runId,
        [switch]$NoReceiptRun,[string]$UnpublishedFile,[string]$TamperedFile,[string]$ForeignSnapshotFile,[string]$ForeignRunFile,[string]$RowTenant=$tenant)
    $capture=Join-Path $root $Name
    $last=Join-Path $capture 'DATA-LAST'
    New-Item -ItemType Directory -Path $last -Force|Out-Null
    foreach($name in $spec.currentFiles){
        $path=Join-Path $last $name
        if($name -ceq $license.snapshot){continue}
        if($name -like '*.csv'){
            $sid=if($name -ceq $ForeignSnapshotFile){'c'*32}else{$snapshotId}
            $rowTenant=if($name -like 'M365_Licenses_Report*'){$RowTenant}else{$tenant}
            [pscustomobject]@{TenantKey=$rowTenant;SnapshotId=$sid}|Export-Csv -LiteralPath $path -NoTypeInformation
        }
    }
    $snapshot=[ordered]@{TenantKey=$tenant;SnapshotId=$snapshotId;SchemaVersion=2;Products=@(1,2,3,4);OtherProducts=@(1,2,3);AssignmentIssues=@{Available=$false;Rows=@()}}
    if(-not $NoReceiptRun){$snapshot.ReceiptRun=@{Producer=$license.producer;RunId=$SnapshotRunId;StartedAtUtc='2026-10-17T22:00:00Z'}}
    Json (Join-Path $last $license.snapshot) $snapshot
    $published=@($license.reportFiles)
    if($WithDecisions){
        foreach($name in $license.decisionFiles){
            $rid=if($name -ceq $ForeignRunFile){'other-run'}else{$runId}
            $sid=if($name -ceq $ForeignSnapshotFile){'c'*32}else{$snapshotId}
            [pscustomobject]@{SnapshotId=$sid;TenantKey=$RowTenant;ReceiptRunId=$rid;SkuPartNumber='SPE_E3'}|Export-Csv -LiteralPath (Join-Path $last $name) -NoTypeInformation
        }
        $published+=@($license.decisionFiles)
    }
    if(-not $NoReceipt){
        $files=foreach($name in $published){
            $hash=(Get-FileHash -LiteralPath (Join-Path $last $name)).Hash
            if($name -ceq $TamperedFile){$hash='0'*64}
            [ordered]@{File=$name;Required=$false;Status='Success';SHA256=$hash;RunId=$runId}
        }
        $optional=foreach($name in @($license.reportFiles)+@($license.decisionFiles)){
            [ordered]@{File=$name;Published=($published -contains $name -and $name -cne $UnpublishedFile)}
        }
        Json (Join-Path $last $license.receipt) ([ordered]@{Owner='SmartInventory-SourceReceipt';ContractVersion='1.2';Producer=$license.producer;TenantKey=$tenant;Status=$ReceiptStatus;RunId=$runId;Files=@($files);OptionalFiles=@($optional)})
    }
    foreach($history in $spec.history){
        $dir=Join-Path $capture "$($history.root)/2026-W40";New-Item -ItemType Directory -Path $dir -Force|Out-Null
        @{TenantKey=$tenant;Id='one'}|Export-Csv -LiteralPath (Join-Path $dir $history.file) -NoTypeInformation
        Json (Join-Path $dir 'manifest.json.txt') @{Week='2026-W40';SnapshotCreatedAtUtc='2026-10-01T00:00:00Z'}
    }
    return $capture
}
function Get-Qualification([string]$Capture){
    $batch=Join-Path $root ('batch-'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $batch|Out-Null
    $entries=@(Copy-PreparedArtifacts -SnapshotRoot $Capture -Destination $batch -TenantKey $tenant)
    Assert-PreparedArtifacts -BatchRoot $batch -Entries $entries -TenantKey $tenant
    [pscustomobject]@{Entries=$entries;Result=(Get-PreparedLicenseReportQualification -SnapshotRoot $Capture -TenantKey $tenant -Artifacts $entries)}
}
function FileState($Result,[string]$Name){@($Result.Files | Where-Object File -ceq $Name)[0].State}

try {
    New-Item -ItemType Directory -Path $root|Out-Null
    Import-Module (Join-Path $PSScriptRoot 'WorkforceLicenseEvidence.psm1') -Force
    $capture=New-Capture 'workforce-decisions' -WithDecisions
    $decisionPath=Join-Path $capture 'DATA-LAST/M365_Licenses_ReportSkuCandidates.csv'
    function Set-WorkforceDecisions($Rows,[switch]$Empty) {
        if ($Empty) {
            '"TenantKey","SnapshotId","ReceiptRunId","UserId","CandidateType","Recommendation"' | Set-Content -LiteralPath $decisionPath
        } else { $Rows | Export-Csv -LiteralPath $decisionPath -NoTypeInformation }
        $receiptPath=Join-Path $capture ('DATA-LAST/'+$license.receipt)
        $receipt=Get-Content $receiptPath -Raw | ConvertFrom-Json
        ($receipt.Files | Where-Object File -eq 'M365_Licenses_ReportSkuCandidates.csv').SHA256=(Get-FileHash $decisionPath).Hash
        Json $receiptPath $receipt
    }
    $row=[pscustomobject]@{TenantKey=$tenant;SnapshotId=$snapshotId;ReceiptRunId=$runId;UserId='contoso|one';CandidateType='Disabled account';Recommendation='Recommended'}
    Set-WorkforceDecisions @($row,$row)
    $signal=Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant
    Check ($signal.State -eq 'Observed' -and $signal.Count -eq 1) 'Workforce must deduplicate accounts across paid SKUs.'
    $row.Recommendation='Review';Set-WorkforceDecisions @($row)
    Check ((Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant).Count -eq 1) 'Review candidates disappeared.'
    $row.Recommendation='Excluded';Set-WorkforceDecisions @($row)
    Check ((Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant).Count -eq 0) 'Excluded candidate counted.'
    Set-WorkforceDecisions @() -Empty
    $signal=Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant
    Check ($signal.State -eq 'Observed' -and $signal.Count -eq 0) 'Qualified empty decisions must be measured zero.'
    $row.TenantKey='foreign';Set-WorkforceDecisions @($row)
    $signal=Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant
    Check ($signal.State -eq 'Not measured' -and $null -eq $signal.Count) 'Foreign tenant must not become zero.'
    $row.TenantKey=$tenant;$row.Recommendation='Recommended';$row.UserId='';Set-WorkforceDecisions @($row)
    Check ($null -eq (Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant).Count) 'Missing candidate key accepted.'
    $captureMissing=New-Capture 'workforce-no-receipt' -WithDecisions -NoReceipt
    Check ($null -eq (Get-WorkforceDisabledLicenseEvidence -DataRoot $captureMissing -TenantKey $tenant).Count) 'Missing receipt became measured zero.'
    # Qualified report with both decision files.
    $q=Get-Qualification (New-Capture 'qualified' -WithDecisions)
    Check ($q.Result.State -eq 'Qualified' -and $q.Result.ReceiptRunId -eq $runId -and $q.Result.SnapshotId -eq $snapshotId) 'Recorded report must be qualified.'
    Check (@($q.Result.Files | Where-Object State -ne 'Qualified').Count -eq 0 -and @($q.Result.Files).Count -eq 7) 'All seven license files must be qualified.'
    Check (@($q.Entries | Where-Object File -in @($license.decisionFiles | ForEach-Object {"Source-$_"})).Count -eq 2) 'Decision files must be copied as optional artifacts.'
    Check (Test-PreparedArtifactName 'Source-M365_Licenses_ReportSkuCapacity.csv') 'Optional decision artifact name refused.'
    # Licences 1.48 output: decision files absent, report still qualified.
    $q=Get-Qualification (New-Capture 'nodecisions')
    Check ($q.Result.State -eq 'Qualified' -and (FileState $q.Result 'M365_Licenses_ReportSkuCandidates.csv') -eq 'Absent') 'Absent decision files must not disqualify the report.'
    # Lineage defects: published batch, Not qualified, decision files never shown alone.
    $cases=@(
        @{Name='noreceipt';Args=@{NoReceipt=$true};Reason='*not part of the capture*'}
        @{Name='failedreceipt';Args=@{ReceiptStatus='Failed'};Reason='*Failed, not Completed*'}
        @{Name='otherrun';Args=@{SnapshotRunId='99999999-0000-0000-0000-000000000000'};Reason='*another licensing receipt run*'}
        @{Name='noreceiptrun';Args=@{NoReceiptRun=$true};Reason='*names no licensing receipt run*'}
        @{Name='unpublished';Args=@{UnpublishedFile='M365_Licenses_ReportGaps.csv'};Reason='*M365_Licenses_ReportGaps.csv*not recorded as published*'}
        @{Name='tampered';Args=@{TamperedFile='M365_Licenses_ReportSummary.csv'};Reason='*M365_Licenses_ReportSummary.csv*hash differs*'}
        @{Name='foreignsnapshot';Args=@{ForeignSnapshotFile='M365_Licenses_ReportCandidates.csv'};Reason='*M365_Licenses_ReportCandidates.csv*another report snapshot*'}
    )
    foreach($case in $cases){
        $arguments=$case.Args
        $q=Get-Qualification (New-Capture $case.Name -WithDecisions @arguments)
        Check ($q.Result.State -eq 'Not qualified' -and $q.Result.Reason -like $case.Reason) "Lineage defect must be Not qualified: $($case.Name) ($($q.Result.Reason))"
        Check (@($q.Result.Files | Where-Object {$_.Kind -eq 'Decision' -and $_.State -eq 'Qualified'}).Count -eq 0) "Decision files shown without a qualified report: $($case.Name)"
    }
    # A decision file of another run is Not qualified on its own; the report stays qualified.
    $q=Get-Qualification (New-Capture 'decisionrun' -WithDecisions -ForeignRunFile 'M365_Licenses_ReportSkuCapacity.csv')
    Check ($q.Result.State -eq 'Qualified' -and (FileState $q.Result 'M365_Licenses_ReportSkuCapacity.csv') -eq 'Not qualified' -and (FileState $q.Result 'M365_Licenses_ReportSkuCandidates.csv') -eq 'Qualified') 'Decision file of another run must be disqualified alone.'
    # Another tenant stays fatal.
    $foreign=New-Capture 'foreigntenant' -WithDecisions -RowTenant 'foreign'
    $failed=$false
    try{Get-Qualification $foreign|Out-Null}catch{$failed=$_.Exception.Message -like '*tenant mismatch*'}
    Check $failed 'Another tenant in license report rows must stay fatal.'
    # Artifact contract compatibility: current and reviewed previous hashes only.
    $contractPath=Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt'
    Check (Test-PreparedArtifactContractCompatible -ContractSHA256 (Get-FileHash -LiteralPath $contractPath).Hash) 'Current artifact contract refused.'
    foreach($hash in @($spec.compatibleContractSHA256)){Check (Test-PreparedArtifactContractCompatible -ContractSHA256 $hash) 'Reviewed previous artifact contract refused.'}
    Check (-not (Test-PreparedArtifactContractCompatible -ContractSHA256 ('F'*64))) 'Unknown artifact contract accepted.'
    'Prepared license report qualification offline checks passed: '+$checks
} finally {
    $full=[IO.Path]::GetFullPath($root);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $full -Leaf) -match '^LicenseQualificationTest-[a-f0-9]{32}$'){
        if(Test-Path -LiteralPath $full){Remove-Item -LiteralPath $full -Recurse -Force}
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDWpl5OHqfRkYfM
# bobwkhUJvnTiCwmBejF32CL3K9l8hqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILkwUMgsmwZPC3L4mUwFE3G/atFYTGseyK4Sd3vrbH9cMA0GCSqG
# SIb3DQEBAQUABIIBgDYzGkZTcVx2/BgMnBQQYdrE+8n0/CS4sdm5qiaY4qijNCTE
# gNkkyFkzduZvSj8f2QqacWIQjOAHB4hyV85KFBg9+Dx/+vSX8vY2ZVqAdVIf23+6
# DVGhuwrZl3qApc1ONn3BAB5zeAzS2Nwcewl1XTzl0dDjWF45W5WoSJ6v8eqnpUTi
# 6XENninsNUDXmyJRg5Wg0Wr+Tn4VwGDssbxrBe5gAH71X3TuB4DdmxMzBIz4IVBi
# 1+4lYcH9GYSC+TWzmHQUHmCg250lgAO/jPG061y1xD/waf65zqeZYJcGGAzyQLr5
# dDaJ8dW9q4nKOLaLx03AkSZ2rAF6NDjF7m3zFSlPC359cSGgvUvZc37vl5gu9GHD
# 606WoVjdnEd0ZxcopM573YjH1N9f4mO/+JggNSRM177c4SkMPrwTe8+7bG0xTHEy
# Z7CAZbni5j827jZt/DLWTzUj8MYpWDkZ9WzyzNq2/0EcZDRJUNXOHcKYWkYCmk0Q
# 6jT/v4kZQUS9H4PLKKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcwOTQx
# NTRaMC8GCSqGSIb3DQEJBDEiBCC1FRu1p1iejAeNDI56oL9L/fNEsjudzqb3Vle9
# qre7FDANBgkqhkiG9w0BAQEFAASCAgCyUA3P6XynJJ/KlfwlZKdnppibEj9AQb+c
# wXUMpBFdqIqxH/3zTLZi1EQN4ptzuYQQ87yRvuLyPYE6mtNfLkrG30RhkBGrkICs
# IZzSh5OnrE8yiwX1hFYn++v2mgLfYkDZyuNCftsHObjCDEXTHbgwA038mJudgmyj
# EdklXGaRP+yijh7+VS/PgOVv/B5/lWdZUaFOtCyuI0C7fNsMZp2a1lFV9xloNVra
# M37m+9/gGlvresn6XMEqj9934tTdVQlmhHyu13PKdt01EwT7T7kpLd1HPxC7Cm8v
# LGokVGnU31hJ8IP9MGq1uADICUCZpzf6wstxn5T2ijG61z4fVMFfLJ+rzRSSq9Ns
# 8eTnyI7ITa2jkX2FyB+k7MrVk9rPlxGlnKn7YwKyj1As1LUGb+lZ6qj2F9HAKNwf
# soXmXv+Sfy41RPRMHqTBI3N2XBR+e+hmsPkUexbytzAd7jT8Cl1pW8OXWf33C05e
# s2DL0fQHlLAyqZKaMMFCkMCn2ME2z8GH4oT50IeyHH1EMgmI6sAUc0hBOQO1GgsM
# GbgpCIfCgGGjhpwxuJggv2DrOwsBe3IcwlO/4RXMKb20xutJC754ugmoPDeCCGD6
# OUrcyV29yDueEEXT3IiPK7fLlOl59hrGq01Tsyt+U9fesznDUsG0LkUqtjrzBdpf
# 1PizvXMfHA==
# SIG # End signature block
