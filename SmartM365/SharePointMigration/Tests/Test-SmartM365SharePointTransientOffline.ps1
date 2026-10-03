<#
.SYNOPSIS
    Offline ShareGate transient batch and pilot review validation.
.VERSION
    1.0.3
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Join-Path $env:TEMP ('SmartM365TransientOffline-' + [guid]::NewGuid().ToString('N'))
$oldModulePath = $env:PSModulePath
$scriptRoot = Join-Path $PSScriptRoot '..\\Scripts\\Diagnostics'
$batchScript = Join-Path $scriptRoot 'SmartM365-SharePointMigration-ShareGateTransientBatch.ps1'
$reviewScript = Join-Path $scriptRoot 'SmartM365-SharePointMigration-ShareGatePilotReview.ps1'
try {
    $project = Join-Path $root 'Migrations\Synthetic'
    $diag = Join-Path $project 'ShareGate\Diagnostics'
    $analysis = Join-Path $diag 'Analysis-Test'
    $witness = Join-Path $diag 'Witness-Test'
    $pilot = Join-Path $diag 'Pilot-Test'
    $qualification = Join-Path $diag 'PathQualification-Test'
    $correction = Join-Path $diag 'PathCorrection-Test'
    $reports = Join-Path $pilot 'Reports'
    $correctionReports = Join-Path $correction 'Reports'
    New-Item -ItemType Directory -Path $analysis,$witness,$reports,$qualification,$correctionReports -Force | Out-Null
    $source = 'https://source.example/comm/'
    $destination = 'https://destination.example/sites/comm/'
    $dominantId = '03a7e397-f234-4a3f-a538-0b79b10cd968'
    $pagesId = 'b9e1059b-08ae-492a-926c-9fbe071756f5'
    $otherId = '4e532988-0dbe-47be-a6ac-a476c3ab97ea'
    $items = @(
        [pscustomobject]@{ Id=88; List='Files'; ListId=$dominantId; Site=$source; Role='PilotA'; Status='Success'; Path='Album/a.jpg' },
        [pscustomobject]@{ Id=185; List='Files'; ListId=$dominantId; Site=$source; Role='PilotB'; Status='Success'; Path='Album/b.jpg' },
        [pscustomobject]@{ Id=467; List='Files'; ListId=$dominantId; Site=$source; Role='PilotC'; Status='Success'; Path='Album/c.jpg' },
        [pscustomobject]@{ Id=868; List='Files'; ListId=$dominantId; Site=$source; Role='PilotD'; Status='Success'; Path='Album/d.jpg' },
        [pscustomobject]@{ Id=1; List='Pages du site'; ListId=$pagesId; Site=$source; Role='PilotSkipped'; Status='Skipped'; Path='Home.aspx' },
        [pscustomobject]@{ Id=1; List='Pages du site'; ListId=$otherId; Site='https://source.example/other/'; Role='SeparateHome'; Status=''; Path='Home.aspx' },
        [pscustomobject]@{ Id=2; List='Files'; ListId=$otherId; Site='https://source.example/other/'; Role='Remaining'; Status=''; Path='Album/Next.jpg' },
        [pscustomobject]@{ Id=3; List='Files'; ListId=$otherId; Site='https://source.example/other/'; Role='Remaining'; Status=''; Path='Autre/Last.jpg' }
    )
    $classified = @($items | ForEach-Object {
        [pscustomobject]@{ SessionId='260930-6'; RowId=('row-' + $_.ListId + '-' + $_.Id); Timestamp='2026-10-03T10:00:00Z';
            RuleId='SG-ACCESS-SOURCE'; State='To fix'; AccessSide='Source'; ItemName=$_.Path;
            ItemKey=($_.Site.TrimEnd('/') + '|' + $_.ListId + '|' + $_.Id); SourceUrl=$_.Site; SourceList=$_.List;
            SourceListId=$_.ListId; SourceItemId=$_.Id; DestinationUrl=$destination; DestinationList=$_.List;
            ObjectType='File'; 'Raw: Source path'=$_.Path; 'Raw: Destination path'=$_.Path }
    })
    $classified += @(1..7 | ForEach-Object {
        [pscustomobject]@{ SessionId='260930-6'; RowId=('no-id-' + $_); Timestamp='2026-10-03T10:00:00Z';
            RuleId='SG-ACCESS-SOURCE'; State='To fix'; AccessSide='Source'; ItemName=('Synthetic ' + $_);
            ItemKey=''; SourceUrl=$source; SourceList=''; SourceListId=''; SourceItemId='';
            DestinationUrl=$destination; DestinationList=''; ObjectType=$(if ($_ -le 6) { 'Site' } else { 'File' });
            'Raw: Source path'=''; 'Raw: Destination path'='' }
    })
    $classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
    $classified | Export-Csv -LiteralPath $classifiedPath -NoTypeInformation -Encoding UTF8
    'Witness' | Set-Content -LiteralPath (Join-Path $witness 'Witness-Results.csv') -Encoding UTF8
    $analysisHash = (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash
    $witnessHash = (Get-FileHash -LiteralPath (Join-Path $witness 'Witness-Results.csv') -Algorithm SHA256).Hash
    $pilotItems = @($items | Select-Object -First 5)
    $planRows = @($pilotItems | ForEach-Object {
        [pscustomobject]@{ SessionId='260930-6'; Role=$_.Role; ItemKey=($_.Site.TrimEnd('/') + '|' + $_.ListId + '|' + $_.Id);
            SourceUrl=$_.Site; SourceList=$_.List; SourceItemId=$_.Id; DestinationUrl=$destination; DestinationList=$_.List }
    })
    $resultRows = @($planRows | ForEach-Object {
        [pscustomobject]@{ SessionId=$_.SessionId; Role=$_.Role; ItemKey=$_.ItemKey; SourceUrl=$_.SourceUrl;
            SourceList=$_.SourceList; SourceItemId=$_.SourceItemId; DestinationUrl=$_.DestinationUrl;
            DestinationList=$_.DestinationList; ReportPath=(Join-Path $reports ('Pilot-{0:D2}.csv' -f ([array]::IndexOf($planRows, $_) + 1)));
            CopySessionId='261003-001' }
    })
    $planRows | Export-Csv -LiteralPath (Join-Path $pilot 'Pilot-Plan.csv') -NoTypeInformation -Encoding UTF8
    $resultRows | Export-Csv -LiteralPath (Join-Path $pilot 'Pilot-Results.csv') -NoTypeInformation -Encoding UTF8
    for ($i=0; $i -lt 5; $i++) {
        [pscustomobject]@{ Status=$pilotItems[$i].Status; 'Source ID'=$pilotItems[$i].Id; 'Source path'=$pilotItems[$i].Path;
            'Destination path'=''; 'Microsoft 365 Import: Status'=$(if ($i -lt 4) { 'Finished' } else { '' }) } |
            Export-Csv -LiteralPath (Join-Path $reports ('Pilot-{0:D2}.csv' -f ($i+1))) -NoTypeInformation -Encoding UTF8
    }
    @("Session=260930-6; AnalysisSHA256=$analysisHash; WitnessSHA256=$witnessHash;",
      '2026-10-03 13:00:00 Starting real item copy 1/5: ID=88',
      '2026-10-03 13:00:30 Finished item 1/5: ID=88',
      '2026-10-03 13:00:30 Starting real item copy 2/5: ID=185',
      '2026-10-03 13:01:00 Finished item 2/5: ID=185',
      '2026-10-03 13:01:00 Starting real item copy 3/5: ID=467',
      '2026-10-03 13:01:30 Finished item 3/5: ID=467',
      '2026-10-03 13:01:30 Starting real item copy 4/5: ID=868',
      '2026-10-03 13:02:00 Finished item 4/5: ID=868') |
        Set-Content -LiteralPath (Join-Path $pilot 'Pilot.log') -Encoding UTF8
    'complete' | Set-Content -LiteralPath (Join-Path $pilot 'Pilot-Summary.json.txt') -Encoding UTF8
    $qualifiedPath = $destination + 'Files/Album/a.jpg'
    [pscustomobject]@{
        SessionId='260930-6'; SourceItemId=88; AnalysisSHA256=$analysisHash;
        SourceUrl=$source; SourceList='Files'; SourceFilePath='Album/a.jpg';
        DestinationUrl=$destination; DestinationList='Files'; DestinationFilePath='Album/a.jpg';
        SourceRead=($source + 'Files/Album/a.jpg'); DestinationBefore=$qualifiedPath; DestinationAfter=$qualifiedPath;
        RootBefore='Absent'; RootAfter='Absent'; CopySessionId='261003-101';
        ShareGateResult='Success'; Qualification='Passed'; Error=''
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $qualification 'PathQualification-Result.json.txt') -Encoding UTF8
    @(
        [pscustomobject]@{'Source ID'=88;'Session ID'='261003-101';Status='Success';'Destination path'='Album/a.jpg';Errors='';Warnings='';'Microsoft 365 Import: Status'=''},
        [pscustomobject]@{'Source ID'='';'Session ID'='261003-101';Status='Success';'Destination path'='';Errors='';Warnings='';'Microsoft 365 Import: Status'='Finished'}
    ) | Export-Csv -LiteralPath (Join-Path $qualification 'ShareGate-Report.csv') -NoTypeInformation -Encoding UTF8
    @('ShareGate module version=26.9.5.0; path=synthetic',
      'Qualification=Passed; ShareGate=Success; Root=Absent;') |
        Set-Content -LiteralPath (Join-Path $qualification 'PathQualification.log') -Encoding UTF8
    . (Join-Path $scriptRoot 'SmartM365-SharePointMigration-PlacementEvidence.ps1')
    $qHash = Get-SmartM365PlacementEvidenceHash -Paths @(
        (Join-Path $qualification 'PathQualification-Result.json.txt'),
        (Join-Path $qualification 'ShareGate-Report.csv'),
        (Join-Path $qualification 'PathQualification.log'))
    $correctionResults = @(
        for ($i=1; $i -le 3; $i++) {
            $item = $items[$i]
            $reportPath = Join-Path $correctionReports ('Item-{0:D2}.csv' -f $i)
            @(
                [pscustomobject]@{'Source ID'=$item.Id;'Session ID'=('261003-10' + $i);Status='Success';'Destination path'=$item.Path;Errors='';Warnings='';'Microsoft 365 Import: Status'=''},
                [pscustomobject]@{'Source ID'='';'Session ID'=('261003-10' + $i);Status='Success';'Destination path'='';Errors='';Warnings='';'Microsoft 365 Import: Status'='Finished'}
            ) | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
            [pscustomobject]@{
                ItemKey=($item.Site.TrimEnd('/') + '|' + $item.ListId + '|' + $item.Id);
                SourceItemId=$item.Id; SourceFilePath=$item.Path; DestinationFilePath=$item.Path;
                DestinationFolder='Album'; DestinationItemUrl=($destination + 'Files/' + $item.Path);
                Status='Success'; RootBefore='Absent'; RootAfter='Absent';
                CopySessionId=('261003-10' + $i); ReportPath=$reportPath; Error=''
            }
        }
    )
    $correctionResults | Export-Csv -LiteralPath (Join-Path $correction 'PathCorrection-Results.csv') -NoTypeInformation -Encoding UTF8
    $correctionPlanHash = 'A' * 64
    [pscustomobject]@{
        SessionId='260930-6'; AnalysisSHA256=$analysisHash; QualificationSHA256=$qHash;
        RunStatus='Completed'; Planned=3; Success=3; Skipped=0; Error=0;
        Unreported=0; NotAttempted=0; PlanSHA256=$correctionPlanHash
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $correction 'PathCorrection-Summary.json.txt') -Encoding UTF8
    @(
        "2026-10-03 10:00:00 ShareGate=26.9.5.0; PlanSHA256=$correctionPlanHash",
        '2026-10-03 10:00:01 Starting item 1/3; ID=185;',
        '2026-10-03 10:00:30 Completed ID=185; Status=Success;',
        '2026-10-03 10:00:31 Starting item 2/3; ID=467;',
        '2026-10-03 10:01:00 Completed ID=467; Status=Success;',
        '2026-10-03 10:01:01 Starting item 3/3; ID=868;',
        '2026-10-03 10:01:30 Completed ID=868; Status=Success;',
        '2026-10-03 10:01:30 Completed all three items;'
    ) | Set-Content -LiteralPath (Join-Path $correction 'PathCorrection.log') -Encoding UTF8
    . (Join-Path $scriptRoot 'SmartM365-SharePointMigration-TransientEvidence.ps1')
    $evidence = Get-SmartM365TransientEvidence -ProjectRoot $project -AnalysisDirectory $analysis -PilotDirectory $pilot -SessionId '260930-6'
    if (@($evidence.AccessItems).Count -ne 8 -or @($evidence.PilotItems).Count -ne 5 -or @($evidence.OutOfBatchRows).Count -ne 7) { throw 'Synthetic evidence count mismatch.' }
    $base = @{ ProjectRoot=$project; AnalysisDirectory=$analysis; WitnessDirectory=$witness;
        PilotDirectory=$pilot; QualificationDirectory=$qualification; PathCorrectionDirectory=$correction;
        SessionId='260930-6'; BatchSize=1 }
    $dry = & $batchScript @base -DryRun 2>&1 | Out-String
    if ($dry -notmatch 'remaining=2; batches=2' -or $dry -notmatch 'Placement proof: qualified=1; corrected=3' -or
        $dry -notmatch 'destinationFolder=Album' -or
        $dry -notmatch 'destinationFolder=Autre' -or $dry -notmatch 'PlanSHA256=([0-9A-F]{64})') { throw "DryRun plan mismatch: $dry" }
    $planHash = $Matches[1]
    $corruptReport = Join-Path $correctionReports 'Item-01.csv'
    $originalBytes = [IO.File]::ReadAllBytes($corruptReport)
    try {
        $tamperedRows = @(Import-Csv -LiteralPath $corruptReport)
        $tamperedRows[0].'Destination path' = 'Wrong/b.jpg'
        $tamperedRows | Export-Csv -LiteralPath $corruptReport -NoTypeInformation -Encoding UTF8
        try { & $batchScript @base -DryRun 2>&1 | Out-Null; throw 'Expected wrong destination-path rejection.' }
        catch { if ($_.Exception.Message -notmatch 'Placement report does not prove') { throw } }
    }
    finally { [IO.File]::WriteAllBytes($corruptReport,$originalBytes) }
    $fakeModuleRoot = Join-Path $root 'Modules\ShareGate'
    New-Item -ItemType Directory -Path $fakeModuleRoot -Force | Out-Null
    @'
@{ RootModule='ShareGate.psm1'; ModuleVersion='99.0.0'; GUID='379ab298-4145-414c-89fb-94c2e44abf55'; FunctionsToExport=@('Connect-Site','Get-List','Get-File','Get-ListItem','New-CopySettings','Copy-Content','Export-Report') }
'@ | Set-Content -LiteralPath (Join-Path $fakeModuleRoot 'ShareGate.psd1') -Encoding UTF8
    @'
function Connect-Site { param([string]$Url,[switch]$Browser) [pscustomobject]@{Url=$Url} }
function Get-List { param($Site,[string]$Name) [pscustomobject]@{Title=$Name; SiteUrl=$Site.Url; RootFolder=('/sites/comm/' + $Name)} }
function Get-File {
    param($List,[string]$Path)
    $modified = if ($List.SiteUrl -like 'https://source.example/*') { [datetime]'2026-10-02T10:00:00' }
                else { [datetime]'2026-10-03T10:00:00Z' }
    [pscustomobject]@{Address=('https://destination.example/sites/comm/' + $List.Title + '/' + $Path); Modified=$modified}
}
function Get-ListItem { param($List,[int]$Id) [pscustomobject]@{Id=$Id; Modified=[datetime]'2026-10-02T10:00:00'} }
function New-CopySettings { param([string]$OnContentItemExists) [pscustomobject]@{Mode=$OnContentItemExists} }
function Copy-Content {
    param($SourceList,$DestinationList,[int[]]$SourceItemId,$CopySettings,[string]$TaskName)
    [pscustomobject]@{SessionId='261003-999'; Successes=0; Warnings=0; Errors=1; Ids=$SourceItemId}
}
function Export-Report {
    param($CopyResult,[string]$Path)
    if ($env:SMART_TRANSIENT_TEST_EXPORT_FAIL -eq '1') { throw 'Synthetic export failure.' }
    @($CopyResult.Ids | ForEach-Object { [pscustomobject]@{ Status='Error'; 'Source ID'=[string]$_; Errors='synthetic 401' } }) |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}
Export-ModuleMember -Function Connect-Site,Get-List,Get-File,Get-ListItem,New-CopySettings,Copy-Content,Export-Report
'@ | Set-Content -LiteralPath (Join-Path $fakeModuleRoot 'ShareGate.psm1') -Encoding UTF8
    $env:PSModulePath = (Split-Path -Parent $fakeModuleRoot) + [IO.Path]::PathSeparator + $oldModulePath
    $expected = @{ ExpectedAnalysisHash=$analysisHash; ExpectedWitnessHash=$witnessHash;
        ExpectedPilotManifestHash=$evidence.PilotManifestSHA256; ExpectedPlanHash=$planHash;
        ExpectedOriginalItemCount=8; ExpectedRemainingItemCount=2; MaxErrorsPerBatch=0 }
    try { & $batchScript @base @expected -Run -ConfirmBatch 2>&1 | Out-Null; throw 'Expected real-copy block.' }
    catch { if ($_.Exception.Message -notmatch 'remains disabled') { throw } }
    if (@(Get-ChildItem -LiteralPath $diag -Directory -Filter 'Transient-*').Count) {
        throw 'Blocked real run created a batch output.'
    }    $reviewOutput = & $reviewScript -ProjectRoot $project -AnalysisDirectory $analysis -PilotDirectory $pilot -SessionId '260930-6' -Run 2>&1 | Out-String
    $reviewRun = @(Get-ChildItem -LiteralPath $diag -Directory -Filter 'PilotReview-*' | Select-Object -First 1)
    if ($reviewRun.Count -ne 1) { throw "Read-only review did not create output: $reviewOutput" }
    $reviewItems = @(Import-Csv -LiteralPath (Join-Path $reviewRun[0].FullName 'PilotReview-Items.csv'))
    if ($reviewItems.Count -ne 5 -or @($reviewItems | Where-Object UrlEvidence -NotLike 'Verified*').Count -ne 0 -or
        @($reviewItems | Where-Object { $_.Result -eq 'Skipped' -and $_.SkipAssessment -like 'Consistent*' }).Count -ne 1) {
        throw 'Read-only review failed to verify URLs or compare Home.aspx Modified dates.'
    }
    Write-Output 'Offline ShareGate batch test passed: four placement proofs, wrong-path rejection, Home.aspx and hors-lot exclusions, folder-aware DryRun, and real-copy block.'
    Write-Output 'Offline ShareGate review test passed: five verified item URLs and Home.aspx Modified comparison.'
}
finally {
    $env:PSModulePath = $oldModulePath
    Remove-Item Env:\SMART_TRANSIENT_TEST_EXPORT_FAIL -ErrorAction SilentlyContinue
    Remove-Module ShareGate -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcnRWC3MRPFsbG
# 2gDGsu8fCl5ZS7dIRluZmnBGDpt//aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKqjrLavvAbhSZlOdYerI0+GG/SCptz3yQcgIxNATmg6MA0GCSqG
# SIb3DQEBAQUABIIBgGbgdzSPOU65DBTdmNdYZ6WqnC5V4Cx2BJsUT1V1aKz7j6IZ
# o/iqPEtM1WtgId+HkdbQnzykjQ9MV4etWqUR1nICgh7aQs6vINBC4sG+b0K/V0zA
# 5n2aFV32+rBEn5TbTaCqCymXzA9GTqr/i0kVvRFWio2mI0BcrlgCNV7MqtJLMClL
# BR/0ZosCshUpTvHyvuZpFFC2dHe2wPCpThPGG8miyiPJWJqrB7Wj+hwN21NBpixM
# PXjCj/pfhK/nmHsEYjC2L9hKs65xUNTOWPST6oRoK3nGrWd4MsXEg22qzDFD0c5C
# DnBedg1GHL5xhOfvWoptR/R4wYwV/ucp82xtQLRxGUaXh3+7xOQlXCjq+xufDd0F
# xSAu4y3cK08iMiiNMO5ERKiAv/5DYbfyecMUL6hB1uZZXkg7cGiwe/0QE0hC+WkQ
# +0wu8GX00fEe3NxUVboop/DvtYFag2TxVyMfH2lblPzKIYg9GwYH1/KbndPI/WUH
# hAMP8CdvanVDwPosTaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxODIy
# NDhaMC8GCSqGSIb3DQEJBDEiBCBl16yqxVzd5yduk6sdBMacEraFSG1k9tu622jj
# vkCFxzANBgkqhkiG9w0BAQEFAASCAgCMvgp31cDjfARnbNuwh+DI63NTLYagQbnH
# 1xoXo0dyUhBLWRE37xwrGDGalNQ2GMILTtVleZ46gRDiHONX80CT2/J7JPPHP7dj
# 8/nqAFjUl/0S7LEu13ApbHsPTPFqbmAkmNSHoOI6MH4VjlJz1rQa92kamHxlNAzu
# bpDjAOSFUYlol6x73fYvFQ/lMSg2PEkCWtafee96qYw/7ZNJCjkkLUEvXHClyiNR
# zC2wE0B9LEQUjoxMEip/T/9uRVfvWV4E2YNHNSv6UkeeQSLGjl5LtdcizU73zZ4I
# 13wdFbdC2U2G0Q+Q+luYOvDh8/UjJAZkfzUe5VSDe8UrcXyWzfaV9Innv7sVD1eM
# 1pAucY+3mCTyYJjMYaxuSWT0wegDkil73TulRJfEu5GHF8LZM1m1L5a4WCe/9QFB
# OCdFSP+IC4LCORBfnCgK6xNJPDecz+RsYDxsRCaiaKWCVzczH+mpmSA4ZElW86CK
# qJo/SsMw0AbQcIi6aX8CWD11MQVtPhhcddCRxKcHy3bf2CWRpdbDWivlA4IMORqF
# q6EzjbyxlCo4jn0tx8cNq2sVh437lukWfS3VIrFy2z0/XRRKepFHnOC/MNfgENYi
# 60tQqILhxDtd3btR+aR3wJq3Y2nKtYU+1ae3giVA14TC1jy81vesOQLpPqAOAYhy
# jeJkVaYAKg==
# SIG # End signature block
