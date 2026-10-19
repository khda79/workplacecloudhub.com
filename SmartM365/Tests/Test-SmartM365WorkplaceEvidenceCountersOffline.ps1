#Requires -Version 7
<#
.SYNOPSIS
    Synthetic entry-point counters and retention diagnostics; no real preparation or transport.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
$ErrorActionPreference='Stop'
$repository=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$entry=Join-Path $repository 'SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1'
$core=Join-Path $repository 'SmartM365/Modules/SmartM365.Core/SmartM365.Core.psm1'
function Read-Ast([string]$Path){
    $t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$t,[ref]$e)
    if($e.Count){throw ($e.Message -join '; ')}
    $ast
}
$ast=Read-Ast $entry
$coreAst=Read-Ast $core
foreach($item in @(
    @{Ast=$ast;Names=@('Invoke-PreparedStep','Register-PreparedRunCsvFiles','Invoke-PostTransferRetention')},
    @{Ast=$coreAst;Names=@('Format-SmartM365LogLine','WriteLog','Complete-SmartM365ExecutionContext','Write-SmartM365CompletionBanner')}
)){
    foreach($name in $item.Names){
        $definition=@($item.Ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$true) | Where-Object Name -eq $name)
        if($definition.Count -ne 1){throw "Expected one function: $name"}
        $functionText=$definition[0].Extent.Text
        if($name -eq 'Invoke-PostTransferRetention'){
            # AST-created functions have no source file; bind only this automatic path.
            $functionText=$functionText.Replace('$PSScriptRoot',("'"+(Split-Path $entry -Parent).Replace("'","''")+"'"))
        }
        . ([scriptblock]::Create($functionText))
    }
}
. (Join-Path $repository 'SmartM365/Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$main=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.TryStatementAst]},$true) | Where-Object {
    @($_.Body.Statements | Where-Object {$_.Extent.Text -match '^\$result\s*=\s*Invoke-PreparedStep\s*\{\s*Invoke-PreparedEvidencePipeline'}).Count -eq 1
})
if($main.Count -ne 1){throw 'Expected one preparation block.'}
$statements=@($main[0].Body.Statements)
$start=0
while($statements[$start].Extent.Text -notmatch '^\$result\s*=\s*Invoke-PreparedStep\s*\{\s*Invoke-PreparedEvidencePipeline'){$start++}
$normal=[scriptblock]::Create(($statements[$start..($statements.Count-1)].Extent.Text -join "`n"))
$transferBlock=[scriptblock]::Create(@($statements | Where-Object {$_.Extent.Text -match '^if\s*\(\$TransferOnly\)'})[0].Extent.Text)
$cleanupText=$main[0].Finally.Extent.Text
$cleanupBlock=[scriptblock]::Create($cleanupText.Substring(1,$cleanupText.Length-2))
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('prepared-counters-'+[guid]::NewGuid().ToString('N'))
$product=Join-Path $root 'product'
$batch=Join-Path $root 'batch'
New-Item -ItemType Directory -Path (Join-Path $product 'scripts'),$batch -Force | Out-Null
$script:checks=0
function Check([bool]$Pass,[string]$Message){if(-not $Pass){throw $Message};$script:checks++}
function Reject([scriptblock]$Action){$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $failed 'Expected malformed receipt rejection'}
function SaveReceipt($Value){$Value | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $batch 'validation.json.txt')}
$validReceipt=@{Passed=$true;SchemaOnly=$false;Files=@(@{File='Plan.csv'},@{File='Findings.csv'})}
SaveReceipt $validReceipt
# Extra artifacts must never inflate the prepared-table count.
'synthetic' | Set-Content -LiteralPath (Join-Path $batch 'Source-Inventory.csv')
@'
param($Tenant,$Mode,$ConfigurationPath,$ExpectedPublishedBatchId,$ExpectedPointerSHA256,[switch]$NoCompletionBanner)
$control=Get-Content -LiteralPath (Join-Path (Split-Path $PSScriptRoot -Parent) 'control.json.txt') -Raw | ConvertFrom-Json
if(-not $NoCompletionBanner){Write-Host 'UNEXPECTED nested completion banner'}
[IO.File]::WriteAllText((Join-Path (Split-Path $PSScriptRoot -Parent) 'banner-flag.txt'),[string][bool]$NoCompletionBanner)
if($control.Throw){throw 'Synthetic retention failure'}
[pscustomobject]@{Status=$control.Status;PlanPath='synthetic-plan';JournalPath='synthetic-journal';Plan=[pscustomobject]@{BlockedReasons=@($control.Reasons)}}
'@ | Set-Content -LiteralPath (Join-Path $product 'scripts/Invoke-PreparedSharePointRetention.ps1')
function RetentionControl([string]$Status,[bool]$Throw=$false){
    @{Status=$Status;Throw=$Throw;Reasons=$(if($Status -eq 'Blocked'){@('Consumer registry is not explicitly qualified.','Retention target mismatch: TenantId')}else{@()})} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $product 'control.json.txt')
}
function Get-SmartM365JsonReadPath {param($Path) $Path}
function Invoke-SmartM365TeamsNotificationFromLog {param($Message,$Level)}
function Invoke-SmartM365SharePointCsvUpload {param($LocalFilePath)}
function Send-SmartM365TeamsNotification {param($Title,$Message,$Level,$Channel,$ResultSummary,$Facts)}
function Send-PreparedSuccessMail {param($BatchId,$CsvFiles,$VerifiedTransfer)}
function Assert-PreparedTransferResult {param($Result,$ExpectedBatchId) if(-not $Result.PointerVerified){throw 'Synthetic unverified pointer'}}
function Invoke-PreparedEvidencePipeline {
    for($i=0;$i -lt $script:childWarnings;$i++){Write-Host ("[2026-01-01 01:00:00] WARNING: Synthetic excluded duplicate group {0}" -f $i)}
    Write-Host 'Ordinary output: WARNING: inside an informational value'
    if($script:pipelineFailure){throw 'Synthetic generator failure'}
    $script:fixture
}
function Send-PreparedEvidenceBatch {
    param($ExpectedBatchId)
    if($script:transferFailure){throw 'Synthetic transfer failure'}
    $script:transferReceipt
}
function Remove-PreparedMappingWorkbooks {
    param($DataRoot,$MappingRoot)
    if($script:cleanupFailure){throw 'Synthetic mapping cleanup failure'}
    Write-Warning 'Synthetic mapping cleanup warning'
}
$globals=@('SmartM365InventoryRun','SmartM365ExecutionSummaryWritten','SmartM365ExecutionStatus','SmartM365ExecutionStartTime','SmartM365WarningCount','SmartM365ErrorCount','SmartM365ScriptName','SmartM365ProfileKey','SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','BasePath','LogTextFile','logTranscriptFile','csvGeneratedPaths','SmartM365SharePointUploadedFiles','SmartM365MailHtmlFiles','EnableSharePointUpload','SmartM365CompletionBannerRunKey')
$saved=@{}
foreach($name in $globals){$value=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue;$saved[$name]=@{Exists=($null -ne $value);Value=$(if($value){$value.Value}else{$null})}}
function Reset-Fixture {
    foreach($name in $globals){Set-Variable -Name $name -Scope Global -Value $null}
    $global:SmartM365WarningCount=0;$global:SmartM365ErrorCount=0
    $global:SmartM365ExecutionSummaryWritten=$false;$global:SmartM365ExecutionStartTime=Get-Date
    $global:SmartM365ScriptName='SyntheticPreparation';$global:BasePath=$root
    $global:LogTextFile=Join-Path $root ('trace-'+[guid]::NewGuid().ToString('N')+'.log')
    $global:csvGeneratedPaths=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $script:SmartM365SourceReceiptUploadPaths=@()
    $script:childWarnings=0;$script:pipelineFailure=$false;$script:transferFailure=$false
    $script:TransferOnly=$false;$script:ValidateOnly=$false;$script:Offline=$true;$script:WorkforceDiagnostic=$false
    $script:batchUploadEnabled=$false;$script:transferParameters=@{};$script:params=@{}
    $script:Tenant='synthetic';$script:scriptName='SyntheticPreparation'
    $script:effective=[pscustomobject]@{TenantKey='synthetic'};$script:output=$root
    $script:ExpectedBatchId='20260101T100000000Z-00000001'
    $script:fixture=[pscustomobject]@{BatchId=$ExpectedBatchId;Files=2;BatchPath=$batch;SourceFiles=3;Identity=@{CsvFiles=3;Rows=4}}
    $script:transferReceipt=[pscustomobject]@{BatchId=$ExpectedBatchId;CsvFiles=2;VerifiedFiles=6;PointerVerified=$true;PointerSHA256=('a'*64);AuditPath='synthetic'}
    $script:smartM365RunGuard=[pscustomobject]@{ReachedTerminal=$false;ExplicitExit=$false;InCleanup=$false}
    SaveReceipt $validReceipt
    RetentionControl 'Preview'
}
function Finish-Fixture([AllowNull()]$Failure=$null){
    Complete-SmartM365ExecutionContext -Status $(if($Failure){'Failed'}else{'Auto'}) -ErrorRecord $Failure
    Get-Content -LiteralPath $global:LogTextFile -Raw
}
try {
    Reset-Fixture
    $value=Invoke-PreparedStep {Write-Warning 'Direct warning';Write-Host "[2026-01-01 01:00:00] WARNING: Child warning`nNormal line";Write-Information 'Information value' -InformationAction Continue;[pscustomobject]@{Value=42}}
    Check ($value.Value -eq 42 -and @($value).Count -eq 1) 'Information/warning relay polluted the producer result'
    Check ($global:SmartM365WarningCount -eq 2) 'Count warning records and flattened child warnings once each'
    Invoke-PreparedStep {Write-Host 'AVERTISSEMENT : Localized child warning'}
    Check ($global:SmartM365WarningCount -eq 3) 'Localized warning prefix is not counted'
    $child=Join-Path $root 'synthetic-child.ps1'
    "Write-Warning 'Repeated native warning'; Write-Warning 'Repeated native warning'; Write-Host 'native normal output'" | Set-Content -LiteralPath $child
    $shell=(Get-Process -Id $PID).Path
    Invoke-PreparedStep {
        & $shell -NoProfile -ExecutionPolicy Bypass -File $child 2>&1 | ForEach-Object {Write-Host ('[2026-01-01 01:00:00] '+[string]$_)}
        if($LASTEXITCODE -ne 0){throw 'Synthetic warning child failed'}
    }
    Check ($global:SmartM365WarningCount -eq 5) 'Native process warnings were dropped or duplicate messages were collapsed'
    Reset-Fixture;$script:childWarnings=8;$script:batchUploadEnabled=$true;RetentionControl 'Blocked'
    . $normal
    Check ($global:csvGeneratedPaths.Count -eq 2) 'Prepared CSV count includes artifacts or missed publication'
    Check ($global:SmartM365WarningCount -eq 9) 'Eight generator warnings plus one blocked-retention warning must survive'
    $text=Finish-Fixture
    Check ($text -match 'GeneratedCsvFiles: 2' -and $text -match 'CSV files : 2') 'Shared summary/banner disagrees with prepared CSV count'
    Check ($text -match 'Warnings\s*: 9' -and $text -match 'COMPLETED WITH WARNINGS') 'Shared completion still reports clean success'
    Check ($text.Contains('Consumer registry is not explicitly qualified.') -and $text.Contains('Retention target mismatch: TenantId')) 'Retention blockers must be present in the main log'
    Check (-not $text.Contains('UNEXPECTED nested completion banner')) 'Embedded retention emitted a premature banner'
    Check ((Get-Content -LiteralPath (Join-Path $product 'banner-flag.txt') -Raw) -eq 'True') 'Wrapper did not suppress the embedded retention banner'
    # Idempotent registration must not double count a batch.
    Register-PreparedRunCsvFiles $script:fixture
    Check ($global:csvGeneratedPaths.Count -eq 2) 'Repeated registration duplicates generated files'
    Reset-Fixture;$script:childWarnings=8;$script:pipelineFailure=$true
    $caught=$null;try{. $normal}catch{$caught=$_;WriteLog -Message $_.Exception.Message -Level ERROR}
    Check ($null -ne $caught -and $global:SmartM365WarningCount -eq 8 -and $global:csvGeneratedPaths.Count -eq 0) 'Failed generation lost warnings or fabricated published CSVs'
    $text=Finish-Fixture $caught
    Check ($text -match 'Status: Failed' -and $text -match 'GeneratedCsvFiles: 0') 'Failed generation summary is incorrect'
    Reset-Fixture;$script:batchUploadEnabled=$true;$script:transferFailure=$true
    $caught=$null;try{. $normal}catch{$caught=$_;WriteLog -Message $_.Exception.Message -Level ERROR}
    Check ($null -ne $caught -and $global:csvGeneratedPaths.Count -eq 2) 'Transfer failure erased already published CSV count'
    Reset-Fixture;$script:ValidateOnly=$true
    . $normal
    Check ($global:csvGeneratedPaths.Count -eq 0) 'Source validation fabricated generated CSVs'
    Reset-Fixture;$script:TransferOnly=$true;$script:batchUploadEnabled=$true
    . $transferBlock
    Check ($global:csvGeneratedPaths.Count -eq 0) 'Transfer-only fabricated CSV generation'
    Reset-Fixture
    Invoke-PostTransferRetention $script:transferReceipt
    Check ($global:SmartM365WarningCount -eq 0) 'Qualified preview was incorrectly counted as a warning'
    RetentionControl 'Preview' $true
    Invoke-PostTransferRetention $script:transferReceipt
    Check ($global:SmartM365WarningCount -eq 1) 'Retention exception was lost or invalidated publication'
    RetentionControl 'PartialOrBlocked'
    Invoke-PostTransferRetention $script:transferReceipt
    Check ($global:SmartM365WarningCount -eq 2) 'Partial retention was reported as clean success'
    foreach($throws in @($false,$true)){
        Reset-Fixture
        . $normal
        $script:cleanupFailure=$throws;$mappingRoot='synthetic';$data=$root;$transcriptStarted=$false;$coreLoaded=$true;$failure=$null
        . $cleanupBlock
        Check ($global:SmartM365WarningCount -eq 1 -and $global:SmartM365ExecutionStatus -eq 'CompletedWithWarnings') 'Entry cleanup warnings were lost at finalization'
        Check ($global:csvGeneratedPaths.Count -eq 2) 'Entry cleanup lost prepared-table counters'
    }
    foreach($mode in 'Failed','SchemaOnly','MissingFile','Duplicate','Unsafe'){
        Reset-Fixture
        $bad=$validReceipt | ConvertTo-Json -Depth 5 | ConvertFrom-Json
        switch($mode){
            Failed {$bad.Passed=$false}
            SchemaOnly {$bad.SchemaOnly=$true}
            MissingFile {$bad.Files=@($bad.Files[0])}
            Duplicate {$bad.Files[1].File=$bad.Files[0].File}
            Unsafe {$bad.Files[1].File='../Outside.csv'}
        }
        SaveReceipt $bad
        Reject {Register-PreparedRunCsvFiles $script:fixture}
        Check ($global:csvGeneratedPaths.Count -eq 0) 'Invalid receipt partially changed generated-file counters'
    }
    $finalText=$main[0].Finally.Extent.Text
    Check ($finalText -match "else\{'Auto'\}") 'Entry point forces Success despite warnings'
    $retentionAst=Read-Ast (Join-Path $repository 'SmartWorkplaceIntelligence/scripts/Invoke-PreparedSharePointRetention.ps1')
    Check ($retentionAst.Extent.Text -match 'if\(-not \$NoCompletionBanner\)') 'Retention entry lacks the embedded-call banner guard'
    Write-Output "PASS: $script:checks synthetic counter and retention-diagnostic checks. No real preparation, network, cleanup or Power BI refresh."
} finally {
    foreach($name in $saved.Keys){if($saved[$name].Exists){Set-Variable -Name $name -Scope Global -Value $saved[$name].Value}else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBUXahrHbhNPy0f
# +gOk9/he/IQOYfeaBReg3K2ATX5IhKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMJZP+sFetLTcCXxreB4gWYrLBjc+NrbszeMPaF95DuHMA0GCSqG
# SIb3DQEBAQUABIIBgFZ5Kwq6Ti4hpLlwltcMl6mVxtR5W5BwVJE8iGHwaIQlNhiy
# UOcmqz+kXv5XHMp3rQcLXXW6QZnKcNuZbVHTxikJvGnjzXbwf98gHZ+XejpqTe7R
# p4PsioTEQuWJ4UHOz+BXRDykTtaE7pvUxDTUA2SWr/uNhUR9emkDeNaL5coIIbEf
# TBaJkJ9mL8vl1+r3py4woWkuYNCBWRTc9L3NAPcw7kApYo8Cv8fkephbNGE1nW1F
# EqIIP6Ru7exczJ06zVDs3zDoH+lxOsdaUJpH+v+iZL4iauFyrdDiIvr2vytRPYj1
# nKA8Z1vqxb69jdPlfAKiKOJNsvdr1fv83t29cEotZqocFe06hkVfv2RxF80NeHB3
# u9f2aQ2CsIFtGuv0gSr97955lziKC4uHAc0RkZIXWX/IvMkSERfhNTGfP0iI2GNf
# HPQTaxNNsWPzWtL4owwfNn5iimu1NebNfKvFDS1iuHNv9Z5XoQ7LdcfBsH/eWlO/
# aeMswPTUYdmxlU0R7KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkyMzQ0
# NDZaMC8GCSqGSIb3DQEJBDEiBCDsFGD0tyu0NzO/o2UIbgEypSuAWOVn3twwZKX6
# HVvhCzANBgkqhkiG9w0BAQEFAASCAgCY+u1ZgG4Tp/xCgEtO//THiJxK8FSuqqeb
# CucESWBexoYr8j9GjRpWtavgg3DwrNERmW/DzDL7yoDgCN7qaCn8EenvXW/oquKx
# UwW5zmPjHgxa0dqfnRMgn2c5MLULLMbaIb6523NXB9rgGxPfs9ZTmTVjtlOdE/RF
# hEtrgjzPQNR2tE0T72A9mvc/YbYu3Wyz5RR2sYLr6+Py97EmjYbuZaW/m6M+1ppU
# oUE3/X4T988Zd0sX8E00uAD8v3ouVGx77RX2FYSFvFJL0QqogcZp+1MGzY269s2x
# qm1RND81TSSpt7d5iMXfvS/1N2uyld+K9wk+m3Tl1f5JFsWd8XcCIUK6T6nGRvG0
# zDiQwdZghqeZaVIJyRlHWw1lLT5T/ncvg71saisGa9yX8j210cjJgubFJ+9Y9NHI
# 0omWQPq6T3M3aQa9tqJQgdAqApLwHJetN+/QjXypHo/AF7ynPoWWFTnMQSHZngyE
# +m7x+9NDJNujW/GYhY3Txd+U8TMCJyC6JDe9EgssEcga07sJcFzcQkHUZONyCQQ4
# IMMm6wPp1RaGlahCPo4PUUs0XsASP3/zdLv/MDx3kbNdxURZHKoeRrwF2esOpVc3
# QYH3Cj3VpYL0FylBwuGnPo0AoBqkco1c8GwkvhfrERXGIz8YPMIRIUEc5/djDOQe
# bbGnvZHjKw==
# SIG # End signature block
