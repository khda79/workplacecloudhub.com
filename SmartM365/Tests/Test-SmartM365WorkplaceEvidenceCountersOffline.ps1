#Requires -Version 7
<#
.SYNOPSIS
    Synthetic entry-point counters and retention diagnostics; no real preparation or transport.
.VERSION
    1.0.1
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
function ConvertTo-SmartM365InventoryConfigLocation {param($Path) $Path}
function Get-SmartM365JsonReadPath {param($Path) $Path}
function Invoke-SmartM365TeamsNotificationFromLog {param($Message,$Level)}
function Invoke-SmartM365SharePointCsvUpload {param($LocalFilePath)}
function Test-SmartM365SourceReceiptMirrorBlocked {param($Path) $false}
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
    # Children attached to a VT console colour their warnings (production job logs: ESC[33;1mWARNING: ...).
    $colored=Join-Path $root 'synthetic-colored-child.ps1'
    "`$PSStyle.OutputRendering='Ansi'; Write-Warning 'Colored native warning'; Write-Host 'colored normal output'" | Set-Content -LiteralPath $colored
    Invoke-PreparedStep {
        & $shell -NoProfile -ExecutionPolicy Bypass -File $colored 2>&1 | ForEach-Object {Write-Host ('[2026-01-01 01:00:00] '+[string]$_)}
        if($LASTEXITCODE -ne 0){throw 'Synthetic colored warning child failed'}
    }
    Check ($global:SmartM365WarningCount -eq 6) 'An ANSI-colored native warning was not counted'
    $escape=[char]27
    Invoke-PreparedStep {Write-Host "[2026-01-01 01:00:00] $escape[33;1mWARNING: Escaped relay warning$escape[0m"}
    Check ($global:SmartM365WarningCount -eq 7) 'A relayed ANSI warning line was not counted'
    $logged=Get-Content -LiteralPath $global:LogTextFile -Raw
    Check ($logged -match 'Colored native warning' -and $logged -match 'Escaped relay warning' -and -not $logged.Contains([string]$escape)) 'ANSI warnings were not logged as plain text'
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCJiN6KjH0MmmDk
# cA+dog5GzK1INDEPwizC6tDwd/+GaqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAd/TEdvUwmx+nbXMz3jxC1
# Lrxd3+gk8WNFoVVZHvNJsTANBgkqhkiG9w0BAQEFAASCAYCJl7Y6hGGp1/FeQrMG
# 9Rc7tqDdEEKwAxehGvJt7r9fps0bDOf6CbXaM7zDL1fwqosQ5sZZ6Cc6zK+4iBcc
# dqtfsZD/L3KtUXGvKJ5ircQy+2mlpw2+qHbVM688gOsmYW2xiLOuOedjQqW3q3St
# e56MI9+XWXlfua83nYUk4bYLZlu60I8gj9Gn2OrkBt3+vNpWwAtv+hZ2+gVTx7mn
# tk4xaGotdZNVr0ODfs/W3HL9Lsu1xeVEICmo8eHxUJnZI6I4ZIvgQ14DeMTrGLl8
# 3FK3q7US9dXZmV4rjoHwngz4HdfER2lt22ONvCfShkOn1fDz87IuIIn0yFMiCPmY
# RDMu/R8fe2FyXn1+WhoypbknX/XnhMp5QrrHbBCmFU6MpfG6ECWnUAta0J1xlfic
# lmMbs+XgILiPNrOuBnjU2E1FTjE2PWw0aLIjl0tyDQN2pVuJxqgANgeSyzkS2vzw
# 8sLdK2I/UqrqX9WKx3Ijrw7Py6emYTCo+J7mqDcCbML+Xxg=
# SIG # End signature block
