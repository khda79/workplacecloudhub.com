#Requires -Version 7
<#
.SYNOPSIS
    Synthetic completion-email regression tests; no tenant or transport is executed.
.VERSION
    1.0.0
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Transport stubs retain production signatures without performing I/O.')]
param()
$ErrorActionPreference='Stop'
$repository=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$entry=Join-Path $repository 'SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1'
function Read-Ast([string]$Path) {
    $tokens=$null; $parseErrors=$null
    $tree=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw ($parseErrors.Message -join '; ')}
    return $tree
}
$ast=Read-Ast $entry
$transferAst=Read-Ast (Join-Path $repository 'SmartWorkplaceIntelligence/scripts/PreparedSharePointTransfer.psm1')
foreach($definition in @(@{Tree=$ast;Name='Send-PreparedSuccessMail'},@{Tree=$transferAst;Name='Assert-PreparedTransferResult'})) {
    $functions=@($definition.Tree.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true) | Where-Object Name -eq $definition.Name)
    if($functions.Count -ne 1){throw "Expected one function: $($definition.Name)"}
    . ([scriptblock]::Create($functions[0].Extent.Text))
}
# Execute the real entry-point tail and transfer branch with stubbed producers and mail.
# This covers placement/guards as well as the HTML builder, without importing Core.
$main=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.TryStatementAst]},$true) | Where-Object {
    @($_.Body.Statements | Where-Object {$_.Extent.Text -match '^\$result\s*=\s*Invoke-PreparedEvidencePipeline'}).Count -eq 1
})
if($main.Count -ne 1){throw 'Expected one preparation try block.'}
$statements=@($main[0].Body.Statements)
$start=0
while($start -lt $statements.Count -and $statements[$start].Extent.Text -notmatch '^\$result\s*=\s*Invoke-PreparedEvidencePipeline'){$start++}
$normal=[scriptblock]::Create(($statements[$start..($statements.Count-1)].Extent.Text -join "`n"))
$transferBranch=@($statements | Where-Object {$_.Extent.Text -match '^if\s*\(\$TransferOnly\)'})
if($transferBranch.Count -ne 1){throw 'Expected one transfer-only branch.'}
$transferOnlyBlock=[scriptblock]::Create($transferBranch[0].Extent.Text)
$script:passed=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:passed++}
function Initialize-Fixture {
    $script:Offline=$false; $script:ValidateOnly=$false; $script:WorkforceDiagnostic=$false
    $script:RepairLegacyHistory=$false; $script:ConvertMetadata=$false; $script:TransferOnly=$false
    $script:batchUploadEnabled=$true; $script:transfer=$null; $script:params=@{}; $script:transferParameters=@{}
    $script:effective=[pscustomobject]@{TenantKey='synthetic<&>'}
    $script:output='C:\Synthetic\DATA-POWERBI<&>'
    $script:scriptName='SmartM365-WorkplaceEvidence-Prepare'
    $script:ExpectedBatchId='20261020T100000000Z-00000001'
    $script:smartM365RunGuard=[pscustomobject]@{ReachedTerminal=$false;ExplicitExit=$false}
    $script:fixture=[pscustomobject]@{BatchId=$script:ExpectedBatchId;Files=2;BatchPath='C:\Synthetic\batch';SourceFiles=3;Identity=@{CsvFiles=3;Rows=4}}
    $script:receipt=[pscustomobject]@{Status='PublishedAndReadBack';BatchId=$script:ExpectedBatchId;CsvFiles=2;ExpectedFiles=6;VerifiedFiles=6;ManifestSHA256=('a'*64);PointerVerified=$true;CsvRecalculated=$false;AuditPath='synthetic'}
    $script:events=[Collections.Generic.List[string]]::new()
    $script:mails=[Collections.Generic.List[object]]::new()
    $script:pipelineFails=$false; $script:transferFails=$false; $script:mailFails=$false
    $script:mailPhase=$null
}
function Invoke-PreparedEvidencePipeline {
    $script:events.Add('prepare')
    if($script:pipelineFails){throw 'Synthetic preparation failure'}
    return $script:fixture
}
function Send-PreparedEvidenceBatch {
    param($ExpectedBatchId)
    $script:events.Add('transfer')
    if($script:transferFails){throw 'Synthetic transfer failure'}
    return $script:receipt
}
function Invoke-PostTransferRetention {param($VerifiedTransfer) $script:events.Add('retention')}
function WriteLog {param($Message,$Level)}
function Send-SmartM365TeamsNotification {param($Title,$Message,$Level,$Channel,$ResultSummary,$Facts) $script:events.Add('teams')}
function SendEmailHtmlReport {
    param($Subject,$BodyHtml,$MailPurpose,[switch]$AllowAttachments)
    $script:events.Add('mail')
    $script:mails.Add(@{}+$PSBoundParameters)
    $script:mailPhase=$phase
    if($script:mailFails){throw 'Synthetic mail failure'}
}
function Reject([scriptblock]$Action,[string]$Pattern) {
    $caught=$null
    try{& $Action}catch{$caught=$_}
    Check ($null -ne $caught -and $caught.Exception.Message -match $Pattern) "Expected failure: $Pattern"
}
$savedStart=Get-Variable SmartM365ExecutionStartTime -Scope Global -ErrorAction SilentlyContinue
$savedLog=Get-Variable LogTextFile -Scope Global -ErrorAction SilentlyContinue
try {
    Set-Variable -Name SmartM365ExecutionStartTime -Scope Global -Value (Get-Date).AddMinutes(-2)
    Set-Variable -Name LogTextFile -Scope Global -Value 'C:\Synthetic\run<&>.log'
    Initialize-Fixture
    . $normal
    Check (($script:events -join ',') -eq 'prepare,transfer,retention,mail,teams') 'Email must follow validated preparation and verified transfer/retention.'
    Check ($script:mails.Count -eq 1) 'One completion email per successful normal run.'
    $mail=$script:mails[0]
    Check ($mail.Subject -eq "$script:scriptName completed - batch $script:ExpectedBatchId") 'Subject identifies the completed batch.'
    Check ($mail.MailPurpose -eq 'Report' -and -not $mail.AllowAttachments) 'Report routing without attachments.'
    Check (-not $mail.ContainsKey('To') -and -not $mail.ContainsKey('From') -and -not $mail.ContainsKey('SendMailMode')) 'Shared private mail configuration remains authoritative.'
    Check ($mail.BodyHtml.Contains('synthetic&lt;&amp;&gt;') -and $mail.BodyHtml.Contains('run&lt;&amp;&gt;.log')) 'Dynamic HTML values must be escaped.'
    Check (-not $mail.BodyHtml.Contains('synthetic<&>')) 'No unescaped tenant value.'
    foreach($text in 'Batch ID','Prepared CSV files','Duration','Output root','Completed at','Verified SharePoint files','Batch files and current pointer read back and verified') {
        Check ($mail.BodyHtml.Contains($text)) "Summary missing: $text"
    }
    Initialize-Fixture; $script:batchUploadEnabled=$false
    . $normal
    Check (($script:events -join ',') -eq 'prepare,mail,teams') 'Disabled upload must not invoke transfer.'
    Check ($script:mails[0].BodyHtml.Contains('Disabled; local batch only') -and -not $script:mails[0].BodyHtml.Contains('Verified SharePoint files')) 'Do not imply cloud publication for local-only preparation.'
    Initialize-Fixture; $script:TransferOnly=$true
    . $transferOnlyBlock
    Check (($script:events -join ',') -eq 'transfer,retention,mail') 'Transfer-only sends one email after verification, without generation.'
    Check ($script:mails.Count -eq 1 -and $script:mails[0].BodyHtml.Contains('Existing-batch transfer (no CSV generation)')) 'Transfer-only summary must identify its operation.'
    foreach($mode in 'Offline','ValidateOnly','WorkforceDiagnostic','RepairLegacyHistory','ConvertMetadata') {
        Initialize-Fixture
        Set-Variable -Name $mode -Value $true -Scope Script
        Send-PreparedSuccessMail -BatchId $script:ExpectedBatchId -CsvFiles 2
        Check ($script:mails.Count -eq 0) "No success mail in $mode mode."
    }
    foreach($mode in 'Offline','ValidateOnly','WorkforceDiagnostic') {
        Initialize-Fixture; $script:batchUploadEnabled=$false
        Set-Variable -Name $mode -Value $true -Scope Script
        . $normal
        Check ($script:mails.Count -eq 0) "Entry-point tail must not notify in $mode mode."
    }
    Initialize-Fixture; $script:pipelineFails=$true
    Reject { . $normal } 'Synthetic preparation failure'
    Check ($script:mails.Count -eq 0) 'No success mail after preparation failure.'
    Initialize-Fixture; $script:transferFails=$true
    Reject { . $normal } 'Synthetic transfer failure'
    Check ($script:mails.Count -eq 0) 'No success mail after transfer failure.'
    foreach($bad in 'pointer','batch','count','receipt') {
        Initialize-Fixture
        switch($bad) {
            pointer {$script:receipt.PointerVerified=$false}
            batch {$script:receipt.BatchId='different-batch'}
            count {$script:receipt.CsvFiles=1}
            receipt {$script:receipt=[pscustomobject]@{Status='PublishedAndReadBack'}}
        }
        Reject { . $normal } 'receipt|CSV counts'
        Check ($script:mails.Count -eq 0) "No success mail for invalid $bad verification."
    }
    Initialize-Fixture; $script:TransferOnly=$true; $script:receipt.PointerVerified=$false
    Reject { . $transferOnlyBlock } 'receipt'
    Check ($script:mails.Count -eq 0) 'Transfer-only must not notify after failed pointer verification.'
    Initialize-Fixture; $script:mailFails=$true
    Reject { . $normal } 'Synthetic mail failure'
    Check ($script:mailPhase -eq 'Success email notification') 'Mail failure must identify the notification phase.'
    Check (($script:events -join ',') -eq 'prepare,transfer,retention,mail') 'Mail failure must not restart preparation or transfer.'
    Write-Output "PASS: $script:passed synthetic success-mail assertions. No preparation, tenant, network or real email."
} finally {
    foreach($original in @(@{Name='SmartM365ExecutionStartTime';Value=$savedStart},@{Name='LogTextFile';Value=$savedLog})) {
        if($original.Value){Set-Variable -Name $original.Name -Scope Global -Value $original.Value.Value}
        else{Remove-Variable -Name $original.Name -Scope Global -ErrorAction SilentlyContinue}
    }
}
