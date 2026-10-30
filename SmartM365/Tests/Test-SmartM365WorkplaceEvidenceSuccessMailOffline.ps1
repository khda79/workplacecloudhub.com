#Requires -Version 7
<#
.SYNOPSIS
    Synthetic completion-email regression tests; no tenant or transport is executed.
.VERSION
    1.0.1
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
    @($_.Body.Statements | Where-Object {$_.Extent.Text -match '^\$result\s*=\s*Invoke-PreparedStep\s*\{\s*Invoke-PreparedEvidencePipeline'}).Count -eq 1
})
if($main.Count -ne 1){throw 'Expected one preparation try block.'}
$statements=@($main[0].Body.Statements)
$start=0
while($start -lt $statements.Count -and $statements[$start].Extent.Text -notmatch '^\$result\s*=\s*Invoke-PreparedStep\s*\{\s*Invoke-PreparedEvidencePipeline'){$start++}
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
function Invoke-PreparedStep {param([scriptblock]$Action) & $Action}
function Register-PreparedRunCsvFiles {param($PreparedResult)}
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
    Check (($script:events -join ',') -eq 'prepare,transfer,mail,teams,retention') 'Email must follow verified transfer and precede post-transfer retention.'
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
    Check (($script:events -join ',') -eq 'transfer,mail,retention') 'Transfer-only sends one email after verification and before retention, without generation.'
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
    Check (($script:events -join ',') -eq 'prepare,transfer,mail') 'Mail failure must not restart preparation or transfer.'
    Write-Output "PASS: $script:passed synthetic success-mail assertions. No preparation, tenant, network or real email."
} finally {
    foreach($original in @(@{Name='SmartM365ExecutionStartTime';Value=$savedStart},@{Name='LogTextFile';Value=$savedLog})) {
        if($original.Value){Set-Variable -Name $original.Name -Scope Global -Value $original.Value.Value}
        else{Remove-Variable -Name $original.Name -Scope Global -ErrorAction SilentlyContinue}
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAk+DDR6uC/4VZv
# 7vw/Ahd0jvzjctskMhWyeT2MXRyge6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC2PxKOq75d6wr8S/MBKUa4
# HZs4CZ25HyDl84HX6yOW7jANBgkqhkiG9w0BAQEFAASCAYBmKNVRCtr0r/mKhkxq
# 18W4T0xctMfkHilHe6ZQ8ZIsuM9JENCoQB22OcvqjAVNppCPehdOjtHP00tjS39p
# l+klOFwky+fYAO5gWh2KhMrTtw9DkQmxWQKXkp4ewjzFoC5ndH65Lyk+H7cWC1b2
# FQIEDO3B2s+O+9GJp4lQoHkbeE27sdT2pbppjt1dR7s6qNOl3eE65zKq68GcgpRd
# 8KqF7cfZYn9rFH3KDGJueVHo5x3YJYMny53Q/zmrhOTSMZgXVzMKAGVztmf3Vtrm
# 2Wo0PK5t4rRbVB97AP0GjqmvEhWjvrLa5g5A4vE6S+y+e0kJWUXG2sda7PnGEsqa
# q5wyOjFLhuJOEiWFGaszeLXOBNy8cOyWYEmbKnKma3eSHwxPQnJ96BCwD5T3lFzC
# 2rU+ozqHuWtuQ9/wsuWc4VF1SH0ESmIX/6cVDvBI9Tay1mNsmvr6DMANDodthRMJ
# E3HNzycQQQEHJzhKbBR2GDfuLrZlzzMbR3iK7JD/JYH28Oc=
# SIG # End signature block
