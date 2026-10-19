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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC8aMfh/9x6ag+W
# TNxVDOVIGdH4h7WoAJXgxDP0ihhEqaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICmc/2WH8PhhVru6pHA0XZ9AxF4ixqGf1IcD3EVH/kxiMA0GCSqG
# SIb3DQEBAQUABIIBgLC+RrIgRVr3vBzdDgsvNnJqT+iLkc87/gYxMsVF43tX6laz
# QcLa8n5yCX7YJEucNpoEAfXb/tp9l6oahvibfvxytuKMGSwIZC+9kFuAQ2Ie1wDW
# j3T0EXRSeKixXK+FVV9rEtVRAM1FNPo4cCVyn+wq8f98hVrhCfpf9OC3vi9K9L0I
# qUJcqCwp5L4rzeNDrOaJvMp0/cVX971pBZBTXUq09TndzIlSVGLc85T11KVKW+uG
# CDeKZbdTzmkc2JUwLCPcG320LDjmWFtBbWK5JLO1P1LnASWBxnrHm04TS37kwT74
# AHHifqJvy5LTsyBcQbeOEHm4LUzqL9lu1yql/LFRHj0WDapfgkw9WioIrwtgFSl5
# N+6Yw5kH3j2xLFSCaAH1Dp5E2S9WUxAvP61okKF5Af8KnDdzIYRgSlk2tPTlfK+n
# P9+7+uHJTKOES5IwzjO/pNdRxa09B0v7bf8TeBQEPc8b+5Z82/ve+MblY2xqnHBm
# 25iBso1uOGs/xjemRKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkyMzQ0
# NDZaMC8GCSqGSIb3DQEJBDEiBCC4iR/sBMFAlauroj7th0fdboHdL6J2nu349T5j
# yg5/SzANBgkqhkiG9w0BAQEFAASCAgBceZ9t/wgIiuT0lPgKCgU2bDGDKytsIuqh
# eNV4m6adHlh1AHEvsMxFvqQvKfjoIDUTRcM0yLSaShS9HYyBqcE9hANPEE1ZSJCt
# Yy3giF/d5LcCW+ieLvhZqvZb0ae70G6pw6qIwJAVYD4Txuux6HJHn3EfFKS2o7+V
# VEn1qEhvJJ/lhHpJg3MnIgjRIwe38334co3T6m2yIbXJoqw4GbrlN/Hgrm9sbNzR
# 0GzISC4p7ZKyijMnm7+s+Pyx9+G4m5nI1XioqWCNCKSVDd/wPjGvCC7OhCxDh8U8
# 0SuZrDRmuGfMf7a/0fVuKMlCxUe+EzYGEEPhKy9ezc/xtUFEovisdHAHrz5b0cMA
# 2p1BZ+nQKyyKLlEjY1Hg/+Q4EgWHVUmDZMW/G6WWAf3KQGioqvjy58K7mDprTOGn
# K+zvMWHCXINnKf/ePxloresE5ld/+e67UINAjOv+alAT7z8jOmejzZ+bDbIyU2qZ
# M4Xn0MMOWHxXraItQhQyXDwsJq+OLnow3Dl3pmwTbkASTZgmsg15YEgrZ+179Jq5
# E9y3/AaLgudbDvN+v+1UwzGKQml/g2BCyOQogX9ySUi9CtYU/akPhHICbpCn1z1V
# +uTqDoAUVxTPs7Hhgx8aZor4PC5PoE3Q/qZ8N74C0dIcSdjLLPUvXOpdsa3MT4nP
# rp9kNBF0dA==
# SIG # End signature block
