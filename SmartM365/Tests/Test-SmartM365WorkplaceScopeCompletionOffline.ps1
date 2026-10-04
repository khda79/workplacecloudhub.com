#Requires -Version 7.0
<#
.SYNOPSIS
Offline transcript lifecycle and truthful trace-publication status regression tests.
.VERSION
1.0.0
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Isolated synthetic Core execution globals in this offline test process.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Transcript mocks and test-owned temporary fixtures only.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Module-scoped transcript closure mock; the real cmdlet is explicitly qualified in the integration case.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='The local transcript integration scenario needs a visible synthetic activity line.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-TraceTest-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryRoot | Out-Null
$results = [Collections.Generic.List[object]]::new()
function Assert-Trace([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Read-TraceFunction([string]$Path, [string[]]$Names) {
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw "Source parser failed: $Path" }
    foreach ($name in $Names) {
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if ($null -eq $node) { throw "Missing function: $name" }
        $node.Extent.Text
    }
}
# Extract only function definitions. No tenant configuration, credentials or collector runs.
$definitions = @(Read-TraceFunction (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psm1') @('Complete-SmartM365ExecutionContext','Write-SmartM365CompletionBanner'))
$definitions += @(Read-TraceFunction (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psm1') @('Complete-SmartM365EvidenceRuntime'))
$module=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
try {
    foreach ($scenario in @(
        @{Name='Success';Enabled=$true;Expected='Success'},
        @{Name='Disabled';Enabled=$false;Expected='Success'},
        @{Name='InitialTranscriptFailure';Enabled=$true;Fail='transcript-1';Expected='CompletedWithWarnings'},
        @{Name='FinalTranscriptFailure';Enabled=$true;Fail='transcript-2';Expected='CompletedWithWarnings'},
        @{Name='FinalLogFailure';Enabled=$true;Fail='log-3';Expected='CompletedWithWarnings'},
        @{Name='ReceiptFailure';Enabled=$true;Fail='receipt-1';Expected='CompletedWithWarnings'},
        @{Name='MailArtifactFailure';Enabled=$true;Fail='mail-1';Expected='CompletedWithWarnings'},
        @{Name='ThrowingUpload';Enabled=$true;Fail='transcript-1';Throw=$true;Expected='CompletedWithWarnings'},
        @{Name='ClosureFailure';Enabled=$true;CloseFailure=$true;Expected='CompletedWithWarnings'},
        @{Name='PriorWarning';Enabled=$true;PriorWarning=$true;Expected='CompletedWithWarnings'},
        @{Name='CollectionFailure';Enabled=$true;CollectionFailure=$true;Expected='Failed'},
        @{Name='RealTranscript';Enabled=$true;RealTranscript=$true;Expected='Success'}
    )) {
        foreach ($option in @('Fail','Throw','CloseFailure','PriorWarning','CollectionFailure','RealTranscript')) {
            if (-not $scenario.ContainsKey($option)) { $scenario[$option]=$false }
        }
        try {
            $scenarioRoot=Join-Path $temporaryRoot $scenario.Name
            New-Item -ItemType Directory -Path $scenarioRoot | Out-Null
            $observed = & $module {
                param($Case,$Folder)
                $script:Case=$Case; $script:Closed=$false; $script:ClosureCalls=0
                $script:Uploads=[Collections.Generic.List[object]]::new()
                $script:Attempts=@{}
                $global:SmartM365ExecutionSummaryWritten=$false
                $global:SmartM365CompletionBannerRunKey=''
                $global:SmartM365ExecutionStatus=''
                $global:SmartM365ExecutionStartTime=Get-Date
                $global:SmartM365WarningCount=if ($Case.PriorWarning) { 1 } else { 0 }
                $global:SmartM365ErrorCount=0
                $global:EnableSharePointUpload=[bool]$Case.Enabled
                $global:SmartM365ScriptName='SmartM365-WorkplaceScope-Inventory'
                $global:SmartM365ProfileKey='test'; $global:SmartM365TenantKey='contoso-test'
                $global:SmartM365OrganizationKey='contoso'; $global:SmartM365EnvironmentKey='test'
                $global:SmartM365TenantId=''; $global:BasePath=$Folder
                $global:LogTextFile=Join-Path $Folder 'run.log'
                $global:logTranscriptFile=Join-Path $Folder 'run.transcript.log'
                $script:Receipt=Join-Path $Folder 'receipt.json.txt'
                $mail=Join-Path $Folder 'mail.html'
                $csv=Join-Path $Folder 'unchanged.csv'
                foreach ($path in @($global:LogTextFile,$global:logTranscriptFile,$script:Receipt,$mail,$csv)) {
                    Set-Content -LiteralPath $path -Value 'synthetic preserved evidence' -Encoding utf8
                }
                $before=(Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash
                $global:csvGeneratedPaths=@($csv)
                $global:SmartM365SharePointUploadedFiles=[Collections.ArrayList]::new()
                $global:SmartM365MailHtmlFiles=@($mail)
                function WriteLog {
                    param($Message,$Level='INFO')
                    if ($Level -eq 'WARNING') { $global:SmartM365WarningCount++ }
                    if ($Level -eq 'ERROR') { $global:SmartM365ErrorCount++ }
                    Add-Content -LiteralPath $global:LogTextFile -Value $Message -Encoding utf8
                }
                function Complete-SmartM365CmdbSourceReceipt {
                    param($Status,$ErrorCount)
                    $script:ReceiptStatus=$Status
                    $script:ReceiptErrors=$ErrorCount
                    $script:Receipt
                }
                function Stop-Transcript {
                    $script:ClosureCalls++
                    if ($script:Case.CloseFailure) { throw 'Synthetic transcript closure failure.' }
                    if ($script:Case.RealTranscript) { Microsoft.PowerShell.Host\Stop-Transcript | Out-Null }
                    $script:Closed=$true
                }
                function Invoke-SmartM365SharePointCsvUpload {
                    param($LocalFilePath)
                    if (-not $global:EnableSharePointUpload) { return }
                    $kind=if ($LocalFilePath -eq $global:logTranscriptFile) {'transcript'} elseif ($LocalFilePath -eq $global:LogTextFile) {'log'} elseif ($LocalFilePath -eq $script:Receipt) {'receipt'} else {'mail'}
                    if ($kind -eq 'transcript' -and -not $script:Closed) { throw 'Attempted upload of active transcript.' }
                    if (-not $script:Attempts.ContainsKey($kind)) { $script:Attempts[$kind]=0 }
                    $script:Attempts[$kind]++
                    $attempt="$kind-$($script:Attempts[$kind])"
                    $content=Get-Content -LiteralPath $LocalFilePath -Raw
                    $script:Uploads.Add([pscustomobject]@{Kind=$kind;Attempt=$attempt;Content=$content})
                    if ($script:Case.Fail -eq $attempt) {
                        if ($script:Case.Throw) { throw 'Synthetic upload exception.' }
                        WriteLog 'Synthetic upload failure.' WARNING
                        return
                    }
                    [void]$global:SmartM365SharePointUploadedFiles.Add($LocalFilePath)
                    [pscustomobject]@{LocalFilePath=$LocalFilePath}
                }
                if ($Case.RealTranscript) {
                    Microsoft.PowerShell.Host\Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
                    Write-Host 'Synthetic collector activity before completion.'
                }
                $status=if ($Case.CollectionFailure) {'Failed'} else {'Success'}
                $errorRecord=if ($Case.CollectionFailure) {[pscustomobject]@{Exception=[Exception]::new('Synthetic collection failure.')}} else {$null}
                Complete-SmartM365EvidenceRuntime -Status $status -ErrorRecord $errorRecord -FailureStage WorkplaceScopeInventory -CloseTranscriptBeforeUpload
                $calls=$script:Uploads.Count
                # Repeated terminal cleanup must not close or upload twice.
                Complete-SmartM365EvidenceRuntime -Status $status -CloseTranscriptBeforeUpload
                [pscustomobject]@{
                    Status=$global:SmartM365ExecutionStatus;Warnings=$global:SmartM365WarningCount;Errors=$global:SmartM365ErrorCount
                    Uploads=@($script:Uploads.ToArray());ClosureCalls=$script:ClosureCalls;CallsBeforeRepeat=$calls
                    Log=(Get-Content -LiteralPath $global:LogTextFile -Raw)
                    Transcript=(Get-Content -LiteralPath $global:logTranscriptFile -Raw)
                    CsvPreserved=($before -ceq (Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash)
                }
            } $scenario $scenarioRoot
            Assert-Trace ($observed.Status -ceq $scenario.Expected) "Incorrect status: $($scenario.Name) -> $($observed.Status)"
            Assert-Trace $observed.CsvPreserved 'Trace publication altered the business CSV.'
            Assert-Trace ($observed.ClosureCalls -eq 1) 'Transcript closure was repeated.'
            Assert-Trace ($observed.Uploads.Count -eq $observed.CallsBeforeRepeat) 'Repeated cleanup republished artifacts.'
            Assert-Trace ($observed.Uploads.Count -le 7) 'Publication is not bounded.'
            $label=switch($scenario.Expected) {'Failed' {'FAILED'} 'CompletedWithWarnings' {'COMPLETED WITH WARNINGS'} default {'SUCCESS'}}
            Assert-Trace ($observed.Log -match "Status\s+: $label") 'Final log banner missing.'
            if ($scenario.CloseFailure) {
                Assert-Trace (@($observed.Uploads | Where-Object Kind -eq transcript).Count -eq 0) 'Closure failure uploaded an active transcript.'
            }
            else {
                Assert-Trace ($observed.Transcript -match "Status\s+: $label") 'Closed transcript lost the final banner.'
                if ($scenario.Expected -eq 'Success' -and $scenario.Enabled) {
                    $last=@($observed.Uploads | Where-Object Kind -eq transcript)[-1]
                    Assert-Trace ($last.Content -match 'Status\s+: SUCCESS') 'Published transcript lacks its completion banner.'
                    Assert-Trace (($last.Content -split 'Status\s+: SUCCESS').Count -eq 2) 'Success banner duplicated in transcript.'
                }
            }
            if (-not $scenario.Enabled) { Assert-Trace ($observed.Uploads.Count -eq 0) 'Disabled external actions uploaded traces.' }
            if ($scenario.Expected -eq 'CompletedWithWarnings') { Assert-Trace ($observed.Warnings -gt 0) 'Publication warning omitted from final counters.' }
            if ($scenario.CollectionFailure) { Assert-Trace ($observed.Errors -eq 1) 'Collection failure was masked or double counted.' }
            $results.Add([pscustomobject]@{Name=$scenario.Name;Passed=$true;Error=''})
        }
        catch { $results.Add([pscustomobject]@{Name=$scenario.Name;Passed=$false;Error=$_.Exception.Message}) }
    }
    $collector=Get-Content -LiteralPath (Join-Path $root 'SmartInventory/M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1') -Raw
    Assert-Trace ($collector -match '-CloseTranscriptBeforeUpload:\$transcriptStarted') 'WorkplaceScope did not delegate transcript ownership.'
    Assert-Trace ($collector -match "-MinimumVersion '1.0.71'") 'WorkplaceScope Core version guard missing.'
    $results | Format-Table -AutoSize
    if (@($results | Where-Object { -not $_.Passed }).Count) { throw 'Offline transcript completion regression failed.' }
    Write-Output ("PASS: {0} offline completion scenarios; no authentication, mail or SharePoint calls." -f $results.Count)
}
finally {
    Remove-Module $module -Force
    # Exact test-owned temporary directory only.
    $resolvedTemporary=[IO.Path]::GetFullPath($temporaryRoot)
    $temporaryParent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if (-not $resolvedTemporary.StartsWith($temporaryParent,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test cleanup target.' }
    Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAh5cc5Mv2kB3Oa
# HXatEklKLcixu6rKhsDr/uGCM9vE9KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILSXxJ6vP66SBqHnvV5xaMzvlnLwcKWG73TJo/Ob5dADMA0GCSqG
# SIb3DQEBAQUABIIBgJcbBbb/8vo7pPvEKCLt24e9UZoLWQAZT+KS25HxheZtRknM
# QvsL83AipK51MdIQkNumNZjyY6AOyCAb6Xi1Pd4K9XEf1p/yMLe82T6d1hbitZJZ
# uW2bn8X+Ib8aS2tswLrpxTAkCVtpXXKcxlpnSjTYkM6JxNDBF7+62fogQB9FzgRd
# XDHISdCs5V1tjyaGL4LvrT/iJo0zUlM28Y0x7mcqyWUNSzDLhKboDZvSM9jozpF7
# CW6h36LKibtza0C9EJnS4LkhO+1bzja/DC3rju72/XHg/9lv/i8QubaSkKz4dj19
# FXT7QNrn301itgopHbjBOpd1GQnAMYfnyouTuD0osMJhhDl2OinsvUHtmREz1U40
# HnEHnqJYJkqLrfl+0/ckrUiaqbANuUZ0s2XqwpWEFCMWOc+XX54tf3trhnt5eA1I
# NVPAovfPrAKbisXI1kuQCrNH8x+4+cMvpHpwbbrUd4sSAjIPz3kgV7xmkx88hKWp
# iOwAFlIMG8Ws341Ux6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQyMjAw
# NTBaMC8GCSqGSIb3DQEJBDEiBCAsN6kyyv/gzSV4JcHRoV2IB3iykEklWc0GAL6g
# 1sWkMzANBgkqhkiG9w0BAQEFAASCAgCbuGND1op1kG4CpClISl/iZKxsLjY/YHzk
# 2AJhe+esZseui1+UtRTKAgxpE5pz4TNY/zZXAw8w+Rf/Lzb0KEMC0zeMbmw+EDLX
# T/3W8zXutx8hCUHs6mGoCc0kyl3S9UuXJtAeAEGwFPs+WfOwONvh+o6opK7pKzOj
# pc3RR1b6ar6fCA1t1tuNGtQs/vqibrmXyGTlnDEmDwgxD8YSh4MEBDbDPVxZF17f
# uj5Eq4xCtD1tDM9cwFE1W5zFw4LfFx8LepIBAaGA5PNepHV29b/4ZeVn82SwrGGl
# DRl07Cd1QjebINRbTmZileIR8dJ8hAa1CLvN38QIUf82lWmQgijYXJlgF6PjQc6Y
# WQZTntghyAQ7OuQhUueTzd1ddzPutH7EhyN01ju1oaqcJK68bmR2Iw0Ux1ApPZOr
# nu0DJFoowZ1/1DZFZdCX/AKeNO63lEKI/k3NRAW0w5ctaZPI+p1QUlnm+RLnBQPy
# ZOC9h2tR6VqNhmCJ/oXt99CKJ23MgqlE8U53iHzGKFhU8L0o04OYJKof4nD/bck+
# DYRsWG99LJ7lSRiEEGXZGGuT0shdCiV1z/eHhVRpZF2VcCzlzcmnd48VWDH6Zupv
# Iup0R0auY8lAuD0cdGxRag0J8Vozdft+I3V6hPnQSmqklDDB470h0LogDTcouHMZ
# QDtfytCFPg==
# SIG # End signature block
