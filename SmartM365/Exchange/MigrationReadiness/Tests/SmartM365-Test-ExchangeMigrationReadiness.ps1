#requires -Version 7.0
<#
.SYNOPSIS
Offline regression tests. All service commands are local fixtures; no authentication or migration.
.VERSION
1.11.17
#>
[CmdletBinding()]
param([string]$OutputRoot = (Join-Path ([IO.Path]::GetTempPath()) ('SEMR-tests-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
$appRoot = Split-Path -Parent $PSScriptRoot
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
Import-Module (Join-Path $appRoot 'SmartM365.ExchangeMigrationReadiness.psm1') -Force
$module = Get-Module SmartM365.ExchangeMigrationReadiness
& $module {
    param($OutputRoot, $appRoot)
    $script:testCount = 0
    function Assert-Test($Condition, $Message) {
        if (-not $Condition) { throw "FAILED: $Message" }
        $script:testCount++
    }
    function Assert-Throws([scriptblock]$Action, [string]$Pattern) {
        $caught = $null
        try { & $Action | Out-Null } catch { $caught = $_ }
        Assert-Test ($caught -and $caught.Exception.Message -match $Pattern) "Expected failure: $Pattern"
    }
    $config = ConvertTo-SemrHashtable (Get-Content (Join-Path $appRoot 'Config/SmartM365-ExchangeMigrationReadiness.local.json.template') -Raw | ConvertFrom-Json)
    foreach ($separator in @(',', ';', "`t")) {
        $csvPath = Join-Path $OutputRoot ('batch-' + [int][char]$separator + '.csv')
        [IO.File]::WriteAllText($csvPath, "EmailAddress${separator}TargetSku`r`nuser@example.invalid${separator}SPE_E3", [Text.UTF8Encoding]::new($true))
        $batch = Import-SemrBatchCsv $csvPath
        Assert-Test ($batch.Rows.Count -eq 1 -and $batch.Rows[0].EmailAddress -eq 'user@example.invalid') 'CSV delimiter and UTF-8 BOM'
    }
    $script:ConnectionState.MicrosoftGraph = $true
    $script:GraphEvidenceByEmail = @{}
    $missing = Get-SemrGraphEvidence 'user@example.invalid'
    Assert-Test ($missing.QueryError -and $null -eq $missing.SourceTimestamp) 'Missing worker payload must not certify an empty directory lookup'
    $script:GraphEvidenceByEmail['user@example.invalid'] = [pscustomobject]@{ SourceTimestamp=Get-Date; Users=@(); LicenseDetails=@(); QueryError='' }
    $empty = Get-SemrGraphEvidence 'user@example.invalid'
    Assert-Test (-not $empty.QueryError -and $empty.Users.Count -eq 0) 'Successful empty Graph lookup stays distinct from missing evidence'

    $policy = [pscustomobject]@{ RequiresLicense=$true; Known=$true; Eligible=$true }
    $capacity = [pscustomobject]@{ Available=$true; Found=$true; AvailableUnits=0 }
    Assert-Test ((Get-SemrLicenseCapacityDecision $policy $capacity $true) -eq 'PASS') 'An already assigned target SKU needs no second seat'
    Assert-Test ((Get-SemrLicenseCapacityDecision $policy $capacity $false) -eq 'FAIL') 'An unassigned target SKU with no free seat remains blocked'
    Assert-Test ((Get-SemrLicenseCapacityDecision $policy $null $false) -eq 'UNKNOWN') 'Uncollected seat capacity remains unknown'
    $capacity.AvailableUnits=1
    Assert-Test ((Get-SemrLicenseCapacityDecision $policy $capacity $false) -eq 'PASS') 'A free target seat is accepted'
    $policy.Eligible=$false
    Assert-Test ((Get-SemrLicenseCapacityDecision $policy $capacity $true) -eq 'FAIL') 'An assigned ineligible SKU is not accepted'

    $script:ConnectionState.ExchangeOnline = $true
    $script:fixtureFailure = ''
    $script:fixtureMissingProperty = $false
    function Get-EXOMailboxPermission {
        [CmdletBinding()]param($Identity)
        if ($script:fixtureFailure -eq 'FullAccess') { throw 'Fixture RBAC denied' }
        foreach ($entry in @(
            @('S-1-5-21-111-222-333-444',$false), @('selfservice@example.invalid',$false),
            @('denied@example.invalid',$true), @('NT AUTHORITY\SELF',$false), @('TEST\Exchange Servers',$false)
        )) { [pscustomobject]@{ AccessRights=@('FullAccess'); User=$entry[0]; IsInherited=$false; Deny=$entry[1] } }
    }
    function Get-RecipientPermission {
        [CmdletBinding()]param($Identity)
        if ($script:fixtureFailure -eq 'SendAs') { throw 'Fixture RBAC denied' }
        [pscustomobject]@{ AccessRights=@('SendAs'); Trustee='S-1-5-21-111-222-333-555'; IsInherited=$false; Deny=$false }
    }
    function Get-EXOMailbox {
        [CmdletBinding()]param($Identity,$Properties,$ResultSize)
        if ($script:fixtureFailure -eq 'SendOnBehalf') { throw 'Fixture RBAC denied' }
        if ($script:fixtureMissingProperty) { return [pscustomobject]@{ DisplayName='Fixture' } }
        [pscustomobject]@{ GrantSendOnBehalfTo=@('delegate@example.invalid') }
    }
    $permissions = @(Get-SemrExchangeOnlinePermission 'user@example.invalid')
    Assert-Test ($permissions.Count -eq 4) 'All three permission types collected, deny and technical identities excluded'
    Assert-Test (@($permissions | Where-Object Delegate -Like 'S-1-5-21-*').Count -eq 2) 'SID delegates preserved'
    Assert-Test (@($permissions | Where-Object Delegate -EQ 'selfservice@example.invalid').Count -eq 1) 'SELF substring is not a technical identity'
    foreach ($failure in @('FullAccess','SendAs','SendOnBehalf')) {
        $script:fixtureFailure = $failure
        Assert-Throws { Get-SemrExchangeOnlinePermission 'user@example.invalid' } 'collection.*(failed|incomplete)'
    }
    $script:fixtureFailure = ''
    $script:fixtureMissingProperty = $true
    Assert-Throws { Get-SemrExchangeOnlinePermission 'user@example.invalid' } 'Send on Behalf collection is incomplete'
    $script:fixtureMissingProperty = $false
    function Test-SemrCommand { param($Name) $Name -ne 'Get-RecipientPermission' }
    Assert-Throws { Get-SemrExchangeOnlinePermission 'user@example.invalid' } 'command is unavailable'
    Remove-Item Function:Test-SemrCommand

    # EXO metadata must establish the tenant; a successful mailbox query cannot do so.
    function Get-ConnectionInformation { [CmdletBinding()]param() @() }
    $session = Get-SemrExchangeOnlineSessionInfo -TenantId '11111111-1111-1111-1111-111111111111' -AllowCommandProbe
    Assert-Test (-not $session.Usable -and -not $session.TenantId) 'Mailbox probe without metadata cannot validate tenant identity'
    function Get-ConnectionInformation {
        [CmdletBinding()]param()
        [pscustomobject]@{ TenantID='11111111-1111-1111-1111-111111111111'; TokenStatus='Active' }
        [pscustomobject]@{ TenantID='22222222-2222-2222-2222-222222222222'; TokenStatus='Active' }
    }
    $session = Get-SemrExchangeOnlineSessionInfo -TenantId '11111111-1111-1111-1111-111111111111'
    Assert-Test (-not $session.Usable) 'Multiple EXO connections require an unambiguous reconnection'
    function Get-ConnectionInformation {
        [CmdletBinding()]param()
        [pscustomobject]@{ TenantID='11111111-1111-1111-1111-111111111111'; TokenStatus='Active' }
    }
    $session = Get-SemrExchangeOnlineSessionInfo -TenantId '11111111-1111-1111-1111-111111111111'
    Assert-Test $session.Usable 'One matching tenant remains reusable'

    $workerAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $appRoot 'SmartM365-ExchangeMigrationReadiness-ExchangeOnPremWorker.ps1'), [ref]$null, [ref]$null)
    foreach ($name in @('ConvertTo-TextArray','Test-WorkerMigrationRelevantDelegate')) {
        $definition = $workerAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        Invoke-Expression $definition.Extent.Text
    }
    $permissionLoop = $workerAst.Find({ param($node) $node -is [Management.Automation.Language.ForEachStatementAst] -and $node.Condition.Extent.Text -eq '@($mailboxPermissionResult.Rows)' }, $true)
    $mailboxPermissionResult = [pscustomobject]@{ Rows=@(
        [pscustomobject]@{ AccessRights=@('FullAccess'); User='denied@example.invalid'; IsInherited=$false; Deny=$true },
        [pscustomobject]@{ AccessRights=@('FullAccess'); User='delegate@example.invalid'; IsInherited=$false; Deny=$false }
    ) }
    $permissions = [Collections.Generic.List[object]]::new()
    Invoke-Expression $permissionLoop.Extent.Text
    Assert-Test ($permissions.Count -eq 1 -and $permissions[0].Delegate -eq 'delegate@example.invalid') 'On-premises Full Access deny is not exported as a grant'

    # Execute the real assessment/export pipeline with all live sources unavailable.
    foreach ($key in @($script:ConnectionState.Keys)) { $script:ConnectionState[$key] = $false }
    function Initialize-SemrLiveSourceConnections {
        param($Config)
        [pscustomobject]@{ ActiveDirectoryMessage='Offline fixture'; ExchangeOnPremisesMessage='Offline fixture'; ExchangeOnPremisesErrorCount=0; ExchangeOnPremisesPartialErrorCount=0; ExchangeOnPremisesFatalErrorCount=0; ExchangeOnPremisesErrors=@(); ExchangeOnPremisesDiagnosticsDirectory='' }
    }
    $assessment = Invoke-SemrAssessment -Batch $batch -Config $config
    Assert-Test ($assessment.AssessmentStatus -eq 'INCOMPLETE') 'Missing mandatory sources make the assessment incomplete'
    Assert-Test ($assessment.Summary[0].Decision -notin @('GO','GO-WARNING')) 'Missing sources cannot produce a positive verdict'
    $export = Export-SemrAssessment -Assessment $assessment -OutputRoot $OutputRoot
    Assert-Test (@(Get-ChildItem $export.RunFolder -Filter '*.csv').Count -eq 8) 'Eight CSV report files'
    $summary = @(Import-Csv $export.SummaryPath)
    Assert-Test ($summary[0].Decision -eq $assessment.Summary[0].Decision -and $summary[0].AssessmentStatus -eq 'INCOMPLETE') 'Summary CSV preserves verdict and status'
    $html = Get-Content $export.HtmlPath -Raw
    Assert-Test ($html.Contains('INCOMPLETE') -and $html.Contains('user@example.invalid')) 'HTML preserves assessment evidence'
    Assert-Test ((ConvertTo-SemrHtmlText '<script>alert(1)</script>') -eq '&lt;script&gt;alert(1)&lt;/script&gt;') 'HTML escapes evidence'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($export.ExcelPath)
    try {
        foreach ($entry in @($zip.Entries | Where-Object FullName -Like '*.xml')) {
            $reader = [IO.StreamReader]::new($entry.Open())
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        Assert-Test (@($zip.Entries | Where-Object FullName -Like 'xl/worksheets/sheet*.xml').Count -eq 9) 'Workbook XML and nine report worksheets are valid'
    } finally { $zip.Dispose() }
    Assert-Test ((Test-SemrMailboxDecisionRegression) -like 'DECISION_SELFTEST_OK*') 'Existing decision regression suite'
    Assert-Test ((Test-SemrReadinessRegression) -like 'READINESS_SELFTEST_OK*') 'Existing readiness regression suite'
    [pscustomobject]@{ Status='PASS'; Checks=$script:testCount; Version=(Get-SemrVersion); ReportFolder=$export.RunFolder } | ConvertTo-Json
} $OutputRoot $appRoot

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDYv1UTLTyAhZaW
# cNJxNp85LWWeNa+e7kLXV7NQ007MvaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIODlvxvdJHoVwzK5TCnirhhkn3qndKkZ0LAvX1TPUBVjMA0GCSqG
# SIb3DQEBAQUABIIBgJ0Gj4ILd4epg34E9huofSy+dNUuhEh56YSzxdoHlwOvOdx0
# Vuqh8l4w6c8kU5JfbOHa8YknsfDeHNPkWRneK9Ns7yDuHrxpb7IZvCEtkDMUFjJL
# D63TuuEpvrSOprxbAfGZ82qBIzymhautwI7foup2oePWLKp/SFvasgL7GUjIF24x
# MYPEeO/GdKm+jqNKBVjyzTwLaMo9JrJ4LmNvA1O9lvVVo6G8ac38SQIyfLXEg4Ra
# D304Qm7jrJ35E38DXI+zlNlmvjt6ToCWdWwyTYEOcX9TJ6RCugusuyKEuam54LHz
# ftWoMT6eKeVTfkstepg97v+xnRwfXeeFC9QIxExKKzb4o5CM+9aVFSLC4Mw1EyBv
# PuHNzjd9Y/szHXPnhxAb9Gr4RjSoiHeOPiOJUo+yWKVVLpG0MWB3lg4JH3vC+JOj
# pcVjYKxGm8LpdN6Pht3n5V6IHiedD60f0GQB2aLyolWiWBdjUpriowRrdd/dLCy1
# 38rgpqDyA82V7YQABqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDgyMTQ5
# MzBaMC8GCSqGSIb3DQEJBDEiBCCpd7+B9hU/k9zHnQOIT/HQU8NPE/nvOzQ20D/l
# iYuiWzANBgkqhkiG9w0BAQEFAASCAgBhGKuKeL/36P+BZ2MW0lAOIrYVmKTxxrn0
# YikYFjkwNTJTeuZDqEOl24W6OaCltMIUNE8CtxIzOAE6f15PSmfwnbUfbzI9ymp0
# 5Dmvu2Zmet+0sSsgrq+jaYf8BE2C1MMFVX8WzXSjM2l/2MUS6peD+QwFeAlQxhWG
# E1ij4GElFRwi75W0KSQ0UBt8K2a8MW6vFrjDHgy4NmBs1mYcpjkawFEuoqyBBjk+
# ifqk1y7ggGd3xGBiuuUPl4EEZ1YkYLOXLlPeHt2tYkNj8ylrX04dv2idkwRxJ0lB
# axegPkuF5UWshJGsIE1RLOSH1Awb0MJEFJ5480YD4KKl1Zl6QypxR2Mn/vMyAxLG
# oWi/+QUcMHrOdrE1Ku8HE0n6FxlsZI2gPnvfm69UCR6HbPHMEH8WRWW52jCZnaVN
# ogW0lQMvaSta4t0qOQrqk0gkws6FF18cRCQ8zdu5uMv/B0af39YJ+Ed8GP0Ys0Bp
# l1Qp8XzKJt7umytxrVUUdL0P0xqOEDNJt8TapmevQYZQuY7a+D8qPf7+Q2m9286z
# Te74QH3ovvtknq0ct8RgcO9x3s9tMh85X5ADjf1ZOokPr6uQuQkU0TSsdOh1zkvy
# QmIS1FNRTA+/r87IrEAS+Xqf1MqVWKjiKpV/iHcKcTo9Bh2zczkJac3NZ3CAHBen
# uw/uNhdp1g==
# SIG # End signature block
