# Version 1.2.0: DirectoryCleanup reads the captured rules after Workforce and the M365 usage signals generator is stubbed; 1.1.0: HybridIdentity reads the captured rules after Workforce; 1.0.1: SharePointMigration fixture generator.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Synthetic rules and simulated child generators only; no production preparation.
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('classification-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture -Force | Out-Null
$script:checks = 0
function Check([bool]$Pass, [string]$Message) {
    if (-not $Pass) { throw $Message }; $script:checks++
}
function Reject([scriptblock]$Action, [string]$Pattern) {
    $message = ''
    try { & $Action | Out-Null } catch { $message = $_.Exception.Message }
    Check ($message -match $Pattern) "Expected '$Pattern', got '$message'"
}
$valid = '{"SchemaVersion":"1.0","RuleVersion":"SYNTHETIC_DO_NOT_LOG","Population":{"Human":["human-a"],"NonHuman":[]}}'
$changed = $valid.Replace('human-a','human-b')
$source = Join-Path $fixture 'rules.local.json.txt'
[IO.File]::WriteAllText($source,$valid)
$module = Import-Module (Join-Path $PSScriptRoot 'PreparedClassification.psm1') -Force -PassThru
try {
    foreach ($path in @($source,$source.Substring(0,$source.Length-4))) {
        $destination = Join-Path $fixture ([guid]::NewGuid().ToString('N') + '.json.txt')
        $output = @(Copy-PreparedClassificationCapture -SourcePath $path -DestinationPath $destination 6>&1)
        $receipt = @($output | Where-Object { $_ -isnot [Management.Automation.InformationRecord] })[0]
        Check ($receipt.Captured.SHA256 -eq (Get-FileHash $source).Hash) 'Captured hash differs'
        Check ($receipt.Captured.Bytes -eq (Get-Item $source).Length) 'Captured length differs'
        Check ($receipt.Source.Path -eq $source) 'Logical/physical resolution differs'
        $messages = ($output | Where-Object { $_ -is [Management.Automation.InformationRecord] }) -join "`n"
        Check ($messages -match 'stage=AfterCopy;' -and $messages -match 'stage=SourceAfterCopy;') 'Capture checkpoints missing'
        Check ($messages -notmatch 'SYNTHETIC_DO_NOT_LOG|human-a|Population') 'Private rule content leaked in logs'
        Check (($receipt | ConvertTo-Json -Depth 5) -notmatch 'SYNTHETIC_DO_NOT_LOG|human-a|Population') 'Receipt contains rule content'
    }
    Reject { Copy-PreparedClassificationCapture -SourcePath $source -DestinationPath $source } 'new, separate destination'
    $missing = Join-Path $fixture 'missing.json.txt'
    Reject { Test-PreparedClassificationCapture -Path $missing -Stage 'BuildStart' } 'BuildStart.*file is missing'
    $empty = Join-Path $fixture 'empty.json.txt'; [IO.File]::WriteAllText($empty,'')
    Reject { Test-PreparedClassificationCapture -Path $empty -Stage 'AfterCopy' } 'AfterCopy.*file is empty'
    $invalid = Join-Path $fixture 'invalid.json.txt'
    foreach ($text in @('{"private":"SYNTHETIC_DO_NOT_LOG",','{"SchemaVersion":"9.0","RuleVersion":"SYNTHETIC_DO_NOT_LOG"}','{"SchemaVersion":"1.0"}')) {
        [IO.File]::WriteAllText($invalid,$text)
        $message = ''
        try { Test-PreparedClassificationCapture -Path $invalid -Stage 'AfterCopy' | Out-Null } catch { $message = $_.Exception.Message }
        Check ($message -match 'JSON/rules are invalid or unreadable') 'Invalid rules accepted'
        Check ($message -notmatch 'SYNTHETIC_DO_NOT_LOG|SchemaVersion|9\.0') 'JSON/rule error leaked content'
    }
    # Mutate at the copy boundary, rather than relying on filesystem timing.
    foreach ($mode in @('MissingCopy','ChangedCopy','ChangedSource')) {
        & $module {
            param($mode,$changed)
            $script:copyMode = $mode; $script:changedRules = $changed
            function script:Copy-Item {
                param([string]$LiteralPath,[string]$Destination,[string]$ErrorAction)
                if ($script:copyMode -ne 'MissingCopy') {
                    Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination -ErrorAction $ErrorAction
                }
                if ($script:copyMode -eq 'ChangedCopy') { [IO.File]::WriteAllText($Destination,$script:changedRules) }
                if ($script:copyMode -eq 'ChangedSource') { [IO.File]::WriteAllText($LiteralPath,$script:changedRules) }
            }
        } $mode $changed
        try {
            [IO.File]::WriteAllText($source,$valid)
            $destination = Join-Path $fixture ($mode + '.json.txt')
            $pattern = switch ($mode) {
                'MissingCopy' { 'AfterCopy.*file is missing' }
                'ChangedCopy' { 'AfterCopy.*SHA256 changed' }
                'ChangedSource' { 'SourceAfterCopy.*SHA256 changed' }
            }
            Reject { Copy-PreparedClassificationCapture -SourcePath $source -DestinationPath $destination } $pattern
        } finally { & $module { Remove-Item Function:script:Copy-Item } }
    }
    [IO.File]::WriteAllText($source,$valid)
    $expected = Test-PreparedClassificationCapture -Path $source -Stage 'Fixture'
    [IO.File]::WriteAllText($source,$valid + ' ')
    Reject { Test-PreparedClassificationCapture -Path $source -Stage 'BeforeBuild' -ExpectedSHA256 $expected.SHA256 -ExpectedBytes $expected.Bytes } 'BeforeBuild.*byte length changed'
    [IO.File]::WriteAllText($source,$changed)
    Reject { Test-PreparedClassificationCapture -Path $source -Stage 'BeforeBuild' -ExpectedSHA256 $expected.SHA256 -ExpectedBytes $expected.Bytes } 'BeforeBuild.*SHA256 changed'
} finally { Remove-Module $module }

# Execute the real build dispatcher with fixture-only generators and the real rule reader.
$harness = Join-Path $fixture 'harness'
$scripts = Join-Path $harness 'SmartWorkplaceIntelligence/scripts'
$common = Join-Path $harness 'SmartM365/SmartInventory/Common'
New-Item -ItemType Directory -Path $scripts,$common -Force | Out-Null
foreach ($name in @('Invoke-PreparedEvidenceBuild.ps1','PreparedClassification.psm1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $scripts $name)
}
foreach ($name in @('SmartM365.AccountClassification.psd1','SmartM365.AccountClassification.psm1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot "../../SmartM365/SmartInventory/Common/$name") -Destination (Join-Path $common $name)
}
[IO.File]::WriteAllText((Join-Path $scripts 'NativeEvidence.psm1'),@'
function Invoke-NativeEvidenceStage { param($LastRoot,$TenantKey,$AuditPath,$OutputRoot) }
Export-ModuleMember -Function Invoke-NativeEvidenceStage
'@)
[IO.File]::WriteAllText((Join-Path $scripts 'New-EnterpriseEstateEvidence.ps1'),@'
param($DataRoot,$OutputRoot)
[IO.File]::WriteAllText((Join-Path $OutputRoot 'Estate.started'),'synthetic')
'@)
[IO.File]::WriteAllText((Join-Path $scripts 'New-DeviceInventoryEvidence.ps1'),@'
param($DataRoot,$OutputPath,$SignalsOutputPath,$DirectorySummaryOutputPath,$SystemEvidenceOutputPath)
[IO.File]::WriteAllText((Join-Path (Split-Path $OutputPath -Parent) 'Devices.started'),'synthetic')
$control = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../control.json.txt') -Raw | ConvertFrom-Json
if ($control.Mode -eq 'Delete') { [IO.File]::Delete($control.Path) }
if ($control.Mode -eq 'Tamper') { [IO.File]::WriteAllText($control.Path,$control.Changed) }
'@)
[IO.File]::WriteAllText((Join-Path $scripts 'New-HybridIdentityEvidence.ps1'),@'
param($DataRoot,$UserEvidencePath,$OutputPath,$AccountClassificationConfigPath)
[IO.File]::WriteAllText((Join-Path (Split-Path $OutputPath -Parent) 'HybridIdentity.started'),[string]$AccountClassificationConfigPath)
'@)
[IO.File]::WriteAllText((Join-Path $scripts 'New-DirectoryCleanupEvidence.ps1'),@'
param($DataRoot,$TenantKey,$UserInventoryPath,$AllDevicesPath,$UserOutputPath,$ComputerOutputPath,$MailboxOutputPath,$QualificationOutputPath,$AccountClassificationConfigPath)
[IO.File]::WriteAllText((Join-Path (Split-Path $UserOutputPath -Parent) 'DirectoryCleanup.started'),[string]$AccountClassificationConfigPath)
'@)
[IO.File]::WriteAllText((Join-Path $scripts 'New-WorkforceIdentityEvidence.ps1'),@'
param($DataRoot,$UserOutputPath,$HistoryOutputPath,$TrendOutputPath,$IdentityOutputPath,$SignalsOutputPath,$SystemEvidencePath,$AccountClassificationConfigPath)
[IO.File]::WriteAllText((Join-Path (Split-Path $UserOutputPath -Parent) 'Workforce.started'),'synthetic')
'@)
foreach ($name in @('New-SharePointMigrationEvidence','New-WindowsMigrationEvidence','New-DeviceLifecycleExperienceTrendEvidence','New-ApplicationInventoryEvidence','New-CollaborationEvidence','New-M365UsageSignalsEvidence','New-ContentStorageEvidence','New-MailboxEvidence','New-DailyTrendsEvidence','New-LicensingEvidence','New-SecurityEvidence','New-BackupResilienceEvidence','New-IntuneOperationalEvidence','New-DataTrustEvidence','New-ExecutiveTrendEvidence')) {
    [IO.File]::WriteAllText((Join-Path $scripts ($name + '.ps1')),'exit 0')
}
$shell = (Get-Process -Id $PID).Path
foreach ($mode in @('Valid','MissingAtStart','TamperedAtStart','Delete','Tamper')) {
    $case = Join-Path $harness $mode
    $raw = Join-Path $case 'raw'; $staging = Join-Path $case 'staging'
    New-Item -ItemType Directory -Path (Join-Path $raw 'DATA-LAST'),$staging -Force | Out-Null
    $rules = Join-Path $case 'AccountClassification.json.txt'
    [IO.File]::WriteAllText($rules,$valid)
    $hash = (Get-FileHash $rules).Hash; $bytes = (Get-Item $rules).Length
    if ($mode -eq 'MissingAtStart') { [IO.File]::Delete($rules) }
    if ($mode -eq 'TamperedAtStart') { [IO.File]::WriteAllText($rules,$changed) }
    @{Mode=$mode;Path=$rules;Changed=$changed} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $scripts '../control.json.txt')
    $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $scripts 'Invoke-PreparedEvidenceBuild.ps1'),'-DataRoot',$raw,'-OutputRoot',$staging,'-AccountClassificationConfigPath',$rules,'-AccountClassificationExpectedSHA256',$hash,'-AccountClassificationExpectedBytes',[string]$bytes,'-NativeEvidenceAuditPath',(Join-Path $case 'unused-audit.json.txt'),'-TenantKey','synthetic-test')
    $output = (& $shell @arguments 2>&1 | Out-String)
    $exitCode = $LASTEXITCODE
    if ($mode -eq 'Valid') {
        Check ($exitCode -eq 0) "Valid build dispatcher failed: $output"
        Check (Test-Path -LiteralPath (Join-Path $staging 'Workforce.started')) 'Valid capture never reached Workforce'
        Check ($output -match 'stage=BuildStart;' -and $output -match 'stage=BeforeWorkforce;' -and $output -match 'stage=BeforeHybridIdentity;') 'Dispatcher checkpoints missing'
        Check ((Get-Content -LiteralPath (Join-Path $staging 'HybridIdentity.started') -Raw) -eq $rules) 'HybridIdentity did not receive the captured rules'
        Check ($output -match 'stage=BeforeDirectoryCleanup;' -and (Get-Content -LiteralPath (Join-Path $staging 'DirectoryCleanup.started') -Raw) -eq $rules) 'DirectoryCleanup did not receive the captured rules after its checkpoint'
    } else {
        Check ($exitCode -ne 0) "Invalid capture accepted: $mode"
        Check (-not (Test-Path -LiteralPath (Join-Path $staging 'Workforce.started'))) "Workforce started after capture loss: $mode"
        Check (-not (Test-Path -LiteralPath (Join-Path $staging 'HybridIdentity.started'))) "HybridIdentity started after capture loss: $mode"
        Check (-not (Test-Path -LiteralPath (Join-Path $staging 'DirectoryCleanup.started'))) "DirectoryCleanup started after capture loss: $mode"
        if ($mode -in @('MissingAtStart','TamperedAtStart')) {
            Check ($output -match 'BuildStart') 'Initial failure stage missing'
            Check (-not (Test-Path -LiteralPath (Join-Path $staging 'Estate.started'))) 'Initial check ran after heavy work'
        } else {
            Check ($output -match 'BeforeWorkforce') 'Delayed failure stage missing'
            Check (Test-Path -LiteralPath (Join-Path $staging 'Devices.started')) 'Delayed fixture never reached its mutation boundary'
        }
    }
    Check ($output -notmatch 'SYNTHETIC_DO_NOT_LOG|human-a|human-b|Population') 'Dispatcher logs leaked private rules'
}
# The receipt remains after owned payload cleanup; neither copied rules nor raw data remain in it.
$pipeline = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
try {
    $work = Join-Path $fixture 'cleanup'; $runId = [guid]::NewGuid().ToString('N'); $run = Join-Path $work $runId
    New-Item -ItemType Directory -Path $run -Force | Out-Null
    @{Owner='PreparedEvidencePipeline/v1';RunId=$runId;DataRoot=$fixture;TenantKey='synthetic-test'} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'run.json.txt')
    $rules = Join-Path $run 'AccountClassification.json.txt'
    & $pipeline {
        param($source,$rules,$run,$work,$runId,$raw)
        $receipt = Copy-PreparedClassificationCapture -SourcePath $source -DestinationPath $rules
        Write-PreparedJson (Join-Path $run 'ClassificationCapture.json.txt') $receipt
        Clear-PreparedRunPayload $work $runId $raw 'synthetic-test'
    } $source $rules $run $work $runId $fixture
    Check (-not (Test-Path -LiteralPath $rules)) 'Owned copied rules survived cleanup'
    $receiptPath = Join-Path $run 'ClassificationCapture.json.txt'
    Check (Test-Path -LiteralPath $receiptPath) 'Metadata receipt lost during cleanup'
    Check ((Get-Content -LiteralPath $receiptPath -Raw) -notmatch 'SYNTHETIC_DO_NOT_LOG|human-a|Population') 'Retained receipt contains rules'
} finally { Remove-Module $pipeline }
Write-Information "PASS: $script:checks classification capture, integrity, privacy, dispatcher and cleanup checks. Synthetic fixtures only." -InformationAction Continue

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAYk7FS3bSxT8cO
# 7rxtCsCSpCMLUaJI33NRHTtYn4WJSqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGDbz66evI2n2IWb+ymqmXFk+5qX2q/aQBMIhlA1qgkXMA0GCSqG
# SIb3DQEBAQUABIIBgFpGlmOmSRVdOSliUTRSBEColGw585vH/QGyLRaOxqSUm5RS
# 7Z7B4nbNHGrFDjdXEam/bU3hZSBHL+JJwj+LuXk8xgh+xnzs71BhCM04LgP8WTgd
# EAagzZRjKUMx29HZSSs71qnLjUk1l44eQgCbhZPJ8BZqmw2Dm5DdWjAjHEqF1bNw
# 2LVUgFiRP1RmIxXlkoRQrlloZYVEAbcljmNg99UBwCzU7xSiPlq7e14u/iKhH7KW
# +ymC6hSlXBq/vJSJ/ATpu9kRxXkvJ94pRl2in6iXnD7vB0ZTSMXIgoxozkiNidRu
# rTNtMFklpZ0feu7Nf9ssxcB1UHxSs6KpMRkWZvMzPJt9Z9xZO+845LRZYmhR5Uxd
# HcOoE5cIOdn5ryIJfkx/3FdKwd6f69uUtG/sIFuxXdV8AmposaKF6mzWcKdnNsnz
# z6evApKGk0guAAKJ6GBkq2AyMMaeDB2EOPa2qU1a6XIgVJRhyLwpjVr6fPngBJnG
# /mq+bhiYmTzBg3FFgqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMjU3
# NTZaMC8GCSqGSIb3DQEJBDEiBCAPOuRShRJ38pZz/cshi7okYHCkGrJXjVuSpsSl
# QdN8qTANBgkqhkiG9w0BAQEFAASCAgAgpCRwDMDtL+5NDl96Q8cdoTU7GIpMTVpn
# WiiT0peZbpCYhk/RSEDYos5SQOweqnuc9oxx+I/LLmtp2j7bXXcnjmc8t3LLh+AT
# uvG3jB94+w/mViy938hz2sfJ8URLnfftYiKKNe925bkezn5mCJm9yHYX2BR2FI1+
# 8Jq/7kLcWtZflQuFNfv1X3bz1FacicKv/jm1oKJpo8wvuzTK+cBXmttKDDQ6lvTK
# T6rRNYXWFNnoQ3emdo+GK1SfM3rrND8SJAsz2P2SKerhdrsCdSnW51yyd/SOqhZV
# mfzDOOL/SLSIzcep38iddE5bIAX2UDpgu2zfMu83coYGUcWSX2KyW4IaW9oKyDYi
# zn+H7suAgpqBCuBJ7TQQA8RKI/6CazT55rIYJMjY+Qc3Po8H8eKgBFT3MNEjnIMR
# svX/qlTBL9heaRpdc1Qp9+ZyaBrWZVEboEm4T5mopbjpzdJrIRidG818llDGlqEp
# RWRfHqbdLrYN9fiujqNTQano9+HR5s3G1CfJF28TjnyDKwRuOqQyB/yYrU6lpgEV
# /f7Qdngneb9e5XoGRujza/sg2zjXgrevRMNm+fWe5UGsIeNgr0SDsqmTf2p2qQfo
# U2Ep0OwiPEqBKN8zF1/7zHCL+jNpuJaK+RcTV+YSBsfrp5Nb7jroK58sPhE9EnF3
# n9lizO4vVw==
# SIG # End signature block
