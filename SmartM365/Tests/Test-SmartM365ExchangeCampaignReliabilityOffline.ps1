#Requires -Version 5.1
<#
.SYNOPSIS
Offline campaign checkpoint, targeting and failure-status regressions.
.VERSION
1.0.0
#>
[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CampaignReview-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
$appRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'Exchange/ExchangeUserNotifications'
[void][IO.Directory]::CreateDirectory($TestRoot)
Import-Module (Join-Path $appRoot 'SmartM365.Communications.psm1') -Force -DisableNameChecking
$script:checks = 0
function Assert-Campaign {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAILED: $Message" }
    $script:checks++
}
function Read-CampaignAst {
    param([string]$Name)
    $path = Join-Path $appRoot "$Name/SmartM365-$Name-NotifyUsers.ps1"
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "Source parse failed: $path" }
    return $ast
}

$path = Join-Path $TestRoot 'registry.csv'
$lock = Enter-SmartM365CommunicationSentRegistryLock -Path $path
try {
    $blocked = $false
    try { $secondLock = Enter-SmartM365CommunicationSentRegistryLock -Path $path; $secondLock.Dispose() } catch { $blocked = $true }
    Assert-Campaign $blocked 'An overlapping writer must be refused'
    $registry = @{}
    Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $path
    Assert-Campaign ((Load-SmartM365CommunicationSentRegistry -Path $path).Count -eq 0) 'Empty registry has a valid schema'
    Register-SmartM365CommunicationSentItem -Registry $registry -Email 'first@example.invalid'
    Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $path
    Register-SmartM365CommunicationSentItem -Registry $registry -Email 'second@example.invalid'
    Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $path
    Assert-Campaign ((Load-SmartM365CommunicationSentRegistry -Path $path).Count -eq 2) 'Atomic replacement preserves earlier deliveries'
    $reader = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $blocked = $false
        try { Save-SmartM365CommunicationSentRegistry -Registry $registry -Path $path } catch { $blocked = $true }
        Assert-Campaign $blocked 'Failed replacement is reported'
    }
    finally { $reader.Dispose() }
    Assert-Campaign ((Load-SmartM365CommunicationSentRegistry -Path $path).Count -eq 2) 'Failed replacement preserves the old registry'
    Assert-Campaign (@(Get-ChildItem $TestRoot -Filter '.sent-registry-*.tmp').Count -eq 0) 'Failed writes clean temporary files'
}
finally { $lock.Dispose() }
$lock = Enter-SmartM365CommunicationSentRegistryLock -Path $path
$lock.Dispose()
Assert-Campaign $true 'Released lock can be acquired again'

# Terminate a real child after its first checkpoint, before any terminal cleanup.
$killRegistry = Join-Path $TestRoot 'killed-child.csv'
$readyPath = Join-Path $TestRoot 'child-ready.txt'
$childCode = @'
$ErrorActionPreference = 'Stop'
Import-Module '@MODULE@' -Force -DisableNameChecking
$lock = Enter-SmartM365CommunicationSentRegistryLock -Path '@REGISTRY@'
$registry = @{}
Register-SmartM365CommunicationSentItem -Registry $registry -Email 'before-stop@example.invalid'
Save-SmartM365CommunicationSentRegistry -Registry $registry -Path '@REGISTRY@'
[IO.File]::WriteAllText('@READY@', 'ready')
Start-Sleep -Seconds 30
'@
$childCode = $childCode.Replace('@MODULE@', (Join-Path $appRoot 'SmartM365.Communications.psm1').Replace("'", "''")).Replace('@REGISTRY@', $killRegistry.Replace("'", "''")).Replace('@READY@', $readyPath.Replace("'", "''"))
$startInfo = New-Object Diagnostics.ProcessStartInfo
$startInfo.FileName = (Get-Process -Id $PID).Path
$startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childCode))
$startInfo.UseShellExecute = $false; $startInfo.CreateNoWindow = $true
$child = [Diagnostics.Process]::Start($startInfo)
try {
    $deadline = [datetime]::UtcNow.AddSeconds(15)
    while (-not (Test-Path $readyPath) -and -not $child.HasExited -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    Assert-Campaign (Test-Path $readyPath) 'Child reaches checkpoint before forced termination'
    $child.Kill(); $child.WaitForExit()
    Assert-Campaign ((Load-SmartM365CommunicationSentRegistry -Path $killRegistry).ContainsKey('before-stop@example.invalid')) 'Forced termination preserves confirmed delivery'
    $lock = Enter-SmartM365CommunicationSentRegistryLock -Path $killRegistry
    $lock.Dispose()
    Assert-Campaign $true 'Process termination releases registry lock'
}
finally {
    if (-not $child.HasExited) { $child.Kill(); $child.WaitForExit() }
    $child.Dispose()
}

# Only extracted recipient and terminal blocks execute. Configuration, authentication,
# inventory reads and operational transports never execute in this suite.
foreach ($name in @('ExchangeMigration', 'ExchangeArchive', 'ExchangeMigrationMailboxSizeReduction')) {
    $ast = Read-CampaignAst $name
    $recipient = $ast.FindAll({param($n) $n -is [Management.Automation.Language.TryStatementAst] -and $n.Body.Extent.Text.Contains('$template = Get-SmartM365CommunicationTemplateContent')}, $true) |
        Sort-Object { $_.Extent.Text.Length } | Select-Object -First 1
    $outer = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] } | Select-Object -Last 1
    Assert-Campaign ($outer.Finally.Extent.Text -match 'registryLock.Dispose') "$name releases its lock on failure and success"
    $lockCommand = $outer.Body.Extent.Text.IndexOf('Enter-SmartM365CommunicationSentRegistryLock')
    $loadCommand = $outer.Body.Extent.Text.IndexOf('Load-SmartM365CommunicationSentRegistry')
    Assert-Campaign ($lockCommand -ge 0 -and $loadCommand -gt $lockCommand) "$name locks before reading registry"
    foreach ($scenario in @('Success', 'SendFailure', 'TeamsFailure', 'CheckpointFailure', 'DryRun')) {
        & {
            Set-StrictMode -Off
            $registry = @{}; $registryCheckpointPending = $false; $preventResend = $true
            $sentRegistryPath = Join-Path $TestRoot ($name + '-' + $scenario + '.csv')
            if ($scenario -eq 'CheckpointFailure') {
                # A directory at the destination makes atomic publication fail.
                [void][IO.Directory]::CreateDirectory($sentRegistryPath)
            }
            $WhatIf = $scenario -eq 'DryRun'; $NoSummaryEmail = $true
            $counters = @{ Files=1; Rows=1; Candidates=1; Sent=0; DryRun=0; Failed=0; Skipped=0; AlreadySent=0; MailboxNotFound=0; RemoteMailbox=0; MailboxUsageUnavailable=0; MailboxUsageBelowThreshold=0; TeamsSent=0; TeamsDryRun=0; TeamsFailed=0 }
            $processedItems = [Collections.ArrayList]::new()
            $email = 'fictional@example.invalid'; $file = [pscustomobject]@{Name='fictional.csv'}
            $config = @{}; $hotlineMap = @{}; $teamsUserMessageByLanguage = @{}
            $freeTextBlock = [pscustomobject]@{Style='';TextHtml=''}
            $exchangeManagementState = [pscustomobject]@{Source='Fixture';Status='Ready'}
            $summaryLevels = [Collections.Generic.List[string]]::new()
            $intraEmailDelayMilliseconds = 0
            function Get-SmartM365CommunicationTemplateContent { [pscustomobject]@{Content='Fictional template'} }
            function Get-SmartM365CommunicationLinkedLogoTokens { [pscustomobject]@{LogoImgTag='';FooterLogoImgTag=''} }
            function Get-SmartM365CommunicationHotline { '' }
            function Get-SmartM365CommunicationConfigValue { '' }
            function Get-LocalizedTextValue { '' }
            function ConvertTo-FreeTextBlock { [pscustomobject]@{Style='';TextHtml=''} }
            function Expand-SmartM365CommunicationTemplate { 'Fictional message' }
            function Assert-SmartM365CommunicationNoUnresolvedToken { }
            function Get-SmartM365CommunicationSubject { 'Fictional subject' }
            function Add-SmartM365CommunicationLogRow { }
            function Send-SmartM365CommunicationMail {
                param([switch]$WhatIf)
                if ($scenario -eq 'SendFailure') { throw 'Synthetic delivery failure' }
                [pscustomobject]@{Sent=(-not $WhatIf);Mode=$(if($WhatIf){'DryRun'}else{'Fixture'})}
            }
            function Get-SmartM365CommunicationTeamsUserMode { 'GraphDelegated' }
            function Send-SmartM365CommunicationTeamsUserMessage {
                # Check durability before even attempting the optional Teams transport.
                if (-not $WhatIf) {
                    Assert-Campaign ((Load-SmartM365CommunicationSentRegistry -Path $sentRegistryPath).ContainsKey($email)) "$name checkpoints before Teams"
                }
                [pscustomobject]@{Sent=($scenario -ne 'TeamsFailure');Mode='Fixture';Error='Synthetic Teams failure'}
            }
            function Send-SmartM365TeamsNotification { param($Level) $summaryLevels.Add($Level) }
            $checkpointThrew = $false
            try { . ([scriptblock]::Create($recipient.Extent.Text)) } catch { $checkpointThrew = $true }
            if ($scenario -eq 'CheckpointFailure') {
                Assert-Campaign $checkpointThrew "$name fails closed on checkpoint failure"
                Assert-Campaign ($counters.Sent -eq 1 -and $processedItems.Count -eq 0) "$name stops before Teams/logging after checkpoint failure"
                return
            }
            Assert-Campaign (-not $checkpointThrew) "$name recipient handling completes for $scenario"
            if ($scenario -in @('Success','TeamsFailure')) {
                Assert-Campaign ((Load-SmartM365CommunicationSentRegistry -Path $sentRegistryPath).ContainsKey($email)) "$name restart retains confirmed email for $scenario"
                Assert-Campaign ($counters.Sent -eq 1) "$name counts successful delivery once"
            }
            else { Assert-Campaign (-not (Test-Path -LiteralPath $sentRegistryPath)) "$name does not persist an unsent email" }
            $tail = [Collections.Generic.List[string]]::new(); $started = $false
            foreach ($statement in $outer.Body.Statements) {
                if ($statement -is [Management.Automation.Language.AssignmentStatementAst] -and $statement.Left.Extent.Text -eq '$summary') { $started = $true }
                if ($started) { $tail.Add($statement.Extent.Text) }
            }
            # Include the real outer catch: verify both the process failure and ERROR summary.
            $terminal = 'try {' + ($tail -join "`n") + '} ' + $outer.CatchClauses[0].Extent.Text
            $terminalError = $false
            try { . ([scriptblock]::Create($terminal)) | Out-Null } catch { $terminalError = $true }
            $expectedFailure = $scenario -in @('SendFailure','TeamsFailure')
            Assert-Campaign ($terminalError -eq $expectedFailure) "$name propagates failure for $scenario"
            if ($expectedFailure) { Assert-Campaign ($summaryLevels.Count -eq 1 -and $summaryLevels[0] -eq 'ERROR') "$name emits ERROR without SUCCESS" }
            elseif ($WhatIf) { Assert-Campaign ($summaryLevels.Count -eq 0) "$name dry run sends no summary notification" }
            else { Assert-Campaign ($summaryLevels.Count -eq 1 -and $summaryLevels[0] -eq 'SUCCESS') "$name successful run remains successful" }
        }
    }
}

$ast = Read-CampaignAst 'ExchangeMigrationMailboxSizeReduction'
foreach ($name in @('Get-LiveLicenseQuotaInfoForGraphUser','Resolve-LiveLicenseQuotaInfo','Get-PrimarySmtpAddressFromGraphUser')) {
    $definition = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name}, $true)
    Set-Item ('Function:' + $name) $definition.Body.GetScriptBlock()
}
$fixtureUser = [pscustomobject]@{Id='fictional';Mail='user@example.invalid';UserPrincipalName='user@example.invalid';DisplayName='Fictional';ProxyAddresses=@('SMTP:user@example.invalid')}
function Get-GraphUserBySmtpAddress { $fixtureUser }
function Get-MgUserLicenseDetail { foreach ($sku in $fixtureSkus) { [pscustomobject]@{SkuPartNumber=$sku} } }
$quotaMap = @{
    SPE_F1=[pscustomobject]@{SkuPartNumber='SPE_F1';ThresholdMB=2048;MaxGb=2;QuotaLabel='2 GB'}
    SPE_E3=[pscustomobject]@{SkuPartNumber='SPE_E3';ThresholdMB=102400;MaxGb=100;QuotaLabel='100 GB'}
}
$allowed = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
[void]$allowed.Add('SPE_F1')
foreach ($fixtureSkus in @(@('SPE_E3'), @('SPE_E3','SPE_F1'), @('spe_f1'))) {
    $actual = Get-LiveLicenseQuotaInfoForGraphUser -User $fixtureUser -QuotaBySku $quotaMap -AllowedSkus $allowed
    $fromList = Resolve-LiveLicenseQuotaInfo -SmtpAddress $fixtureUser.Mail -QuotaBySku $quotaMap -AllowedSkus $allowed
    $expected = if ($fixtureSkus -contains 'SPE_F1') { 'SPE_F1' } else { '' }
    Assert-Campaign ([string]$actual.SkuPartNumber -eq $expected) 'Live inventory respects configured audience and quota'
    Assert-Campaign ([string]$fromList.SkuPartNumber -eq $expected) 'Live CSV-list lookup respects configured audience and quota'
}
[pscustomobject]@{Status='PASS';Checks=$script:checks;Engine=$PSVersionTable.PSVersion.ToString();TestRoot=$TestRoot}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBZqtaDvJQ8bPQk
# aOI1rAq5og3BSV0gP6yvfU1aRFTBvaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAaaBT7rhKTQR39FVfCwdEDH172Xb4dlmBbTt3aOiPNqMA0GCSqG
# SIb3DQEBAQUABIIBgCmY9fLbhzwpH5Y4Y8U0mbw/devbxrlHoIqRLz2GkKGtAZj1
# ZHdtWYwtuNxAD3vAndCQ/QCrxEumREBz1BcNUvBDn73ALANg0wVMIjcwS3GP8JmX
# TMPcDHO301Adtjo1Z3zo0bfnhCuD6WF/ceDG2oaCrClpUwwb8cATLmlntw4NwLaV
# en5ZT9S5jN14HnC7A7F177fO0LSpilRmj8feOpJESwbcvwmCyTOBr9jTEnTKO/80
# +zCN60pZkaExPYCEla3NCe8f0UXA1xkVR67bMHKeIvLHOEtjmB06D2jbX7ehZF7d
# iHWdslgGR1mXbVVr0k1nAxtKj3Bgovukgfmb5aLxS+rPODZokVGoz+HGhtlgOMV9
# ZaxyCz04Qsu5N0+0U9oHtCZcqZZPofVxu6IJFJrxfsBY1leSaipVcNa5Wig1P8yj
# RhaeyLrQYm2145rwsXKG4wnTT3w0Z9m9AFNkhhm9fdYfhBSvsp1W/5WEeSpEEZMv
# ZMdLa/vLUT3pvB1dQ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzE0
# MzRaMC8GCSqGSIb3DQEJBDEiBCCfetGXMCxASzWl0pDI4OUG7CcX+cLUCEiH8SEn
# eeK7QjANBgkqhkiG9w0BAQEFAASCAgCdNWdH8BIoJGEUyXWYv7fQzw9WWuOabp3T
# pKmynGf01IcpUvlGm8GodFkMAGMlb9DRr+pUrc84RLKgjnS3QhfBQtM5Ay2vts8a
# RzFy3UndGISd1aR+7O1Ys4Z9Saz+YzNrij3nWr0v8Bh7JEEzusY0e81AJJaEnfet
# 7HP925182+csIIY6QpbDB5WMn9dwqB0JzOX+GTwpwA+NuEUtI0YHvH+ENbgQb8jx
# c7ZjsM+6WQqCIHNK47u71L0PYpOFUM2ulGJpbY7BuFB5D99EbjlN7qjACrb1wWYU
# qGZG85JTIQhB0CoKqb+ThH5Q+hmEJ/hWXqPBQyVovz/KMYGX3xENaq81/nGlqCSD
# AKcNy1G8Rh96bcv1UlRW9triyqIv+LHIpvZyibWMcyGEehbwWpAg1gm9q9V1RF21
# uglUCsfrnPQx5TVaD1zIpoXTQBmE/6SLquWZY2dwcKGsj3xvG0CeF7DUzeTe35Xp
# 8PZvEG/4rCSE0OYEpvoCVTKOLv4D6wbln6SAMSVkBSKWL6MMVLx6btadpHYP0poZ
# EjwcGjr59B0EaIYnLGIear32zRILBWTtBueeLfqgqtc/9ABV4hSULSuiPa+ZtVa0
# rnG704xNGQPUwXsXwy536HfsYdRzCEN1Duz2MBJAjctW/4jXSQzH/DsJv8rnTlOC
# CWtKZk5sUw==
# SIG # End signature block
