#Requires -Version 7.0
<#
.SYNOPSIS
Offline regression checks for CMDB source paths, terminal exit codes and Autopatch startup.
.DESCRIPTION
Executes selected AST statements with synthetic configuration and mocked external
actions. Child processes execute only extracted cleanup blocks, never collectors.
No tenant configuration, authentication, acquisition, upload or mail is used.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$results = [Collections.Generic.List[object]]::new()
function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Case([string]$Name, [scriptblock]$Body) {
    try { & $Body; $results.Add([pscustomobject]@{Name=$Name;Passed=$true;Error=''}) }
    catch { $results.Add([pscustomobject]@{Name=$Name;Passed=$false;Error=$_.Exception.Message}) }
}
function Read-Ast([string]$RelativePath) {
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root $RelativePath),[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    $ast
}
function Import-Function($Ast,[string]$Name) {
    $node=$Ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true)
    if ($null -eq $node) { throw "Missing function: $Name" }
    Set-Item -Path "Function:script:$Name" -Value ([scriptblock]::Create($node.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
}
$producers=@(
    'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1',
    'SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1',
    'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1'
)
foreach ($producer in $producers | Select-Object -Skip 1) {
    $ast=Read-Ast $producer
    Import-Function $ast 'Get-ScriptLocalConfigValue'
    Import-Function $ast 'Resolve-SmartM365ConfigValue'
    $assignment=@($ast.EndBlock.Statements | Where-Object {
        $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -eq '$LatestCsvFolderPath'
    })
    Assert-True ($assignment.Count -eq 1) 'Latest source path must be explicitly initialized once.'
    $guard=@($ast.EndBlock.Statements | Where-Object {
        $_ -is [Management.Automation.Language.IfStatementAst] -and $_.Extent.Text -like '*LatestCsvFolderPath must resolve*'
    })
    Assert-True ($guard.Count -eq 1) 'Missing explicit source-path configuration gate.'
    $pathBlock=[scriptblock]::Create($assignment[0].Extent.Text + "`n" + $guard[0].Extent.Text)
    foreach ($marker in @('__USE_GLOBAL__','USE_GLOBAL','','   ')) {
        Case "$(Split-Path $producer -Leaf) resolves latest root from global / '$marker'" {
            $script:SmartM365GlobalConfig=[pscustomobject]@{LatestCsvFolderPath='C:\Synthetic\DATA-LAST'}
            $ScriptLocalConfig=[pscustomobject]@{LatestCsvFolderPath=$marker}
            . $pathBlock
            Assert-True ($LatestCsvFolderPath -ceq 'C:\Synthetic\DATA-LAST') 'Inherited root was lost.'
        }
    }
    Case "$(Split-Path $producer -Leaf) preserves explicit latest-root override" {
        $script:SmartM365GlobalConfig=[pscustomobject]@{LatestCsvFolderPath='C:\Synthetic\DATA-LAST'}
        $ScriptLocalConfig=[pscustomobject]@{LatestCsvFolderPath='C:\SyntheticOverride\DATA-LAST'}
        . $pathBlock
        Assert-True ($LatestCsvFolderPath -ceq 'C:\SyntheticOverride\DATA-LAST') 'Override was lost.'
    }
    foreach ($missing in @('', '{{UnknownRoot}}\DATA-LAST')) {
        Case "$(Split-Path $producer -Leaf) rejects missing or unresolved source root / '$missing'" {
            $script:SmartM365GlobalConfig=[pscustomobject]@{LatestCsvFolderPath=$missing}
            $ScriptLocalConfig=[pscustomobject]@{LatestCsvFolderPath='__USE_GLOBAL__'}
            $caught=$null
            try { . $pathBlock } catch { $caught=$_ }
            Assert-True ($null -ne $caught -and $caught.Exception.Message -like '*LatestCsvFolderPath must resolve*') 'Invalid path reached publication.'
        }
    }
    Case "$(Split-Path $producer -Leaf) validates latest root before environment and acquisition" {
        $init=$ast.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'InitializeScriptEnvironment'},$true)
        Assert-True ($guard[0].Extent.StartOffset -lt $init.Extent.StartOffset) 'Invalid paths were not stopped before initialization.'
    }
}

# Real process exit propagation, but only the actual terminal cleanup block runs.
$childHarness=@'
$ErrorActionPreference='Stop'
$globalError=$null; $script:ExitCode=0; $connectedGraphInThisRun=$false
$runReachedTerminalState=$true; $TaskName='Synthetic'; $MaxItems=0; $hardwareRows=@()
$DomainWorker=$false; $ReportOnly=$false; $DuplicateAnalysisOnly=$false; $TargetDomains=@()
$global:SmartM365ExecutionStatus=''; $global:SmartM365ErrorCount=0
$global:SmartM365WarningCount=0; $global:logTranscriptFile=$null; $global:LogTextFile=$null
$global:LogPath=$null; $global:RetentionMaxLogs=30; $OutputPath='C:\Synthetic'
function WriteLog { param($Message,$Level) Write-Output $Message }
function Clear-SmartM365AdReferenceXlsxCache {}
function Stop-Transcript {}
function Update-SmartM365TimestampedTranscript { param($Path) }
function Remove-SmartM365TimestampedFilesOlderThan { param($FolderPath,$FilePattern,$RetentionDays,[switch]$RequireCurrentRunPublication,$LogFile) }
function RemoveOldFiles { param($Path,$Filter,$KeepCount,$LogFile) }
function Set-SmartM365CmdbSourceScope { param($CompleteScope,$Scope) }
function Complete-SmartM365ExecutionContext {
    param($Status,$ErrorRecord,$FailureStage)
    if ($script:mockSummaryFailure) { throw 'Synthetic summary failure.' }
    $global:SmartM365ExecutionStatus=if ($ErrorRecord -or $Status -eq 'Failed' -or $global:SmartM365ErrorCount -gt 0) {'Failed'} else {'Success'}
    if ($ErrorRecord) { Write-Output ('Original error: '+$ErrorRecord.Exception.Message) }
    Write-Output 'SUMMARY_REACHED'
}
$script:mockSummaryFailure=$false
'@
foreach ($producer in $producers) {
    $ast=Read-Ast $producer
    $terminal=@($ast.EndBlock.Statements | Where-Object {
        $_ -is [Management.Automation.Language.TryStatementAst] -and $null -ne $_.Finally -and
        $_.Finally.Extent.Text -like '*Complete-SmartM365ExecutionContext*'
    })
    Assert-True ($terminal.Count -eq 1) 'Terminal cleanup is ambiguous.'
    foreach ($scenario in @('Success','CaughtFailure','LoggedFailure','SummaryFailure','Interrupted')) {
        if ($scenario -eq 'Interrupted' -and $producer -notlike '*/IntuneInventory/*') { continue }
        Case "$(Split-Path $producer -Leaf) terminal process exit / $scenario" {
            $setup=switch ($scenario) {
                'CaughtFailure' { '$globalError=[Management.Automation.ErrorRecord]::new([InvalidOperationException]::new("Synthetic export failure."),"Synthetic",[Management.Automation.ErrorCategory]::InvalidOperation,$null)' }
                'LoggedFailure' { '$global:SmartM365ErrorCount=1' }
                'SummaryFailure' { '$script:mockSummaryFailure=$true' }
                'Interrupted' { '$runReachedTerminalState=$false' }
                default { '' }
            }
            $code=$childHarness+"`n"+$setup+"`ntry {} finally "+$terminal[0].Finally.Extent.Text+"`nexit 0"
            $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
            $observed=@(& (Join-Path $PSHOME 'pwsh.exe') -NoProfile -EncodedCommand $encoded 2>&1)
            $observedExit=$LASTEXITCODE
            $expected=if ($scenario -eq 'Success') {0} else {1}
            Assert-True ($observedExit -eq $expected) ("Expected exit {0}, got {1}: {2}" -f $expected,$observedExit,($observed -join "`n"))
            if ($scenario -ne 'SummaryFailure') {
                Assert-True (($observed -join "`n") -like '*SUMMARY_REACHED*') 'Exit bypassed terminal summary.'
            }
            if ($scenario -eq 'CaughtFailure') {
                Assert-True (($observed -join "`n") -like '*Original error: Synthetic export failure.*') 'Cleanup hid the original exception.'
            }
        }
    }
}

$autopatch=Read-Ast 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
$main=@($autopatch.EndBlock.Statements | Where-Object {
    $_ -is [Management.Automation.Language.TryStatementAst] -and $_.Body.Extent.Text -like '*Start-CoreSmartM365CmdbSourceReceipt*'
})
Assert-True ($main.Count -eq 1) 'Receipt startup is outside protected initialization.'
Case 'Autopatch imports Core before starting its receipt or Graph acquisition' {
    $commands=@($main[0].Body.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst]},$true))
    $names=@($commands | ForEach-Object GetCommandName)
    Assert-True ([array]::IndexOf($names,'Import-SmartM365CorePreflight') -lt [array]::IndexOf($names,'Start-CoreSmartM365CmdbSourceReceipt')) 'Receipt precedes module import.'
    Assert-True ([array]::IndexOf($names,'Start-CoreSmartM365CmdbSourceReceipt') -lt [array]::IndexOf($names,'Connect-GraphSession')) 'Graph acquisition precedes proof startup.'
}
foreach ($scenario in @('Success','ModuleFailure','ReceiptFailure')) {
    Case "Autopatch actual main block with mocked external actions / $scenario" {
        Remove-Item Function:script:Start-CoreSmartM365CmdbSourceReceipt -ErrorAction SilentlyContinue
        $script:scenario=$scenario; $script:events=[Collections.Generic.List[string]]::new()
        $script:CompletionStatus='Success'; $script:CompletionError=$null; $script:TranscriptStarted=$false
        $script:GraphTransientRetryCount=0; $global:SmartM365WarningCount=0; $global:SmartM365ErrorCount=0
        $OutputFolder='C:\Synthetic'; $LogFolder='C:\Synthetic'; $TranscriptFile='C:\Synthetic\transcript.log'
        $LogFile='C:\Synthetic\run.log'; $LatestCsvFolderPath='C:\Synthetic\DATA-LAST'
        $ScriptName='Synthetic'; $ScriptVersion='test'; $StartTime=Get-Date
        $IncludeFeatureUpdates=$true; $IncludeQualityUpdates=$true; $MaxItems=0
        $SummaryCsvPath='summary.csv'; $SummaryLatestCsvPath='summary-latest.csv'
        $DetailCsvPath='detail.csv'; $DetailLatestCsvPath='detail-latest.csv'
        $PolicyCsvPath='policy.csv'; $PolicyLatestCsvPath='policy-latest.csv'
        $SummaryColumns=@(); $DetailColumns=@(); $PolicyColumns=@()
        function Ensure-Folder { param($Path) }
        function Start-Transcript { param($Path,[switch]$Force) }
        function Stop-Transcript {}
        function Test-Ps7 {}
        function Import-RequiredModule { param($Name) }
        function Import-SmartM365CorePreflight {
            $script:events.Add('Module')
            if ($script:scenario -eq 'ModuleFailure') { throw 'Synthetic module failure.' }
            Set-Item Function:script:Start-CoreSmartM365CmdbSourceReceipt -Value {
                param($ScriptPath,$SourceRootPath)
                $script:events.Add('Receipt')
                if ($script:scenario -eq 'ReceiptFailure') { throw 'Synthetic receipt failure.' }
            }
        }
        function Connect-GraphSession { $script:events.Add('Graph') }
        function Invoke-CoreSmartM365Preflight { param($ScriptName,$RequiredModules,$OutputPaths,$RequiredGraphApplicationPermissions,$GraphProbeUris) }
        function Get-FeatureUpdatePolicyMap { @{} }
        function Get-QualityUpdatePolicyMap { @{} }
        function Import-ExportedCsv { param($ReportName,$Select,$Filter) }
        function Group-AlertSummary { param($Details) }
        function Publish-CoreSmartM365Csv { param($Data,$TimestampedPath,$LatestPath,$Columns) $script:events.Add('Publish') }
        function Remove-CoreSmartM365TimestampedFilesOlderThan { param($FolderPath,$FilePattern,$RetentionDays,$LogFile) }
        function Disconnect-MgGraph {}
        function Write-Log { param($Message,$Level) if ($Level -eq 'ERROR') { $global:SmartM365ErrorCount++ } }
        function Set-CoreSmartM365CmdbSourceScope { param($CompleteScope,$Scope) }
        function Complete-CoreSmartM365ExecutionContext { param($Status,$ErrorRecord) $script:events.Add('Summary') }
        $caught=$null
        try { & ([scriptblock]::Create($main[0].Extent.Text)) } catch { $caught=$_ }
        if ($scenario -eq 'Success') {
            Assert-True ($null -eq $caught) ('Unexpected startup error: '+$caught)
            Assert-True (($script:events -join '|') -ceq 'Module|Receipt|Graph|Publish|Publish|Publish|Summary') 'Acquisition order or empty publication changed.'
        }
        else {
            Assert-True ($null -ne $caught -and $script:CompletionStatus -eq 'Failed') 'Startup failure was swallowed.'
            Assert-True (-not $script:events.Contains('Graph') -and -not $script:events.Contains('Publish')) 'Failed initialization reached acquisition or publication.'
            Assert-True ($script:events.Contains('Summary')) 'Protected failure bypassed summary.'
        }
    }
}
$results | Format-Table -AutoSize
$failed=@($results | Where-Object { -not $_.Passed })
if ($failed.Count) { $failed | Format-List; throw "$($failed.Count) CMDB runtime repair tests failed." }
Write-Host ("PASS: {0} offline runtime repair checks; production actions=0." -f $results.Count)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD4ahfDycdVxS9M
# NPsWP7tI/kg7T+rU7jbwEXxbk2TkQKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICqqyk8M8/MZueYgbAQmKr2HFV17JkwDNFsp7tj+rHMtMA0GCSqG
# SIb3DQEBAQUABIIBgDzcLYzn5ZxGCdWEm6JpIZIzDmvnj6xvdOiQs31EqX+9MIIU
# 2LOlwRmzWk6svb9/7TZZGvHhjiL3CwgqLA3WIURqOu++hbkeKXSgaer+wppHGNQ0
# 6+ShPaAgu1vsuBhDAy7DYgOL11tPZ3Y6XXyd6I5MIKQPWhW2160sagPAcSG5Hrxj
# ovkIHr1Fvqagrcw6x0KNm9rUxbWmETYZ/LOm6HiSAvQqOnjCDtIBrOwTpzbiMHz8
# IRHnk8Re3e+2Vb9s+KH5gHyCDf2oB7rRsT7/DvDOtwh3/1xR8Is4rZFFJE8g9jAZ
# jg3G+9dDT5emDTINi6V9qGPvnUbCyOchvrlrz3u2DtYAPWbpvm5p8YDmqZ6C4zfP
# JrzWCkhPEuBvMiMdO1zWLApZ2ujd4phMNM7UM5QdTOndSvofGfvMr35ANltbcDnq
# 0cT/kxWie0mnKRlL8Xd37AIeFXr8CsJxl+Zr+uOZpvO7c7pj9Hv8kPcgdYZnoowo
# JqoIx5AmPz8NGQk+mKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNDE2
# MzNaMC8GCSqGSIb3DQEJBDEiBCCoARtyi/krasmNhX4Ra2XftezFeSSSU0wX7s7y
# qecHrzANBgkqhkiG9w0BAQEFAASCAgCczu4+gb0UjYdohZXsqeQK/Z+iu4Me461+
# TUBpdESrD62Qh4hUjYyAkHQlb3pydCgnJjowvcs9ETnPFU+BedhYpWut29tX0r9Z
# h+SvuN8e4YiXNoHmPfq5US06Av9k1sKss5jNrtWZhJBfSqB8yJOScwAZOWZWdic/
# XD5zU0u1+B64qZHz+41a4bG4Tk5w/w9Y9NYG0fFKgQMoAK3rZuzkJ/INkKs+1kxN
# OGSpa1pksiVmUbg+3kOdF/stXtA/FFS8PkHp9SBQ4gJy+aAv7UkNsHqTkdrHGsNW
# JCfh8lYvTY1ra/G87rgHoQgUBIEwSR0sVMWzkyM7D4KLglY8EGj4B3qPR+ch76tW
# 6LRpAcUrIcMk0esOruRni1bxFM6HMvE7fSSeoMBPIAPNTF2P1HkAvcRWg+RD05Ug
# vIjThVj4DURCRMfgH7wKtjITt8R2F6sUhGThjb3fMu50XbYsWJ8T61ZO0KXZJ9Nn
# OHrg8BwgTCU0O+zuL+MV9EAmO2m54DmthtJBIlrzVVas5BQqCKLfGrU7K4lS64Wb
# GS2QeZ8Dw/PcSRjTPPkjQ6j5fEvZnoV40N2er5ePTezftMl18MxUsMMBYqEuF6of
# V2Ext0XEQbmyMWHX+YCkCQacZpxcPnnIsJFkXFv7DmAiv0vkpH4eM2BHmwayzUKq
# 5xTWcBo+0g==
# SIG # End signature block
