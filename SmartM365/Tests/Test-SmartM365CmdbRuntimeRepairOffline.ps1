#Requires -Version 7.0
<#
.SYNOPSIS
Offline regression checks for CMDB source paths, terminal exit codes and Autopatch startup.
.DESCRIPTION
Executes selected AST statements with synthetic configuration and mocked external
actions. Child processes execute only extracted cleanup blocks, never collectors.
No tenant configuration, authentication, acquisition, upload or mail is used.
.VERSION
1.0.2
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
function Get-InventoryRunBody($Ast) {
    # Collector bodies run inside the shared inventory run lifecycle guard when the collector has one.
    $outer=@($Ast.EndBlock.Statements | Where-Object {
        $_ -is [Management.Automation.Language.TryStatementAst] -and $null -ne $_.Finally -and
        $_.Finally.Extent.Text -match 'Complete-SmartM365InventoryRun'
    })
    Assert-True ($outer.Count -le 1) 'Inventory run lifecycle guard is ambiguous.'
    if ($outer.Count -eq 1) { return $outer[0].Body.Statements }
    $Ast.EndBlock.Statements
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
    $assignment=@((Get-InventoryRunBody $ast) | Where-Object {
        $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -eq '$LatestCsvFolderPath'
    })
    Assert-True ($assignment.Count -eq 1) 'Latest source path must be explicitly initialized once.'
    $guard=@((Get-InventoryRunBody $ast) | Where-Object {
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
function Stop-SmartM365RunTranscript {}
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
$smartM365RunGuard=[pscustomobject]@{InCleanup=$false;ReachedTerminal=$true;ErrorRecord=$null}
'@
foreach ($producer in $producers) {
    $ast=Read-Ast $producer
    $terminal=@((Get-InventoryRunBody $ast) | Where-Object {
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
$main=@((Get-InventoryRunBody $autopatch) | Where-Object {
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
        $smartM365RunGuard=[pscustomobject]@{InCleanup=$false;ReachedTerminal=$false;ErrorRecord=$null}
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
        function Enable-CoreSmartM365DailyCaptureRetention { param($OutputRoot) $null=$OutputRoot }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAAoVT1U+LmTLFx
# YbWhq/yS5XyXRG9KwRugtvGP5OEGAqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOqEk2p/oP+tauN+2HNu9g+GRmDSkS4HsCtbouPHem7kMA0GCSqG
# SIb3DQEBAQUABIIBgDdtLrS6x3JeX8hoUcB1DLS0t8Q0ttcnORKXV9Y1HsLfQcD3
# YYeuofezZyLOwo4hGNZ8FkJ61tH3+Ho2byJUX98NvlqE//dCN3dtsN50EKwOT2uH
# n2wH0ybdHCZsUaRrpRTD9j5C8k1u0r3MHSUjBgl/AUoUoXdH7xmye55JxkF+PEE4
# FF0dN1/UGcQhG8rZjguIFS5Qz5o2qFsJrRF9yQcJNF/8mOU8psOZDcJjCechZvNg
# 9jUK/Y+dwc5xLu0YWKLc1nckUAi33JX0I7WpRftiG5tjK0S94X+dXezyeNKF27ro
# DdJgXuDz63LbCoq8Lk5U4tON5gn+8+VC2mp3n+gXiLVyacVETtZgjKKXAa6/WH4S
# GnR+dDAbOqpQ3/YDiWGPzl1p0a+3UBiyR1GzJXkg+n+C8x/DHIQvsPZPv9Gl6NW6
# GGXFWmBqQ4sOhJDTPERWzUYYvGKjj9/Npp1aIrNZypR+iAv8nc1f8IgN11qp3/f1
# jhO6MzMA3Hs48A+KtKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAxMzA2
# MThaMC8GCSqGSIb3DQEJBDEiBCDjTDcZY3yyhHCE/ULpvuEUe1m8D8Ho5YgbL2eg
# Gx7yhTANBgkqhkiG9w0BAQEFAASCAgAVyNrxrIhnWQgGSEpfhWlLgHA9/L+5zUeW
# 3e5+MxN82DaW17UCclexDeAtLKF8W0WbywadRbdJIRlE3r4KQwvN7ivS22dYDcvU
# O8D6tBMKHr/ZiEsbDEgv2FFj5+uTI+gozo/d4YU6LVZKRESAJ7ByNqVTdwVSwh33
# AGPYHVhtqfuCUXgAXuk6I1Zu8V70UCNt54txrFJsLJIgXSXc8CJ824YwUr8NcArP
# vHqgYbjkNA5KSzJRSVoBCQQGWrqT0IwlNbHqYuef8Pv6h+m7eKIFpw2p7S8vnyog
# 6Wwknl2PoR4R26ZapKx/JhH/VxEYtKJhkS5zDHrMIsLa2EGBfF+yPF3+nv7YaVhJ
# D1Ft6PCbtj6VLcposj4dhjPAdJMUCf4RJuPg3dPwdcsWK6/46JPO8VsJOtUw55mX
# iJ/Mo7vwFt4b9WnBRtEfH6lTg64gnUuoHPbsmd0o97U0URhuABajqkQuzpY8equ7
# 2It5EG8dnKUuGbY48sCTZunH0TcRErFPV6FfMBx3fgAPep4yPfBti2gWJz8c4XPL
# Lasacs8rEvj1cFUTbj3Qq21FXmupfc6WWBDKTCdIki0Dtxq3hpRxZRQncv9Unnnc
# 26cAvjL5RBYoQY0DiR2qt0KcUcmJJP8XgK2lLbCLQBOJLjTglDZhwtQw+7P2Xhdy
# rRKrSVsFyQ==
# SIG # End signature block
