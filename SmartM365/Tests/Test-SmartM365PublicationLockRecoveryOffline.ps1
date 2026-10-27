#Requires -Version 5.1
<#
.SYNOPSIS
Offline publication lock recovery and terminal exit-code regression tests.
.VERSION
1.0.0
.NOTES
Synthetic TEMP fixtures only. Network, retention and notification helpers are replaced.
Real Windows file handles exercise sharing violations; sleeps are recorded, not elapsed.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic identity and terminal state are confined to fresh child processes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline stubs retain the Core helper signatures.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='The actual Devices exit predicate is parsed and executed against these fixture variables.')]
param([ValidateSet('', 'Success', 'Transient', 'Lock', 'Scope', 'ReceiptWrite', 'Cancelled')][string]$ChildCase='', [switch]$Lifecycle)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$core=Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.107' -Force -PassThru
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-PublicationLock-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$script:checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Hash([string]$Path){(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash}
& $core {
    $script:ProbeLogs=New-Object 'Collections.Generic.List[string]'
    $script:ProbeSleeps=New-Object 'Collections.Generic.List[int]'
    $script:ProbeUploads=New-Object 'Collections.Generic.List[string]'
    $script:ReleaseLock=$null
    function script:WriteLog {
        param($Message,$Level)
        $script:ProbeLogs.Add([string]$Message)
        if($Level -eq 'ERROR'){$global:SmartM365ErrorCount++}
        if($Level -eq 'WARNING'){$global:SmartM365WarningCount++}
    }
    function script:Start-Sleep {
        param([int]$Seconds)
        $script:ProbeSleeps.Add($Seconds)
        if($script:ReleaseLock){$script:ReleaseLock.Dispose();$script:ReleaseLock=$null}
    }
    function script:Invoke-SmartM365SharePointCsvUpload {param($LocalFilePath) $script:ProbeUploads.Add([string]$LocalFilePath)}
    function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv {param($SourceFiles,$TimestampedPath) throw 'Unexpected history call'}
    function script:Invoke-SmartM365DailyCaptureRetention {param($Context,$Proof)}
    function script:Assert-SmartM365CsvDataCompleteness {param($Data,$Columns,$BaseFileName,$TimestampedPath,$LatestPath)}
    function script:Write-SmartM365CompletionBanner {param($Status,$ScriptName,$StartedAt,$EndedAt,$WarningCount,$ErrorCount,$GeneratedCsvFiles)}
}
$lock=$null
try {
    if($ChildCase){
        $latest=Join-Path $temporary 'DATA-LAST'
        $output=Join-Path $temporary 'DATA-ALL'
        New-Item -ItemType Directory -Path $latest,$output | Out-Null
        $global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='contoso'
        $global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
        $global:SmartM365ExecutionSummaryWritten=$false
        $global:SmartM365ErrorCount=0;$global:SmartM365WarningCount=0
        $producer=Join-Path $root 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1'
        $files=@('Intune_DeviceHardware_All.csv','Intune_Devices_Inventory.csv','Intune_ManagedDevices_All.csv')
        function Stage([string]$Label){
            $global:csvGeneratedPaths=@()
            Start-SmartM365CmdbSourceReceipt -ScriptPath $producer -SourceRootPath $latest
            Enable-SmartM365StagedSourcePublication
            foreach($file in $files){
                Export-SmartM365Csv -Data @([pscustomobject]@{Id='contoso-device';Name=$Label}) -Columns @('Id','Name') -TimestampedPath (Join-Path $output ($Label+'-'+$file)) -LatestPath (Join-Path $latest $file) -NoWeeklyHistory | Out-Null
            }
            Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope 'CMDB:managed,hardware'
        }
        Stage 'baseline'
        Complete-SmartM365SourceReceipt -Status Completed | Out-Null
        $receipt=Join-Path $latest 'SmartInventory_SmartM365-Devices-Inventory.current.json.txt'
        $runPath=Join-Path $latest 'SmartInventory_SmartM365-Devices-Inventory.run.json.txt'
        $before=Hash $receipt
        Stage 'candidate'
        $global:SmartM365ErrorCount=0;$global:SmartM365WarningCount=0
        & $core {$script:ProbeUploads.Clear();$script:ProbeSleeps.Clear()}
        if($ChildCase -in @('Lock','Transient')){
            $lock=[IO.File]::Open((Join-Path $latest $files[1]),'Open','Read','None')
            if($ChildCase -eq 'Transient'){& $core {param($h) $script:ReleaseLock=$h} $lock}
        }
        if($ChildCase -eq 'Scope'){Set-SmartM365CmdbSourceScope -CompleteScope $false -Scope 'CMDB:managed,hardware'}
        if($ChildCase -eq 'ReceiptWrite'){
            & $core {function script:Write-SmartM365CmdbReceiptDocument {param($Context,$Document,[switch]$RunState) throw 'Synthetic receipt write failure'}}
        }
        $status=if($ChildCase -eq 'Cancelled'){'Cancelled'}else{'Success'}
        if($Lifecycle){
            $state=& $core {param($p) Start-SmartM365InventoryRun -ScriptPath $p} $producer
            $state.Initialized=$true
            Complete-SmartM365ExecutionContext -Status $status
            & $core {param($s) Complete-SmartM365InventoryRun -State $s} $state
            $exitCode=$state.ExitCode
        }else{
            Complete-SmartM365ExecutionContext -Status $status
            # Execute the real Devices collector's terminal exit-code predicate (no collector body).
            $tokens=$null;$errors=$null
            $ast=[Management.Automation.Language.Parser]::ParseFile($producer,[ref]$tokens,[ref]$errors)
            $node=$ast.FindAll({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if ($null -ne $completionError -or $completionStatus')},$true) | Select-Object -Last 1
            if(-not $node){throw 'Devices terminal exit-code predicate missing'}
            $completionError=$null;$completionStatus=$status;$script:ExitCode=0
            . ([scriptblock]::Create($node.Extent.Text))
            $exitCode=$script:ExitCode
        }
        if($lock){$lock.Dispose();$lock=$null}
        if($ChildCase -ne 'ReceiptWrite'){
            $run=Get-Content -LiteralPath $runPath -Raw | ConvertFrom-Json
            if($ChildCase -in @('Success','Transient')){
                Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication) 'Successful publication not qualified'
                $proof=Get-Content -LiteralPath $receipt -Raw | ConvertFrom-Json
                foreach($file in $proof.Files){Check ((Hash (Join-Path $latest $file.File)) -eq $file.SHA256) 'Receipt SHA mismatch'}
            }else{
                Check ((Hash $receipt) -eq $before) 'Failed/cancelled run replaced the previous receipt'
                Check ($run.Status -eq 'Failed') 'Run not recorded as failed'
                $csvUploads=@(& $core {$script:ProbeUploads.ToArray()} | Where-Object {$_ -like '*.csv'})
                Check ($csvUploads.Count -eq 0) 'Failed/cancelled CSVs uploaded'
                if($ChildCase -eq 'Lock'){
                    Check $run.UnqualifiedPublication 'Partially promoted run lost its fence'
                    Check ($run.Error.Contains($files[1]) -and $run.Error.Contains('attempt=6/6')) 'Failure lacks filename/attempt'
                    Check ((@(& $core {$script:ProbeSleeps.ToArray()}) -join ',') -eq '2,4,8,16,30') 'Unexpected retry budget'
                    $rejected=$false
                    try{Assert-SmartM365SourcePublication -ReceiptPath $receipt -Proof (Get-Content $receipt -Raw | ConvertFrom-Json)}catch{$rejected=$true}
                    Check $rejected 'Reader accepted a partial set'
                }
            }
        }
        $expected=if($ChildCase -in @('Success','Transient')){0}elseif($ChildCase -eq 'Cancelled'){130}else{1}
        Check ($exitCode -eq $expected) "Wrong exit code for $ChildCase : $exitCode instead of $expected"
        if($expected -eq 1){Check ($global:SmartM365ExecutionStatus -eq 'Failed' -and $global:SmartM365ErrorCount -gt 0) 'Terminal status/error count not failed'}
        "CHILD PASS: $ChildCase lifecycle=$Lifecycle exit=$exitCode checks=$script:checks"
    }else{
        $source=Join-Path $temporary 'stage.csv';$target=Join-Path $temporary 'current.csv'
        [IO.File]::WriteAllText($source,'new');[IO.File]::WriteAllText($target,'old')
        $expected=Hash $source
        $lock=[IO.File]::Open($target,'Open','Read','None')
        & $core {param($handle) $script:ReleaseLock=$handle} $lock
        & $core {param($s,$d) Move-SmartM365StagedSourceFile $s $d} $source $target
        $lock.Dispose();$lock=$null
        Check ((Hash $target) -eq $expected -and -not (Test-Path $source)) 'Transient lock did not recover'
        Check ((@(& $core {$script:ProbeSleeps.ToArray()}) -join ',') -eq '2') 'Transient lock not retried once'
        Check ((& $core {$script:ProbeLogs -join "`n"}).Contains($target)) 'Retry log lacks destination'
        # Absent destination uses Move. Missing source is a non-lock error: no retry.
        [IO.File]::WriteAllText($source,'move')
        $newTarget=Join-Path $temporary 'new.csv'
        & $core {param($s,$d) Move-SmartM365StagedSourceFile $s $d} $source $newTarget
        Check ([IO.File]::ReadAllText($newTarget) -eq 'move') 'Move failed'
        & $core {$script:ProbeSleeps.Clear()}
        $message=''
        try{& $core {param($s,$d) Move-SmartM365StagedSourceFile $s $d} $source $target}catch{$message=$_.Exception.Message}
        Check ($message.Contains($source) -and $message.Contains($target) -and $message.Contains('HResult=')) 'Permanent error lacks paths/code'
        Check (@(& $core {$script:ProbeSleeps.ToArray()}).Count -eq 0) 'Non-lock failure retried'
        # Artifacts use the same retry path and still verify SHA256 after promotion.
        [IO.File]::WriteAllText($source,'artifact')
        $lock=[IO.File]::Open($target,'Open','Read','None')
        & $core {
            param($s,$d,$handle)
            $script:ReleaseLock=$handle
            function script:Write-SmartM365SourceRunState {param($Context,$Status) if(-not $Context.UnqualifiedPublication -or $Status -ne 'Publishing'){throw 'Fence missing'}}
            $context=@{SourceRoot=[IO.Path]::GetDirectoryName($d);StagedFiles=@{};StagedArtifacts=@{};PublicationStarted=$false;UnqualifiedPublication=$false}
            $context.StagedArtifacts[$d]=$s
            Publish-SmartM365StagedSourceSet -Context $context -Records @(@{File='not-staged.csv'})
        } $source $target $lock
        $lock.Dispose();$lock=$null
        Check ([IO.File]::ReadAllText($target) -eq 'artifact') 'Artifact recovery failed'
        # Fresh processes exercise actual exit propagation, independent from this test's globals.
        $engine=(Get-Process -Id $PID).Path
        foreach($case in @('Success','Transient','Lock','Scope','ReceiptWrite')){
            foreach($mode in @($false,$true)){
                $arguments=@('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$PSCommandPath,'-ChildCase',$case)
                if($mode){$arguments+='-Lifecycle'}
                $output=& $engine @arguments 2>&1
                $code=$LASTEXITCODE
                Check ($code -eq $(if($case -in @('Success','Transient')){0}else{1}) -and ($output -join "`n") -match 'CHILD PASS:') "Child failed: $case lifecycle=$mode code=$code $output"
            }
        }
        $output=& $engine -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $PSCommandPath -ChildCase Cancelled -Lifecycle 2>&1
        Check ($LASTEXITCODE -eq 130 -and ($output -join "`n") -match 'CHILD PASS:') "Cancellation changed: $output"
        "PASS: $script:checks lock recovery/child-process checks; PowerShell $($PSVersionTable.PSVersion)."
    }
}finally{
    if($lock){$lock.Dispose()}
    & $core {if($script:SmartM365CmdbSourceContext -and $script:SmartM365CmdbSourceContext.Lock){$script:SmartM365CmdbSourceContext.Lock.Dispose()}}
    Remove-Module $core -Force -ErrorAction SilentlyContinue
    $resolved=[IO.Path]::GetFullPath($temporary)
    $tempBase=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($tempBase,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notlike 'SmartM365-PublicationLock-*'){throw 'Unsafe test cleanup path'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
if($ChildCase){exit $exitCode}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAanInQD8IrYpK2
# VLo/VCGxLiVGiCExz86X5qG5DQGYkaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMm3W2KlKybLLaW0+Ej0EiaFGDxVKdQrHFuTga9wRb4aMA0GCSqG
# SIb3DQEBAQUABIIBgHCuiTb19oyKo2u8zbvnxRBreSTMtES+C8w/YkmN6zW5zEQM
# tDg6OpShhIQfQOsFaPda3KaibAhXJZBZFV9vmz6wP0TkhsMQTsL6GPFx/sQ5305Y
# N8m+7M9V5KhfmDOi41egZwhpNKuc8H+3L1kXFoXCa6lYLNpbsB9Z1S7OYMW4jNKj
# 1XyfoU6x0RpxWR69k/80S98TPheC+onbZHqPWwVxDT/FNhtvP/epLK8FdBe8VKjS
# 3ZTEBNlD9+yZvaHbXtc5Kwvwgzs2zpmHBaOXPTK1pFvwoTQQqRsaS226s7S+9LrU
# yTuLAij0VwDP3m+eOICjLcll7rsNQFMmWwOS07hHYz6LSWX/xJK3fxDxFOywW4hN
# h21MTJhOTRQ6gKAUbKnMZwi/vgWXnyR/dpNITIAM+emsOVR3NWG7t+zyh5oCSKvO
# Rwvh99gqYocW6H42Da4MiiUz+RR3/iqcypxRgFw2+oYOVINaK58FbwNKr8t+2UXa
# sH4xBYFqr3H56wL4tKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMDQ3
# MzBaMC8GCSqGSIb3DQEJBDEiBCD5JE56fvU4GvU6ebwabiTaRv8IR3M9lt51Y4ac
# 4h5KJDANBgkqhkiG9w0BAQEFAASCAgCLlesmuNpa/FteYQFd43uG1vvMGZ0cD8lE
# I8elHiSy89/uKWqrX7Fy/py1jASgdSnzdNz1sabAHwlOB7Ncx+HbGgCe5mQSPSSH
# BpUGsU1G6UNz/jecy8JZQHy5YnR/Zq+nDwYm1j0XPFLUO5n2EAveXg8lLc+yofiE
# PUv8LJhSZzbHuFy59AhxPhbIOnKqfCWgPxV7r1vpdX+Nu8SLJmfEp3e3B3fVg9E6
# Umk/A6SP5Mu8rIcqPc8x+h0HGUTh/oPf3/jGgKFQIULZ5KGQtnGDHcwmrCJOWukr
# vZ5lDrOrmOb4kB8D9uNoyWJDZq+bOL0S/uNREogbDwSrLm+6PZDkEa5XUxi6ByWB
# 4RWxBRPJ4CmAmyRU7zpAmJiqIIdxp8XABsQWQtCwVejYSm4waA3POsLFhiENeMG3
# HLwimjTt+OcU9WCHVI31rNqO2T6gOMJTAefhqpnRopwfYwdykeGb595g3veWJqU7
# QawYorUWsw2iAtFTwdKMcUAMYx1neUBWxEJAh+NShXnEIRcYdjOWk9mNa9/d3G4P
# msxGYcp3BfXx5E8m5+XqSNeoRyMTQsPeH7QkQce71m6MyFtXHQhi5TVONK09lTL7
# ShBZYzT8YMkXMkXe/wFy1U25rxtKCec/bjKD1NiB+zepBOvvDF/3BWoxn9Z7oiBg
# 9npc9wzrUA==
# SIG # End signature block
