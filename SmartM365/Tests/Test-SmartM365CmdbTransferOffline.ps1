#Requires -Version 7.0
<# .SYNOPSIS Offline transport fault tests. Synthetic fixtures and callbacks only. #>
[CmdletBinding()]
param([string]$PythonCommand='python')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CmdbTransferTest-'+[guid]::NewGuid().ToString('N'))
$python=(Get-Command $PythonCommand -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
$base=Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartInventory/PreparedEvidence'
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Rejected([scriptblock]$Action,[string]$Pattern){
    $caught=$null
    try {& $Action | Out-Null} catch {$caught=$_}
    Check ($null -ne $caught -and $caught.Exception.Message -match $Pattern) "Expected rejection: $Pattern; got $caught"
}
try {
    $null=New-Item -Path $root -ItemType Directory
    & $python -B (Join-Path $PSScriptRoot 'test_cmdb_transfer.py') --fixture-root $root
    if($LASTEXITCODE -ne 0){throw 'Synthetic fixture generation failed.'}
    Import-Module (Join-Path $base 'SmartM365-CmdbSharePointTransfer.psm1') -Force
    $prepared=Join-Path $root 'DATA-POWERBI-CMDB'
    $manifest=Join-Path $prepared 'current.json.txt'
    $hash=(Get-FileHash -LiteralPath $manifest -Algorithm SHA256).Hash
    $identity=@{TenantKey='synthetic';OrganizationKey='test';EnvironmentKey='test';TenantId='synthetic-tenant'}
    $parameters=@{PreparedRoot=$prepared;Identity=$identity;PythonPath=$python;ExpectedManifestSHA256=$hash}
    $baseline=@{}
    foreach($file in Get-ChildItem -LiteralPath $prepared){$baseline[$file.Name]=[IO.File]::ReadAllBytes($file.FullName)}
    $state=@{Uploads=[Collections.Generic.List[string]]::new();Cloud=@{};FailUpload='';FailDownload='';Corrupt='';Mutate='';ThrowUpload=$false}
    $upload={
        param($local,$name)
        $state.Uploads.Add($name)
        if($state.ThrowUpload){throw 'Synthetic upload exception'}
        if($state.FailUpload -eq $name){return $null}
        $state.Cloud[$name]=[IO.File]::ReadAllBytes($local)
        if($state.Mutate -and $name -eq $state.Mutate){[IO.File]::AppendAllText($local,'changed')}
        return $true
    }.GetNewClosure()
    $download={
        param($destination,$name)
        if($state.FailDownload -eq $name){return $null}
        [IO.File]::WriteAllBytes($destination,$state.Cloud[$name])
        if($state.Corrupt -eq $name){[IO.File]::AppendAllText($destination,'corrupt')}
        return $true
    }.GetNewClosure()
    function Reset-TestState {
        foreach($name in $baseline.Keys){[IO.File]::WriteAllBytes((Join-Path $prepared $name),$baseline[$name])}
        $state.Uploads.Clear();$state.Cloud.Clear()
        $state.FailUpload='';$state.FailDownload='';$state.Corrupt='';$state.Mutate='';$state.ThrowUpload=$false
    }
    function Transfer {Send-SmartM365CmdbPreparedSnapshot @parameters -UploadFile $upload -DownloadFile $download}
    $plan=Get-SmartM365CmdbTransferPlan @parameters
    Check ($plan.CsvFiles -eq 46 -and @($plan.Files).Count -eq 47) 'Expected 46 CSVs plus manifest.'
    Check ($plan.Files[-1].Name -ceq 'current.json.txt') 'Manifest must be last.'
    Check ((Resolve-SmartM365CmdbSharePointFolder 'SYNTHETIC/DATA') -ceq 'SYNTHETIC/DATA/DATA-POWERBI-CMDB') 'Wrong isolated target.'
    foreach($target in 'SYNTHETIC/DATA-POWERBI','SYNTHETIC/DATA-LAST','SYNTHETIC/DATA/../DATA','__USE_GLOBAL__','SYNTHETIC//DATA'){
        Rejected {Resolve-SmartM365CmdbSharePointFolder $target} 'resolved SharePoint DATA root'
    }
    $result=Transfer
    Check ($result.Status -eq 'PublishedAndReadBack' -and $result.VerifiedFiles -eq 47 -and $result.ManifestVerified) 'Expected verified publication.'
    Check ($state.Uploads.Count -eq 47 -and $state.Uploads[-1] -ceq 'current.json.txt') 'Pointer not last.'
    foreach($name in $baseline.Keys){Check ((Get-FileHash -LiteralPath (Join-Path $prepared $name)).Hash -eq $plan.Files.Where({$_.Name -ceq $name})[0].SHA256) "Changed local file: $name"}
    $first=$plan.Files[0].Name
    foreach($scenario in 'FailUpload','FailDownload','Corrupt','Mutate'){
        Reset-TestState
        $state[$scenario]=$first
        Rejected {Transfer} $(if($scenario -eq 'FailUpload'){'upload was not confirmed'}elseif($scenario -eq 'FailDownload'){'read-back was not confirmed'}else{'byte mismatch'})
        Check (-not $state.Uploads.Contains('current.json.txt')) "Manifest uploaded after $scenario."
    }
    Reset-TestState;$state.ThrowUpload=$true
    Rejected {Transfer} 'Synthetic upload exception'
    Check (-not $state.Uploads.Contains('current.json.txt')) 'Manifest uploaded after exception.'
    Reset-TestState;$state.Corrupt='current.json.txt'
    Rejected {Transfer} 'byte mismatch'
    # Failure after pointer upload is explicitly not reported as verified success.
    Check ($state.Uploads[-1] -eq 'current.json.txt') 'Expected final-pointer readback failure.'
    Reset-TestState
    $parameters.ExpectedManifestSHA256='0'*64
    Rejected {Transfer} 'Expected manifest'
    Check ($state.Uploads.Count -eq 0) 'Hash pin failure uploaded artifacts.'
    $parameters.ExpectedManifestSHA256=$hash
    $identity.TenantId='foreign'
    Rejected {Transfer} 'identity mismatch'
    Check ($state.Uploads.Count -eq 0) 'Foreign identity uploaded artifacts.'
    $identity.TenantId='synthetic-tenant'
    $lock=[IO.File]::Open((Join-Path $root '.cmdb-preparation.lock'),[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::ReadWrite)
    try {$lock.Lock(0,1);Rejected {Transfer} 'process|portion|verrou|acc.*s|locked|utilis'} finally {$lock.Unlock(0,1);$lock.Dispose()}
    Check ($state.Uploads.Count -eq 0) 'Concurrent lock uploaded artifacts.'
    # Prove that the PowerShell lock is also recognized by the real Python generator.
    $lock=[IO.File]::Open((Join-Path $root '.cmdb-preparation.lock'),[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::ReadWrite)
    try {
        $lock.Lock(0,1)
        $code='import sys; from pathlib import Path; sys.path.insert(0,sys.argv[1]); from cmdb_prepare import PublicationLock; lock=PublicationLock(Path(sys.argv[2])); lock.__enter__()'
        $message=& $python -B -c $code $base (Join-Path $root '.cmdb-preparation.lock') 2>&1
        Check ($LASTEXITCODE -ne 0 -and ($message -join '') -match 'owns the publication lock') 'Python preparation did not recognize transfer lock.'
    } finally {$lock.Unlock(0,1);$lock.Dispose()}
    # Acquisition expiry is not bypassed by a newly written manifest.
    $doc=Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json
    foreach($record in $doc.SourceEvidence.Files){$record.StartedAtUtc=[DateTimeOffset]::UtcNow.AddDays(-20).ToString('o');$record.CompletedAtUtc=$record.StartedAtUtc}
    [IO.File]::WriteAllText($manifest,($doc | ConvertTo-Json -Depth 40))
    $parameters.ExpectedManifestSHA256=(Get-FileHash -LiteralPath $manifest).Hash
    Rejected {Transfer} 'Stale acquisition'
    Check ($state.Uploads.Count -eq 0) 'Expired evidence uploaded artifacts.'
    Reset-TestState
    # Verify the shared deadline guard, without sleeping or tenant access.
    $module=Get-Module SmartM365-CmdbSharePointTransfer
    Rejected {& $module {Assert-SmartM365CmdbTransferFreshness -Plan ([pscustomobject]@{EarliestSourceExpiryUtc=[DateTimeOffset]::UtcNow.AddSeconds(-1).ToString('o')})}} 'expired during transfer'
    foreach($path in @((Join-Path $base 'SmartM365-CmdbEvidence-Publish.ps1'),(Join-Path $base 'SmartM365-CmdbSharePointTransfer.psm1'))){
        $tokens=$null;$errors=$null
        $null=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
        Check ($errors.Count -eq 0) "PowerShell parse failed: $path"
    }
    $tokens=$null;$errors=$null
    $wrapperAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-CmdbEvidence-Publish.ps1'),[ref]$tokens,[ref]$errors)
    $offlineBranch=$wrapperAst.Find({param($node)
        $node -is [Management.Automation.Language.IfStatementAst] -and
        $node.Clauses[0].Item1.Extent.Text -ceq '$ValidateOnly'
    },$true)
    Check ($null -ne $offlineBranch) 'Missing offline validation branch.'
    $offlineCommands=@($offlineBranch.Clauses[0].Item2.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst]},$true) |
        ForEach-Object {$_.GetCommandName()})
    Check ($offlineCommands -contains 'Get-SmartM365CmdbTransferPlan' -and
           -not @($offlineCommands | Where-Object {$_ -match 'SharePoint|Send-|Connect-|Invoke-(RestMethod|WebRequest)'}).Count) 'ValidateOnly contains external actions.'
    Check ($wrapperAst.Extent.Text.Contains('$global:EnableSharePointUpload=$false') -and
           $wrapperAst.Extent.Text.Contains('$script:SmartM365TeamsNotificationInProgress=$true') -and
           $wrapperAst.Extent.Text.Contains('$script:SmartM365TeamsNotificationInProgress=$previous')) 'Automatic external callbacks must be disabled and restored.'
    $firstSend=$wrapperAst.Extent.Text.IndexOf('Send-SmartM365CmdbPreparedSnapshot @parameters')
    Check ($firstSend -gt $wrapperAst.Extent.Text.IndexOf('} else {', $offlineBranch.Extent.StartOffset)) 'Transport must remain in the non-validation branch.'
    # Execute the unchanged wrapper in a child process against an isolated fake
    # runtime. External APIs throw if reached; no operational configuration or
    # Core module is loaded. This tests native argument binding and route output.
    $fakeSmart=Join-Path $root 'runtime/SmartM365'
    $fakePrepared=Join-Path $fakeSmart 'SmartInventory/PreparedEvidence'
    $fakeCore=Join-Path $fakeSmart 'Modules/SmartM365.Core'
    $fakeConfig=Join-Path $fakeSmart 'Config'
    foreach($directory in $fakePrepared,$fakeCore,$fakeConfig){$null=New-Item -Path $directory -ItemType Directory -Force}
    foreach($name in 'SmartM365-CmdbEvidence-Publish.ps1','SmartM365-CmdbSharePointTransfer.psm1','cmdb_transfer.py','cmdb_freshness.py','cmdb-prepared-contract.json.txt'){
        Copy-Item -LiteralPath (Join-Path $base $name) -Destination (Join-Path $fakePrepared $name)
    }
    $settings=@{LatestCsvFolderPath=(Join-Path $root 'DATA-LAST');LogAllRootPath=(Join-Path $root 'LOG-ALL');
        PythonCommand=$python;TenantKey='synthetic';OrganizationKey='test';EnvironmentKey='test';TenantId='synthetic-tenant';
        SharePointTargetFolderPath='SYNTHETIC/CSV';SharePointSiteHostname='synthetic.invalid';SharePointSitePath='/sites/synthetic';
        SharePointLibraryDisplayName='Documents';AppId='synthetic-app';Thumb='synthetic-certificate'}
    [IO.File]::WriteAllText((Join-Path $fakeConfig 'fixture.json.txt'),($settings | ConvertTo-Json))
    [IO.File]::WriteAllText((Join-Path $fakeConfig 'SmartM365-TenantContext.ps1'), @'
function Initialize-SmartM365TenantContext {
    param($Tenant,$StartPath)
    Get-Content -LiteralPath (Join-Path $PSScriptRoot 'fixture.json.txt') -Raw | ConvertFrom-Json
}
'@)
    [IO.File]::WriteAllText((Join-Path $fakeCore 'SmartM365.Core.psd1'), "@{RootModule='SmartM365.Core.psm1';ModuleVersion='1.0.77';GUID='0a5e1631-8547-4088-8abf-8d2c99f0cd61'}")
    [IO.File]::WriteAllText((Join-Path $fakeCore 'SmartM365.Core.psm1'), @'
$script:Fixture=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../Config/fixture.json.txt') -Raw | ConvertFrom-Json -AsHashtable
function Read-SmartM365JsonConfig {param($Path,[switch]$Required) return $script:Fixture}
function Resolve-SmartM365ConfigValue {param($Value) return $Value}
function ConvertTo-SmartM365SharePointDataRootPath {param($TargetFolderPath) return $TargetFolderPath -replace '/CSV$','/DATA'}
function InitializeScriptEnvironment {
    param($OutputPathInit,$LogFileName,$CallerScriptPath)
    $null=New-Item -Path $OutputPathInit -ItemType Directory -Force
    $global:logTranscriptFile=Join-Path $OutputPathInit 'fixture-transcript.log'
}
function WriteLog {param($Message,$Level) Write-Output "$Level $Message"}
function Complete-SmartM365ExecutionContext {
    param($Status,$ErrorRecord,$FailureStage)
    if($global:EnableSharePointUpload -or $global:EnableTeamsNotifications -or -not $script:SmartM365TeamsNotificationInProgress){throw 'Offline callbacks were not suppressed.'}
    Write-Output "OFFLINE_COMPLETION $Status"
}
function Invoke-SmartM365SharePointCsvUpload {throw 'External upload reached in offline mode.'}
function Invoke-SmartM365SharePointFileDownload {throw 'External download reached in offline mode.'}
function Connect-MgGraph {throw 'Authentication reached in offline mode.'}
'@)
    $wrapperOutput=& pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $fakePrepared 'SmartM365-CmdbEvidence-Publish.ps1') -Tenant test -PreparedRootPath $prepared -ExpectedManifestSHA256 $hash -ValidateOnly 2>&1
    Check ($LASTEXITCODE -eq 0) ('Offline wrapper execution failed: '+($wrapperOutput -join ' '))
    Check (($wrapperOutput -join ' ') -match 'SYNTHETIC/DATA/DATA-POWERBI-CMDB' -and
           ($wrapperOutput -join ' ') -match 'OFFLINE_COMPLETION Success') 'Offline wrapper target/completion missing.'
    Check (-not ($wrapperOutput -join ' ' -match 'External (upload|download)|Authentication reached')) 'Offline wrapper reached an external callback.'
    Write-Host "PASS: $checks offline CMDB transfer checks. No tenant calls, collectors or live shared writes."
} finally {
    $resolved=[IO.Path]::GetFullPath($root)
    if((Split-Path $resolved -Parent).TrimEnd('\','/') -cne ([IO.Path]::GetTempPath()).TrimEnd('\','/') -or
       (Split-Path $resolved -Leaf) -notmatch '^SmartM365-CmdbTransferTest-[a-f0-9]{32}$'){throw 'Unsafe fixture cleanup target.'}
    if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB0shL9zIk77x77
# HaTxiRnEkNXTRM2sFxDn2SXP6REFQ6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOuir1oks62iJxm0q6FBCFQrqhC1kbqqx/e8R8dxYKV2MA0GCSqG
# SIb3DQEBAQUABIIBgDjAexTdHvdOynlPa8qFNoqJ3zTCX7vtB4j8xhQnvz8DW0DQ
# 4Pn+uI5QZ+iIvwsoSo1YLhYPGTNbDojD1iAxb500gTT006bIJP1qaFGfqUqlvDAH
# 2eiRsMcMgI9dXH1QUOe0TudJX76oR5bbPyTdXTN767eyRHUMX4gvqxbEbVh5EqhY
# TEmqwRYEmqWMv5kYfAzKtMcFzpMQ4qbPdlF0fQGGMbt4TTqpv8OTkH38oiW4lkGr
# OHCxyOKVh3XjQ7p7pgjPUv95e/18MaAFf2+sq5TscpHUhghObeHPs1eV7Dmb3X9q
# yVr8/203xpEFTaFNABEcu/hUkfbZvdYarDmTpZr8gVbNLvuVb0BqR1mvYHKcM8JL
# QbseRBIkGPqe8u1BvzorVFJqONzGNBAfUe9R8K3PPKwj/oyNb96lv4dQCtKbjUi3
# c2vgCEEpPALcpGsbUiwtqgx6hkbMRtLKwIiL7OqvYEP5gOvh8FGDTvnBNU2kV0rJ
# tMD/JME9Xm/LIxTBAaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDYyMjQz
# MjJaMC8GCSqGSIb3DQEJBDEiBCDndINWKmcQowqKCT9268CA7c5BVYGhMzALA9NV
# 1fZHVzANBgkqhkiG9w0BAQEFAASCAgAkEgKOQTdKK066G+IEEB/Y+Huj8RBEWXxp
# u68A2dyqFJOyAcoR1XEeAUwJwUwzzE/Kq4CQa2UtBGVOR0lKcmYEiUsJ4IL5pvbA
# YrzQjtpHxakn1lhJONOKGlsvtA1guAtQ+tES5nQrNj3OgHHsfFjw0rW9AGcdldw+
# olxSu1sT1aFNJOW3IcZxmAN+YTaL2ZvqC5+MchGuQX/5wZ4a4aAeN6KKKUCy4V3P
# IXsyHqd1Ry9XPjUcdsSXoXtkyXVWQzcXdfPg0PgZl+eFyaCc7/btfYL8NBKpSdoa
# 1HDqM/5UxfnjIwMEF+QisNMDD9RKmuLjgbTxybTkmR6wu2ksSt3tO5kaPXR0+cnq
# 8Bm/t4uDYY4pHd73ve+otbueSnacFWsGC5+cwW1cGW40ld8GLqolFs5fZdnuQWlw
# rlFrVW9NaaLV856AjtP+ap3qxk5B3DN8wZjBW59wTUB2doUQCZ6gB/mnbTLl+iZd
# 9FkEKEwdRSrWsUXIA26J6QvCMKHrB3uN13fGUbHlIgtIr832uKqAY/lXZW9wXfg0
# 9tJrCIaVAKE134zVseuKI6hprGM1u/NZEs05c//yRztCSp5GAt3HvSMS+CTSOW9d
# djaZhIoHT/j51iWq8CWR2og2agbAs0NUbKYYDN9KYRyymCYe0EDPhl3FMtAoSvaC
# qa7b9CNq1A==
# SIG # End signature block
