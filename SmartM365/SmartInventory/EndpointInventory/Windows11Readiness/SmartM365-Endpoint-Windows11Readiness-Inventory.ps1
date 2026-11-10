<#
.SYNOPSIS
    Collects Windows 11 hardware-readiness evidence from an explicit list of PCs.
.DESCRIPTION
    Version 1.0.1. Separate from Intune eligibility and SWI. The target list defines
    this run's scope, not the whole tenant. PowerShell remoting and existing rights
    are required for remote targets; localhost uses the current process identity.
    A partial run produces timestamped evidence and errors but never promotes latest
    or weekly history. The Microsoft result and diagnostic log are retained per row.
    Windows PowerShell 5.1 is used on endpoints; this inventory entry requires PS 7.
.PARAMETER ComputerListPath
    Text file containing one DNS computer name per line; # starts a comment line.
.PARAMETER Tenant
    SmartM365 tenant profile; defaults to test.
.PARAMETER TimeoutSeconds
    Remote timeout per target. Local execution has no remote timeout.
.PARAMETER Credential
    Optional existing remoting credential; never saved. Cannot be used for localhost.
#>
[CmdletBinding()]
param(
    [string]$Tenant='test',
    [Parameter(Mandatory)][string]$ComputerListPath,
    [ValidateRange(10,3600)][int]$TimeoutSeconds=120,
    [pscredential]$Credential
)
$ErrorActionPreference='Stop'
$ScriptVersion='1.0.1'
$root=$PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Config/SmartM365-TenantContext.ps1') -PathType Leaf)) {
    $parent=Split-Path $root -Parent
    if ($parent -eq $root) { $root=$null; break }; $root=$parent
}
if (-not $root) { throw 'SmartM365 root not found.' }
. (Join-Path $root 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
$runLock=$null
try {
. (Join-Path $root 'Config/SmartM365-TenantContext.ps1')
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null
Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -Force -ErrorAction Stop
$script:CompletionStatus='Failed'; $script:CompletionError=$null
try {
    if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 or later is required.' }
    $ScriptLocalConfig=Read-SmartM365ScriptConfiguration -ScriptPath $PSCommandPath
    function Get-ReadinessConfig([string]$Name,$Default) {
        Get-SmartM365ConfigurationValue -Name $Name -LocalConfig $ScriptLocalConfig -EffectiveConfig $global:SmartM365GlobalConfig -DefaultValue $Default -EffectiveFallback Blank
    }
    $outputRoot=[string](Get-ReadinessConfig 'ScriptCsvLogFolderPath' '')
    $latestRoot=[string](Get-ReadinessConfig 'LatestCsvFolderPath' '')
    if (-not $outputRoot -or -not $latestRoot -or -not $global:SmartM365TenantKey) { throw 'Tenant output paths or TenantKey are missing.' }
    $global:RetentionMaxCSV=[int](Get-ReadinessConfig 'RetentionMaxCSV' 30)
    $global:RetentionMaxLogs=[int](Get-ReadinessConfig 'RetentionMaxLogs' 30)
    $global:EnableSharePointUpload=[bool](Get-ReadinessConfig 'EnableSharePointUpload' $false)
    foreach ($name in @('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath','LogAllRootPath')) {
        Set-Variable -Scope Global -Name $name -Value (Get-ReadinessConfig $name '')
    }
    $outputRoot=InitializeScriptEnvironment -OutputPathInit $outputRoot -LogFileName 'SmartM365-Endpoint-Windows11Readiness-Inventory'
    try { $runLock=[IO.File]::Open((Join-Path $outputRoot '.windows11-readiness.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None) } catch { throw 'Another readiness run holds the output lock, or the lock cannot be acquired.' }
    . (Join-Path $root 'Modules/SmartM365.Core/SmartM365-SharePointUploadPolicy.ps1')
    Initialize-SmartM365CollectorUploadContext -LocalConfig $ScriptLocalConfig -GlobalConfig $global:SmartM365GlobalConfig -DataAllRoot $outputRoot -DataLastRoot $latestRoot -LogAllRoot $global:LogPath -ExternalActionsAllowed $true
    $collectorPath=Join-Path $PSScriptRoot 'SmartM365-Collect-Windows11Readiness.ps1'
    $microsoftPath=Join-Path $PSScriptRoot 'SmartM365-Windows11-HardwareReadiness-Microsoft.ps1'
    $expectedHash='93A1270CF92A9259E6CC98C58E91C51DE2B3D204D3D10E946804772915C1103F'
    if ((Get-FileHash -LiteralPath $microsoftPath -Algorithm SHA256).Hash -ne $expectedHash) { throw 'Readiness script hash mismatch. Review the approved script before changing its pinned hash.' }
    $targets=@(Get-Content -LiteralPath $ComputerListPath -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -Unique)
    if ($targets.Count -eq 0) { throw 'The target list is empty.' }
    foreach ($target in $targets) { if ($target -notmatch '^(?=.{1,253}$)[a-zA-Z0-9](?:[a-zA-Z0-9.-]*[a-zA-Z0-9])?$') { throw 'The target list contains an invalid computer name.' } }
    $scopeText=($targets | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join "`n"
    $sha=[Security.Cryptography.SHA256]::Create()
    try { $scopeHash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($scopeText)))).Replace('-','') } finally { $sha.Dispose() }
    $collector=Get-Content -LiteralPath $collectorPath -Raw -Encoding UTF8
    $ms=[Convert]::ToBase64String([IO.File]::ReadAllBytes($microsoftPath))
    $remote={ param($collectorText,$msText) & ([scriptblock]::Create($collectorText)) -ReadinessScriptBase64 $msText }
    $rows=New-Object 'System.Collections.Generic.List[object]'; $errors=New-Object 'System.Collections.Generic.List[object]'
    $runId=[guid]::NewGuid().ToString('N'); $stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
    foreach ($target in $targets) {
        WriteLog -Message ("Checking target {0}" -f $target) -Level INFO
        $job=$null
        try {
            if ($target -ieq 'localhost' -or $target -ieq $env:COMPUTERNAME) {
                if ($Credential) { throw 'Credentials cannot be applied to local collection.' }
                $result=@(& ([scriptblock]::Create($collector)) -ReadinessScriptBase64 $ms)
            } else {
                $invoke=@{ComputerName=$target;ScriptBlock=$remote;ArgumentList=@($collector,$ms);AsJob=$true;ErrorAction='Stop'}
                if ($Credential) { $invoke.Credential=$Credential }
                $job=Invoke-Command @invoke
                if (-not (Wait-Job -Job $job -Timeout $TimeoutSeconds)) { Stop-Job -Job $job -ErrorAction SilentlyContinue; throw "Timed out after $TimeoutSeconds seconds." }
                $result=@(Receive-Job -Job $job -ErrorAction Stop)
                if ($job.State -ne 'Completed') { throw 'Remote collection did not complete.' }
            }
            if ($result.Count -ne 1 -or -not $result[0].PSObject.Properties['pc.GLOBALResult']) { throw 'Collection did not return one valid result.' }
            if ($result[0].'pc.GLOBALResult' -eq 'FAILED TO RUN') { throw 'Microsoft readiness script reported FAILED TO RUN.' }
            $row=[ordered]@{ SourceFile=$target; RunId=$runId; ScopeHash=$scopeHash; ScopeTargetCount=$targets.Count; SourceScriptSha256=$expectedHash }
            foreach ($prop in $result[0].PSObject.Properties) {
                if ($prop.Name -like 'pc.*' -or $prop.Name -in @('ComputerDomain','MicrosoftReturnCode','MicrosoftReason','MicrosoftLogging')) { $row[$prop.Name]=$prop.Value }
            }
            $rows.Add([pscustomobject]$row)
        } catch {
            $errors.Add([pscustomobject][ordered]@{RunId=$runId;ScopeHash=$scopeHash;Target=$target;TimestampUtc=[DateTime]::UtcNow.ToString('o');Error=$_.Exception.Message})
            WriteLog -Message ("Target {0} failed: {1}" -f $target,$_.Exception.Message) -Level WARNING
        } finally { if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } }
    }
    $base='Endpoint_Windows11_HardwareReadiness'
    $columns=@('SourceFile','RunId','ScopeHash','ScopeTargetCount','SourceScriptSha256','pc.Date','pc.name','pc.Manufacturer','pc.Model','pc.WinVersion','pc.WINEDT','pc.Storage','pc.Storage_Check','pc.Memory','pc.Memory_Check','pc.TPM','pc.TPM_Check','pc.Processor','pc.Processor_Check','pc.SecureBoot','pc.SecureBoot_Check','pc.CPU','pc.CPU_Check','pc.GLOBALResult','ComputerDomain','MicrosoftReturnCode','MicrosoftReason','MicrosoftLogging')
    $errorCols=@('RunId','ScopeHash','Target','TimestampUtc','Error')
    $resultPath=Join-Path $outputRoot ("{0}_{1}.csv" -f $base,$stamp)
    $errorsPath=Join-Path $outputRoot ("{0}_Errors_{1}.csv" -f $base,$stamp)
    if ((Test-Path -LiteralPath $resultPath) -or (Test-Path -LiteralPath $errorsPath)) { throw 'Run output already exists; retry later.' }
    Export-SmartM365Csv -Data $errors.ToArray() -TimestampedPath $errorsPath -Columns $errorCols -NoWeeklyHistory | Out-Null
    if ($errors.Count -eq 0) {
        Export-SmartM365Csv -Data $rows.ToArray() -TimestampedPath $resultPath -LatestPath (Join-Path $latestRoot "$base.csv") -Columns $columns | Out-Null
        $script:CompletionStatus=if (@($rows | Where-Object { $_.'pc.GLOBALResult' -in @('UNDETERMINED','FAILED TO RUN') }).Count) { 'CompletedWithWarnings' } else { 'Success' }
    } else {
        Export-SmartM365Csv -Data $rows.ToArray() -TimestampedPath $resultPath -Columns $columns -NoWeeklyHistory | Out-Null
        $script:CompletionStatus='CompletedWithWarnings'
    }
    WriteLog -Message ("Readiness collection: targets={0}; results={1}; errors={2}; scope={3}; complete={4}" -f $targets.Count,$rows.Count,$errors.Count,$scopeHash,($errors.Count -eq 0)) -Level INFO
    if ($errors.Count -gt 0) { throw "Incomplete collection: $($errors.Count) of $($targets.Count) targets failed. Latest and weekly history were not promoted." }
} catch {
    $script:CompletionStatus='Failed'; $script:CompletionError=$_
    try { WriteLog -Message ("Windows 11 readiness inventory failed: {0}" -f $_.Exception.Message) -Level ERROR } catch { Write-Warning $_.Exception.Message }
    try { Send-SmartM365Mail -Subject 'Windows 11 readiness inventory - ERROR' -BodyHtml ([System.Net.WebUtility]::HtmlEncode($_.Exception.Message)) -MailPurpose Error } catch { Write-Warning 'Error mail could not be sent.' }
    throw
} finally {
    Complete-SmartM365ExecutionContext -Status $script:CompletionStatus -ErrorRecord $script:CompletionError
}

    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    try { Complete-SmartM365InventoryRun -State $smartM365RunGuard } finally { if ($runLock) { $runLock.Dispose() } }
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if ($smartM365RunGuard.ExitCode -ne 0) { exit $smartM365RunGuard.ExitCode }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC6vb2CeSAgEC5e
# STnONISVMFN9QFMDOYsBgU7Vn8iDz6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDvPhOHfN0W19mV8dyvWoDxZdVl8S0c7XwE4PhWx4cR5MA0GCSqG
# SIb3DQEBAQUABIIBgAq3bqRI0JARxUc5r1CWbhEAqLbHQAn/r16DzLXaydhs5hMk
# WbIzRzAU3vO4eDgcWa4f3IBN/eH1NhQNqTQcOt7veu9ccQnuEl9glBUc5s1zeOVx
# vnywCYlDny57J/3z6ywTMxFdFMZkDmaNlCxcTydwAwp4NSE+6hts4LPbewx3YZo0
# Yz6gMJ7wADwSo0P2TW7w0NuLjWjiGdojtE/To6dGTXSPSTS38U+eaH+q72a48wki
# OpDnx+Zl/DWS4tjeDBMZXQNukEokdNHhfOc3OOeyxYS6PqTBhwYH2/FadpkXTZA2
# xrP3GMh8k894R2RogoH2O5iuliekzcs3x7fIdKGK1hJA2MAMKtf7W7pKgJ5oIUGF
# o646ia2MaQ4Vr562s2bannWSLM7qXwDELuh761Xmu7bDMBFCK/qJlgK+BOdHgEd2
# rX5qs7PAzBi88pAp8TV+abqVosMiyjBhUEOZ5FrCCOM93PDjKtztYvhCgwoDOM0p
# DGGdlIg7BrKK0AJa4qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDcxMDQw
# NDdaMC8GCSqGSIb3DQEJBDEiBCDrGEJ87U5lb1rPXPWwplm/n5A80upZAhqNV6Bt
# UPgLvjANBgkqhkiG9w0BAQEFAASCAgAKKZPmDfezgEZphG9EXD5ikz/JNBPggMs5
# tBLuKdwZIDQZcQKYy9KOuAHkivsHRNDK/ryG3D/Wvvmf84oYW32UGcg2OvPI76Lh
# ZlCDVsP9bLjoOv6lNEV3yq2UCKMoauTjHpktKf0EwYHJJbaieujVlgJu4auh/9B8
# ntm1fN5jM97SXtH6EORlauOJYtfGKc0QhMFVqFojDWOTwzBiiJLbvJ6a0kr4I2u1
# RCsFrIq6lMneZAs35py4cOAUgjnzK0iqSg7Z8Du2ikLcTR1nxEnnI7DG3EX4+lut
# 2pjJFltADvGc9SjUv16CTjn6CI30OuWXSexlDKhhAL5PQM2byRnLUzkpbzEuTEmD
# cApCHWBbTVPF0/K8I4gXLxuVFaQegPK/G5MpUyqUFaItZiwU5s0veUfeuZ4S+aI1
# Um7sxnJyr+f4OPTCG0FMqjTt8yfNOds9sFRRmnz21aSTQebcOoUwW05lMpBwucGb
# 3n/2kfyecuY3I3LVR94yWbYuQSMoYSIdtNE4dika8hEQ01cuAS4ehsir24AFtjHv
# m9CI+D8BUE5ZNn/txh2LTVHf30f+pVzUESaexnl3G42CyG9+eNxbzWGzkb67yWEZ
# lnpTbcsQBCm3k/lDPMUP6p+bpafOK98oiRk2PHJw34Hw2Wo0YB+8Xc6ashylcDeQ
# HVs+x95i7A==
# SIG # End signature block
