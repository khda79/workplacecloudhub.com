#Requires -Version 7.0
<#
.SYNOPSIS
Publish an existing, pinned current CMDB snapshot to its separate SharePoint folder.
.VERSION
0.1.2
.NOTES
Manual only. ValidateOnly is offline. No collection, CSV generation, history,
Intelligence transfer, notification, Power BI refresh or report switch.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Shared Core runtime toggles are saved and restored; transport is explicitly enabled only in its callbacks.')]
param([string]$Tenant='test',[string]$PreparedRootPath,
      [ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedManifestSHA256,[switch]$ValidateOnly)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$ScriptVersion='0.1.2'
$failure=$null; $initialized=$false; $transcript=$false; $core=$null; $guardInstalled=$false
$previousGuard=$false; $saved=@{}; $warning=$false
try {
    if(-not $ValidateOnly -and -not $ExpectedManifestSHA256){throw 'Publication requires ExpectedManifestSHA256 from the reviewed snapshot.'}
    $smartRoot=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $smartRoot 'Config/SmartM365-TenantContext.ps1')
    $effective=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.87' -ErrorAction Stop
    $core=Get-Module SmartM365.Core
    $previousGuard=& $core {
        $previous=Get-Variable -Name SmartM365TeamsNotificationInProgress -Scope Script -ErrorAction SilentlyContinue
        if($null -ne $previous){$previous.Value}else{$false}
    }
    & $core {$script:SmartM365TeamsNotificationInProgress=$true}
    $guardInstalled=$true
    $config=Read-SmartM365JsonConfig -Path (Join-Path $PSScriptRoot 'SmartM365-CmdbEvidence-Publish.local.json.txt') -Required
    function Resolve-PublicationConfig([string]$Name){
        $value=$config[$Name]
        if($null -eq $value -or $value -in @('','__USE_GLOBAL__','USE_GLOBAL')){$value=$effective.$Name}
        return & $core {param($text) Resolve-SmartM365ConfigValue -Value $text} ([string]$value)
    }
    $latest=[IO.Path]::GetFullPath((Resolve-PublicationConfig 'LatestCsvFolderPath')).TrimEnd('\','/')
    if((Split-Path $latest -Leaf) -cne 'DATA-LAST'){throw 'LatestCsvFolderPath must resolve to authoritative SmartInventory DATA-LAST.'}
    $expectedRoot=Join-Path (Split-Path $latest -Parent) 'DATA-POWERBI-CMDB'
    $root=if($PreparedRootPath){[IO.Path]::GetFullPath($PreparedRootPath).TrimEnd('\','/')}else{$expectedRoot}
    if($root -ine $expectedRoot){throw 'PreparedRootPath must be the DATA-POWERBI-CMDB sibling of configured DATA-LAST.'}
    Import-Module (Join-Path $PSScriptRoot 'SmartM365-CmdbSharePointTransfer.psm1') -Force
    $dataCloudRoot=ConvertTo-SmartM365SharePointDataRootPath -TargetFolderPath (Resolve-PublicationConfig 'SharePointTargetFolderPath')
    $cloudRoot=Resolve-SmartM365CmdbSharePointFolder -NormalizedDataRoot $dataCloudRoot
    $cloud=@{Enabled=$true;TenantId=[string]$effective.TenantId}
    foreach($pair in @(@('SiteHostname','SharePointSiteHostname'),@('SitePath','SharePointSitePath'),
                       @('LibraryDisplayName','SharePointLibraryDisplayName'),@('AppId','AppId'),@('Thumbprint','Thumb'))){
        $cloud[$pair[0]]=Resolve-PublicationConfig $pair[1]
        if([string]::IsNullOrWhiteSpace($cloud[$pair[0]]) -or $cloud[$pair[0]] -match '__USE_GLOBAL__|\{\{'){
            throw "Unresolved SharePoint publication configuration: $($pair[1])."
        }
    }
    $identity=@{}
    foreach($field in 'TenantKey','OrganizationKey','EnvironmentKey','TenantId'){$identity[$field]=[string]$effective.$field}
    $pythonName=if($config['PythonCommand']){[string]$config['PythonCommand']}else{'python'}
    $python=(Get-Command $pythonName -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $pythonVersion=& $python -B -c 'import sys; print(".".join(map(str,sys.version_info[:3])))'
    if($LASTEXITCODE -ne 0 -or [version]$pythonVersion -lt [version]'3.10'){throw 'Python 3.10+ is required.'}
    # Keep automatic log transfers and notification callbacks disabled in both
    # modes. Only the 47 explicit artifact callbacks can write to SharePoint.
    foreach($name in 'EnableSharePointUpload','EnableTeamsNotifications','SmtpServer','From','To','ErrorMailTo'){
        $variable=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
        $saved[$name]=@{Exists=($null -ne $variable);Value=$(if($variable){$variable.Value}else{$null})}
    }
    $global:EnableSharePointUpload=$false; $global:EnableTeamsNotifications=$false
    $global:SmtpServer=''; $global:From=''; $global:To=''; $global:ErrorMailTo=''
    $logRoot=& $core {param($value) Resolve-SmartM365ConfigValue -Value $value} ([string]$effective.LogAllRootPath)
    if([string]::IsNullOrWhiteSpace($logRoot) -or -not [IO.Path]::IsPathFullyQualified($logRoot) -or $logRoot -match '\{\{|__USE_GLOBAL__'){
        throw 'An absolute resolved LogAllRootPath is required.'
    }
    InitializeScriptEnvironment -OutputPathInit (Join-Path $logRoot 'Publication/CMDB') -LogFileName 'SmartM365-CmdbEvidence-Publish' -CallerScriptPath $PSCommandPath | Out-Null
    $initialized=$true
    Start-SmartM365RunTranscript -Path $global:logTranscriptFile -Append | Out-Null
    $transcript=$true
    WriteLog -Message "CMDB publication $ScriptVersion. Prepared='$root'; SharePoint='$cloudRoot'; ValidateOnly=$ValidateOnly. No collection or report switch." -Level INFO
    $parameters=@{PreparedRoot=$root;Identity=$identity;PythonPath=$python;ExpectedManifestSHA256=$ExpectedManifestSHA256}
    if($ValidateOnly){
        $result=Get-SmartM365CmdbTransferPlan @parameters
        $result | Select-Object Status,CsvFiles,ManifestSHA256,ContractVersion,GeneratedAtUtc,EarliestSourceExpiryUtc,FreshnessWarnings,
            @{Name='SharePointTarget';Expression={$cloudRoot}} | ConvertTo-Json -Depth 8
    } else {
        $transitionRoot=Join-Path $logRoot 'Publication/CMDB/SharePointTransition'
        $relocated=Move-SmartM365CmdbTransitionState -PreparedRoot $root -StateRoot $transitionRoot -Identity $identity
        if($relocated){WriteLog -Message "Relocated $relocated inactive SharePoint transition artifacts outside the CMDB cohort." -Level INFO}
        $parameters['UploadFile']={
            param($localPath,$relativePath)
            if((Get-SmartM365SharePointRelativeFilePath -LocalFilePath $localPath) -cne $relativePath){
                throw 'CMDB transfer must not be nested under raw or log folders.'
            }
            SmartM365.Core\Invoke-SmartM365SharePointCsvUpload -LocalFilePath $localPath -TargetFolderPath $cloudRoot -EnsureParentFolders -JsonTransitionStateFolderPath $transitionRoot @cloud
        }.GetNewClosure()
        $parameters['DownloadFile']={
            param($destination,$relativePath)
            SmartM365.Core\Invoke-SmartM365SharePointFileDownload -LocalFilePath $destination -SharePointRelativePath $relativePath -TargetFolderPath $cloudRoot -Force @cloud
        }.GetNewClosure()
        $parameters['Progress']={param($name,$index,$total) WriteLog -Message "CMDB transfer $index/${total}: $name" -Level INFO}
        $result=Send-SmartM365CmdbPreparedSnapshot @parameters
        $result | ConvertTo-Json -Depth 8
    }
    foreach($item in @($result.FreshnessWarnings)){$warning=$true;WriteLog -Message $item.Message -Level WARNING}
    WriteLog -Message $(if($ValidateOnly){'CMDB transfer validation completed offline. Nothing was uploaded.'}else{'CMDB snapshot uploaded and read back. Power BI qualification remains separate.'}) -Level SUCCESS

    $smartM365RunGuard.ReachedTerminal=$true
} catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_
$failure=$_;throw} finally {
    $smartM365RunGuard.InCleanup=$true

    try {
        if($initialized){Complete-SmartM365ExecutionContext -Status $(if($failure){'Failed'}elseif($warning){'CompletedWithWarnings'}else{'Success'}) -ErrorRecord $failure -FailureStage 'CmdbPublication'}
    } finally {
        try {if($transcript){Stop-SmartM365RunTranscript | Out-Null}} finally {
            if($guardInstalled){& $core {param($previous) $script:SmartM365TeamsNotificationInProgress=$previous} $previousGuard}
            foreach($name in $saved.Keys){
                if($saved[$name].Exists){Set-Variable -Name $name -Scope Global -Value $saved[$name].Value}
                else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}
            }
        }
    }
}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBydITv9AbGVLjt
# Gd9SPjVBFE+6A/2kuORtWlWjqr1fFaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIG4bgd/98an/pDZUa4TFxPdbeAOuSOTJPwMF4natRa/yMA0GCSqG
# SIb3DQEBAQUABIIBgIvH5GtC8DrXK18hNFiXwevrZQucArfmKtygPOsbAKD2a6d9
# fJnc8qwwOo5QjOnleCsGjuA9VcbDXf6vWSpwtqt38D2skrEC/eJMLuw30xaJFQJ8
# exw83/AuTgJpRnmbYm/TdynIstI8n2X5TCwynmQJ8ws5uToN8OCrmIV0HxWTETeo
# 7a5venduKkNx8GtJTMX3AQ2tiZBSXbop0uSPxWZHVyJeG2lXQ+pUu6D5dSCkO6Lw
# djGE/O6CoSx4XWlZT6r8yB7uXB5ec6hL6cKbXGsyP2u+9rAIXR3xcHd30gCnfibB
# XAzcgA1Bqa0f1ASrCKfuUSuwlWFa9B58xvNz4LPlg2XbNk61623Hym+mozWQbevZ
# xYxX0KtnQlwPeVbqVEteAeZMqeVlH1V+XyfM0yazkbibL88T+5oXYigcLjvfFr/0
# 8owe446TqGqR3VPffz6xIC/jwoBhXkv1inKjpf3HFtxtHpMKrXexBpIxcFh/jQ4f
# r6D/OHZ+77xfFmtbiKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTQxNTMw
# NDFaMC8GCSqGSIb3DQEJBDEiBCAF1dEHOWMlAuF9WSA3rWCMBOu+Eqcwp2fpohSd
# M0m7DDANBgkqhkiG9w0BAQEFAASCAgCttjvxANzbwqbEOU8MaDFnu7EyL23pODK1
# CPo6VfRUIuUTSfU7vxasbirBMMkm4gyw77R8ZQ6ZbKgIr7l5n+5D+xX8melkffFz
# vESA9N2kFI1e6/gGK7Pt76+ffKh4oQrHI1+ogWIPe1xvkieHha0N+E6DbLbG8+Fn
# 2jcCBfaj266A/ZdhyGn8AWoF6KJAJkXGVupv2w630r7EnSMl/mrIgJA792hntAhS
# w7FNdUVC4EOr7teQ5K88oz+m51o16IimtXf/ZEwOFxAMVuL7Y2vYCIHeSYXwNFKN
# k85KcDM4LGJp6mQ+LtfRnDSaVLwtq2Kb+TWNo2iVJRmJ8lGSPzPotPJwPBlUwBOX
# 5aigWr+mH5+CZYy507HUkQ3Nueu3rS3WfMPHtBLuOi8yYdKfji7n4fWyRS3RqJd5
# VXJurXQWbgim67K/eQZuHj3BiSyt3EBHff8L2Z/cYadHA6N8viwDOWrk8OUuQTCi
# ZMWF/kTx85QPCgQQbWDkIIh1PbiHjgrEVde7EpT7TpHGbQwB4Eftxe5NfQ26x85+
# 23+/vaDfTk9AJowxnjh8dpCjmZxmEiwq2V7G50RdRGcK+Yq+mBC7Cb3TqgTf59Ad
# ZP3u3i70tKhx774eNYqP67SzVDSXi9ZLpKzv5znFcl179TK+Yje4MdfJXbakdcmP
# nre/jGuvfQ==
# SIG # End signature block
