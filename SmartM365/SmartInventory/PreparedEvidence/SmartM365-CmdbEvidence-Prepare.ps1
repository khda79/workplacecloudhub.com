#Requires -Version 7.0
<#
.SYNOPSIS
Prepare the current-only CMDB reporting tables from proven SmartInventory CSVs.
.VERSION
0.3.13
.NOTES
Local preparation by default. -Publish explicitly transfers the newly validated
snapshot through the existing SharePoint publisher. -ValidateOnly never uploads.
Normal run logs follow EnableSharePointUpload independently of the prepared batch.
No collector, history or Power BI refresh is invoked.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Existing SmartM365.Core operational contract; offline toggles are saved and restored.')]
param([string]$Tenant='test',[string]$SourceRootPath,[switch]$ValidateOnly,[switch]$Publish)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$script:Version='0.3.13'
$failure=$null; $runtimeInitialized=$false; $transcriptStarted=$false
$core=$null; $previousTeamsGuard=$false; $teamsGuardInstalled=$false
$savedOfflineGlobals=@{}; $preparationWarning=$false; $logUploadEnabled=$false
function Get-SmartM365CmdbLogUploadEnabled {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.PSModuleInfo]$CoreModule,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Configuration,
          [switch]$ValidateOnly)
    if($ValidateOnly){return $false}
    $value=& $CoreModule {
        param($config)
        Get-ModuleLocalConfigValue -Config ([pscustomobject]$config) -Name 'EnableSharePointUpload' -DefaultValue $false
    } $Configuration
    if($value -is [bool]){return $value}
    $parsed=$false
    if([bool]::TryParse([string]$value,[ref]$parsed)){return $parsed}
    throw 'EnableSharePointUpload must resolve to a Boolean for CMDB run logs.'
}
function Invoke-SmartM365CmdbPreparedPublication {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$PreparationResult,
          [Parameter(Mandatory)][string]$PreparedRoot,
          [Parameter(Mandatory)][string]$TenantProfile,
          [Parameter(Mandatory)][scriptblock]$PublisherInvoker)
    if($PreparationResult.Status -notin @('Prepared','PreparedWithCleanupWarning')){
        throw 'Publication requires a successfully prepared current snapshot.'
    }
    $manifest=Join-Path $PreparedRoot 'current.json.txt'
    $item=Get-Item -LiteralPath $manifest -Force -ErrorAction Stop
    if($item.PSIsContainer -or $item.Attributes -band [IO.FileAttributes]::ReparsePoint){
        throw 'Linked or non-file CMDB manifest refused.'
    }
    # Capture this generation, not a manually pasted or previous batch hash.
    # The publisher independently locks and validates all 46 files and freshness.
    $hash=(Get-FileHash -LiteralPath $manifest -Algorithm SHA256).Hash
    $exitCode=& $PublisherInvoker $TenantProfile $PreparedRoot $hash
    if($exitCode -isnot [int] -or $exitCode -ne 0){
        throw 'CMDB SharePoint publication failed. Prepared local output is retained; no upload success is inferred.'
    }
}
function Resolve-SmartM365CmdbPreparationSourcePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.PSModuleInfo]$CoreModule,
          [AllowEmptyString()][string]$SourceRootPath,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Configuration,
          [AllowEmptyString()][string]$LatestCsvFolderPath)
    # Explicit override, then local configuration, then the effective tenant profile.
    # Resolve inherited tokens before GetFullPath can anchor them to the current directory.
    $candidate=$SourceRootPath
    if([string]::IsNullOrWhiteSpace($candidate)){
        $localPath=[string]$Configuration['LatestCsvFolderPath']
        $candidate=if(-not [string]::IsNullOrWhiteSpace($localPath) -and
            $localPath.Trim() -notin @('__USE_GLOBAL__','USE_GLOBAL')){
            $localPath
        }else{$LatestCsvFolderPath}
    }
    $resolved=[string](& $CoreModule {param($value) Resolve-SmartM365ConfigValue -Value $value} $candidate)
    $resolved=$resolved.Trim()
    if([string]::IsNullOrWhiteSpace($resolved) -or $resolved -in @('__USE_GLOBAL__','USE_GLOBAL') -or
       $resolved -match '\{\{|\}\}' -or -not [IO.Path]::IsPathFullyQualified($resolved)){
        throw 'CMDB preparation requires a fully resolved absolute DATA-LAST source path; no source or output directory was created.'
    }
    $source=[IO.Path]::GetFullPath($resolved).TrimEnd([char[]]@('\','/'))
    if((Split-Path $source -Leaf) -ne 'DATA-LAST'){
        throw 'Use the authoritative SmartInventory DATA-LAST, not DATA-POWERBI.'
    }
    return $source
}
function Resolve-SmartM365CmdbPreparationLogPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.PSModuleInfo]$CoreModule,
          [AllowEmptyString()][string]$LogRootPath)
    $resolved=& $CoreModule {param($value) Resolve-SmartM365ConfigValue -Value $value} $LogRootPath
    if([string]::IsNullOrWhiteSpace($resolved) -or $resolved -in @('__USE_GLOBAL__','USE_GLOBAL') -or
       $resolved -match '\{\{|\}\}' -or -not [IO.Path]::IsPathFullyQualified($resolved)){
        throw 'CMDB preparation requires a fully resolved absolute LogAllRootPath; no log directory was created.'
    }
    return Join-Path $resolved 'Preparation/CMDB'
}
try {
    $smartRoot=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $smartRoot 'Config/SmartM365-TenantContext.ps1')
    $effective=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.80' -ErrorAction Stop
    # WriteLog can notify Teams using module-local settings, independent of the
    # global toggle. Suppress its existing callback for this preparation invocation.
    # This is in-memory only: no module/config file is changed.
    $core=Get-Module SmartM365.Core
    $previousTeamsGuard=& $core {
        $previous=Get-Variable -Name SmartM365TeamsNotificationInProgress -Scope Script -ErrorAction SilentlyContinue
        if($null -ne $previous){$previous.Value}else{$false}
    }
    & $core { $script:SmartM365TeamsNotificationInProgress=$true }
    $teamsGuardInstalled=$true
    $config=Read-SmartM365JsonConfig -Path (Join-Path $PSScriptRoot 'SmartM365-CmdbEvidence-Prepare.local.json.txt') -Required
    $logUploadEnabled=Get-SmartM365CmdbLogUploadEnabled -CoreModule $core -Configuration $config -ValidateOnly:$ValidateOnly
    $source=Resolve-SmartM365CmdbPreparationSourcePath -CoreModule $core -SourceRootPath $SourceRootPath -Configuration $config -LatestCsvFolderPath ([string]$effective.LatestCsvFolderPath)
    $output=Join-Path (Split-Path $source -Parent) 'DATA-POWERBI-CMDB'
    $pythonName=if($config['PythonCommand']){[string]$config['PythonCommand']}else{'python'}
    # Keep logs in LOG-ALL; initialization must not create the protected output.
    foreach($name in @('EnableSharePointUpload','EnableTeamsNotifications','SmtpServer','From','To','ErrorMailTo',
            'SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath',
            'AppId','TenantId','Thumb','Thumbprint')){
        $variable=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
        $savedOfflineGlobals[$name]=@{Exists=($null -ne $variable);Value=$(if($null -ne $variable){$variable.Value}else{$null})}
    }
    $global:EnableSharePointUpload=$false
    $global:EnableTeamsNotifications=$false
    $global:SmtpServer=''; $global:From=''; $global:To=''; $global:ErrorMailTo=''
    if($logUploadEnabled){
        foreach($name in @('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath','AppId','Thumb')){
            $value=& $core {
                param($config,$key)
                Get-ModuleLocalConfigValue -Config ([pscustomobject]$config) -Name $key -DefaultValue ''
            } $config $name
            Set-Variable -Name $name -Scope Global -Value ([string]$value)
        }
        $global:TenantId=[string]$effective.TenantId
        $global:Thumbprint=$global:Thumb
    }
    $logBase=Resolve-SmartM365CmdbPreparationLogPath -CoreModule $core -LogRootPath ([string]$effective.LogAllRootPath)
    InitializeScriptEnvironment -OutputPathInit $logBase -LogFileName 'SmartM365-CmdbEvidence-Prepare' -CallerScriptPath $PSCommandPath | Out-Null
    $runtimeInitialized=$true
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    $transcriptStarted=$true
    WriteLog -Message "CMDB preparation $script:Version. Source='$source'; Output='$output'; ValidateOnly=$ValidateOnly; Publish=$Publish. ValidateOnly never uploads." -Level INFO
    WriteLog -Message "Run log SharePoint upload enabled: $logUploadEnabled; prepared batch publication requires -Publish and successful generation." -Level INFO
    if(-not $ValidateOnly){
        Import-Module (Join-Path $PSScriptRoot 'SmartM365-CmdbSharePointTransfer.psm1') -Force
        $identity=@{}
        foreach($field in 'TenantKey','OrganizationKey','EnvironmentKey','TenantId'){$identity[$field]=[string]$effective.$field}
        $transitionRoot=Join-Path (Split-Path (Split-Path $logBase -Parent) -Parent) 'Publication/CMDB/SharePointTransition'
        $relocated=Move-SmartM365CmdbTransitionState -PreparedRoot $output -StateRoot $transitionRoot -Identity $identity
        if($relocated){WriteLog -Message "Relocated $relocated inactive SharePoint transition artifacts outside the CMDB cohort." -Level INFO}
    }
    $python=Get-Command $pythonName -CommandType Application -ErrorAction Stop | Select-Object -First 1
    $pythonVersion=& $python.Source -c 'import sys; print(".".join(map(str,sys.version_info[:3])))'
    if($LASTEXITCODE -ne 0 -or [version]$pythonVersion -lt [version]'3.10'){throw 'Python 3.10+ is required.'}
    $arguments=@((Join-Path $PSScriptRoot 'cmdb_prepare.py'),'--source',$source,'--output',$output,
        '--tenant-key',[string]$effective.TenantKey,'--organization-key',[string]$effective.OrganizationKey,
        '--environment-key',[string]$effective.EnvironmentKey,'--tenant-id',[string]$effective.TenantId)
    if($ValidateOnly){$arguments+='--validate-only'}
    $resultText=& $python.Source @arguments 2>&1
    if($LASTEXITCODE -ne 0){
        $diagnostic=@($resultText | ForEach-Object {[string]$_} | Where-Object {-not [string]::IsNullOrWhiteSpace($_)} | Select-Object -Last 1) -join ''
        throw "CMDB preparation rejected its source or output contract. Last validated output is preserved. $diagnostic"
    }
    $result=$resultText | ConvertFrom-Json -ErrorAction Stop
    WriteLog -Message ([string]$resultText) -Level INFO
    foreach($warning in @($result.FreshnessWarnings) + @($result.CoverageWarnings) + @($result.ApplicationWarnings) + @($result.LicenseWarnings)){
        $preparationWarning=$true
        WriteLog -Message ([string]$warning.Message) -Level WARNING
    }
    if($result.Status -eq 'PreparedWithCleanupWarning'){
        $preparationWarning=$true
        WriteLog -Message "Validated output published locally, but transient rollback cleanup needs review: '$($result.CleanupRequired)'." -Level WARNING
    }
    if($Publish -and -not $ValidateOnly){
        $publisherScript=Join-Path $PSScriptRoot 'SmartM365-CmdbEvidence-Publish.ps1'
        $publisherInvoker={
            param($tenantProfile,$preparedRoot,$manifestHash)
            # A separate PS7 process prevents preparation's offline globals from
            # leaking into the publisher's independently resolved tenant config.
            & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -ExecutionPolicy Bypass -File $publisherScript -Tenant $tenantProfile -PreparedRootPath $preparedRoot -ExpectedManifestSHA256 $manifestHash | Out-Host
            return [int]$LASTEXITCODE
        }.GetNewClosure()
        Invoke-SmartM365CmdbPreparedPublication -PreparationResult $result -PreparedRoot $output -TenantProfile $Tenant -PublisherInvoker $publisherInvoker
        WriteLog -Message 'CMDB preparation and verified SharePoint publication completed. No collection, history or Power BI refresh.' -Level SUCCESS
    } else {
        WriteLog -Message 'CMDB local preparation completed. No collection, history, report switch or prepared-batch publication.' -Level SUCCESS
    }
} catch { $failure=$_; throw } finally {
    try {
        if($runtimeInitialized){
            # Enable only completion-time traces; the prepared batch has its own publisher.
            $global:EnableSharePointUpload=$logUploadEnabled -and -not $ValidateOnly
            Complete-SmartM365ExecutionContext -Status $(if($failure){'Failed'}elseif($preparationWarning){'CompletedWithWarnings'}else{'Success'}) -ErrorRecord $failure -FailureStage 'CmdbPreparation' -CloseTranscriptBeforeUpload
            $transcriptStarted=$false
        }
    } finally {
        try {if($transcriptStarted){Stop-Transcript | Out-Null}}
        finally {
            if($teamsGuardInstalled){& $core {param($previous) $script:SmartM365TeamsNotificationInProgress=$previous} $previousTeamsGuard}
            foreach($name in $savedOfflineGlobals.Keys){
                $previous=$savedOfflineGlobals[$name]
                if($previous.Exists){Set-Variable -Name $name -Scope Global -Value $previous.Value}
                else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}
            }
        }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDy/O9MUx3jxCjh
# gEvOGnD4rd+7ibVRs36/t17CUHEUQqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKPc07Kjg5PSENm8PUFbOlA8RXYLnC3yFcmuYBPk7jgnMA0GCSqG
# SIb3DQEBAQUABIIBgCkbMJH3iuzNOgV4vI2fKgZ9IEm7mQaLqc6iD7jm0lg8NZ8A
# nSjJOceZx39/gsT4z9ruh2ZSfW71c9i7354lGZcbUo7t7cQe4kllEUcFcJ56f4/H
# /xhLhM2h5Vs5xeNWm3qytBVviI3hpNKw0CSzgFOQdtqRVL3j2fYZhMcAx9THPDg8
# R+VBpVL+ee7kul/ZScvWWmzErYxFqLMdowPdLUqBOESyUHKXutiCWzGoewMx0pml
# VQ4YdDXXpIdknMP98/BjFMJ1d0oWmAwhHpPeem5uuNnrVOFpdM1W92sUelTg1vtb
# 3plL2JOmYi3OnzS6jI9ZmNHf0Rly86ynNloB1g70XNuu8I7R7uczCLT5bpSBsgam
# E9cBB+XrdKRajTcFbdSrm6vJ8CZmdOZF7ShuFSS81UTmACLfJgieFmTeu1E8gT6L
# Ph2jfN32xlKDygqV79jllsZEhap/WYuTNKg2cZ+8ZeTarqNwQIKTB7f4lOGi7zap
# JwgnUF0N4K6sX7F9xqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgwOTI4
# NThaMC8GCSqGSIb3DQEJBDEiBCAd6F5BC6nv6a7eWgmPawHWbC093XZIG2DkS35T
# jIEmkDANBgkqhkiG9w0BAQEFAASCAgCm0LKwoUoSxB24o24DhD6Z4sGaSPFfDRpY
# jSzdcARMok1/2Zwic68BuPllRQGwpbUK9GBzVuyfLFj5P72OOkc1JsPgSHKoFvM/
# Kszm60hKPVIkNrBdgF+yQsKkDmyhdVHY1VbbvdjQUHPLwpMX4fkHpLHTUtUQ+if6
# rXzQHvDsCr8jU4wseuk4V5I0QeMdkeuhf/bOeDh6TbMLxw0jS+rFZEN27jRhvs67
# sN9NMPm7wmYjiMLevndtZdkjDcw/98M+09JNPn6QPVKSGbT+Z+zlJJ1CEpyE1zb8
# 5NrMHr7w3JBmA8ugmMmSZcARLu6q2VOyD9XoknkBNaTFrkW03qBwrUwifgaDcO/T
# 77FIbiF3TxiX0O+UJUYRT+nZy9sKV0IM8xFcQlOo7V6gayAIAXhPOz0ciu0wH3ze
# +dIYzfCv/oMJNJszViIz/QYqh8pfAmLubVf24v9yfTlF9jSbZDZf/28rFx1GZAuV
# GfD7oe/M9xOd3LzBAeD2au5RTteIKcZYdKyQBDYRXIY3JERq+g9VyKEwZgKTxyMr
# mvmnjfNk0roEFCQ77gw8+KuBikgSxTtBcFsYrau3idliV0VrJgVtjGTEaQWLCDTD
# Q2PIgbFoiwY4ktrk6auvFPuZDMKbmzRdD2gYz5fscaEZ3BW6hM+V0IQsI6RHpO0c
# Had/boeGpw==
# SIG # End signature block
