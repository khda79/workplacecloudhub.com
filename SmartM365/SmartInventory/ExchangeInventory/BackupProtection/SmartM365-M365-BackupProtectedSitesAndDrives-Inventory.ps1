#Requires -Version 7.0
<#
.SYNOPSIS
Collects Microsoft 365 Backup protected SharePoint site and OneDrive protection units.
.PARAMETER InteractiveAuth
Uses delegated interactive Microsoft Graph authentication. Enabled by default, as in the protected mailbox
inventory, because Microsoft 365 Backup Storage answers app-only calls from a client that is not a registered
Backup service app with HTTP 403 AppNotRegistered. Use -InteractiveAuth:$false only to request app-only
certificate authentication explicitly.
.VERSION
1.0.9
.REQUIREMENTS
Microsoft Graph delegated permission BackupRestore-Configuration.Read.All (default), or the application
permission for the explicit app-only mode.
#>
[CmdletBinding()]
param([string]$Tenant='test',[switch]$InteractiveAuth=$true,[switch]$IncludeNonProtected,[switch]$ValidateOnly,[switch]$SelfTest,[switch]$EnableConfiguredExternalActions,[string]$OutputPath,[string]$LatestCsvFolderPath,[ValidateRange(0,10000000)][int]$MaxItems=0)
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
$ErrorActionPreference='Stop';$ScriptVersion='1.0.9';$script:Runtime=$null;$script:CompletionStatus='Success';$script:CompletionError=$null

function ConvertTo-BackupProtectionRows{
    param([object[]]$Units,[ValidateSet('SharePointSite','OneDrive')][string]$Workload,[string]$RunId,[string]$CollectedAtUtc,[switch]$KeepNonProtected)
    return @($Units|Where-Object{$KeepNonProtected-or [string](Get-SmartM365EvidenceProperty $_ @('status'))-ieq'protected'}|ForEach-Object{[pscustomobject][ordered]@{RunId=$RunId;CollectedAtUtc=$CollectedAtUtc;CollectionStatus='Observed';IsPartialInventory=$false;SourceScriptVersion=$ScriptVersion;Workload=$Workload;ProtectionUnitId=[string](Get-SmartM365EvidenceProperty $_ @('id'));DirectoryObjectId=[string](Get-SmartM365EvidenceProperty $_ @('siteId','driveId','userId'));DisplayName=[string](Get-SmartM365EvidenceProperty $_ @('siteName','displayName'));WebUrl=[string](Get-SmartM365EvidenceProperty $_ @('siteWebUrl','webUrl'));Email=[string](Get-SmartM365EvidenceProperty $_ @('email'));Status=[string](Get-SmartM365EvidenceProperty $_ @('status'));CreatedDateTime=[string](Get-SmartM365EvidenceProperty $_ @('createdDateTime'));LastModifiedDateTime=[string](Get-SmartM365EvidenceProperty $_ @('lastModifiedDateTime'));RawIdentityJson=ConvertTo-SmartM365EvidenceJson $_}})
}

# The Graph error code sits in the response body (ErrorDetails), not in the exception message "Forbidden".
function Get-BackupStorageErrorCode{
    param([Parameter(Mandatory)][Management.Automation.ErrorRecord]$ErrorRecord)
    $texts=[Collections.Generic.List[string]]::new()
    if($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ErrorDetails.Message)){$texts.Add([string]$ErrorRecord.ErrorDetails.Message)}
    $exception=$ErrorRecord.Exception
    while($null -ne $exception){$texts.Add([string]$exception.Message);$exception=$exception.InnerException}
    foreach($text in $texts){
        $match=[regex]::Match($text,'"code"\s*:\s*"(?<Code>[^"]+)"')
        if($match.Success){return $match.Groups['Code'].Value}
        if($text -match '(?i)\bAppNotRegistered\b'){return 'AppNotRegistered'}
    }
    return ''
}

function Get-BackupStorageAccessFailureMessage{
    param([string]$ErrorCode,[string]$Uri,[string]$AuthMode,[string]$ClientId,[string]$OriginalMessage)
    $client=if([string]::IsNullOrWhiteSpace($ClientId)){'<unknown>'}else{$ClientId}
    if($ErrorCode -ine 'AppNotRegistered'){
        $codeText=if([string]::IsNullOrWhiteSpace($ErrorCode)){''}else{" Graph error code: $ErrorCode."}
        return ("Microsoft 365 Backup Storage access probe failed for {0} ({1} client {2}): {3}{4}" -f $Uri,$AuthMode,$client,$OriginalMessage,$codeText)
    }
    $cause="Microsoft 365 Backup Storage refused {0} with HTTP 403 AppNotRegistered ('Service app is not found') for {1} client application {2}. Probable cause: this client is not a registered Backup Storage service app (solutions/backupRestore/serviceApps). This code does not mean that the Backup service is disabled or that BackupRestore-Configuration.Read.All is missing." -f $Uri,$AuthMode,$client
    if($AuthMode -ieq 'AppOnly'){
        $action="Action: run with the default delegated authentication (do not pass -InteractiveAuth:`$false), as the protected mailbox inventory does, or use the job M365-BackupProtectedSitesAndDrives-Interactive. Do not register this inventory application as a Backup service app: registration requires BackupRestore-Control.ReadWrite.All, and activation makes it the Microsoft 365 Backup controller (billing, 7 to 30 day grace period)."
    }else{
        $action='Action: check the Microsoft 365 Backup Storage service apps and the delegated sign-in with a Microsoft 365 Backup administrator; the delegated client was refused as well.'
    }
    return "$cause $action"
}

# Read-only probe of both endpoints before collection, so that an access refusal is explained, not only "Forbidden".
function Assert-BackupStorageAccess{
    param([Parameter(Mandatory)][string[]]$Uris,[string]$AuthMode,[string]$ClientId,[scriptblock]$Invoker)
    foreach($uri in $Uris){
        try{
            $requestParameters=@{Uri=$uri;Operation='Microsoft 365 Backup Storage access probe'}
            if($Invoker){$requestParameters.Invoker=$Invoker;$requestParameters.NoSleep=$true}
            Invoke-SmartM365EvidenceGraphRequest @requestParameters|Out-Null
            if(-not $Invoker){Write-SmartM365EvidenceLog -Message ("Microsoft 365 Backup Storage access probe OK: {0}" -f $uri)}
        }catch{
            if($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]){throw}
            $message=Get-BackupStorageAccessFailureMessage -ErrorCode (Get-BackupStorageErrorCode -ErrorRecord $_) -Uri $uri -AuthMode $AuthMode -ClientId $ClientId -OriginalMessage $_.Exception.Message
            throw [InvalidOperationException]::new($message,$_.Exception)
        }
    }
}

if($SelfTest){Import-Module (Join-Path $PSScriptRoot '..\..\Common\SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.23' -Force;$test=ConvertTo-BackupProtectionRows @([pscustomobject]@{id='1';siteId='s';status='protected'},[pscustomobject]@{id='2';status='unprotected'}) SharePointSite r n;if($test.Count-ne 1 -or $test[0].DirectoryObjectId-ne's'){throw 'Backup protection self-test failed.'}
    # The Graph request helper reads the HTTP status through Core, as in a real run.
    Import-Module -Name (Join-Path $smartM365LifecycleRoot 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.121' -Global -ErrorAction Stop
    $denied=[Management.Automation.ErrorRecord]::new([Exception]::new('Response status code does not indicate success: Forbidden (Forbidden).'),'GraphForbidden',[Management.Automation.ErrorCategory]::PermissionDenied,$null)
    $denied.ErrorDetails=[Management.Automation.ErrorDetails]::new('{"error":{"code":"AppNotRegistered","message":"Service app is not found. Please register this service app."}}')
    if((Get-BackupStorageErrorCode -ErrorRecord $denied) -ne 'AppNotRegistered'){throw 'Backup protection self-test failed: AppNotRegistered code not read from the response body.'}
    $probeFailure=$null
    try{Assert-BackupStorageAccess -Uris @('https://graph.microsoft.test/probe') -AuthMode AppOnly -ClientId 'client-test' -Invoker {param($m,$u,$b) throw $denied}}catch{$probeFailure=$_}
    if($null -eq $probeFailure -or $probeFailure.Exception.Message -notmatch 'AppNotRegistered' -or $probeFailure.Exception.Message -notmatch 'client-test' -or $probeFailure.Exception.Message -notmatch 'M365-BackupProtectedSitesAndDrives-Interactive'){throw 'Backup protection self-test failed: AppNotRegistered refusal not explained.'}
    $other=[Management.Automation.ErrorRecord]::new([Exception]::new('Response status code does not indicate success: Forbidden (Forbidden).'),'GraphForbidden',[Management.Automation.ErrorCategory]::PermissionDenied,$null)
    $other.ErrorDetails=[Management.Automation.ErrorDetails]::new('{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges."}}')
    $otherFailure=$null
    try{Assert-BackupStorageAccess -Uris @('https://graph.microsoft.test/probe') -AuthMode AppOnly -ClientId 'client-test' -Invoker {param($m,$u,$b) throw $other}}catch{$otherFailure=$_}
    if($null -eq $otherFailure -or $otherFailure.Exception.Message -notmatch 'Authorization_RequestDenied' -or $otherFailure.Exception.Message -match 'serviceApps'){throw 'Backup protection self-test failed: other refusals must keep their own Graph code.'}
    'PASS: Backup sites and drives offline contract';$smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return}
try{
    $tenantContextPath=Join-Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) 'Config\SmartM365-TenantContext.ps1';. $tenantContextPath
    $effectiveConfig=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $PSScriptRoot '..\..\Common\SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.23' -Force
    $script:Runtime=Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effectiveConfig -DefaultOutputRelativePath 'M365\Backup\ProtectedSitesAndDrives' -OutputPath $OutputPath -LatestCsvFolderPath $LatestCsvFolderPath -ValidateOnly:$ValidateOnly -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath -ReadOnly:($ValidateOnly -or $SelfTest) -ScopeParameters @{IncludeNonProtected=$IncludeNonProtected}
    Connect-SmartM365EvidenceGraph $script:Runtime @('BackupRestore-Configuration.Read.All') -InteractiveAuth:$InteractiveAuth
    $graphContext=Get-MgContext
    $authMode=if($InteractiveAuth){'Delegated'}else{'AppOnly'}
    $clientId=''
    if($null -ne $graphContext){
        if(-not [string]::IsNullOrWhiteSpace([string]$graphContext.AuthType)){$authMode=[string]$graphContext.AuthType}
        $clientId=[string]$graphContext.ClientId
    }
    Write-SmartM365EvidenceLog -Message ("Microsoft Graph authentication: {0}; client application: {1}." -f $authMode,$(if($clientId){$clientId}else{'<unknown>'}))
    $preflightParameters=@{ScriptName=[IO.Path]::GetFileNameWithoutExtension($PSCommandPath);OutputPaths=@($script:Runtime.OutputPath,$script:Runtime.LatestCsvFolderPath)}
    if($authMode -ieq 'AppOnly'){$preflightParameters.RequiredGraphApplicationPermissions=@('BackupRestore-Configuration.Read.All')}
    Invoke-SmartM365Preflight @preflightParameters|Out-Null
    $siteUri='https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.siteProtectionUnit'
    $driveUri='https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.driveProtectionUnit'
    Assert-BackupStorageAccess -Uris @(($siteUri+'?$top=1'),($driveUri+'?$top=1')) -AuthMode $authMode -ClientId $clientId
    $sites=@(Get-SmartM365EvidenceGraphCollection -Uri $siteUri -MaxItems $MaxItems)
    $drives=@(Get-SmartM365EvidenceGraphCollection -Uri $driveUri -MaxItems $MaxItems)
    if($ValidateOnly){Write-SmartM365EvidenceLog -Message ("Backup protection endpoints available; sites={0}, drives={1}. No CSV published." -f $sites.Count,$drives.Count) -Level SUCCESS;$smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return}
    $siteRows=ConvertTo-BackupProtectionRows $sites SharePointSite $script:Runtime.RunId $script:Runtime.CollectedAtUtc -KeepNonProtected:$IncludeNonProtected
    $driveRows=ConvertTo-BackupProtectionRows $drives OneDrive $script:Runtime.RunId $script:Runtime.CollectedAtUtc -KeepNonProtected:$IncludeNonProtected
    $columns=@('RunId','CollectedAtUtc','CollectionStatus','IsPartialInventory','SourceScriptVersion','Workload','ProtectionUnitId','DirectoryObjectId','DisplayName','WebUrl','Email','Status','CreatedDateTime','LastModifiedDateTime','RawIdentityJson')
    Export-SmartM365EvidenceDataset $script:Runtime 'M365_Backup_ProtectedSharePointSites' $siteRows $columns|Out-Null
    Export-SmartM365EvidenceDataset $script:Runtime 'M365_Backup_ProtectedOneDriveAccounts' $driveRows $columns|Out-Null

    $smartM365RunGuard.ReachedTerminal=$true
}catch{
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_
$script:CompletionStatus='Failed';$script:CompletionError=$_;throw}finally{
    $smartM365RunGuard.InCleanup=$true
try{Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false}catch{};if($script:Runtime){Complete-SmartM365EvidenceRuntime -Status $script:CompletionStatus -ErrorRecord $script:CompletionError -FailureStage 'BackupSitesAndDrivesInventory'}}
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUiHmYPinfQkFu
# BrplCYy+Kzylmdt02iyzKpq7+5dJ7qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHr/k+sICLd5ibqtwqzEPOZZojMpQ3R/TnVr7OQrc8a3MA0GCSqG
# SIb3DQEBAQUABIIBgKwl07vtdAxkqwTh7/JYB+NJSwaKy1+RwVDWFwzYRVy9XN0a
# gdZ73ZXYkc+fFI21K6ao04QdfH1tmg5ZIgK9Ts+SfxwFSzHOstO2CuIsoCuiIrqg
# l0a7VULqT53YDKOIg539FM7vN7gD5tarEZwO/n/ppPB66B3f0aFPisdPusiEz7O1
# RcKj/6F/fSTcQ3VVNsTVNLsYQPqZERxZP112aPAfMYEjHKfefFdLV/EbBZCUy5v5
# vcYKTX1N43xTettYKs6epgSDdX122luGfANXbGtcYfALCRGKum2MpYm9yBgeqMZm
# +Gizyf1UUkHM++uH1bvjIeVgO8zYaN9mlNTggYywIe7gF+fSpeLqQydIVyJbSOx4
# VM6OSw5PPJYF4OjmX7ZmIZgSgy1gJN8DjkJDAK3HvkOfRKa6HN/fzn953OZayFcB
# 3bNFEjbq3bh/0LxiKosvw68Ox9iQo7GMt7fI2FQShsbejqFxfjnIzyb3EiTRU3fL
# cprKrzdasVmxMsy5BqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTAxMzAx
# MzBaMC8GCSqGSIb3DQEJBDEiBCCdIE2nIRQcGqygEelxZuvlIXWIHfpqITHZQ5kp
# DcEvIzANBgkqhkiG9w0BAQEFAASCAgAQxfBMmp4V83nmZV62JsOBQBPPpC3TeduX
# /Ds/Oo8YQD0ExnSuycUjdZ9LYirWlu6m5jdBQA82t/WO9e4O2qmonPAJRg8pq4rN
# ab1dSPg3hd5yhC1M+Qem6vkBIIY59UwsC3PcAFIpAcOTN6VBqRAGP/85X6Pcd6lW
# PSgQLajoP4+l622k6yw5MpbjQUqPSQcA0ZPt5muxQKclWo88rPNK40sAky4zR4Sh
# tMhlNfPTYMREM1owJku/PwyOx989L4cvRrA8X6UX0+3200Si2amd1RTpennVsSwT
# X53riX5tOuZMvH9DYsmd24Qny7GkHKMxUnORZIf8g+ktUmMTc7jrybikkDekLBKc
# TiQa7xBULpxy1oBS/VMe4hjkTnT86qwcfghmJlpkmDGIgW25XJyFIMDWy6y94jSu
# Op0Bw1qPMO37hn/CwYbeb/D+jcReUtK7Lg70+7pcTZdp4MeH0Ha04EDdAJvcEvwv
# 4CCYuPVpm0XiAYkNnmk04sblYazxiZtLMtLq84MDnGBa2SQOkRYl5FLvBqiI2K+U
# Kw+UkDn+nP540xYdoU00iyGRRT8qp/tBnZ9l8XfUJNKUU1ZnBo+Hvua5Ae7Rkuun
# lyCUiZADoTqNizwOeXAoAqabyhN84w0HZ89ilNOuas0UILHzKonfhviMnVOJ4Qhj
# +XrWMf7+bA==
# SIG # End signature block
