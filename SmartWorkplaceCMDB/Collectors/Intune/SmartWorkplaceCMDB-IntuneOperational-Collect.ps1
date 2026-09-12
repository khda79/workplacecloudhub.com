<#
.SYNOPSIS
Collects read-only Intune operational inventory for SmartWorkplaceCMDB.

.DESCRIPTION
Publishes separate raw snapshots for Autopilot devices, detected applications,
configuration policies, and Windows feature/quality update policies. The
collector uses explicit Microsoft Graph fields and supports one synthetic JSON
fixture containing the four source families.

.VERSION
1.0.0
#>
[CmdletBinding(DefaultParameterSetName='Graph')]
param(
    [Alias('ProfileKey')][string]$Tenant='default',
    [string]$OrganizationKey,[string]$EnvironmentKey,[string]$TenantKey,[string]$TenantId,
    [string]$DataRootPath,[string]$DataAllRootPath,[string]$LatestOutputRootPath,[string]$LogRootPath,
    [string]$GlobalConfigPath,[string]$TenantConfigPath,
    [Parameter(ParameterSetName='Fixture',Mandatory)][string]$InputJsonPath,
    [ValidateRange(0,2147483647)][int]$MaxItems=0,
    [switch]$NoConfigWrite,[switch]$ValidateOnly
)

$ScriptVersion='1.0.0'
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0

function Get-ConfigSection { param([System.Collections.IDictionary]$Configuration,[string]$Name) if($Configuration.Contains($Name) -and $Configuration[$Name] -is [System.Collections.IDictionary]){return $Configuration[$Name]};return [ordered]@{} }
function Get-ConfigText { param([System.Collections.IDictionary]$Configuration,[string]$Name) if($Configuration.Contains($Name)){return ([string]$Configuration[$Name]).Trim()};return '' }
function Get-Value { param([AllowNull()]$Object,[string]$Name) return Get-SmartWorkplaceCMDBGraphObjectValue $Object $Name }
function Get-Text { param([AllowNull()]$Value) if($null-eq $Value){return ''};return ([string]$Value-replace"`r`n|`n|`r",' ').Trim() }
function Get-DateText { param([AllowNull()]$Value,[string]$Field,[string]$Key) if($null-eq $Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return ''};$date=[datetimeoffset]::MinValue;if(-not[datetimeoffset]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$date)){throw "$Field '$Value' is invalid for '$Key'."};return $date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ',[Globalization.CultureInfo]::InvariantCulture) }
function Get-IntegerText { param([AllowNull()]$Value,[string]$Field,[string]$Key) if($null-eq $Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return ''};$number=0;if(-not[int]::TryParse([string]$Value,[ref]$number)-or$number-lt 0){throw "$Field '$Value' is invalid for '$Key'."};return $number }
function Get-ListText { param([AllowNull()]$Value) return (@(@($Value)|ForEach-Object{Get-Text $_}|Where-Object{-not[string]::IsNullOrWhiteSpace($_)}|Sort-Object -Unique)-join';') }

$scriptRoot=Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot=Split-Path -Parent (Split-Path -Parent $scriptRoot)
$core=Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$graph=Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Graph\SmartWorkplaceCMDB.Graph.psd1'
$rawContractPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
Import-Module $core -Force
Import-Module $graph -Force
$bound=@{};foreach($key in $PSBoundParameters.Keys){$bound[$key]=$PSBoundParameters[$key]}
$context=Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly -or $NoConfigWrite -or $PSCmdlet.ParameterSetName -eq 'Fixture')
$paths=Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -ExplicitDataRoot:([bool]$DataRootPath) -NoWrite:$ValidateOnly
$contract=Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$tableNames=@('Intune_AutopilotDevices.csv','Intune_DetectedApps.csv','Intune_ConfigurationPolicies.csv','Intune_WindowsUpdatePolicies.csv')
$tables=@{};$latest=@{}
foreach($name in $tableNames){$matches=@($contract.tables | Where-Object name -eq $name);if($matches.Count -ne 1){throw "Raw contract definition missing or duplicated: $name"};$tables[$name]=$matches[0];$latest[$name]=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$matches[0].area) $name)))}

$mode=if($ValidateOnly){'Validate'}elseif($PSCmdlet.ParameterSetName -eq 'Fixture'){'Fixture'}else{'Collect'}
$runtime=Start-SmartWorkplaceCMDBExecutionContext -Context $context -ScriptPath $PSCommandPath -ScriptVersion $ScriptVersion -Mode $mode -NoWrite:$ValidateOnly
$executionError=$null
try{
    $configuration=Get-ConfigSection $context.Configuration 'MicrosoftGraph'
    $clientId=Get-ConfigText $configuration 'ClientId';$thumbprint=Get-ConfigText $configuration 'CertificateThumbprint'
    $fixture=$null;$readiness=$null
    if($PSCmdlet.ParameterSetName -eq 'Fixture'){$InputJsonPath=[IO.Path]::GetFullPath($InputJsonPath);$fixture=Get-Content -Raw -LiteralPath $InputJsonPath|ConvertFrom-Json}
    else{$readiness=Test-SmartWorkplaceCMDBGraphAppOnlyReadiness -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint}
    if($ValidateOnly){
        [pscustomobject]@{Status='Valid';ScriptVersion=$ScriptVersion;SourceMode=if($fixture){'OfflineJson'}else{'MicrosoftGraphAppOnly'};RawContractVersion=[string]$contract.contractVersion;RequiredGraphPermissions='DeviceManagementServiceConfig.Read.All;DeviceManagementApps.Read.All;DeviceManagementConfiguration.Read.All';OutputCount=$tableNames.Count}|Format-List
        return
    }
    if($fixture){
        $autopilot=@($fixture.autopilotDevices);$apps=@($fixture.detectedApps);$policies=@($fixture.configurationPolicies);$feature=@($fixture.featureUpdatePolicies);$quality=@($fixture.qualityUpdatePolicies)
    }else{
        $autopilot=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=999' -RequiredPermission 'DeviceManagementServiceConfig.Read.All' -MaxItems $MaxItems)
        $apps=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$select=id,displayName,version,publisher,deviceCount,platform&$top=999' -RequiredPermission 'DeviceManagementApps.Read.All' -MaxItems $MaxItems)
        $policies=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name,description,platforms,technologies,templateReference,createdDateTime,lastModifiedDateTime&$top=100' -RequiredPermission 'DeviceManagementConfiguration.Read.All' -MaxItems $MaxItems)
        $feature=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?$top=100' -RequiredPermission 'DeviceManagementConfiguration.Read.All' -MaxItems $MaxItems)
        $quality=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles?$top=100' -RequiredPermission 'DeviceManagementConfiguration.Read.All' -MaxItems $MaxItems)
    }
    if($MaxItems -gt 0){$autopilot=@($autopilot|Select-Object -First $MaxItems);$apps=@($apps|Select-Object -First $MaxItems);$policies=@($policies|Select-Object -First $MaxItems);$feature=@($feature|Select-Object -First $MaxItems);$quality=@($quality|Select-Object -First $MaxItems)}
    $collected=[datetime]::UtcNow.ToString('o')
    $rows=@{}
    $rows['Intune_AutopilotDevices.csv']=@($autopilot|ForEach-Object{$id=Get-Text(Get-Value $_ 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Autopilot response missing id.'};[pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';AutopilotDeviceId=$id;DisplayName=Get-Text(Get-Value $_ 'displayName');SerialNumber=Get-Text(Get-Value $_ 'serialNumber');Manufacturer=Get-Text(Get-Value $_ 'manufacturer');Model=Get-Text(Get-Value $_ 'model');GroupTag=Get-Text(Get-Value $_ 'groupTag');PurchaseOrderIdentifier=Get-Text(Get-Value $_ 'purchaseOrderIdentifier');EnrollmentState=Get-Text(Get-Value $_ 'enrollmentState');LastContactedDateTime=Get-DateText (Get-Value $_ 'lastContactedDateTime') 'lastContactedDateTime' $id;UserPrincipalName=Get-Text(Get-Value $_ 'userPrincipalName');AzureAdDeviceId=Get-Text(Get-Value $_ 'azureActiveDirectoryDeviceId');ManagedDeviceId=Get-Text(Get-Value $_ 'managedDeviceId');SourceCollectedDateTime=$collected}})
    $rows['Intune_DetectedApps.csv']=@($apps|ForEach-Object{$id=Get-Text(Get-Value $_ 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Detected application response missing id.'};[pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';AppId=$id;DisplayName=Get-Text(Get-Value $_ 'displayName');Version=Get-Text(Get-Value $_ 'version');Publisher=Get-Text(Get-Value $_ 'publisher');DeviceCount=Get-IntegerText (Get-Value $_ 'deviceCount') 'deviceCount' $id;Platform=Get-Text(Get-Value $_ 'platform');SourceCollectedDateTime=$collected}})
    $rows['Intune_ConfigurationPolicies.csv']=@($policies|ForEach-Object{$id=Get-Text(Get-Value $_ 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Configuration policy response missing id.'};$template=Get-Value $_ 'templateReference';[pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';PolicyId=$id;DisplayName=Get-Text(Get-Value $_ 'name');Description=Get-Text(Get-Value $_ 'description');Platforms=Get-ListText(Get-Value $_ 'platforms');Technologies=Get-ListText(Get-Value $_ 'technologies');TemplateId=Get-Text(Get-Value $template 'templateId');TemplateFamily=Get-Text(Get-Value $template 'templateFamily');CreatedDateTime=Get-DateText (Get-Value $_ 'createdDateTime') 'createdDateTime' $id;LastModifiedDateTime=Get-DateText (Get-Value $_ 'lastModifiedDateTime') 'lastModifiedDateTime' $id;SourceCollectedDateTime=$collected}})
    $updateRows=New-Object System.Collections.Generic.List[object]
    foreach($item in $feature){$id=Get-Text(Get-Value $item 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Feature update policy response missing id.'};$updateRows.Add([pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';PolicyType='Feature';PolicyId=$id;DisplayName=Get-Text(Get-Value $item 'displayName');TargetVersion=Get-Text(Get-Value $item 'featureUpdateVersion');ReleaseDateTime='';DaysUntilForcedReboot='';CreatedDateTime=Get-DateText (Get-Value $item 'createdDateTime') 'createdDateTime' $id;LastModifiedDateTime=Get-DateText (Get-Value $item 'lastModifiedDateTime') 'lastModifiedDateTime' $id;SourceCollectedDateTime=$collected})}
    foreach($item in $quality){$id=Get-Text(Get-Value $item 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Quality update policy response missing id.'};$updateRows.Add([pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';PolicyType='Quality';PolicyId=$id;DisplayName=Get-Text(Get-Value $item 'displayName');TargetVersion='';ReleaseDateTime=Get-DateText (Get-Value $item 'expeditedUpdateReleaseDateTime') 'expeditedUpdateReleaseDateTime' $id;DaysUntilForcedReboot=Get-IntegerText (Get-Value $item 'daysUntilForcedReboot') 'daysUntilForcedReboot' $id;CreatedDateTime=Get-DateText (Get-Value $item 'createdDateTime') 'createdDateTime' $id;LastModifiedDateTime=Get-DateText (Get-Value $item 'lastModifiedDateTime') 'lastModifiedDateTime' $id;SourceCollectedDateTime=$collected})}
    $rows['Intune_WindowsUpdatePolicies.csv']=@($updateRows.ToArray())
    foreach($name in $tableNames){$keyColumn=if($name -eq 'Intune_AutopilotDevices.csv'){'AutopilotDeviceId'}elseif($name -eq 'Intune_DetectedApps.csv'){'AppId'}else{'PolicyId'};$duplicate=@($rows[$name]|Group-Object $keyColumn|Where-Object Count -gt 1);if($name -eq 'Intune_WindowsUpdatePolicies.csv'){$duplicate=@($rows[$name]|Group-Object {"$($_.PolicyType)|$($_.PolicyId)"}|Where-Object Count -gt 1)};if($duplicate.Count){throw "Duplicate keys returned for ${name}: $($duplicate.Name -join ', ')"}}
    $published=@()
    foreach($name in $tableNames){
        $run=$null
        try{$run=Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($latest[$name]) -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems;$stamp=[datetime]::UtcNow;$base=[IO.Path]::GetFileNameWithoutExtension($name);$history=Join-Path $paths.DataAllRootPath ('Intune\Operational\{0}\{1}\{2}_{3}.csv' -f $stamp.ToString('yyyy'),$stamp.ToString('MM'),$base,$stamp.ToString('yyyyMMdd-HHmmssfff'));Publish-SmartWorkplaceCMDBSourceCsv -Run $run -InputObject @($rows[$name]) -Columns @($tables[$name].columns|ForEach-Object{[string]$_}) -HistoryPath $history -LatestPath $latest[$name] -ContractPath $rawContractPath -ContractTableName $name|Out-Null;$published+=$latest[$name]}catch{if($null -ne $run){Complete-SmartWorkplaceCMDBSourceCollection -Run $run -Failed};throw}
    }
    Write-Information ("SmartWorkplaceCMDB Intune operational collection completed. Autopilot={0}; Apps={1}; ConfigurationPolicies={2}; UpdatePolicies={3}."-f$rows['Intune_AutopilotDevices.csv'].Count,$rows['Intune_DetectedApps.csv'].Count,$rows['Intune_ConfigurationPolicies.csv'].Count,$rows['Intune_WindowsUpdatePolicies.csv'].Count) -InformationAction Continue
    [pscustomobject]@{Status='Completed';ScriptVersion=$ScriptVersion;AutopilotCount=$rows['Intune_AutopilotDevices.csv'].Count;DetectedAppCount=$rows['Intune_DetectedApps.csv'].Count;ConfigurationPolicyCount=$rows['Intune_ConfigurationPolicies.csv'].Count;UpdatePolicyCount=$rows['Intune_WindowsUpdatePolicies.csv'].Count;PublishedPath=$published}
}catch{$executionError=$_;throw}finally{Complete-SmartWorkplaceCMDBExecutionContext -RuntimeContext $runtime -ErrorRecord $executionError}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC4I/RW4ebKnhVc
# FPVk2zSlUcUcIEJVYf7xXpnOZ98lKqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDgD6To4sk7fGOvYNmmKVjqFZiTJLy+9xu9tuqZVAWTJMA0GCSqG
# SIb3DQEBAQUABIIBgJvB/As/qvHcXlebT3EDTwXP23lrqVhFygkY/2UO5ZtPcW/C
# cReSEG74JzAYj+IYZU6aRWJyD3CqubM7wLkNBZZJ5Dh6TkuUY3zGHeFzvkRBXUlh
# K+MfuC3VFG4z5z0Dnd5F67lwlBzGz5Irm0PrF10/r0BhYe0AlLxd9kl97vGMMz1T
# AqySioM8IDglTv8lDD/duMXUSVyJhwDvQ/qAwSbqIGEgxzXQsBLoK3a/sczt62oE
# 38k3AYfSSXjn87Wt4Z7If01RHoZzxJMfU7tXfDj4nls6FHPmlbHhUWM1ezr36R5D
# 9J05TkeiiikrjXWlPaH0zvat4YwK7dnN9YdPTV5xjNB4JGR1U7tHA05UqCg9hAOA
# PU8UKY9qsBlAxpEBPql30kUWPXgcCXYV3HP86xLk0UJY89C5AG+9h3lY9eL7A6hp
# oiDKCcr5ReQlVla+jjpg9FybK1/hqV8lWUZ30rfRTS9p0B70nxCn5AeIqGxP8oQj
# 2l2tF3qQG1Ij1ZHxMaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxODIx
# MTZaMC8GCSqGSIb3DQEJBDEiBCD6jsb3qFuxoQeLbn1qbsQ8egcB2cQcnFAHEViP
# ij86DTANBgkqhkiG9w0BAQEFAASCAgCLRp6aohZQgG1QBHcO2da3swLpUtgDgS5c
# Y/1U5NTqKm2RPnogw/0JVQQaK1Qyd516h2FZO8J8FhBfAas0hpISMJ92vkb+kE54
# fZtKAIc4QdoN3wrYJa86K8NaldYc12Hsxo+ZqPlXZuGqfRVdVc5fMfLhJKi8uLp3
# m943QqHyLrmI/1fHo1318hAE9MApEcPPLyX4P793UCFyJnwz0+iBKtK5cXiMsrE2
# +m4BoA53XOzoYYAhHyt1t3zzBD2P05L+RE43a3Ipd3hPC9kOh2LXUYW9smSVrQiy
# XlV3N8Z+LPmXcuW7AhGx39oKUicb/bZ0SQ7/pmVIMaf2T53G2DnWk4Ia7CbEU/qj
# 6IinZK90t5ml5XwdCh8LJO7YJtDENZBH+Ut7fxSQ3jTjGpleN9xQN09V2iO+Hki/
# TH53Mb4VTdSQkxWQVs4ZBXQ8mjuXp1iu8ls+Uh29MPk7Oai6JcB1VKddHeoRuIdo
# UnrGbhcdDErP1WNws7+VRg/uC9+352/ecutC/Ren/DkE4KkJXKqvF/ncknbuqvgv
# oJw/3rVHwvMKVjsumpaxdkPSUNdpdYRSIQ72mqBUWGhMH3hGSrN1RYR3PGoe/W1+
# dewYzP5yXTz8JvlqEWTuJIG5jj8ByFbaXzFMlcGnNhWTL4bCw8U/PVTWUfTgC9Bi
# MfjNlGNmhg==
# SIG # End signature block
