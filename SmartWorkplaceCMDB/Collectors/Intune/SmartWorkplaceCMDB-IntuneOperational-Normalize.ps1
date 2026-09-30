<#
.SYNOPSIS
Normalizes Intune operational inventory into dedicated Power BI tables.

.VERSION
1.3.1
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant='default',
    [string]$OrganizationKey,[string]$EnvironmentKey,[string]$TenantKey,[string]$TenantId,
    [string]$DataRootPath,[string]$DataAllRootPath,[string]$LatestOutputRootPath,[string]$LogRootPath,
    [string]$GlobalConfigPath,[string]$TenantConfigPath,
    [switch]$NoConfigWrite,[switch]$ValidateOnly
)

$ScriptVersion='1.3.1'
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0

function Get-HeaderStatus { param([string]$Path,[string[]]$Columns) if(-not (Test-Path -LiteralPath $Path -PathType Leaf)){return 'Missing'};$line=Get-Content -LiteralPath $Path -TotalCount 1;$actual=if([string]::IsNullOrWhiteSpace($line)){@()}else{@($line.Split(',') | ForEach-Object {$_.Trim().Trim('"')})};if(($actual -join [char]31) -ceq ($Columns -join [char]31)){return 'Valid'};return 'Incompatible' }
function Get-KeyText { param([AllowNull()]$Value) if($null -eq $Value){return ''};return ([string]$Value).Trim().ToLowerInvariant() }
function ConvertTo-CsvField { param([AllowNull()]$Value) return '"'+(([string]$Value)-replace'"','""')+'"' }
function Assert-SourceSnapshotStatus {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)]$Paths)
    $statePath=$Path+'.status.json.txt'
    if(-not(Test-Path -LiteralPath $statePath -PathType Leaf)){$statePath=$Path+'.status.json'}
    if(-not(Test-Path -LiteralPath $statePath -PathType Leaf)){return}
    $state=Get-Content -Raw -LiteralPath $statePath|ConvertFrom-Json
    foreach($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')){if([string]$state.$name-ne[string]$Paths.$name){throw 'Source evidence identity mismatch.'}}
    if($state.Status-ne'Completed'-or$state.SHA256-ne(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash){throw "Unusable source snapshot: '$Path'. Recollect before normalization."}
}
function Export-DeviceApplicationFactStreaming {
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)]$Paths
    )
    Assert-SourceSnapshotStatus -Path $InputPath -Paths $Paths
    $folder=Split-Path -Parent $OutputPath;if(-not(Test-Path -LiteralPath $folder)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}
    $tempPath='{0}.tmp.{1}.csv'-f$OutputPath,[guid]::NewGuid().ToString('N')
    $writer=$null;$count=0;$keys=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try{
        $writer=[IO.StreamWriter]::new($tempPath,$false,[Text.UTF8Encoding]::new($true))
        $writer.WriteLine((@($Columns|ForEach-Object{ConvertTo-CsvField $_})-join','))
        Import-Csv -LiteralPath $InputPath|ForEach-Object{
            $row=$_
            foreach($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')){if([string]$row.$name-ne[string]$Paths.$name){throw "CSV tenant identity mismatch for '$name'."}}
            $appId=Get-KeyText $row.AppId;$managedDeviceId=Get-KeyText $row.ManagedDeviceId
            if(-not$appId-or-not$managedDeviceId){throw 'Empty AppId or ManagedDeviceId in Intune_DetectedAppDeviceRelationships.csv.'}
            $relationshipKey=Get-KeyText $row.RelationshipKey;$expectedKey=$appId+'|'+$managedDeviceId
            if($relationshipKey-cne$expectedKey){throw "RelationshipKey does not match its exact AppId and ManagedDeviceId: '$relationshipKey'."}
            if(-not$keys.Add($relationshipKey)){throw "Duplicate operational key found in Intune_DetectedAppDeviceRelationships.csv: '$relationshipKey'."}
            $values=[ordered]@{
                TenantKey=$Paths.TenantKey;OrganizationKey=$Paths.OrganizationKey;EnvironmentKey=$Paths.EnvironmentKey;TenantId=$Paths.TenantId
                TenantDeviceApplicationKey=('{0}|device-application|{1}|{2}'-f$Paths.TenantKey,$appId,$managedDeviceId)
                TenantApplicationKey=('{0}|detected-app|{1}'-f$Paths.TenantKey,$appId)
                ManagedDeviceId=[string]$row.ManagedDeviceId;AppId=[string]$row.AppId;SourceCollectedDateTime=[string]$row.SourceCollectedDateTime
            }
            $writer.WriteLine((@($Columns|ForEach-Object{ConvertTo-CsvField $values[$_]})-join','));$count++
            if($count%500000-eq 0){Write-Information("Intune application-device normalization: {0} exact relation(s)."-f$count)-InformationAction Continue}
        }
        $writer.Dispose();$writer=$null
        Move-Item -LiteralPath $tempPath -Destination $OutputPath -Force
    }finally{if($writer){$writer.Dispose()};if(Test-Path -LiteralPath $tempPath){Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue}}
    return $count
}

function Export-CuratedCsvAtomic {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$InputObject,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)]$Identity
    )
    $folder=Split-Path -Parent $OutputPath
    if(-not(Test-Path -LiteralPath $folder)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}
    $tempPath='{0}.tmp.{1}.csv'-f$OutputPath,[guid]::NewGuid().ToString('N')
    try{
        Export-SmartWorkplaceCMDBCsv -InputObject $InputObject -Path $tempPath -Columns $Columns @Identity
        if((Get-HeaderStatus $tempPath $Columns)-ne'Valid'){throw "Curated operational staging output failed validation: $tempPath"}
        Move-Item -LiteralPath $tempPath -Destination $OutputPath -Force
    }finally{if(Test-Path -LiteralPath $tempPath){Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue}}
}

$scriptRoot=Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot=Split-Path -Parent (Split-Path -Parent $scriptRoot)
$module=Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
Import-Module $module -Force
$bound=@{};foreach($key in $PSBoundParameters.Keys){$bound[$key]=$PSBoundParameters[$key]}
$context=Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths=$context.Paths
$rawContract=Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$curatedContract=Get-SmartWorkplaceCMDBTableContract -Path $curatedContractPath
$mappings=@(
    @{Raw='Intune_AutopilotDevices.csv';Curated='FactAutopilotDevice.csv';Key='AutopilotDeviceId';Prefix='autopilot'},
    @{Raw='Intune_DetectedApps.csv';Curated='DimDetectedApplication.csv';Key='AppId';Prefix='app'},
    @{Raw='Intune_DetectedAppDeviceRelationships.csv';Curated='FactDeviceApplication.csv';Key='RelationshipKey';Prefix='device-application'},
    @{Raw='Intune_ConfigurationPolicies.csv';Curated='DimIntuneConfigurationPolicy.csv';Key='PolicyId';Prefix='configuration-policy'},
    @{Raw='Intune_WindowsUpdatePolicies.csv';Curated='DimWindowsUpdatePolicy.csv';Key='PolicyId';Prefix='update-policy'}
)
$definitions=@()
foreach($mapping in $mappings){
    $raw=@($rawContract.tables | Where-Object name -eq $mapping.Raw);$target=@($curatedContract.tables | Where-Object name -eq $mapping.Curated)
    if($raw.Count -ne 1 -or $target.Count -ne 1){throw "Operational contract definition missing or duplicated: $($mapping.Raw) / $($mapping.Curated)"}
    $input=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$raw[0].area) ([string]$raw[0].name))))
    $output=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$target[0].area) ([string]$target[0].name))))
    $inputStatus=Get-HeaderStatus $input @($raw[0].columns|ForEach-Object{[string]$_});$outputStatus=Get-HeaderStatus $output @($target[0].columns|ForEach-Object{[string]$_})
    if($inputStatus -eq 'Incompatible'){throw "Operational source CSV contract is incompatible: $($mapping.Raw)"}
    $definitions+=@{Mapping=$mapping;Raw=$raw[0];Target=$target[0];Input=$input;Output=$output;InputStatus=$inputStatus;OutputStatus=$outputStatus}
}
if($ValidateOnly){foreach($definition in $definitions){if($definition.InputStatus -eq 'Valid'){Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $definition.Input -Paths $paths|Out-Null}};[pscustomobject]@{Status='Valid';ScriptVersion=$ScriptVersion;DatasetCount=$definitions.Count;RawContractVersion=[string]$rawContract.contractVersion;CuratedContractVersion=[string]$curatedContract.contractVersion}|Format-List;return}
$identity=@{TenantKey=$paths.TenantKey;OrganizationKey=$paths.OrganizationKey;EnvironmentKey=$paths.EnvironmentKey;TenantId=$paths.TenantId}
$published=@()
foreach($definition in $definitions){
    if($definition.InputStatus -eq 'Missing'){throw "Required raw operational inventory is missing: $($definition.Input)"}
    if($definition.Mapping.Raw -eq 'Intune_DetectedAppDeviceRelationships.csv'){
        $count=Export-DeviceApplicationFactStreaming -InputPath $definition.Input -OutputPath $definition.Output -Columns @($definition.Target.columns|ForEach-Object{[string]$_}) -Paths $paths
        if((Get-HeaderStatus $definition.Output @($definition.Target.columns|ForEach-Object{[string]$_})) -ne 'Valid'){throw "Curated operational output failed validation: $($definition.Output)"}
        $published+=[pscustomobject]@{Table=$definition.Mapping.Curated;Count=$count;Path=$definition.Output}
        continue
    }
    $rows=@(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $definition.Input -Paths $paths)
    $keyName=[string]$definition.Mapping.Key
    $duplicates=@(if($definition.Mapping.Raw -eq 'Intune_WindowsUpdatePolicies.csv'){$rows | Group-Object {"$($_.PolicyType)|$($_.PolicyId)"} | Where-Object Count -gt 1}else{$rows | Group-Object $keyName | Where-Object Count -gt 1})
    if($duplicates.Count){throw "Duplicate operational keys found in $($definition.Mapping.Raw): $($duplicates.Name -join ', ')"}
    $targetRows=@($rows|ForEach-Object{
        $row=$_
        $key=Get-KeyText $row.$keyName;if([string]::IsNullOrWhiteSpace($key)){throw "Empty $keyName in $($definition.Mapping.Raw)."}
        $properties=[ordered]@{}
        switch($definition.Mapping.Raw){
            'Intune_AutopilotDevices.csv'{$properties.TenantAutopilotDeviceKey=('{0}|autopilot|{1}'-f$paths.TenantKey,$key);foreach($name in @('AutopilotDeviceId','DisplayName','SerialNumber','Manufacturer','Model','GroupTag','EnrollmentState','LastContactedDateTime','AzureAdDeviceId','ManagedDeviceId','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
            'Intune_DetectedApps.csv'{$properties.TenantApplicationKey=('{0}|detected-app|{1}'-f$paths.TenantKey,$key);foreach($name in @('AppId','SourceApplicationKey','DisplayName','Version','Publisher','DeviceCount','ReportedDeviceCount','ExactRelatedDeviceCount','RelationshipCoverageStatus','Platform','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
            'Intune_DetectedAppDeviceRelationships.csv'{$appId=Get-KeyText $row.AppId;$managedDeviceId=Get-KeyText $row.ManagedDeviceId;$properties.TenantDeviceApplicationKey=('{0}|device-application|{1}|{2}'-f$paths.TenantKey,$appId,$managedDeviceId);$properties.TenantApplicationKey=('{0}|detected-app|{1}'-f$paths.TenantKey,$appId);foreach($name in @('ManagedDeviceId','AppId','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
            'Intune_ConfigurationPolicies.csv'{$properties.TenantPolicyKey=('{0}|configuration-policy|{1}'-f$paths.TenantKey,$key);foreach($name in @('PolicyId','DisplayName','Description','Platforms','Technologies','TemplateId','TemplateFamily','CreatedDateTime','LastModifiedDateTime','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
            'Intune_WindowsUpdatePolicies.csv'{$type=Get-KeyText $row.PolicyType;$properties.TenantUpdatePolicyKey=('{0}|update-policy|{1}|{2}'-f$paths.TenantKey,$type,$key);foreach($name in @('PolicyType','PolicyId','DisplayName','TargetVersion','ReleaseDateTime','DaysUntilForcedReboot','CreatedDateTime','LastModifiedDateTime','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
        }
        [pscustomobject]$properties
    })
    Export-CuratedCsvAtomic -InputObject $targetRows -OutputPath $definition.Output -Columns @($definition.Target.columns|ForEach-Object{[string]$_}) -Identity $identity
    if((Get-HeaderStatus $definition.Output @($definition.Target.columns|ForEach-Object{[string]$_})) -ne 'Valid'){throw "Curated operational output failed validation: $($definition.Output)"}
    $published+=[pscustomobject]@{Table=$definition.Mapping.Curated;Count=$targetRows.Count;Path=$definition.Output}
}
Write-Information ("SmartWorkplaceCMDB Intune operational normalization completed. Tables={0}; Rows={1}." -f $published.Count, (($published | Measure-Object Count -Sum).Sum)) -InformationAction Continue
[pscustomobject]@{Status='Completed';ScriptVersion=$ScriptVersion;Published=$published;RawContractVersion=[string]$rawContract.contractVersion;CuratedContractVersion=[string]$curatedContract.contractVersion}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDsEXeb0Ierv8ug
# AlPZCiKLKdwWEN0N1OxnesrZqo5qyKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCvKNFsZ85CU7cPqSjoj5zC
# PdgA2jBorOnp/EN/kX6mcDANBgkqhkiG9w0BAQEFAASCAYAI3guHGl8217VaQGvF
# +81GisNVZK6x/Fw1RvgGbyT66OhWvxORVr2LkmrfcW1hBb1DEGkw/b6nIr9Pnu1p
# Q39roIsea5MjqnbYMfPjoyXJfWyCFI7yUr2EnB7npIHPbGuiZ9jXDyNz6lvzajv0
# RmKw4jBk88P5OKqZY0Ek++eEGTXUPKR1RlsGeAyrF5aDFffKJ/as+ZAslwGsmtWp
# rEjyastwUPB7iuMgU7W413wyGgU3ISf/WQ/FCyzjpi6BPOsK8oDTc2UQIJXSapka
# 8c52Jg5oh0GnqT6UbYynD/ksJlH7AgvUa9KdqIovjwbCPjmSUKxdU+Qamhu+al/Q
# vHoxlQDcJ7a0+CwbMlhid0QMQb2nlvSADHPioTmUQU/nKkUPA61AKIuPR+vQGNQq
# UyilribHWZmaePNcSlS0rYLWwRi2qsC9PiIWz0ZTEeMGfhPl6v6QsfUu6PBiHSDb
# a8N35fJdm3e8UCPDMmpb0RQY7TthRDQZsExhvQttB9sEOY8=
# SIG # End signature block
