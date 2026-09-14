<#
.SYNOPSIS
Collects SharePoint site usage and Microsoft Teams membership for SmartWorkplaceCMDB.

.DESCRIPTION
Uses Microsoft Graph application authentication. SharePoint site population and
activity come from the D180 usage report. Teams are enumerated as Microsoft 365
groups provisioned as teams; each team's members and owners are read explicitly.
No SharePoint member count is inferred because the usage report does not contain
that grain. Offline JSON is supported for deterministic tests.

.VERSION
1.1.1
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

$ScriptVersion='1.1.1'
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0

function Get-ConfigSection { param([Collections.IDictionary]$Configuration,[string]$Name) if($Configuration.Contains($Name)-and$Configuration[$Name]-is[Collections.IDictionary]){return $Configuration[$Name]};return [ordered]@{} }
function Get-ConfigText { param([Collections.IDictionary]$Configuration,[string]$Name) if($Configuration.Contains($Name)){return ([string]$Configuration[$Name]).Trim()};return '' }
function Get-Text { param([AllowNull()]$Value) if($null-eq$Value){return ''};return ([string]$Value-replace"`r`n|`n|`r",' ').Trim() }
function Get-Value { param([AllowNull()]$Object,[string[]]$Names) foreach($name in $Names){if($null-eq$Object){continue};$value=if($Object-is[Collections.IDictionary]){if($Object.Contains($name)){$Object[$name]}else{$null}}else{$property=$Object.PSObject.Properties[$name];if($property){$property.Value}else{$null}};if($null-ne$value){return $value}};return $null }
function Get-IntegerText { param([AllowNull()]$Value,[string]$Field,[string]$Key) if($null-eq$Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return ''};$number=[int64]0;if(-not[int64]::TryParse([string]$Value,[Globalization.NumberStyles]::Integer,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)-or$number-lt 0){throw "$Field '$Value' is invalid for '$Key'."};return $number }
function Get-DateText { param([AllowNull()]$Value,[string]$Field,[string]$Key) if($null-eq$Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return ''};$date=[datetimeoffset]::MinValue;if(-not[datetimeoffset]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$date)){throw "$Field '$Value' is invalid for '$Key'."};return $date.ToUniversalTime().ToString('o') }
function Get-BoolText { param([AllowNull()]$Value) if($null-eq$Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return ''};$parsed=$false;if([bool]::TryParse([string]$Value,[ref]$parsed)){return $parsed};if([string]$Value-eq'0'){return $false};if([string]$Value-eq'1'){return $true};throw "Boolean value '$Value' is invalid." }

function Get-ConnectedGraphCollection {
    param([Parameter(Mandatory)][string]$Uri,[int]$Limit=0)
    $items=New-Object System.Collections.Generic.List[object]
    $next=$Uri
    while(-not[string]::IsNullOrWhiteSpace($next)){
        $response=Invoke-SmartWorkplaceCMDBGraphRequestWithRetry -Uri $next
        Assert-SmartWorkplaceCMDBCollectionPage -Response $response
        foreach($item in @(Get-SmartWorkplaceCMDBGraphObjectValue $response 'value')){if($null-ne$item){$items.Add($item);if($Limit-gt 0-and$items.Count-ge$Limit){break}}}
        if($Limit-gt 0-and$items.Count-ge$Limit){break}
        $next=[string](Get-SmartWorkplaceCMDBGraphObjectValue $response '@odata.nextLink')
    }
    return @($items.ToArray())
}

function Get-ConnectedGraphReport {
    param([Parameter(Mandatory)][string]$Endpoint,[Parameter(Mandatory)][string]$TemporaryFolder)
    $path=Join-Path $TemporaryFolder ($Endpoint+'.csv')
    $uri="https://graph.microsoft.com/v1.0/reports/$Endpoint(period='D180')"
    for($attempt=1;$attempt -le 4;$attempt++){
        try{Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $path -ErrorAction Stop|Out-Null;return @(Import-Csv -LiteralPath $path)}
        catch{if($attempt -ge 4){throw};$status=Get-SmartWorkplaceCMDBGraphRetryStatusCode $_;if($status -notin @(408,429,500,502,503,504)){throw};Start-Sleep -Seconds (Get-SmartWorkplaceCMDBGraphRetryDelaySeconds $_ $attempt)}
    }
}

$scriptRoot=Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot=Split-Path -Parent (Split-Path -Parent $scriptRoot)
$core=Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$graph=Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Graph\SmartWorkplaceCMDB.Graph.psd1'
$contractPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
Import-Module $core -Force
Import-Module $graph -Force
$bound=@{};foreach($key in $PSBoundParameters.Keys){$bound[$key]=$PSBoundParameters[$key]}
$context=Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly-or$NoConfigWrite-or$PSCmdlet.ParameterSetName-eq'Fixture')
$paths=Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths -Fixture:($PSCmdlet.ParameterSetName-eq'Fixture') -MaxItems $MaxItems -ExplicitDataRoot:([bool]$DataRootPath) -NoWrite:$ValidateOnly
$contract=Get-SmartWorkplaceCMDBTableContract -Path $contractPath
$tableNames=@('M365_SharePointSites.csv','M365_Teams.csv','M365_TeamMembers.csv')
$tables=@{};$latest=@{}
foreach($name in $tableNames){$match=@($contract.tables|Where-Object { $_.name -eq $name });if($match.Count -ne 1){throw "Raw contract definition missing or duplicated: $name"};$tables[$name]=$match[0];$latest[$name]=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$match[0].area) $name)))}
$mode=if($ValidateOnly){'Validate'}elseif($PSCmdlet.ParameterSetName-eq'Fixture'){'Fixture'}else{'Collect'}
$runtime=Start-SmartWorkplaceCMDBExecutionContext -Context $context -ScriptPath $PSCommandPath -ScriptVersion $ScriptVersion -Mode $mode -NoWrite:$ValidateOnly
$executionError=$null;$connected=$false;$temporaryFolder=$null
try{
    $configuration=Get-ConfigSection $context.Configuration 'MicrosoftGraph'
    $clientId=Get-ConfigText $configuration 'ClientId';$thumbprint=Get-ConfigText $configuration 'CertificateThumbprint'
    $fixture=$null;$readiness=$null
    if($PSCmdlet.ParameterSetName-eq'Fixture'){$InputJsonPath=[IO.Path]::GetFullPath($InputJsonPath);$fixture=Get-Content -Raw -LiteralPath $InputJsonPath|ConvertFrom-Json}
    else{$readiness=Test-SmartWorkplaceCMDBGraphAppOnlyReadiness -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint}
    if($ValidateOnly){[pscustomobject]@{Status='Valid';ScriptVersion=$ScriptVersion;SourceMode=if($fixture){'OfflineJson'}else{'MicrosoftGraphAppOnly'};RequiredGraphPermissions='Reports.Read.All;Group.Read.All;TeamMember.Read.All;Team.ReadBasic.All';RawContractVersion=[string]$contract.contractVersion;OutputCount=$tableNames.Count}|Format-List;return}

    $siteSource=@();$teamSource=@()
    if($fixture){$siteSource=@($fixture.sharePointSites);$teamSource=@($fixture.teams)}else{
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Connect-MgGraph -TenantId $readiness.TenantId -ClientId $readiness.ClientId -CertificateThumbprint $readiness.CertificateThumbprint -ContextScope Process -NoWelcome -ErrorAction Stop|Out-Null;$connected=$true
        $mgContext=Get-MgContext -ErrorAction Stop;if([string]$mgContext.TenantId-ne$readiness.TenantId){throw 'Microsoft Graph connected to an unexpected tenant.'}
        $temporaryFolder=Join-Path ([IO.Path]::GetTempPath()) ('SmartWorkplaceCMDB-Collaboration-'+[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $temporaryFolder -Force|Out-Null
        $siteSource=@(Get-ConnectedGraphReport -Endpoint 'getSharePointSiteUsageDetail' -TemporaryFolder $temporaryFolder)
        $activityRows=@(Get-ConnectedGraphReport -Endpoint 'getTeamsTeamActivityDetail' -TemporaryFolder $temporaryFolder)
        $activityById=@{};foreach($row in $activityRows){$id=Get-Text(Get-Value $row @('Team Id','Team ID','TeamId'));if($id){$activityById[$id.ToLowerInvariant()]=$row}}
        $filter=[uri]::EscapeDataString("resourceProvisioningOptions/Any(x:x eq 'Team')")
        $groups=@(Get-ConnectedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName,visibility,createdDateTime&`$top=999" -Limit $MaxItems)
        $teamList=New-Object System.Collections.Generic.List[object]
        foreach($group in $groups){
            $teamId=Get-Text(Get-Value $group @('id'));if(-not$teamId){throw 'A Teams group is missing id.'}
            $members=@(Get-ConnectedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$teamId/members?`$select=id,displayName,userPrincipalName,userType&`$top=999")
            $owners=@(Get-ConnectedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$teamId/owners?`$select=id&`$top=999")
            $teamDetail=Invoke-SmartWorkplaceCMDBGraphRequestWithRetry -Uri "https://graph.microsoft.com/v1.0/teams/${teamId}?`$select=isArchived"
            $ownerIds=@{};foreach($owner in $owners){$ownerId=Get-Text(Get-Value $owner @('id'));if($ownerId){$ownerIds[$ownerId.ToLowerInvariant()]=$true}}
            $activity=if($activityById.ContainsKey($teamId.ToLowerInvariant())){$activityById[$teamId.ToLowerInvariant()]}else{$null}
            $teamList.Add([pscustomobject]@{id=$teamId;displayName=Get-Value $group @('displayName');visibility=Get-Value $group @('visibility');createdDateTime=Get-Value $group @('createdDateTime');lastActivityDate=Get-Value $activity @('Last Activity Date');isArchived=Get-Value $teamDetail @('isArchived');owners=$owners;members=$members;ownerIds=$ownerIds})
        }
        $teamSource=@($teamList.ToArray())
    }
    if($MaxItems-gt0){$siteSource=@($siteSource|Select-Object -First $MaxItems);$teamSource=@($teamSource|Select-Object -First $MaxItems)}

    $collected=[datetime]::UtcNow.ToString('o')
    $siteRows=New-Object System.Collections.Generic.List[object]
    foreach($site in $siteSource){
        $siteId=Get-Text(Get-Value $site @('siteId','Site Id','Site ID'));$url=Get-Text(Get-Value $site @('siteUrl','Site URL','Site Url'));$template=Get-Text(Get-Value $site @('rootWebTemplate','Root Web Template'))
        if(-not$siteId){throw 'A SharePoint usage-report row is missing Site Id.'}
        if($template-eq'SPSPERS'-or$url-match'-my\.sharepoint\.com/personal/'){continue}
        $siteRows.Add([pscustomobject][ordered]@{SourceSystem='Microsoft365Reports';SiteId=$siteId;SiteUrl=$url;SiteName=Get-Text(Get-Value $site @('siteName','Site Name','displayName'));OwnerPrincipalName=Get-Text(Get-Value $site @('ownerPrincipalName','Owner Principal Name','Owner Display Name','Owner'));LastActivityDate=Get-DateText (Get-Value $site @('lastActivityDate','Last Activity Date')) 'LastActivityDate' $siteId;StorageUsedBytes=Get-IntegerText (Get-Value $site @('storageUsedBytes','Storage Used (Byte)','Storage Used (Bytes)')) 'StorageUsedBytes' $siteId;StorageAllocatedBytes=Get-IntegerText (Get-Value $site @('storageAllocatedBytes','Storage Allocated (Byte)','Storage Allocated (Bytes)')) 'StorageAllocatedBytes' $siteId;RootWebTemplate=$template;IsDeleted=Get-BoolText(Get-Value $site @('isDeleted','Is Deleted'));SourceCollectedDateTime=$collected})
    }
    $teamRows=New-Object System.Collections.Generic.List[object];$memberRows=New-Object System.Collections.Generic.List[object]
    foreach($team in $teamSource){
        $teamId=Get-Text(Get-Value $team @('id','teamId','TeamId'));if(-not$teamId){throw 'A team is missing id.'}
        $members=@(Get-Value $team @('members')|Where-Object { $null -ne $_ });$owners=@(Get-Value $team @('owners')|Where-Object { $null -ne $_ })
        $ownerIds=@{};foreach($owner in $owners){$id=Get-Text(Get-Value $owner @('id','userId'));if($id){$ownerIds[$id.ToLowerInvariant()]=$true}}
        $seen=@{};$guestCount=0;$unresolvedMemberCount=0
        foreach($member in $members){$userId=Get-Text(Get-Value $member @('id','userId'));if(-not$userId){$unresolvedMemberCount++;continue};$key=($teamId+'|'+$userId).ToLowerInvariant();if($seen.ContainsKey($key)){continue};$seen[$key]=$true;$userType=Get-Text(Get-Value $member @('userType'));if($userType-eq'Guest'){$guestCount++};$memberRows.Add([pscustomobject][ordered]@{SourceSystem='MicrosoftTeams';RelationshipKey=$key;TeamId=$teamId;UserId=$userId;UserPrincipalName=Get-Text(Get-Value $member @('userPrincipalName'));UserType=$userType;Role=if($ownerIds.ContainsKey($userId.ToLowerInvariant())){'Owner'}else{'Member'};SourceCollectedDateTime=$collected})}
        $teamRows.Add([pscustomobject][ordered]@{SourceSystem='MicrosoftTeams';TeamId=$teamId;DisplayName=Get-Text(Get-Value $team @('displayName'));Visibility=Get-Text(Get-Value $team @('visibility'));CreatedDateTime=Get-DateText (Get-Value $team @('createdDateTime')) 'CreatedDateTime' $teamId;LastActivityDate=Get-DateText (Get-Value $team @('lastActivityDate')) 'LastActivityDate' $teamId;OwnerCount=$owners.Count;MemberCount=$seen.Count;GuestCount=$guestCount;UnresolvedMemberCount=$unresolvedMemberCount;MembershipCoverageStatus=if($unresolvedMemberCount-eq 0){'Complete'}else{'Partial'};IsArchived=Get-BoolText(Get-Value $team @('isArchived'));SourceCollectedDateTime=$collected})
    }
    $partialTeams=@($teamRows.ToArray()|Where-Object { $_.MembershipCoverageStatus -eq 'Partial' })
    if($partialTeams.Count -gt 0){Write-Warning ("Microsoft Graph returned {0} Team(s) with {1} member object(s) lacking an immutable id. Exact links were preserved; affected Team member counts are marked Partial."-f$partialTeams.Count,(($partialTeams|Measure-Object UnresolvedMemberCount -Sum).Sum))}
    $rows=@{'M365_SharePointSites.csv'=@($siteRows.ToArray());'M365_Teams.csv'=@($teamRows.ToArray());'M365_TeamMembers.csv'=@($memberRows.ToArray())}
    foreach($name in $tableNames){$key=if($name -eq 'M365_SharePointSites.csv'){'SiteId'}elseif($name -eq 'M365_Teams.csv'){'TeamId'}else{'RelationshipKey'};$duplicates=@($rows[$name]|Group-Object $key|Where-Object { $_.Count -gt 1 });if($duplicates.Count){throw "Duplicate keys returned for ${name}: $($duplicates.Name -join ', ')"}}
    $published=@();foreach($name in $tableNames){$run=$null;try{$run=Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($latest[$name]) -Fixture:($PSCmdlet.ParameterSetName-eq'Fixture') -MaxItems $MaxItems;$stamp=[datetime]::UtcNow;$base=[IO.Path]::GetFileNameWithoutExtension($name);$history=Join-Path $paths.DataAllRootPath ('M365\Collaboration\{0}\{1}\{2}_{3}.csv'-f$stamp.ToString('yyyy'),$stamp.ToString('MM'),$base,$stamp.ToString('yyyyMMdd-HHmmssfff'));Publish-SmartWorkplaceCMDBSourceCsv -Run $run -InputObject @($rows[$name]) -Columns @($tables[$name].columns|ForEach-Object{[string]$_}) -HistoryPath $history -LatestPath $latest[$name] -ContractPath $contractPath -ContractTableName $name|Out-Null;$published+=$latest[$name]}catch{if($run){Complete-SmartWorkplaceCMDBSourceCollection -Run $run -Failed};throw}}
    Write-Information ("SmartWorkplaceCMDB collaboration collection completed. SharePointSites={0}; Teams={1}; TeamMembers={2}."-f$siteRows.Count,$teamRows.Count,$memberRows.Count) -InformationAction Continue
    [pscustomobject]@{Status='Completed';ScriptVersion=$ScriptVersion;SharePointSiteCount=$siteRows.Count;TeamCount=$teamRows.Count;TeamMemberCount=$memberRows.Count;PublishedPath=$published}
}catch{$executionError=$_;throw}finally{if($connected){Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null};if($temporaryFolder-and(Test-Path -LiteralPath $temporaryFolder)-and(Split-Path -Leaf $temporaryFolder)-like'SmartWorkplaceCMDB-Collaboration-*'){Remove-Item -LiteralPath $temporaryFolder -Recurse -Force -ErrorAction SilentlyContinue};Complete-SmartWorkplaceCMDBExecutionContext -RuntimeContext $runtime -ErrorRecord $executionError}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCYewO4ohjuEhad
# WCOlcCMUZi5h8wcpZo5bxFMvaiIU6aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFgxMWeevrA/rp672siVsUyNYk185hDUhsD52WyeU31QMA0GCSqG
# SIb3DQEBAQUABIIBgC11X8lj9fJHpayjw2HzNMa+dKHYMJUxrTgjf9Z9KCjo85U0
# 8gYHQUC8nz+rS1dIF/Otcf3I6rjsl8U1Vb4Vv3tC1/UV8PeMqWxe8lyMzi7Yuiaa
# vbbNpqqqOTsKFg+H3mzNUijz5bHrUEwrUTdn/BOCs3rWNPqMS0TbUaqCy1ennVdA
# mV2jFtG1/HqvLy8mKHN0bZ3avmHE1D4oYBQx9zHw7l7vTTvJCHAP002HgDlh0CmY
# Gi+GJCZ6v9UoNrZ4TkmBAYPzXToDdnkSrx07UZFXZhY7U+xHFtP7wuqAViY1dWt7
# Ph0iExypALQHXppDdhAAw6Zx8bBxjxmvzzAL3eBPDzZ39zTlDvU4SZNPeRMp6WSP
# BZNKNrYm4S/JNAQocoZjT9asaCsgeiqj0sJOCFTDF5U9Aesc+5jhWsjCzv/RwPtm
# hWhBe4BkRbYSKFpswX1tDhTa2ZPGsOMHzc8zzqnKcEllCdsBqwRMVRpqi/jBygeW
# RXe3kL+JLQKRN3sxvaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQwNDI0
# NTVaMC8GCSqGSIb3DQEJBDEiBCDUYXtDhokvbCDbWo9YlC9v0DwaNggXNMKgul8M
# mpszTDANBgkqhkiG9w0BAQEFAASCAgCeTQkXkkMtDjP4Z8mt1Kp7VZPt8HNK5Wfn
# f8mX+QBKHc51sZgdNIsseFNLjGGqY3E2Mf9CYy+YTE/O9pAX6/yybMMMrFdSoC7F
# 8SUy7UNqItdbekCut3/Nma3PJkS9W1/Ohsl5gqUf+X5DBMqEwjZovFC864eInBm1
# +Ejh2ofU0aIGhO5G1hBfNrtY3Os6UrWCH11RkeHxz3VIKurDNa7H185KuTcfRWpc
# QLfvzGZ3Bixl56ABt/gJ7rzLiLi8ihJcGBCO5gACffcpFHKAEsEtU3HM9j9UmNnS
# 2XK1GlRNooCMg4V/hrjbkfXNWtbi3pVqa5Lm4GWVpKHRmv1TH3qgIMtbUys6LdND
# WYIpidwPvHyAYcAIZle0fMTYvjmQMHDfzJPJVAIeec3f1cu8haZgoGgohLcXv6Bg
# R6TmhMFSEQWaPHbPpOGiQEAQZWkQj0DtY2kLDBo8/jh0wBLGpW5KOJpe9L/XTgNJ
# FIhl20t0S+x7S07L3VFskLna0pt8k2x8z/uTsu7h+QSdLsOlwHxGsHvLiH/VptTO
# RYNl6HgvbpY5/SeZ6WNq4IpqMPA4KlmG/CFG6IpaD5S4G5aUhxt16bQ5/mLeixUQ
# 3HL1MV/PRllM3mfTlsBo8xM38gYLkn3ZcE3sWUtlYC2Yno0lO3jgehkQoFPSUTvs
# p4a7ka2vmQ==
# SIG # End signature block
