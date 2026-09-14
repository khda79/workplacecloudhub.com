<#
.SYNOPSIS
Collects read-only Intune operational inventory for SmartWorkplaceCMDB.

.DESCRIPTION
Publishes separate raw snapshots for Autopilot devices, detected applications,
configuration policies, and Windows feature/quality update policies. The
collector uses explicit Microsoft Graph fields and supports one synthetic JSON
fixture containing the four source families.

.VERSION
1.3.1
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

$ScriptVersion='1.3.1'
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$script:ExactApplicationIdCache=[Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)

function Get-ConfigSection { param([System.Collections.IDictionary]$Configuration,[string]$Name) if($Configuration.Contains($Name) -and $Configuration[$Name] -is [System.Collections.IDictionary]){return $Configuration[$Name]};return [ordered]@{} }
function Get-ConfigText { param([System.Collections.IDictionary]$Configuration,[string]$Name) if($Configuration.Contains($Name)){return ([string]$Configuration[$Name]).Trim()};return '' }
function Get-Value { param([AllowNull()]$Object,[string]$Name) return Get-SmartWorkplaceCMDBGraphObjectValue $Object $Name }
function Get-Text { param([AllowNull()]$Value) if($null-eq $Value){return ''};return ([string]$Value-replace"`r`n|`n|`r",' ').Trim() }
function Get-DateText { param([AllowNull()]$Value,[string]$Field,[string]$Key) if($null-eq $Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return ''};$date=[datetimeoffset]::MinValue;if(-not[datetimeoffset]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$date)){throw "$Field '$Value' is invalid for '$Key'."};return $date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ',[Globalization.CultureInfo]::InvariantCulture) }
function Get-IntegerText { param([AllowNull()]$Value,[string]$Field,[string]$Key) if($null-eq $Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return ''};$number=0;if(-not[int]::TryParse([string]$Value,[ref]$number)-or$number-lt 0){throw "$Field '$Value' is invalid for '$Key'."};return $number }
function Get-ListText { param([AllowNull()]$Value) return (@(@($Value)|ForEach-Object{Get-Text $_}|Where-Object{-not[string]::IsNullOrWhiteSpace($_)}|Sort-Object -Unique)-join';') }
function Get-PreferredText { param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,[Parameter(Mandatory)][string]$Field) $values=@($Rows|ForEach-Object{[string]$_.$Field}|Where-Object{-not[string]::IsNullOrWhiteSpace($_)});if($values.Count-eq 0){return ''};$ranked=@($values|Group-Object{$_.ToLowerInvariant()}|Sort-Object @{Expression='Count';Descending=$true},@{Expression='Name';Ascending=$true});return [string]$ranked[0].Group[0] }
function ConvertTo-CsvField { param([AllowNull()]$Value) return '"'+(([string]$Value)-replace'"','""')+'"' }
function Get-ExactApplicationId {
    param(
        [Parameter(Mandatory)][string]$SourceApplicationKey,
        [AllowEmptyString()][string]$DisplayName,
        [AllowEmptyString()][string]$Publisher,
        [AllowEmptyString()][string]$Version,
        [AllowEmptyString()][string]$Platform
    )
    $parts=@($SourceApplicationKey,$DisplayName,$Publisher,$Version,$Platform)|ForEach-Object{Get-Text $_}
    if([string]::IsNullOrWhiteSpace($parts[0])){throw 'An exact detected-application identity cannot be built without SourceApplicationKey.'}
    $payload=($parts|ForEach-Object{('{0}:{1}'-f$_.Length,$_ )})-join'|'
    if($script:ExactApplicationIdCache.ContainsKey($payload)){return [string]$script:ExactApplicationIdCache[$payload]}
    $algorithm=[Security.Cryptography.SHA256]::Create()
    try{$result=([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload)))).Replace('-','').ToLowerInvariant();$script:ExactApplicationIdCache[$payload]=$result;return $result}
    finally{$algorithm.Dispose()}
}
function Get-BatchRetryDelaySeconds { param([AllowNull()]$Response,[int]$Attempt) $seconds=0;$retryAfter='';try{$headers=Get-Value $Response 'headers';if($headers -is [Collections.IDictionary]){foreach($key in $headers.Keys){if([string]$key-ieq'Retry-After'){$retryAfter=[string]$headers[$key];break}}}elseif($headers){$property=$headers.PSObject.Properties['Retry-After'];if($property){$retryAfter=[string]$property.Value}}}catch{};if([int]::TryParse($retryAfter,[ref]$seconds)-and$seconds-gt 0){return [Math]::Min($seconds,300)};$exponent=[Math]::Max(0,([int]$Attempt-1));$backoff=15.0*[Math]::Pow(2.0,[double]$exponent);return [int][Math]::Min($backoff,300.0) }
function Get-AppDeviceRelationBatch {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Apps)
    $result=@{};$pending=@($Apps);$attempt=0
    while($pending.Count-gt 0){
        $attempt++
        if($attempt-gt 5){throw "Intune detected-application batch did not complete after five attempts for $($pending.Count) application(s). The last-valid snapshot was preserved."}
        $requests=New-Object System.Collections.Generic.List[object];$requestApps=@{};$requestId=1
        foreach($app in $pending){$id=[string]$requestId;$appId=Get-Text(Get-Value $app 'id');$requestApps[$id]=$app;$requests.Add([ordered]@{id=$id;method='GET';url="/deviceManagement/detectedApps/$appId/managedDevices?`$select=id&`$top=999"});$requestId++}
        $body=@{requests=$requests}|ConvertTo-Json -Depth 6 -Compress
        $response=Invoke-SmartWorkplaceCMDBGraphRequestWithRetry -Uri 'https://graph.microsoft.com/v1.0/$batch' -MaximumRetryCount 5 -BaseDelaySeconds 15 -MaximumDelaySeconds 300 -RequestScript {param($uri)Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json' -ErrorAction Stop}
        $responses=@{};foreach($item in @(Get-Value $response 'responses')){$responses[(Get-Text(Get-Value $item 'id'))]=$item}
        $retry=New-Object System.Collections.Generic.List[object];$retryDelay=0
        foreach($id in $requestApps.Keys){
            $app=$requestApps[$id];$appId=Get-Text(Get-Value $app 'id');$item=$responses[$id]
            if($null-eq$item){$retry.Add($app);$retryDelay=[Math]::Max($retryDelay,(Get-BatchRetryDelaySeconds -Response $null -Attempt $attempt));continue}
            $status=[int](Get-Value $item 'status')
            if($status-eq 200){
                $devices=New-Object System.Collections.Generic.List[object];$itemBody=Get-Value $item 'body';foreach($device in @(Get-Value $itemBody 'value')){if($null-ne$device){$devices.Add($device)}}
                $next=Get-Text(Get-Value $itemBody '@odata.nextLink');while(-not[string]::IsNullOrWhiteSpace($next)){$page=Invoke-SmartWorkplaceCMDBGraphRequestWithRetry -Uri $next -MaximumRetryCount 8 -BaseDelaySeconds 15 -MaximumDelaySeconds 300;Assert-SmartWorkplaceCMDBCollectionPage -Response $page;foreach($device in @(Get-Value $page 'value')){if($null-ne$device){$devices.Add($device)}};$next=Get-Text(Get-Value $page '@odata.nextLink')}
                $result[$appId]=$devices.ToArray();continue
            }
            if($status-in@(408,429,500,502,503,504)){$retry.Add($app);$retryDelay=[Math]::Max($retryDelay,(Get-BatchRetryDelaySeconds -Response $item -Attempt $attempt));continue}
            throw "Intune detected-application batch failed for AppId '$appId' with HTTP $status. The last-valid snapshot was preserved."
        }
        $pending=$retry.ToArray();if($pending.Count-gt 0){Write-Warning("Microsoft Graph throttled or deferred {0} detected-application request(s). Batch retry {1}/5 in {2} second(s)."-f$pending.Count,$attempt,$retryDelay);Start-Sleep -Seconds $retryDelay}
    }
    return $result
}

function Get-AppInventoryRawExport {
    param([Parameter(Mandatory)][string]$OutputFolder)

    $body = @{
        reportName = 'AppInvRawData'
        filter = ''
        select = @('ApplicationKey','ApplicationName','ApplicationPublisher','ApplicationVersion','DeviceId','DeviceName','Platform')
        format = 'csv'
        snapshotId = ''
    } | ConvertTo-Json -Depth 5 -Compress
    $job = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
        -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs' `
        -MaximumRetryCount 5 -BaseDelaySeconds 15 -MaximumDelaySeconds 300 `
        -RequestScript { param($uri) Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json' -ErrorAction Stop }
    $jobId = Get-Text (Get-Value $job 'id')
    if ([string]::IsNullOrWhiteSpace($jobId)) { throw 'Intune AppInvRawData export did not return a job id.' }
    Write-Information ("Intune AppInvRawData export requested. JobId={0}." -f $jobId) -InformationAction Continue

    $deadline = [datetimeoffset]::UtcNow.AddMinutes(30)
    do {
        Start-Sleep -Seconds 5
        $job = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
            -Uri ("https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs/{0}" -f $jobId) `
            -MaximumRetryCount 8 -BaseDelaySeconds 5 -MaximumDelaySeconds 120
        $status = Get-Text (Get-Value $job 'status')
        Write-Information ("Intune AppInvRawData export status: {0}." -f $status) -InformationAction Continue
        if ($status -eq 'failed') { throw "Intune AppInvRawData export '$jobId' failed." }
    } while ($status -ne 'completed' -and [datetimeoffset]::UtcNow -lt $deadline)
    if ($status -ne 'completed') { throw "Intune AppInvRawData export '$jobId' did not complete within 30 minutes." }

    $downloadUrl = Get-Text (Get-Value $job 'url')
    if ([string]::IsNullOrWhiteSpace($downloadUrl)) { throw "Completed Intune AppInvRawData export '$jobId' did not return a download URL." }
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    $zipPath = Join-Path $OutputFolder ($jobId + '.zip')
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -ErrorAction Stop
    Expand-Archive -LiteralPath $zipPath -DestinationPath $OutputFolder -Force
    $csvFiles = @(Get-ChildItem -LiteralPath $OutputFolder -File -Filter '*.csv')
    if ($csvFiles.Count -ne 1) { throw "Intune AppInvRawData export '$jobId' produced $($csvFiles.Count) CSV files; exactly one is required." }
    $expected = @('ApplicationKey','ApplicationName','ApplicationPublisher','ApplicationVersion','DeviceId','DeviceName','Platform')
    $header = Get-Content -LiteralPath $csvFiles[0].FullName -TotalCount 1 -ErrorAction Stop
    $actual = @($header.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
    if (($actual -join [char]31) -cne ($expected -join [char]31)) { throw "Intune AppInvRawData export '$jobId' returned an unexpected CSV schema." }
    return $csvFiles[0].FullName
}

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
$tableNames=@('Intune_AutopilotDevices.csv','Intune_DetectedApps.csv','Intune_DetectedAppDeviceRelationships.csv','Intune_ConfigurationPolicies.csv','Intune_WindowsUpdatePolicies.csv')
$tables=@{};$latest=@{}
foreach($name in $tableNames){$matches=@($contract.tables | Where-Object name -eq $name);if($matches.Count -ne 1){throw "Raw contract definition missing or duplicated: $name"};$tables[$name]=$matches[0];$latest[$name]=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$matches[0].area) $name)))}

$mode=if($ValidateOnly){'Validate'}elseif($PSCmdlet.ParameterSetName -eq 'Fixture'){'Fixture'}else{'Collect'}
$runtime=Start-SmartWorkplaceCMDBExecutionContext -Context $context -ScriptPath $PSCommandPath -ScriptVersion $ScriptVersion -Mode $mode -NoWrite:$ValidateOnly
$executionError=$null
try{
    $configuration=Get-ConfigSection $context.Configuration 'MicrosoftGraph'
    $clientId=Get-ConfigText $configuration 'ClientId';$thumbprint=Get-ConfigText $configuration 'CertificateThumbprint'
    $fixture=$null;$readiness=$null;$preparedAppDeviceCsvPath='';$appDeviceRelationCount=0;$relationshipCountMismatchCount=0;$exactDeviceCounts=@{};$collected=[datetime]::UtcNow.ToString('o')
    if($PSCmdlet.ParameterSetName -eq 'Fixture'){$InputJsonPath=[IO.Path]::GetFullPath($InputJsonPath);$fixture=Get-Content -Raw -LiteralPath $InputJsonPath|ConvertFrom-Json}
    else{$readiness=Test-SmartWorkplaceCMDBGraphAppOnlyReadiness -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint}
    if($ValidateOnly){
        [pscustomobject]@{Status='Valid';ScriptVersion=$ScriptVersion;SourceMode=if($fixture){'OfflineJson'}else{'MicrosoftGraphAppOnly'};RawContractVersion=[string]$contract.contractVersion;RequiredGraphPermissions='DeviceManagementServiceConfig.Read.All;DeviceManagementApps.Read.All;DeviceManagementApps.ReadWrite.All;DeviceManagementManagedDevices.Read.All;DeviceManagementConfiguration.Read.All';OutputCount=$tableNames.Count}|Format-List
        return
    }
    if($fixture){
        $autopilot=@($fixture.autopilotDevices);$apps=@($fixture.detectedApps);$appDeviceRelations=@($fixture.detectedAppDeviceRelationships);$policies=@($fixture.configurationPolicies);$feature=@($fixture.featureUpdatePolicies);$quality=@($fixture.qualityUpdatePolicies)
        foreach($app in $apps){$exactDeviceCounts[(Get-Text(Get-Value $app 'id'))]=0};foreach($group in @($appDeviceRelations|Group-Object{Get-Text(Get-Value $_ 'appId')})){$exactDeviceCounts[[string]$group.Name]=@($group.Group|ForEach-Object{Get-Text(Get-Value $_ 'managedDeviceId')}|Where-Object{$_}|Sort-Object -Unique).Count}
    }else{
        $autopilot=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=999' -RequiredPermission 'DeviceManagementServiceConfig.Read.All' -MaxItems $MaxItems)
        # A complete collection is sourced from AppInvRawData below because it carries
        # the exact application-device grain. Enumerating detectedApps first would be
        # redundant, would be discarded, and is heavily throttled on large tenants.
        $apps=@()
        if($MaxItems-gt0){
            $apps=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$select=id,displayName,version,publisher,deviceCount,platform&$top=999' -RequiredPermission 'DeviceManagementApps.Read.All' -MaxItems $MaxItems)
        }
        $preparedFolder=Join-Path ([IO.Path]::GetTempPath()) ('SmartWorkplaceCMDB\Prepared\{0}'-f[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $preparedFolder -Force|Out-Null
        $preparedAppDeviceCsvPath=Join-Path $preparedFolder 'Intune_DetectedAppDeviceRelationships.csv'
        $relationColumns=@($tables['Intune_DetectedAppDeviceRelationships.csv'].columns|ForEach-Object{[string]$_})
        $writer=[IO.StreamWriter]::new($preparedAppDeviceCsvPath,$false,[Text.UTF8Encoding]::new($true))
        $appReadiness=Test-SmartWorkplaceCMDBGraphAppOnlyReadiness -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $appConnected=$false
        try{
            Connect-MgGraph -TenantId $appReadiness.TenantId -ClientId $appReadiness.ClientId -CertificateThumbprint $appReadiness.CertificateThumbprint -ContextScope Process -NoWelcome -ErrorAction Stop|Out-Null
            $appConnected=$true
            $writer.WriteLine((@($relationColumns|ForEach-Object{ConvertTo-CsvField $_})-join','))
            if($MaxItems-eq 0){
                $exportFolder=Join-Path $preparedFolder 'AppInvRawData';$exportCsv=''
                for($exportAttempt=1;$exportAttempt-le2;$exportAttempt++){
                    try{$exportCsv=Get-AppInventoryRawExport -OutputFolder $exportFolder;break}
                    catch{
                        if($exportAttempt-ge2-or[string]$_.Exception.Message-notmatch'(?i)\bService\s*Unavailable\b'){throw}
                        Write-Warning 'The Intune export job remained unavailable after bounded request retries. One fresh export job will be requested in 60 seconds.'
                        if(Test-Path -LiteralPath $exportFolder){Remove-Item -LiteralPath $exportFolder -Recurse -Force}
                        Start-Sleep -Seconds 60
                    }
                }
                if([string]::IsNullOrWhiteSpace($exportCsv)){throw 'Intune AppInvRawData export did not produce a reusable CSV.'}
                $applicationStats=@{};$pairKeys=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);$duplicatePairCount=0
                Import-Csv -LiteralPath $exportCsv | ForEach-Object {
                    $sourceApplicationKey=Get-Text $_.ApplicationKey;$managedDeviceId=Get-Text $_.DeviceId
                    if(-not$sourceApplicationKey-or-not$managedDeviceId){throw 'Intune AppInvRawData returned a row without ApplicationKey or DeviceId.'}
                    $displayName=Get-Text $_.ApplicationName;$publisher=Get-Text $_.ApplicationPublisher;$version=Get-Text $_.ApplicationVersion;$platform=Get-Text $_.Platform
                    $appId=Get-ExactApplicationId -SourceApplicationKey $sourceApplicationKey -DisplayName $displayName -Publisher $publisher -Version $version -Platform $platform
                    $pairKey=($appId+'|'+$managedDeviceId).ToLowerInvariant();if(-not$pairKeys.Add($pairKey)){$duplicatePairCount++;return}
                    if(-not$applicationStats.ContainsKey($appId)){$applicationStats[$appId]=[pscustomobject]@{id=$appId;sourceApplicationKey=$sourceApplicationKey;displayName=$displayName;publisher=$publisher;version=$version;platform=$platform;deviceCount=0}}
                    $stat=$applicationStats[$appId]
                    if($stat.sourceApplicationKey-cne$sourceApplicationKey-or$stat.displayName-cne$displayName-or$stat.publisher-cne$publisher-or$stat.version-cne$version-or$stat.platform-cne$platform){throw "A SHA-256 collision was detected while deriving exact application identity for source ApplicationKey '$sourceApplicationKey'."}
                    $stat.deviceCount=[int]$stat.deviceCount+1;$appDeviceRelationCount++
                    $values=@($paths.TenantKey,$paths.OrganizationKey,$paths.EnvironmentKey,$paths.TenantId,'MicrosoftIntune',$pairKey,$appId,$managedDeviceId,$collected);$writer.WriteLine((@($values|ForEach-Object{ConvertTo-CsvField $_})-join','))
                    if($appDeviceRelationCount%250000-eq 0){Write-Information ("Intune AppInvRawData processing: {0} exact application-device relations." -f $appDeviceRelationCount) -InformationAction Continue}
                }
                if($duplicatePairCount-gt 0){Write-Warning ("Intune AppInvRawData contained {0} duplicate application-device row(s); exact duplicate pairs were retained once." -f $duplicatePairCount)}
                $apps=[object[]]$applicationStats.Values
                foreach($app in $apps){$exactDeviceCounts[[string]$app.id]=[int]$app.deviceCount}
            }else{
            $relationApps=@($apps|Group-Object{Get-Text(Get-Value $_ 'id')}|ForEach-Object{$_.Group[0]});$batchNumber=0;$batchCount=[Math]::Ceiling($relationApps.Count/20)
            for($offset=0;$offset-lt$relationApps.Count;$offset+=20){
                $batchNumber++;$last=[Math]::Min($offset+19,$relationApps.Count-1);$batch=@($relationApps[$offset..$last]);$batchMap=Get-AppDeviceRelationBatch -Apps $batch
                foreach($app in $batch){
                    $appId=Get-Text(Get-Value $app 'id');$deviceIds=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                    foreach($device in @($batchMap[$appId])){$managedDeviceId=Get-Text(Get-Value $device 'id');if(-not$managedDeviceId){throw "Detected application '$appId' returned a managed device without id."};[void]$deviceIds.Add($managedDeviceId)}
                    $reported=Get-IntegerText(Get-Value $app 'deviceCount') 'deviceCount' $appId
                    $exactDeviceCounts[$appId]=$deviceIds.Count
                    foreach($managedDeviceId in $deviceIds){$relationshipKey=($appId+'|'+$managedDeviceId).ToLowerInvariant();$values=@($paths.TenantKey,$paths.OrganizationKey,$paths.EnvironmentKey,$paths.TenantId,'MicrosoftIntune',$relationshipKey,$appId,$managedDeviceId,$collected);$writer.WriteLine((@($values|ForEach-Object{ConvertTo-CsvField $_})-join','));$appDeviceRelationCount++}
                }
                if($batchNumber-eq 1-or$batchNumber%25-eq 0-or$batchNumber-eq$batchCount){Write-Information ("Intune application-device collection: batch {0}/{1}; applications {2}/{3}; exact relations {4}." -f $batchNumber,$batchCount,($last+1),$relationApps.Count,$appDeviceRelationCount) -InformationAction Continue}
                Start-Sleep -Milliseconds 350
            }
            }
        }finally{if($writer){$writer.Dispose()};if($appConnected){Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null}}
        $appDeviceRelations=@()
        $policies=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name,description,platforms,technologies,templateReference,createdDateTime,lastModifiedDateTime&$top=100' -RequiredPermission 'DeviceManagementConfiguration.Read.All' -MaxItems $MaxItems)
        $feature=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?$top=100' -RequiredPermission 'DeviceManagementConfiguration.Read.All' -MaxItems $MaxItems)
        $quality=@(Invoke-SmartWorkplaceCMDBGraphPagedRequest -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint -Uri 'https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles?$top=100' -RequiredPermission 'DeviceManagementConfiguration.Read.All' -MaxItems $MaxItems)
    }
    if($MaxItems -gt 0){$autopilot=@($autopilot|Select-Object -First $MaxItems);$apps=@($apps|Select-Object -First $MaxItems);$appIds=@{};foreach($app in $apps){$appIds[(Get-Text(Get-Value $app 'id')).ToLowerInvariant()]=$true};$appDeviceRelations=@($appDeviceRelations|Where-Object{$appIds.ContainsKey((Get-Text(Get-Value $_ 'appId')).ToLowerInvariant())});$policies=@($policies|Select-Object -First $MaxItems);$feature=@($feature|Select-Object -First $MaxItems);$quality=@($quality|Select-Object -First $MaxItems)}
    $rows=@{}
    $rows['Intune_AutopilotDevices.csv']=@($autopilot|ForEach-Object{$id=Get-Text(Get-Value $_ 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Autopilot response missing id.'};[pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';AutopilotDeviceId=$id;DisplayName=Get-Text(Get-Value $_ 'displayName');SerialNumber=Get-Text(Get-Value $_ 'serialNumber');Manufacturer=Get-Text(Get-Value $_ 'manufacturer');Model=Get-Text(Get-Value $_ 'model');GroupTag=Get-Text(Get-Value $_ 'groupTag');PurchaseOrderIdentifier=Get-Text(Get-Value $_ 'purchaseOrderIdentifier');EnrollmentState=Get-Text(Get-Value $_ 'enrollmentState');LastContactedDateTime=Get-DateText (Get-Value $_ 'lastContactedDateTime') 'lastContactedDateTime' $id;UserPrincipalName=Get-Text(Get-Value $_ 'userPrincipalName');AzureAdDeviceId=Get-Text(Get-Value $_ 'azureActiveDirectoryDeviceId');ManagedDeviceId=Get-Text(Get-Value $_ 'managedDeviceId');SourceCollectedDateTime=$collected}})
    $rows['Intune_DetectedApps.csv']=@($apps|ForEach-Object{$id=Get-Text(Get-Value $_ 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Detected application response missing id.'};$sourceApplicationKey=Get-Text(Get-Value $_ 'sourceApplicationKey');if([string]::IsNullOrWhiteSpace($sourceApplicationKey)){$sourceApplicationKey=$id};$reported=Get-IntegerText (Get-Value $_ 'deviceCount') 'deviceCount' $id;$hasExact=$exactDeviceCounts.ContainsKey($id);$exact=if($hasExact){[int]$exactDeviceCounts[$id]}else{''};$coverage=if(-not$hasExact){'NotCollected'}elseif([string]::IsNullOrWhiteSpace([string]$reported)-or[int64]$reported-eq[int64]$exact){'Complete'}else{'ReconciledCountMismatch'};[pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';AppId=$id;SourceApplicationKey=$sourceApplicationKey;DisplayName=Get-Text(Get-Value $_ 'displayName');Version=Get-Text(Get-Value $_ 'version');Publisher=Get-Text(Get-Value $_ 'publisher');DeviceCount=if($hasExact){$exact}else{$reported};ReportedDeviceCount=$reported;ExactRelatedDeviceCount=$exact;RelationshipCoverageStatus=$coverage;Platform=Get-Text(Get-Value $_ 'platform');SourceCollectedDateTime=$collected}})
    if([string]::IsNullOrWhiteSpace($preparedAppDeviceCsvPath)){$rows['Intune_DetectedAppDeviceRelationships.csv']=@($appDeviceRelations|ForEach-Object{$appId=Get-Text(Get-Value $_ 'appId');$managedDeviceId=Get-Text(Get-Value $_ 'managedDeviceId');if(-not$appId-or-not$managedDeviceId){throw 'A detected application-device relationship is missing appId or managedDeviceId.'};[pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';RelationshipKey=($appId+'|'+$managedDeviceId).ToLowerInvariant();AppId=$appId;ManagedDeviceId=$managedDeviceId;SourceCollectedDateTime=$collected}});$appDeviceRelationCount=$rows['Intune_DetectedAppDeviceRelationships.csv'].Count}else{$rows['Intune_DetectedAppDeviceRelationships.csv']=@()}
    $detectedAppGroups=@($rows['Intune_DetectedApps.csv']|Group-Object AppId)
    $duplicateDetectedAppGroups=@($detectedAppGroups|Where-Object Count -gt 1)
    if($duplicateDetectedAppGroups.Count -gt 0){
        $conflictingDetectedAppGroupCount=0
        $collapsedApps=New-Object System.Collections.Generic.List[object]
        foreach($group in @($detectedAppGroups|Sort-Object Name)){
            $metadataConflict=$false
            foreach($field in @('DisplayName','Version','Publisher','Platform')){$distinct=@($group.Group|ForEach-Object{([string]$_.$field).ToLowerInvariant()}|Sort-Object -Unique);if($distinct.Count-gt 1){$metadataConflict=$true}}
            $deviceCounts=@($group.Group|ForEach-Object{if(-not[string]::IsNullOrWhiteSpace([string]$_.DeviceCount)){[int]$_.DeviceCount}})
            $reportedDeviceCounts=@($group.Group|ForEach-Object{if(-not[string]::IsNullOrWhiteSpace([string]$_.ReportedDeviceCount)){[int]$_.ReportedDeviceCount}})
            $exactDeviceCountsForGroup=@($group.Group|ForEach-Object{if(-not[string]::IsNullOrWhiteSpace([string]$_.ExactRelatedDeviceCount)){[int]$_.ExactRelatedDeviceCount}})
            if(@($deviceCounts|Sort-Object -Unique).Count-gt 1-or@($reportedDeviceCounts|Sort-Object -Unique).Count-gt 1-or@($exactDeviceCountsForGroup|Sort-Object -Unique).Count-gt 1){$metadataConflict=$true}
            if($metadataConflict){$conflictingDetectedAppGroupCount++}
            $collapsedApps.Add([pscustomobject][ordered]@{
                SourceSystem='MicrosoftIntune';AppId=[string]$group.Name
                SourceApplicationKey=Get-PreferredText -Rows @($group.Group) -Field 'SourceApplicationKey'
                DisplayName=Get-PreferredText -Rows @($group.Group) -Field 'DisplayName'
                Version=Get-PreferredText -Rows @($group.Group) -Field 'Version'
                Publisher=Get-PreferredText -Rows @($group.Group) -Field 'Publisher'
                DeviceCount=if($deviceCounts.Count-gt 0){@($deviceCounts|Measure-Object -Maximum)[0].Maximum}else{''}
                ReportedDeviceCount=if($reportedDeviceCounts.Count-gt 0){@($reportedDeviceCounts|Measure-Object -Maximum)[0].Maximum}else{''}
                ExactRelatedDeviceCount=if($exactDeviceCountsForGroup.Count-gt 0){@($exactDeviceCountsForGroup|Measure-Object -Maximum)[0].Maximum}else{''}
                RelationshipCoverageStatus=Get-PreferredText -Rows @($group.Group) -Field 'RelationshipCoverageStatus'
                Platform=Get-PreferredText -Rows @($group.Group) -Field 'Platform'
                SourceCollectedDateTime=$collected
            })
        }
        Write-Warning ("Microsoft Graph returned {0} duplicate detected-application key(s), including {1} with conflicting attributes. Canonical text values and the maximum DeviceCount were retained."-f$duplicateDetectedAppGroups.Count,$conflictingDetectedAppGroupCount)
        $rows['Intune_DetectedApps.csv']=@($collapsedApps.ToArray())
    }
    $relationshipCountMismatchCount=@($rows['Intune_DetectedApps.csv']|Where-Object RelationshipCoverageStatus -eq 'ReconciledCountMismatch').Count
    if($relationshipCountMismatchCount-gt 0){Write-Warning ("Intune reported a different aggregate deviceCount for {0} detected application(s). DeviceCount was reconciled to the exact managed-device relations; ReportedDeviceCount preserves the source aggregate." -f $relationshipCountMismatchCount)}
    $rows['Intune_ConfigurationPolicies.csv']=@($policies|ForEach-Object{$id=Get-Text(Get-Value $_ 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Configuration policy response missing id.'};$template=Get-Value $_ 'templateReference';[pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';PolicyId=$id;DisplayName=Get-Text(Get-Value $_ 'name');Description=Get-Text(Get-Value $_ 'description');Platforms=Get-ListText(Get-Value $_ 'platforms');Technologies=Get-ListText(Get-Value $_ 'technologies');TemplateId=Get-Text(Get-Value $template 'templateId');TemplateFamily=Get-Text(Get-Value $template 'templateFamily');CreatedDateTime=Get-DateText (Get-Value $_ 'createdDateTime') 'createdDateTime' $id;LastModifiedDateTime=Get-DateText (Get-Value $_ 'lastModifiedDateTime') 'lastModifiedDateTime' $id;SourceCollectedDateTime=$collected}})
    $updateRows=New-Object System.Collections.Generic.List[object]
    foreach($item in $feature){$id=Get-Text(Get-Value $item 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Feature update policy response missing id.'};$updateRows.Add([pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';PolicyType='Feature';PolicyId=$id;DisplayName=Get-Text(Get-Value $item 'displayName');TargetVersion=Get-Text(Get-Value $item 'featureUpdateVersion');ReleaseDateTime='';DaysUntilForcedReboot='';CreatedDateTime=Get-DateText (Get-Value $item 'createdDateTime') 'createdDateTime' $id;LastModifiedDateTime=Get-DateText (Get-Value $item 'lastModifiedDateTime') 'lastModifiedDateTime' $id;SourceCollectedDateTime=$collected})}
    foreach($item in $quality){$id=Get-Text(Get-Value $item 'id');if([string]::IsNullOrWhiteSpace($id)){throw 'Quality update policy response missing id.'};$updateRows.Add([pscustomobject][ordered]@{SourceSystem='MicrosoftIntune';PolicyType='Quality';PolicyId=$id;DisplayName=Get-Text(Get-Value $item 'displayName');TargetVersion='';ReleaseDateTime=Get-DateText (Get-Value $item 'expeditedUpdateReleaseDateTime') 'expeditedUpdateReleaseDateTime' $id;DaysUntilForcedReboot=Get-IntegerText (Get-Value $item 'daysUntilForcedReboot') 'daysUntilForcedReboot' $id;CreatedDateTime=Get-DateText (Get-Value $item 'createdDateTime') 'createdDateTime' $id;LastModifiedDateTime=Get-DateText (Get-Value $item 'lastModifiedDateTime') 'lastModifiedDateTime' $id;SourceCollectedDateTime=$collected})}
    $rows['Intune_WindowsUpdatePolicies.csv']=@($updateRows.ToArray())
    foreach($name in $tableNames){$keyColumn=if($name -eq 'Intune_AutopilotDevices.csv'){'AutopilotDeviceId'}elseif($name -eq 'Intune_DetectedApps.csv'){'AppId'}elseif($name -eq 'Intune_DetectedAppDeviceRelationships.csv'){'RelationshipKey'}else{'PolicyId'};$duplicate=@($rows[$name]|Group-Object $keyColumn|Where-Object Count -gt 1);if($name -eq 'Intune_WindowsUpdatePolicies.csv'){$duplicate=@($rows[$name]|Group-Object {"$($_.PolicyType)|$($_.PolicyId)"}|Where-Object Count -gt 1)};if($duplicate.Count){throw "Duplicate keys returned for ${name}: $($duplicate.Name -join ', ')"}}
    $published=@()
    foreach($name in $tableNames){
        $run=$null
        try{$run=Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($latest[$name]) -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems;$stamp=[datetime]::UtcNow;$base=[IO.Path]::GetFileNameWithoutExtension($name);$history=Join-Path $paths.DataAllRootPath ('Intune\Operational\{0}\{1}\{2}_{3}.csv' -f $stamp.ToString('yyyy'),$stamp.ToString('MM'),$base,$stamp.ToString('yyyyMMdd-HHmmssfff'));$publishParameters=@{Run=$run;Columns=@($tables[$name].columns|ForEach-Object{[string]$_});HistoryPath=$history;LatestPath=$latest[$name];ContractPath=$rawContractPath;ContractTableName=$name};if($name-eq'Intune_DetectedAppDeviceRelationships.csv'-and-not[string]::IsNullOrWhiteSpace($preparedAppDeviceCsvPath)){$publishParameters.InputCsvPath=$preparedAppDeviceCsvPath}else{$publishParameters.InputObject=@($rows[$name])};Publish-SmartWorkplaceCMDBSourceCsv @publishParameters|Out-Null;$published+=$latest[$name]}catch{if($null -ne $run){Complete-SmartWorkplaceCMDBSourceCollection -Run $run -Failed};throw}
    }
    Write-Information ("SmartWorkplaceCMDB Intune operational collection completed. Autopilot={0}; Apps={1}; AppDeviceRelations={2}; ConfigurationPolicies={3}; UpdatePolicies={4}."-f$rows['Intune_AutopilotDevices.csv'].Count,$rows['Intune_DetectedApps.csv'].Count,$appDeviceRelationCount,$rows['Intune_ConfigurationPolicies.csv'].Count,$rows['Intune_WindowsUpdatePolicies.csv'].Count) -InformationAction Continue
    [pscustomobject]@{Status='Completed';ScriptVersion=$ScriptVersion;AutopilotCount=$rows['Intune_AutopilotDevices.csv'].Count;DetectedAppCount=$rows['Intune_DetectedApps.csv'].Count;DetectedAppDeviceRelationshipCount=$appDeviceRelationCount;ConfigurationPolicyCount=$rows['Intune_ConfigurationPolicies.csv'].Count;UpdatePolicyCount=$rows['Intune_WindowsUpdatePolicies.csv'].Count;PublishedPath=$published}
}catch{$executionError=$_;Write-Information ("Collector failure stack: {0}"-f$_.ScriptStackTrace) -InformationAction Continue;throw}finally{if(-not[string]::IsNullOrWhiteSpace($preparedAppDeviceCsvPath)){Remove-Item -LiteralPath (Split-Path -Parent $preparedAppDeviceCsvPath) -Recurse -Force -ErrorAction SilentlyContinue};Complete-SmartWorkplaceCMDBExecutionContext -RuntimeContext $runtime -ErrorRecord $executionError}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCFNkmWFdM2uW87
# 1bWL3vPbrHovS1a+rPpDtvKzxz9vdqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICYdsxH40GaPgqmuuLxc6D+t/Xhdy+JKa1yk09MacWIeMA0GCSqG
# SIb3DQEBAQUABIIBgJGlEy6KLlaEbPHpWG+9X40P97lm28k3+VZYHvS/LBnNu6Kl
# i7BjEKBf4Dzw4b4kbr/NoyNozdGFzVr7JCo2v8GB6aSapFozX/45Ug0k0hlxqKL+
# MfQ6T4jpC0dOZN6sRlIUtCeTQFTKESD9cKpwqkn+mr0/L7N6wS2g/5x1OoaembCb
# QeXm1xjr/U3mh1+yKG0Vxm/VZqEORqZQt836DwzdWfb122Z3M+CoF30LqO2SQv+W
# u+1IihjRQgtSeVcUcDCPuioU8QmFpTEkhtOMC5ZXpn5jVFWiaaQ2TZCWqcMuEATk
# xH/CDYE936APCXIoGkKeLT8OFw/Pt7N1DASkTaNFoS0fErBWzjGyzxJtVF8+2bHS
# d+bS6EZCU0qOS3xJHETr3VwWZJkC6QO9t4h+ufY3Dvo1+3xMhTjQfv52OwWV23Zd
# /rwccmPyKNPdoC83PIRdks06WRzs1nN8/mf6WA71KDPFfD2KGzetwGL9HC60ojpr
# bMhiaGtb2HqP9Zoz7qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQwODE3
# NThaMC8GCSqGSIb3DQEJBDEiBCCVHHayGfPVBYm4jmhYnwTUj+xgBOULq/VxysEA
# tHT1rTANBgkqhkiG9w0BAQEFAASCAgBHpC/5l6p9rb+ZSmy5E21aB2g85D43kRKk
# z3zKuOKFVS2V1iu1Fl1pT0zva523VTa3l9ZLwrAxV8EnA/xtDcfRjcsMm1BMcvSH
# tls5F/o2YKCW9HvLY9XU9Q4taWyJcZVUzAnfLV7frcdmfNk7UATOPdxn+ZP/hNic
# tzhw9L0fa9MSDWBLZ3arDuoZUrTEbZy807f8uHzA+pApAkkK4/x8RvK7NK7SlnUD
# J1QVv1GZbAyfBVllPGuBLAxvR+WlKidQmqLkBoRaXMTKhJTyr0vWjDpe9HUYg3Q9
# GXch+cD00jHuncrSKxIv5bEb7w7wTKoKTkU4C/9AGp3RP6cFhrG1Nm+JgnD5zvQa
# 6VDJPfc8RMpsUugtZMLHyOUxSLqjmjajSgypb/VyC17pTWb3YMe7alXbFeWj7VRf
# mcU5Ltn15aRHEZ2h/cNuoVjCMxjNcFVBB/gOMdo3R2kZXdXKappG+EGGVh1z5f2m
# 4xlQxQz/IwUI9Nk8Y2kjzcuslYZP4vvWGPxr38zSjj0Q28UU9b9C7H4PLxx+YZBl
# jY5Benb9o6XeNyV/N6CxJ+Vpnrr531R0zY8msHpN8C788e/LQ0mF2dnm/9AJjbeu
# 6I19XpFD6im5ECYIhHpledGzHstWFE9Gse95J7Nh6SwDM9GvxFITITd3cqqYLYjC
# FW75qTI7uA==
# SIG # End signature block
