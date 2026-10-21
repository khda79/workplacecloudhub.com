#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Teams tenant inventory with CSV exports and HTML alert summary.
.VERSION
 0.46
.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; ImportExcel.
    Minimum Graph application permissions: Team.ReadBasic.All; TeamMember.Read.All; Channel.ReadBasic.All; Group.Read.All; Reports.Read.All; Sites.Read.All.
    Optional: ChannelMember.Read.All is required only when private/shared channel member or owner expansion is enabled.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, ImportExcel, SmartM365.Core.psd1
    Minimum application permissions: Team.ReadBasic.All, TeamMember.Read.All, Channel.ReadBasic.All, Group.Read.All, Reports.Read.All, Sites.Read.All.
    Optional: ChannelMember.Read.All for private/shared channel owners when -IncludeChannelOwners is used.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='SmartM365.Core uses global execution context variables for logs and generated CSV tracking.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='Final console status is intentional for command-line use.')]
[CmdletBinding()]
param(
    [string]$Tenant='test',
    [int]$InactiveDays=180,
    [int]$MaxTeams=0,
    [switch]$DryRun,
    [switch]$AlwaysSend,
    [switch]$MailOnly,
    [switch]$AppendHistory,
    [double]$QuotaCriticalPercent=90,
    [int]$MinOwners=2,
    [int]$GuestWarningThreshold=25,
    [switch]$RequireSensitivityLabel,
    [switch]$IncludeChannelOwners,
    [string]$OutputPath,
    [int]$MaxItems = 0
)
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
if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
    foreach ($smartM365LimitName in @('TopUsers','TopMailboxes','MaxDevices','MaxSites','MaxTeams','MaxApps','MaxPolicies','Limit','MaxPages')) {
        $smartM365LimitVariable = Get-Variable -Name $smartM365LimitName -Scope Script -ErrorAction SilentlyContinue
        if ($smartM365LimitVariable -and -not $PSBoundParameters.ContainsKey($smartM365LimitName) -and $null -ne $smartM365LimitVariable.Value) {
            Set-Variable -Name $smartM365LimitName -Value ([int]$MaxItems) -Scope Script
        }
    }
}
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
$ScriptVersion="0.46"
$ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$TaskName = $ScriptBaseName
$RunStarted=Get-Date; $RunDateUtc=$RunStarted.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture); $RunId=[guid]::NewGuid().ToString(); $CurrentOperation='Initialize'
$TeamsRows=New-Object 'System.Collections.Generic.List[object]'; $MembersRows=New-Object 'System.Collections.Generic.List[object]'; $ChannelsRows=New-Object 'System.Collections.Generic.List[object]'; $GuestsRows=New-Object 'System.Collections.Generic.List[object]'; $Alerts=New-Object 'System.Collections.Generic.List[object]'; $GeneratedCsvPaths=New-Object 'System.Collections.Generic.List[string]'
if($PSVersionTable.PSVersion.Major -lt 7){throw 'This script requires PowerShell 7 or later.'}
$tenantContextPath=&{ $d=$PSScriptRoot; while($d){ foreach($c in @((Join-Path $d 'SmartM365-TenantContext.ps1'),(Join-Path $d 'Config\SmartM365-TenantContext.ps1'))){ if(Test-Path -LiteralPath $c){return $c} }; $p=Split-Path $d -Parent; if([string]::IsNullOrWhiteSpace($p)-or$p-eq$d){break}; $d=$p }; throw 'SmartM365-TenantContext.ps1 not found.' }
. $tenantContextPath
$TenantContext=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ctxDir=Split-Path $tenantContextPath -Parent; $SmartM365Root=if((Split-Path $ctxDir -Leaf)-ieq 'Config'){Split-Path $ctxDir -Parent}else{$ctxDir}
Import-Module -Name (Join-Path $SmartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.99' -Force -ErrorAction Stop
$LocalConfigPath=Join-Path $PSScriptRoot "$ScriptBaseName.local.json"; $LocalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $LocalConfigPath; $LocalTemplatePath=(Get-SmartM365JsonTemplateName -Path $LocalConfigPath)
if(-not(Test-Path -LiteralPath $LocalConfigPath)){Initialize-SmartM365LocalJsonFromTemplate -Path $LocalConfigPath -TemplatePath $LocalTemplatePath -ConfigDescription 'script local configuration'|Out-Null}
$ScriptConfig=Get-Content -LiteralPath $LocalConfigPath -Raw|ConvertFrom-Json
$ScriptConfig=Sync-SmartM365JsonConfigWithTemplate -Config $ScriptConfig -Path $LocalConfigPath -TemplatePath $LocalTemplatePath
function Resolve-ConfigToken{param([AllowNull()][object]$Value) if($Value -isnot [string]){return $Value}; $r=$Value; for($i=0;$i-lt 10;$i++){ $m=[regex]::Matches($r,'\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}'); if($m.Count-eq 0){break}; foreach($x in $m){$p=$TenantContext.PSObject.Properties[$x.Groups['Name'].Value]; if($p-and$null-ne$p.Value){$r=$r.Replace($x.Value,[string]$p.Value)}}}; $r}
function Get-ConfigValue{param([string]$Name,[AllowNull()][object]$DefaultValue) $p=$ScriptConfig.PSObject.Properties[$Name]; if($p-and$null-ne$p.Value){ if($p.Value -isnot [string] -or ($p.Value.Trim() -and $p.Value.Trim() -notin @('__USE_GLOBAL__','USE_GLOBAL'))){return Resolve-ConfigToken $p.Value}}; $c=$TenantContext.PSObject.Properties[$Name]; if($c-and$null-ne$c.Value){return Resolve-ConfigToken $c.Value}; Resolve-ConfigToken $DefaultValue}
function IsoUtc {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try {
        $dateValue = if ($Value -is [datetimeoffset]) {
            [datetimeoffset]$Value
        }
        elseif ($Value -is [datetime]) {
            [datetimeoffset]([datetime]$Value)
        }
        else {
            [datetimeoffset]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
        }
        return $dateValue.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
    }
    catch { return '' }
}
function Num{param([AllowNull()][object]$Value) if($null-eq$Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return ''}; try{([double]$Value).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)}catch{[string]$Value}}
function Prop{param([AllowNull()][object]$Object,[string[]]$Names) if($null-eq$Object){return $null}; foreach($n in $Names){$p=$Object.PSObject.Properties[$n]; if($p){return $p.Value}}; $null}
function JoinVals{param([AllowNull()][object[]]$Values) @($Values|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_)}) -join '; '}

function Test-TeamsGraphProperty {
    param([AllowNull()][object]$Object,[Parameter(Mandatory)][string]$Name)
    if($null-eq$Object){return $false}
    if($Object -is [System.Collections.IDictionary]){return $Object.Contains($Name)}
    return $null -ne $Object.PSObject.Properties[$Name]
}
function Get-TeamsGraphPropertyValue {
    param([AllowNull()][object]$Object,[Parameter(Mandatory)][string]$Name)
    if($null-eq$Object){return $null}
    if($Object -is [System.Collections.IDictionary]){return $Object[$Name]}
    $property=$Object.PSObject.Properties[$Name]
    if($property){return $property.Value}
    return $null
}

function Invoke-TeamsDailySummaryMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MarkerPath,
        [Parameter(Mandatory)][scriptblock]$SendAction,
        [switch]$ForceSend
    )

    $today = (Get-Date).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $markerParent = Split-Path -Path $MarkerPath -Parent
    if (-not (Test-Path -LiteralPath $markerParent -PathType Container)) {
        New-Item -Path $markerParent -ItemType Directory -Force | Out-Null
    }

    $lockPath = "$MarkerPath.lock"
    $lockStream = $null
    try {
        try {
            $lockStream = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        }
        catch [IO.IOException] {
            WriteLog -Message "Daily Teams summary email is already being evaluated by another run: $lockPath" -Level INFO
            return $false
        }

        $lastSentDate = if (Test-Path -LiteralPath $MarkerPath -PathType Leaf) {
            [string](Get-Content -LiteralPath $MarkerPath -Raw -ErrorAction SilentlyContinue)
        }
        else { '' }
        if ($lastSentDate.Trim() -eq $today -and -not $ForceSend) {
            WriteLog -Message "Daily Teams summary email already sent for $today; email skipped." -Level INFO
            return $false
        }

        $null = & $SendAction
        [IO.File]::WriteAllText($MarkerPath, $today, [Text.UTF8Encoding]::new($false))
        return $true
    }
    finally {
        if ($null -ne $lockStream) { $lockStream.Dispose() }
    }
}

function Add-Alert{param([string]$TeamId,[string]$TeamDisplayName,[ValidateSet('Warning','Critical')][string]$Status,[string]$Check,[AllowNull()][object]$NumericValue,[string]$TextValue,[string]$Threshold,[string]$Details) [void]$Alerts.Add([pscustomobject]@{TeamId=$TeamId;TeamDisplayName=$TeamDisplayName;Status=$Status;Check=$Check;NumericValue=(Num $NumericValue);TextValue=$TextValue;Threshold=$Threshold;Details=$Details})}
function Get-TeamsMailStatus {
    param([hashtable]$Summary,[bool]$PartialInventory)
    $total=[int]$Summary.TotalTeams
    if ($total -gt 0 -and -not $PartialInventory -and (([double]$Summary.OwnerlessTeams / $total) -ge 0.1 -or ([double]$Summary.HighQuotaTeams / $total) -ge 0.1)) { return 'Critical' }
    if (($Summary.CriticalCount + $Summary.WarningCount) -gt 0) { return 'Warning' }
    return 'OK'
}
$TeamsChannelRequestHeaders=@{Prefer='include-unknown-enum-members'}
function Get-TeamsRetryDelay {
    param(
        [AllowNull()][object]$Headers,
        [AllowNull()][object]$ErrorRecord,
        [int]$DefaultSeconds,
        [int]$MaximumSeconds = 300
    )
    if ($null -eq $Headers -and $null -ne $ErrorRecord) {
        try { $Headers = $ErrorRecord.Exception.Response.Headers } catch {}
    }
    $value = $null
    if ($null -ne $Headers) {
        try { $value = @($Headers.GetValues('Retry-After') | Select-Object -First 1)[0] } catch {}
        if ($null -eq $value -and $Headers -is [Collections.IDictionary]) {
            foreach ($name in @('Retry-After', 'retry-after')) {
                if ($Headers.Contains($name)) { $value = $Headers[$name]; break }
            }
        }
        if ($null -eq $value) {
            foreach ($name in @('Retry-After', 'RetryAfter')) {
                $property = $Headers.PSObject.Properties[$name]
                if ($property) { $value = $property.Value; break }
            }
        }
    }
    if ($null -eq $value -and $null -ne $ErrorRecord) {
        try { $value = $ErrorRecord.Exception.Data['Retry-After'] } catch {}
    }
    $seconds = 0
    if ($null -ne $value -and [int]::TryParse([string]$value, [ref]$seconds) -and $seconds -gt 0) {
        return [Math]::Min($seconds, $MaximumSeconds)
    }
    $retryDate = [datetimeoffset]::MinValue
    if ($null -ne $value -and [datetimeoffset]::TryParse([string]$value, [ref]$retryDate)) {
        $seconds = [int][Math]::Ceiling(($retryDate.ToUniversalTime() - [datetimeoffset]::UtcNow).TotalSeconds)
        if ($seconds -gt 0) { return [Math]::Min($seconds, $MaximumSeconds) }
    }
    return [Math]::Min([Math]::Max(1,$DefaultSeconds),$MaximumSeconds)
}

function Invoke-Graph {
    param(
        [string]$Uri,
        [string]$Operation = 'Graph request',
        [string]$OutputFilePath = '',
        [hashtable]$Headers = @{},
        [ValidateRange(1, 10)][int]$MaxAttempts = 5
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $request = @{ Method = 'GET'; Uri = $Uri; ErrorAction = 'Stop' }
            if ($OutputFilePath) { $request.OutputFilePath = $OutputFilePath }
            if ($Headers.Count -gt 0) { $request.Headers = $Headers }
            return Invoke-MgGraphRequest @request
        }
        catch {
            $statusCode = $null
            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
            $transient = $statusCode -in @(408, 409, 429, 500, 502, 503, 504) -or
                [string]$_.Exception.Message -match '(?i)throttl|TooManyRequests|temporarily unavailable|timeout|timed out'
            if (-not $transient -or $attempt -ge $MaxAttempts) { throw }
            $fallbackDelay = [Math]::Min(300, [Math]::Pow(2, $attempt) * 5)
            $delay = Get-TeamsRetryDelay -ErrorRecord $_ -DefaultSeconds ([int]$fallbackDelay) -MaximumSeconds 300
            WriteLog -Message ("$Operation transient/throttled. Status=$statusCode; attempt $attempt/$MaxAttempts; retry in $delay s.") -Level WARNING
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-GraphCollection {
    param(
        [string]$Uri,
        [string]$Operation,
        [hashtable]$Headers = @{},
        [scriptblock]$RequestInvoker
    )
    $items = New-Object 'System.Collections.Generic.List[object]'
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $next = $Uri
    $page = 0
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        if (-not $visited.Add($next)) { throw "$Operation returned a repeated @odata.nextLink; collection is incomplete." }
        $page++
        $response = if ($null -ne $RequestInvoker) { & $RequestInvoker $next } else { Invoke-Graph -Uri $next -Operation "$Operation page $page" -Headers $Headers }
        if (-not (Test-TeamsGraphProperty -Object $response -Name 'value')) {
            throw "$Operation page $page returned an invalid Graph collection response without a value property."
        }
        foreach ($item in @(Get-TeamsGraphPropertyValue -Object $response -Name 'value')) { if ($null -ne $item) { [void]$items.Add($item) } }
        $next = [string](Get-TeamsGraphPropertyValue -Object $response -Name '@odata.nextLink')
    }
    return $items.ToArray()
}

function Invoke-TeamsGraphBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Requests,
        [ValidateRange(1,10)][int]$MaxAttempts=5
    )

    if($Requests.Count-eq 0){return @()}
    $batchUri="https://graph.microsoft.com/v1.0/"+'$batch'
    $requestById=@{}
    foreach($request in $Requests){$requestById[[string]$request.id]=$request}
    $pending=@($Requests)
    $completed=[System.Collections.Generic.List[object]]::new()

    for($attempt=1;$attempt-le$MaxAttempts-and$pending.Count-gt 0;$attempt++){
        $body=@{requests=$pending}|ConvertTo-Json -Depth 8
        try{
            $batchResponse=Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $body -ContentType 'application/json' -ErrorAction Stop
        }catch{
            # 401/403 and other definitive answers fail at once; only transient transport errors retry.
            if($attempt-ge$MaxAttempts-or-not(Test-SmartM365GraphRetryableError -ErrorRecord $_).Retry){throw}
            $fallbackDelay=[Math]::Min(60,[Math]::Pow(2,$attempt)*5)
            $delay=Get-TeamsRetryDelay -ErrorRecord $_ -DefaultSeconds ([int]$fallbackDelay) -MaximumSeconds 60
            WriteLog -Message ("Teams Graph batch transport failed; attempt {0}/{1}; retry in {2} s: {3}" -f $attempt,$MaxAttempts,$delay,$_.Exception.Message) -Level INFO
            Start-Sleep -Seconds $delay
            continue
        }

        $retry=[System.Collections.Generic.List[object]]::new()
        $retryAfter=0
        foreach($response in @($batchResponse.responses)){
            $status=[int]$response.status
            if($status-in@(429,500,502,503,504)-and$attempt-lt$MaxAttempts){
                $original=$requestById[[string]$response.id]
                if($null-ne$original){[void]$retry.Add($original)}
                $fallbackDelay=[Math]::Min(60,[Math]::Pow(2,$attempt)*5)
                $retryAfter=[Math]::Max($retryAfter,(Get-TeamsRetryDelay -Headers $response.headers -DefaultSeconds $fallbackDelay))
            }else{
                [void]$completed.Add($response)
            }
        }

        $pending=@($retry)
        if($pending.Count-gt 0){
            WriteLog -Message ("Teams Graph batch has {0} transient sub-request(s); attempt {1}/{2}; retry in {3} s." -f $pending.Count,$attempt,$MaxAttempts,$retryAfter) -Level INFO
            Start-Sleep -Seconds $retryAfter
        }
    }

    return $completed.ToArray()
}
function Get-TeamsBatchSeed {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Teams)

    $result = @{}
    for ($offset = 0; $offset -lt $Teams.Count; $offset += 3) {
        $last = [math]::Min($offset + 2, $Teams.Count - 1)
        $requests = [System.Collections.Generic.List[object]]::new()
        $requestMap = @{}
        $requestId = 1

        foreach ($team in @($Teams[$offset..$last])) {
            $teamId = [string]$team.id
            $result[$teamId] = [pscustomobject]@{ Details=$null; Owners=$null; Members=$null; Channels=$null; Drive=$null; Site=$null }
            $specs = @(
                @{ Kind='Details'; Url="/teams/${teamId}?`$select=id,isArchived" },
                @{ Kind='Owners'; Url="/groups/$teamId/owners/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" },
                @{ Kind='Members'; Url="/groups/$teamId/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" },
                @{ Kind='Channels'; Headers=$TeamsChannelRequestHeaders; Url="/teams/$teamId/channels" },
                @{ Kind='Drive'; Url="/groups/$teamId/sites/root/drive?`$select=quota,webUrl" },
                @{ Kind='Site'; Url="/groups/$teamId/sites/root?`$select=id,webUrl" }
            )
            foreach ($spec in $specs) {
                $localId = [string]$requestId
                $requestMap[$localId] = [pscustomobject]@{ TeamId=$teamId; Kind=$spec.Kind }
                $batchRequest=@{id=$localId;method='GET';url=$spec.Url}
                if($spec.ContainsKey('Headers') -and $spec.Headers){$batchRequest.headers=$spec.Headers}
                [void]$requests.Add($batchRequest)
                $requestId++
            }
        }

        try {
            $batchResponses = @(Invoke-TeamsGraphBatch -Requests $requests.ToArray())
        }
        catch {
            WriteLog -Message ("Teams Graph batch failed for team offset {0}; sequential fallback will be used: {1}" -f $offset,$_.Exception.Message) -Level WARNING
            continue
        }

        foreach ($response in $batchResponses) {
            $meta = $requestMap[[string]$response.id]
            if ([int]$response.status -ne 200) {
                WriteLog -Message ("Teams Graph batch sub-request failed: TeamId={0}; Kind={1}; HTTP={2}. Sequential fallback will be used." -f $meta.TeamId,$meta.Kind,$response.status) -Level WARNING
                continue
            }
            if ($meta.Kind -in @('Owners','Members','Channels')) {
                if (-not (Test-TeamsGraphProperty -Object $response.body -Name 'value')) {
                    WriteLog -Message ("Teams Graph batch sub-request returned an invalid collection body: TeamId={0}; Kind={1}. Sequential fallback will be used." -f $meta.TeamId,$meta.Kind) -Level WARNING
                    continue
                }
                $items = [System.Collections.Generic.List[object]]::new()
                foreach ($item in @(Get-TeamsGraphPropertyValue -Object $response.body -Name 'value')) { if ($null -ne $item) { [void]$items.Add($item) } }
                $nextLink=[string](Get-TeamsGraphPropertyValue -Object $response.body -Name '@odata.nextLink')
                if (-not [string]::IsNullOrWhiteSpace($nextLink)) {
                    $continuationHeaders=if($meta.Kind-eq'Channels'){$TeamsChannelRequestHeaders}else{@{}}
                    foreach ($item in @(Get-GraphCollection -Uri $nextLink -Operation ("Get Teams {0} continuation" -f $meta.Kind) -Headers $continuationHeaders)) { [void]$items.Add($item) }
                }
                $result[$meta.TeamId].($meta.Kind) = @($items)
            }
            else {
                if ($null -eq $response.body) {
                    WriteLog -Message ("Teams Graph batch sub-request returned an empty body: TeamId={0}; Kind={1}. Sequential fallback will be used." -f $meta.TeamId,$meta.Kind) -Level WARNING
                    continue
                }
                $result[$meta.TeamId].($meta.Kind) = $response.body
            }
        }
    }
    return $result
}

function Get-TeamsChannelSiteId {
    param([AllowNull()][object]$DriveItem)
    $parent = Get-TeamsGraphPropertyValue -Object $DriveItem -Name 'parentReference'
    return [string](Get-TeamsGraphPropertyValue -Object $parent -Name 'siteId')
}

function Get-TeamsDriveSiteId {
    param([AllowNull()][object]$Drive)
    $ids = Get-TeamsGraphPropertyValue -Object $Drive -Name 'sharepointIds'
    $siteGuid = [guid]::Empty
    $webGuid = [guid]::Empty
    if (-not [guid]::TryParse([string](Get-TeamsGraphPropertyValue -Object $ids -Name 'siteId'),[ref]$siteGuid) -or
        -not [guid]::TryParse([string](Get-TeamsGraphPropertyValue -Object $ids -Name 'webId'),[ref]$webGuid)) { return '' }
    $siteUrl = [string](Get-TeamsGraphPropertyValue -Object $ids -Name 'siteUrl')
    if (-not $siteUrl) { $siteUrl = [string](Get-TeamsGraphPropertyValue -Object $Drive -Name 'webUrl') }
    $uri = $null
    if (-not [uri]::TryCreate($siteUrl,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -ne 'https') { return '' }
    return ('{0},{1},{2}' -f $uri.Host.ToLowerInvariant(),$siteGuid.ToString('D'),$webGuid.ToString('D'))
}

function Add-TeamsChannelSiteReferences {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ChannelRows)

    foreach ($row in $ChannelRows) {
        $type = [string]$row.MembershipType
        $row | Add-Member -NotePropertyMembers @{
            SharePointSiteId = ''
            SharePointSiteLinkStatus = if ($type -in @('','standard')) { 'NotApplicable' } else { 'Unavailable' }
        } -Force
    }
    $targets = @($ChannelRows | Where-Object { [string]$_.MembershipType -in @('private','shared') })
    $driveRows = @{}
    $folderStatus = @{}
    $driveStatus = @{}
    for ($offset = 0; $offset -lt $targets.Count; $offset += 20) {
        $last = [math]::Min($offset + 19,$targets.Count - 1)
        $requests = [System.Collections.Generic.List[object]]::new()
        $requestMap = @{}
        $requestId = 1
        foreach ($row in @($targets[$offset..$last])) {
            $id = [string]$requestId
            $requestMap[$id] = $row
            $channelId = [uri]::EscapeDataString([string]$row.ChannelId)
            [void]$requests.Add(@{id=$id;method='GET';url="/teams/$($row.TeamId)/channels/$channelId/filesFolder"})
            $requestId++
        }
        try { $responses = @(Invoke-TeamsGraphBatch -Requests $requests.ToArray()) }
        catch {
            WriteLog -Message ("Teams channel site batch failed at offset {0}: {1}" -f $offset,$_.Exception.Message) -Level WARNING
            continue
        }
        foreach ($response in $responses) {
            $row = $requestMap[[string]$response.id]
            if ($null -eq $row) { continue }
            $status = [string][int]$response.status
            $folderStatus[$status] = 1 + [int]$folderStatus[$status]
            if ($status -ne '200') { continue }
            $siteId = Get-TeamsChannelSiteId -DriveItem $response.body
            if (-not $siteId) { $siteId = Get-TeamsDriveSiteId -Drive $response.body }
            if ($siteId) {
                $row.SharePointSiteId = $siteId
                $row.SharePointSiteLinkStatus = 'Collected'
                continue
            }
            $parent = Get-TeamsGraphPropertyValue -Object $response.body -Name 'parentReference'
            $driveId = [string](Get-TeamsGraphPropertyValue -Object $parent -Name 'driveId')
            if (-not $driveId) { continue }
            if (-not $driveRows.ContainsKey($driveId)) { $driveRows[$driveId] = [System.Collections.Generic.List[object]]::new() }
            [void]$driveRows[$driveId].Add($row)
        }
    }
    $driveIds = @($driveRows.Keys | Sort-Object)
    for ($offset = 0; $offset -lt $driveIds.Count; $offset += 20) {
        $last = [math]::Min($offset + 19,$driveIds.Count - 1)
        $requests = [System.Collections.Generic.List[object]]::new()
        $requestMap = @{}
        $requestId = 1
        foreach ($driveId in @($driveIds[$offset..$last])) {
            $id = [string]$requestId
            $requestMap[$id] = $driveId
            [void]$requests.Add(@{id=$id;method='GET';url=("/drives/{0}?`$select=id,sharepointIds,webUrl" -f [uri]::EscapeDataString($driveId))})
            $requestId++
        }
        try { $responses = @(Invoke-TeamsGraphBatch -Requests $requests.ToArray()) }
        catch {
            WriteLog -Message ("Teams channel drive metadata batch failed at offset {0}: {1}" -f $offset,$_.Exception.Message) -Level WARNING
            continue
        }
        foreach ($response in $responses) {
            $driveId = $requestMap[[string]$response.id]
            if (-not $driveId) { continue }
            $status = [string][int]$response.status
            $driveStatus[$status] = 1 + [int]$driveStatus[$status]
            if ($status -ne '200') { continue }
            $siteId = Get-TeamsDriveSiteId -Drive $response.body
            if (-not $siteId) { continue }
            foreach ($row in $driveRows[$driveId]) {
                $row.SharePointSiteId = $siteId
                $row.SharePointSiteLinkStatus = 'Collected'
            }
        }
    }
    $unavailable = @($targets | Where-Object { $_.SharePointSiteLinkStatus -ne 'Collected' }).Count
    $folderCodes = @($folderStatus.Keys | Sort-Object | ForEach-Object { '{0}:{1}' -f $_,$folderStatus[$_] }) -join ','
    $driveCodes = @($driveStatus.Keys | Sort-Object | ForEach-Object { '{0}:{1}' -f $_,$driveStatus[$_] }) -join ','
    WriteLog -Message ("Teams channel site links: {0}/{1} resolved; {2} unavailable; filesFolder HTTP=[{3}]; drive metadata HTTP=[{4}]." -f ($targets.Count - $unavailable),$targets.Count,$unavailable,$folderCodes,$driveCodes) -Level $(if ($unavailable -gt 0) { 'WARNING' } else { 'INFO' })
}

function Get-ReportRow {
    param([string]$ReportName, [string]$Period = 'D180')
    $temporaryPath = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-$ReportName-$([guid]::NewGuid().ToString('N')).csv")
    try {
        Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/reports/{0}(period='{1}')" -f $ReportName, $Period) -Operation $ReportName -OutputFilePath $temporaryPath | Out-Null
        if (-not (Test-Path -LiteralPath $temporaryPath -PathType Leaf) -or (Get-Item -LiteralPath $temporaryPath).Length -eq 0) {
            throw "Report $ReportName produced no CSV content."
        }
        return @(Import-Csv -LiteralPath $temporaryPath -ErrorAction Stop)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { try { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction Stop } catch { Write-Warning "Temporary CSV cleanup failed: $($_.Exception.Message)" } }
    }
}

function Write-TeamsCsvAtomically {
    param([object[]]$Rows, [string[]]$Columns, [Parameter(Mandatory)][string]$Path)
    $parent = Split-Path -Path $Path -Parent
    if ([string]::IsNullOrWhiteSpace($parent)) { $parent = (Get-Location).Path }
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -Path $parent -ItemType Directory -Force -ErrorAction Stop | Out-Null }
    $extension = [IO.Path]::GetExtension($Path)
    if ([string]::IsNullOrWhiteSpace($extension)) { $extension = '.tmp' }
    $temporaryPath = Join-Path $parent ('{0}.{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($Path), [guid]::NewGuid().ToString('N'), $extension)
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    try {
        if (@($Rows).Count -eq 0) {
            $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ','
            Set-Content -LiteralPath $temporaryPath -Value $header -Encoding $encoding -ErrorAction Stop
        }
        else {
            $Rows | Select-Object -Property $Columns | Add-SmartM365TenantKey |
                Export-Csv -LiteralPath $temporaryPath -NoTypeInformation -Encoding $encoding -ErrorAction Stop
        }
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { try { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction Stop } catch { Write-Warning "Temporary CSV cleanup failed: $($_.Exception.Message)" } }
    }
}

function Export-InventoryCsv {
    param(
        [object[]]$Rows,
        [string[]]$Columns,
        [string]$TimestampedPath,
        [string]$LatestPath,
        [string]$HistoryPath
    )
    $Columns = @('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId') + @($Columns | Where-Object { $_ -inotmatch '^(TenantKey|OrganizationKey|EnvironmentKey|TenantId)$' })
    Assert-SmartM365CsvDataCompleteness -Data $Rows -Columns $Columns -TimestampedPath $TimestampedPath -LatestPath $LatestPath
    Write-TeamsCsvAtomically -Rows $Rows -Path $TimestampedPath -Columns $Columns
    # DATA-LAST copy through the shared atomic copy: Publishing fence, receipt registration and source
    # staging, so a running or failed run never changes DATA-LAST.
    $latestWritePath = if (Get-Command -Name Resolve-SmartM365SourcePublicationPath -ErrorAction SilentlyContinue) { Resolve-SmartM365SourcePublicationPath -Path $LatestPath } else { $LatestPath }
    Copy-SmartM365FileAtomically -SourcePath $TimestampedPath -DestinationPath $latestWritePath -VerifyHash
    [void]$GeneratedCsvPaths.Add($TimestampedPath)
    # WeeklyHistory reads this run's copy, never the previous run's DATA-LAST file.
    [void]$GeneratedCsvPaths.Add($latestWritePath)
    if (-not $global:csvGeneratedPaths) { $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase) }
    [void]$global:csvGeneratedPaths.Add($TimestampedPath)
    [void]$global:csvGeneratedPaths.Add($LatestPath)
    if ($DryRun) {
        WriteLog -Message 'DryRun enabled: SharePoint CSV upload skipped.' -Level INFO
    }
    else {
        $timestampedUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $TimestampedPath
        if ($latestWritePath -cne $LatestPath) { Add-SmartM365StagedPublicationAction -Path $LatestPath -SharePointUpload }
        else { Invoke-SmartM365SharePointCsvUpload -LocalFilePath $LatestPath | Out-Null }
        # Same seven-day retention on SharePoint as on the server, once this run's copy is uploaded.
        if ($timestampedUpload) { Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath $TimestampedPath -RetentionDays 7 | Out-Null }
    }
    if ($AppendHistory -and $HistoryPath -and $Rows.Count -gt 0) {
        $historyParent = Split-Path $HistoryPath -Parent
        if (-not (Test-Path -LiteralPath $historyParent)) { New-Item -Path $historyParent -ItemType Directory -Force | Out-Null }
        if (Test-Path -LiteralPath $HistoryPath) { Repair-SmartM365CsvTenantKeySchema -Path $HistoryPath -Delimiter ',' -Encoding UTF8 | Out-Null }
        Add-SmartM365CsvRowsAtomically -Data @($Rows | Select-Object -Property $Columns) -Path $HistoryPath -Columns $Columns -Encoding utf8BOM
    }
}
function Get-TeamsCsvColumnNames {
    param([Parameter(Mandatory)][string]$Path)
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        if ($parser.EndOfData) { return @() }
        return @($parser.ReadFields())
    }
    finally { $parser.Dispose() }
}

function Import-TeamsMailCsvSet {
    param([Parameter(Mandatory)][string]$Folder,[Parameter(Mandatory)][string]$ExpectedTenantKey)
    $names=@('Teams','Members','Channels','Guests')
    $sets=@{}; $files=New-Object 'System.Collections.Generic.List[object]'
    $sourceRunId=''; $sourceRunDate=''; $sourceTenantId=''; $primaryWriteTime=$null
    foreach($name in $names){
        $path=Join-Path $Folder "M365_Teams_$name.csv"
        if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw "MailOnly requires the latest CSV: $path"}
        $headers=@(Get-TeamsCsvColumnNames -Path $path)
        $requiredColumns=@('TenantKey','RunId','RunDateUtc')
        if($name -eq 'Teams'){$requiredColumns+=@('TeamId','TeamDisplayName','OwnerCount','GuestCount','InactiveDays','StorageQuotaPercent','IsArchived','Visibility','Status','Details','Threshold','ActivityEvidenceStatus','StorageCollectionStatus')}
        foreach($required in $requiredColumns){if($headers -notcontains $required){throw "MailOnly CSV $path is missing $required."}}
        $beforeHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $rows=@(Import-Csv -LiteralPath $path -ErrorAction Stop)
        $afterHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        if($beforeHash -ne $afterHash){throw "MailOnly CSV changed while being read: $path"}
        $sets[$name]=$rows
        $file=Get-Item -LiteralPath $path
        if($name -eq 'Teams'){
            if($rows.Count -eq 0){throw 'MailOnly requires at least one Teams row to establish the source run.'}
            $primaryWriteTime=$file.LastWriteTimeUtc
        }
        $files.Add([pscustomobject]@{Path=$path;WorksheetName=$name;TableName="Teams$name";LastWriteTimeUtc=$file.LastWriteTimeUtc;Empty=($rows.Count -eq 0);Sha256=$afterHash})
        foreach($row in $rows){
            $key=[string]$row.TenantKey; $id=[string]$row.RunId; $date=[string]$row.RunDateUtc
            if([string]::IsNullOrWhiteSpace($key) -or $key -ine $ExpectedTenantKey -or [string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($date)){throw "MailOnly CSV $path contains a row with missing or unexpected tenant/run identity."}
            if(-not $sourceRunId){$sourceRunId=$id; $sourceRunDate=$date; $sourceTenantId=[string]$row.TenantId}
            if($id -ne $sourceRunId -or $date -ne $sourceRunDate -or ([string]$row.TenantId -ne $sourceTenantId)){throw "MailOnly CSV files do not belong to one source run: $path"}
        }
    }
    foreach($file in $files){
        if($file.Empty -and [math]::Abs(($file.LastWriteTimeUtc-$primaryWriteTime).TotalMinutes) -gt 30){throw "MailOnly empty CSV is not timestamp-consistent with the Teams CSV: $($file.Path)"}
    }
    $parsedDate=[datetimeoffset]::MinValue
    if(-not [datetimeoffset]::TryParse($sourceRunDate,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$parsedDate)){throw "MailOnly source RunDateUtc is invalid: $sourceRunDate"}
    [pscustomobject]@{Rows=$sets;Files=$files.ToArray();RunId=$sourceRunId;RunDateUtc=$sourceRunDate;Started=$parsedDate.UtcDateTime}
}

function Ensure-TeamsImportExcelModule {
    [CmdletBinding()]
    param()
    Initialize-SmartM365ImportExcel -InstallIfMissing | Out-Null
}
function New-TeamsTimestampedWorkbook {
    param(
        [Parameter(Mandatory)][object[]]$CsvFiles,
        [Parameter(Mandatory)][string]$Path
    )

    Initialize-SmartM365ImportExcel -InstallIfMissing | Out-Null
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    foreach ($csv in $CsvFiles) {
        $rows = @(Import-Csv -LiteralPath $csv.Path)
        $isEmpty = $rows.Count -eq 0
        if ($isEmpty) {
            $placeholder = [ordered]@{}
            foreach ($column in @(Get-TeamsCsvColumnNames -Path $csv.Path)) { $placeholder[$column] = '' }
            $rows = @([pscustomobject]$placeholder)
        }
        $rows | Export-Excel -Path $Path -WorksheetName $csv.WorksheetName -TableName $csv.TableName -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
        if ($isEmpty) {
            $package = Open-ExcelPackage -Path $Path
            try { $package.Workbook.Worksheets[$csv.WorksheetName].DeleteRow(2) }
            finally { Close-ExcelPackage -ExcelPackage $package }
        }
    }
    return $Path
}
function Add-TeamsWorkbookOverview {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][hashtable]$Summary,[Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Alerts,[Parameter(Mandatory)][string]$Status)
    $overview = @(
        [pscustomobject]@{Metric='Inventory status';Value=$Status}
        [pscustomobject]@{Metric='Teams inventoried';Value=$Summary.TotalTeams}
        [pscustomobject]@{Metric='Teams without owner';Value=$Summary.OwnerlessTeams}
        [pscustomobject]@{Metric='Inactive teams';Value=$Summary.InactiveTeams}
        [pscustomobject]@{Metric='Teams above storage quota threshold';Value=$Summary.HighQuotaTeams}
        [pscustomobject]@{Metric='Active teams';Value=$Summary.ActiveTeams}
        [pscustomobject]@{Metric='Archived teams';Value=$Summary.ArchivedTeams}
        [pscustomobject]@{Metric='Public teams';Value=$Summary.PublicTeams}
        [pscustomobject]@{Metric='Teams with guests';Value=$Summary.TeamsWithGuests}
        [pscustomobject]@{Metric='Object-level critical findings';Value=$Summary.CriticalCount}
        [pscustomobject]@{Metric='Object-level warning findings';Value=$Summary.WarningCount}
    )
    $overview | Export-Excel -Path $Path -WorksheetName 'Summary' -TableName 'TeamsSummary' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    $findings = @($Alerts | Select-Object Status,TeamDisplayName,Check,NumericValue,TextValue,Threshold,Details)
    if ($findings.Count -gt 0) {
        $findings | Export-Excel -Path $Path -WorksheetName 'Findings' -TableName 'TeamsFindings' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    }
    else {
        [pscustomobject]@{Status='';TeamDisplayName='';Check='';NumericValue='';TextValue='';Threshold='';Details=''} | Export-Excel -Path $Path -WorksheetName 'Findings' -TableName 'TeamsFindings' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
        $package = Open-ExcelPackage -Path $Path
        try { $package.Workbook.Worksheets['Findings'].DeleteRow(2) }
        finally { Close-ExcelPackage -ExcelPackage $package }
    }
}

function New-TeamsSharePointLinksHtml {
    param([Parameter(Mandatory)][string[]]$Paths)

    $links = foreach ($path in $Paths) {
        $record = Get-SmartM365SharePointUploadRecordForLocalFile -FilePath $path
        if (-not $record -or [string]::IsNullOrWhiteSpace([string]$record.WebUrl)) {
            WriteLog -Message ("Mail link omitted because no SharePoint WebUrl is available for {0}." -f [IO.Path]::GetFileName($path)) -Level WARNING
            continue
        }
        $name = [Net.WebUtility]::HtmlEncode([IO.Path]::GetFileName($path))
        $url = [Net.WebUtility]::HtmlEncode([string]$record.WebUrl)
        "<li style='margin:0 0 6px;'><a href='$url' style='color:#075985;text-decoration:underline;'>$name</a></li>"
    }
    if (@($links).Count -eq 0) { return '' }
    return "<div class='card'><h2>Timestamped exports</h2><p>SharePoint links for this run:</p><ul>$($links -join '')</ul></div>"
}
function ConvertTo-HtmlReport {
    param([object[]]$AlertRows,[object[]]$TeamRows=@(),[hashtable]$Summary,[string]$Worst,[datetime]$Started,[datetime]$Ended,[string]$FileLinksHtml='',[string]$SourceNote='')
    $color=@{OK='#107c10';Warning='#ff8c00';Critical='#d13438'}
    $affected=@($AlertRows|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.TeamId)}|Select-Object -ExpandProperty TeamId -Unique).Count
    $activityUnavailable=@($TeamRows|Where-Object{$_.ActivityEvidenceStatus -ne 'Observed'}).Count
    $storageUnavailable=@($TeamRows|Where-Object{$_.StorageCollectionStatus -ne 'Collected'}).Count
    $refreshDates=@($TeamRows|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.ActivityReportRefreshDate)}|ForEach-Object{[string]$_.ActivityReportRefreshDate}|Sort-Object -Unique)
    $refreshText=if($refreshDates.Count -eq 0){'Not available'}elseif($refreshDates.Count -eq 1){$refreshDates[0]}else{"$($refreshDates[0]) to $($refreshDates[-1])"}
    $categories=@(
        @{Check='Owners';Priority='High';Label='No owner';Action='Assign owners';Text='NoOwner'},
        @{Check='Owners';Priority='Review';Label='Owners below minimum';Action='Add another owner';Text='LowOwnerCount'},
        @{Check='StorageQuotaPercent';Priority='High';Label='Storage above threshold';Action='Review storage and quota'},
        @{Check='Inactivity';Priority='Review';Label='Inactive team';Action='Confirm use before archiving'},
        @{Check='GuestCount';Priority='Review';Label='Guest count above threshold';Action='Review membership'},
        @{Check='PublicSensitiveLabel';Priority='Review';Label='Public team with sensitivity label';Action='Review visibility and label'},
        @{Check='MissingSensitivityLabel';Priority='Review';Label='Missing sensitivity label';Action='Review labeling'}
    )
    $categoryRows=foreach($category in $categories){
        $categoryAlerts=@($AlertRows|Where-Object{$_.Check -eq $category.Check -and (-not $category.ContainsKey('Text') -or $_.TextValue -eq $category.Text)})
        if($categoryAlerts.Count -gt 0){
            $label=$category.Label
            if($category.Check -eq 'StorageQuotaPercent' -and [string]$categoryAlerts[0].Threshold -match '(?<value>[\d.]+) percent'){$label="Storage above $($Matches.value)%"}
            if($category.Check -eq 'Inactivity' -and [string]$categoryAlerts[0].Threshold -match '(?<value>\d+) days'){$label="Inactive for over $($Matches.value) days"}
            ,@($category.Priority,$label,$categoryAlerts.Count,$category.Action)
        }
    }
    $categoryHtml=if(@($categoryRows).Count -gt 0){New-SmartM365EmailTableHtml -Headers @('Priority','Finding','Teams','Suggested review') -Rows @($categoryRows)}else{'<p>No action was flagged by the configured checks.</p>'}
    $rankedAlerts=@($AlertRows|Sort-Object @{Expression={if($_.Status -eq 'Critical'){0}else{1}}},@{Expression={switch($_.Check){'Owners'{0}'StorageQuotaPercent'{1}'Inactivity'{2}default{3}}}},TeamDisplayName)
    $topAlerts=New-Object 'System.Collections.Generic.List[object]'
    $seenTeams=@{}
    foreach($alert in $rankedAlerts){
        $key=[string]$alert.TeamId
        if($seenTeams.ContainsKey($key)){continue}
        $seenTeams[$key]=$true
        $topAlerts.Add($alert)
        if($topAlerts.Count -eq 5){break}
    }
    $topRows=foreach($alert in $topAlerts){
        $finding=switch([string]$alert.Check){'Owners'{if($alert.TextValue -eq 'NoOwner'){'No owner'}else{'Owners below minimum'}}'StorageQuotaPercent'{'Storage above threshold'}'Inactivity'{'Inactive team'}'GuestCount'{'Guest count above threshold'}'PublicSensitiveLabel'{'Public team with sensitivity label'}'MissingSensitivityLabel'{'Missing sensitivity label'}default{[string]$alert.Check}}
        $measured=if(-not[string]::IsNullOrWhiteSpace([string]$alert.NumericValue)){[string]$alert.NumericValue}else{[string]$alert.TextValue}
        ,@([string]$alert.TeamDisplayName,$finding,$measured,[string]$alert.Threshold)
    }
    $topHtml=if($topRows){New-SmartM365EmailTableHtml -Headers @('Team','Finding','Measured','Threshold') -Rows @($topRows)}else{'<p>No priority team to review.</p>'}
    $cards=@(
        [pscustomobject]@{Label='Teams inventoried';Value=([int]$Summary.TotalTeams).ToString('N0');Detail='teams processed';Background='#f8fafc';Border='#dbe3ef';Accent='#0f172a';Span=1},
        [pscustomobject]@{Label='Teams to review';Value=$affected.ToString('N0');Detail='distinct teams';Background='#eff6ff';Border='#bfdbfe';Accent='#1d4ed8';Span=1},
        [pscustomobject]@{Label='Without owner';Value=([int]$Summary.OwnerlessTeams).ToString('N0');Detail='owner required';Background='#fff7ed';Border='#fed7aa';Accent='#9a3412';Span=1},
        [pscustomobject]@{Label='Above storage threshold';Value=([int]$Summary.HighQuotaTeams).ToString('N0');Detail='measured quotas';Background='#faf5ff';Border='#e9d5ff';Accent='#7e22ce';Span=1},
        [pscustomobject]@{Label='Inactive teams';Value=([int]$Summary.InactiveTeams).ToString('N0');Detail='dated activity';Background='#f0fdfa';Border='#99f6e4';Accent='#0f766e';Span=1},
        [pscustomobject]@{Label='Teams with guests';Value=([int]$Summary.TeamsWithGuests).ToString('N0');Detail='guest members';Background='#f0fdf4';Border='#bbf7d0';Accent='#166534';Span=1},
        [pscustomobject]@{Label='Activity unavailable';Value=$activityUnavailable.ToString('N0');Detail='no activity report row';Background='#eef2ff';Border='#c7d2fe';Accent='#4338ca';Span=2}
    )
    $sb=[Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!doctype html><html><head><meta charset="utf-8"></head><body style="font-family:Segoe UI,Arial,sans-serif;background:#f5f8fb;color:#1f2937;padding:24px;">')
    [void]$sb.AppendLine(("<div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h1>Microsoft Teams Inventory <span style='color:#fff;border-radius:999px;padding:4px 10px;background:{0};font-size:14px;'>{1}</span></h1><p>{2} teams processed; {3} distinct teams require review. Findings can overlap. The complete detail is in the attached workbook.</p></div>" -f $color[$Worst],$Worst,$Summary.TotalTeams,$affected))
    [void]$sb.AppendLine("<div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h2>Summary</h2>$(New-SmartM365EmailKpiGridHtml -Cards $cards -RowSizes @(4,3))</div>")
    [void]$sb.AppendLine("<div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h2>Action required</h2>$categoryHtml<h3>First teams to review</h3>$topHtml</div>")
    [void]$sb.AppendLine(("<div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h2>Evidence coverage</h2><p>Activity report refresh: {0}<br>Teams without an activity row: {1}<br>Teams without measured storage: {2}<br>{3}</p></div>" -f [Net.WebUtility]::HtmlEncode($refreshText),$activityUnavailable,$storageUnavailable,[Net.WebUtility]::HtmlEncode($SourceNote)))
    if (-not [string]::IsNullOrWhiteSpace($FileLinksHtml)) { [void]$sb.AppendLine($FileLinksHtml) }
    $runMachine=if($SourceNote){'Not recorded in CSV'}else{$env:COMPUTERNAME}
    $runEnd=if($SourceNote){'Not recorded in CSV'}else{IsoUtc $Ended}
    $runDuration=if($SourceNote){'Not recorded in CSV'}else{(New-TimeSpan -Start $Started -End $Ended).ToString()}
    [void]$sb.AppendLine(("<div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;color:#64748b;font-size:12px;'><b>Run details</b><br>RunId: {0}<br>Machine: {1}<br>Started UTC: {2}<br>Ended UTC: {3}<br>Duration: {4}<br>Teams processed: {5}</div>" -f [Net.WebUtility]::HtmlEncode([string]$RunId),[Net.WebUtility]::HtmlEncode($runMachine),(IsoUtc $Started),[Net.WebUtility]::HtmlEncode($runEnd),[Net.WebUtility]::HtmlEncode($runDuration),$Summary.TotalTeams))
    [void]$sb.AppendLine('</body></html>')
    $sb.ToString()
}
if([string]::IsNullOrWhiteSpace($OutputPath)){$OutputPath=[string](Get-ConfigValue 'TeamsInventoryCsvLogFolderPath' '{{DataAllRootPath}}\M365\Teams\Inventory')}
$LatestCsvFolderPath=[string](Get-ConfigValue 'LatestCsvFolderPath' $TenantContext.LatestCsvFolderPath); $WeeklyHistoryFolderPath=[string](Get-ConfigValue 'WeeklyHistoryFolderPath' (Join-Path $OutputPath 'WeeklyHistory')); $WeeklyHistoryRetentionWeeks=[int](Get-ConfigValue 'WeeklyHistoryRetentionWeeks' 52); $EnableWeeklyHistory=[bool](Get-ConfigValue 'EnableWeeklyHistory' $true)
 if(-not$PSBoundParameters.ContainsKey('RequireSensitivityLabel')){$RequireSensitivityLabel=[bool](Get-ConfigValue 'RequireSensitivityLabel' $false)}; if(-not$PSBoundParameters.ContainsKey('IncludeChannelOwners')){$IncludeChannelOwners=[bool](Get-ConfigValue 'IncludeChannelOwners' $false)}
 $ShowMailLinks=[bool](Get-ConfigValue 'ShowMailLinks' $false)
$global:RetentionMaxCSV=[int](Get-ConfigValue 'RetentionMaxCSV' 30); $global:RetentionMaxLogs=[int](Get-ConfigValue 'RetentionMaxLogs' 30)
$global:EnableSharePointUpload=[bool](Get-ConfigValue 'EnableSharePointUpload' $false); $global:SharePointSiteHostname=[string](Get-ConfigValue 'SharePointSiteHostname' ''); $global:SharePointSitePath=[string](Get-ConfigValue 'SharePointSitePath' ''); $global:SharePointLibraryDisplayName=[string](Get-ConfigValue 'SharePointLibraryDisplayName' 'Documents'); $global:SharePointTargetFolderPath=[string](Get-ConfigValue 'SharePointTargetFolderPath' '')
$AppId=[string](Get-ConfigValue 'AppId' ''); $TenantId=[string](Get-ConfigValue 'TenantId' ''); $OrgDomain=[string](Get-ConfigValue 'OrgDomain' ''); $Thumb=[string](Get-ConfigValue 'Thumbprint' (Get-ConfigValue 'Thumb' ''))
$global:AppId=$AppId; $global:TenantId=$TenantId; $global:OrgDomain=$OrgDomain; $global:Thumb=$Thumb; $global:Thumbprint=$Thumb
$teamColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','Description','Visibility','CreatedDateTimeUtc','Classification','SensitivityLabel','IsArchived','OwnerCount','MemberCount','GuestCount','StandardChannelCount','PrivateChannelCount','SharedChannelCount','LastActivityDateUtc','InactiveDays','StorageUsedGB','StorageQuotaGB','StorageQuotaPercent','Status','NumericValue','TextValue','Threshold','Details')
$memberColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','UserId','DisplayName','UserPrincipalName','Mail','UserType','Role','Status','NumericValue','TextValue','Threshold','Details')
$teamColumns += @('MemberCollectionStatus','OwnerCollectionStatus','ChannelCollectionStatus','StorageCollectionStatus','ActivityEvidenceStatus','ActivityReportRefreshDate','CollectedAtUtc','SharePointSiteId','SharePointSiteUrl','SharePointSiteLinkStatus','IsPartialInventory')
$channelColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','ChannelId','ChannelDisplayName','MembershipType','CreatedDateTimeUtc','PrivateChannelOwners','Status','NumericValue','TextValue','Threshold','Details')
$channelColumns += @('SharePointSiteId','SharePointSiteLinkStatus')
$guestColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','UserId','DisplayName','UserPrincipalName','Mail','Status','NumericValue','TextValue','Threshold','Details')
try{
 $CurrentOperation='Initialize script environment'; $OutputPath=InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $ScriptBaseName; Start-SmartM365RunTranscript -Path $global:logTranscriptFile -Append|Out-Null; Write-SmartM365LoadedModuleVersions; WriteLog -Message "Starting $TaskName"
 $CurrentOperation='Check ImportExcel startup health'; Ensure-TeamsImportExcelModule
 if($MailOnly){
  $CurrentOperation='Validate existing Teams CSVs'
  if($MaxTeams -gt 0 -or $MaxItems -gt 0){throw 'MailOnly cannot be combined with MaxTeams or MaxItems.'}
  $tenantKey=if($TenantContext.PSObject.Properties['TenantKey']){[string]$TenantContext.TenantKey}else{$Tenant}
  $snapshot=Import-TeamsMailCsvSet -Folder $LatestCsvFolderPath -ExpectedTenantKey $tenantKey
  $teamArray=@($snapshot.Rows.Teams); $memberArray=@($snapshot.Rows.Members); $channelArray=@($snapshot.Rows.Channels); $guestArray=@($snapshot.Rows.Guests)
  $RunId=$snapshot.RunId; $RunStarted=$snapshot.Started
  $thresholdText=[string]$teamArray[0].Threshold
  $sourceInactiveDays=if($thresholdText -match 'inactive <= (?<value>\d+) days'){[int]$Matches.value}else{$InactiveDays}
  $sourceQuotaPercent=if($thresholdText -match 'storage <= (?<value>[\d.]+) percent'){[double]::Parse($Matches.value,[Globalization.CultureInfo]::InvariantCulture)}else{$QuotaCriticalPercent}
  foreach($row in $teamArray){
   $flags=@([string]$row.Details -split '; ')
   if($flags -contains 'No owner'){Add-Alert $row.TeamId $row.TeamDisplayName Critical Owners $row.OwnerCount 'NoOwner' $row.Threshold 'Team has no owner.'}
   if($flags -contains 'Owner count below threshold'){Add-Alert $row.TeamId $row.TeamDisplayName Warning Owners $row.OwnerCount 'LowOwnerCount' $row.Threshold 'Team has too few owners.'}
   if($flags -contains 'Inactive team'){Add-Alert $row.TeamId $row.TeamDisplayName Warning Inactivity $row.InactiveDays $row.LastActivityDateUtc $row.Threshold 'Team is a candidate for archival.'}
   if($flags -contains 'Storage quota critical'){Add-Alert $row.TeamId $row.TeamDisplayName Critical StorageQuotaPercent $row.StorageQuotaPercent $row.StorageUsedGB $row.Threshold 'SharePoint storage quota usage is above threshold.'}
   if($flags -contains 'High guest count'){Add-Alert $row.TeamId $row.TeamDisplayName Warning GuestCount $row.GuestCount 'Guests' $row.Threshold 'Team has many external guests.'}
   if($flags -contains 'Public team with sensitivity label'){Add-Alert $row.TeamId $row.TeamDisplayName Warning PublicSensitiveLabel 1 $row.SensitivityLabel $row.Threshold 'Public team has a sensitivity/classification label.'}
   if($flags -contains 'Missing sensitivity label'){Add-Alert $row.TeamId $row.TeamDisplayName Warning MissingSensitivityLabel 0 'NoLabel' $row.Threshold 'Tenant policy expects labels.'}
  }
  $alertArray=$Alerts.ToArray()
  $summary=@{TotalTeams=$teamArray.Count;ActiveTeams=@($teamArray|Where-Object{$_.IsArchived-ne'True' -and ([string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-or [double]$_.InactiveDays-le$sourceInactiveDays)}).Count;InactiveTeams=@($teamArray|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-and [double]$_.InactiveDays-gt$sourceInactiveDays}).Count;ArchivedTeams=@($teamArray|Where-Object{$_.IsArchived-eq'True'}).Count;PublicTeams=@($teamArray|Where-Object{$_.Visibility-eq'Public'}).Count;PrivateTeams=@($teamArray|Where-Object{$_.Visibility-eq'Private'}).Count;TeamsWithGuests=@($teamArray|Where-Object{[int]$_.GuestCount-gt 0}).Count;OwnerlessTeams=@($teamArray|Where-Object{[int]$_.OwnerCount-eq 0}).Count;HighQuotaTeams=@($teamArray|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.StorageQuotaPercent)-and [double]$_.StorageQuotaPercent-gt$sourceQuotaPercent}).Count;CriticalCount=@($alertArray|Where-Object Status -eq Critical).Count;WarningCount=@($alertArray|Where-Object Status -eq Warning).Count;MemberRows=$memberArray.Count;ChannelRows=$channelArray.Count;GuestRows=$guestArray.Count}
  $worst=Get-TeamsMailStatus -Summary $summary -PartialInventory $true
  $CurrentOperation='Build Teams mail from existing CSVs'
  Ensure-TeamsImportExcelModule
  $stamp='{0}_{1}' -f (Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss_fff',[Globalization.CultureInfo]::InvariantCulture),[guid]::NewGuid().ToString('N').Substring(0,8)
  $workbookPath=Join-Path $OutputPath "M365_Teams_Inventory_MailOnly_$stamp.xlsx"
  New-TeamsTimestampedWorkbook -CsvFiles $snapshot.Files -Path $workbookPath|Out-Null
  foreach($file in $snapshot.Files){if((Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -ne $file.Sha256){throw "MailOnly CSV changed before the email was prepared: $($file.Path)"}}
  Add-TeamsWorkbookOverview -Path $workbookPath -Summary $summary -Alerts $alertArray -Status $worst
  $sourceNote="Mail reconstructed from existing CSVs collected $($snapshot.RunDateUtc). No new inventory collection; source scope and freshness were not reverified."
  $subject="[SNAPSHOT][$($worst.ToUpperInvariant())] Microsoft Teams Inventory | $($summary.OwnerlessTeams) ownerless | $($summary.HighQuotaTeams) above storage threshold"
  $html=ConvertTo-HtmlReport -AlertRows $alertArray -TeamRows $teamArray -Summary $summary -Worst $worst -Started $RunStarted -Ended $RunStarted -SourceNote $sourceNote
  if($DryRun){WriteLog -Message 'DryRun enabled: MailOnly summary email skipped.' -Level INFO}else{$dailySummaryMarkerPath=Join-Path (Split-Path $global:LogTextFile -Parent) "$ScriptBaseName-DailySummary-LastSent.txt"; $dailySummarySent=Invoke-TeamsDailySummaryMail -MarkerPath $dailySummaryMarkerPath -ForceSend:$AlwaysSend -SendAction {Send-SmartM365Mail -Subject $subject -BodyHtml $html -Attachments @($workbookPath) -AllowAttachments -SuppressAttachmentLinks}; if($dailySummarySent){WriteLog -Message "MailOnly Teams summary email sent from source run $RunId." -Level SUCCESS}}
  try{Stop-SmartM365RunTranscript|Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile}catch{$null=$_}
  Complete-SmartM365ExecutionContext -Status Success
  $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
 }
 Start-SmartM365CmdbSourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $LatestCsvFolderPath -ReadOnly:$DryRun
 # DATA-LAST changes only when this run completes; DryRun holds no receipt.
 Enable-SmartM365StagedSourcePublication
 $CurrentOperation='Connect Microsoft Graph'; Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true; $conn=Connect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -AppId $AppId -Thumbprint $Thumb -TenantId $TenantId -Organization $OrgDomain -GraphScopes @('Team.ReadBasic.All','TeamMember.Read.All','Channel.ReadBasic.All','Group.Read.All','Reports.Read.All'); if(-not$conn.GraphConnected){throw 'Microsoft Graph app-only connection failed.'}
 $CurrentOperation='Ensure ImportExcel module'; Ensure-TeamsImportExcelModule
 $CurrentOperation='Run preflight'; Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('Microsoft.Graph.Authentication','ImportExcel') -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('Team.ReadBasic.All','TeamMember.Read.All','Channel.ReadBasic.All','Group.Read.All','Reports.Read.All','Sites.Read.All') -GraphProbeUris @('https://graph.microsoft.com/v1.0/organization','https://graph.microsoft.com/v1.0/groups?$top=1')|Out-Null
 $CurrentOperation='Load tenant metadata'; $org=Invoke-Graph -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName' -Operation 'Get organization'; $TenantName=[string]@($org.value)[0].displayName; if([string]::IsNullOrWhiteSpace($TenantName)){$TenantName=$Tenant}
 $activityById=@{}; foreach($r in (Get-ReportRow -ReportName 'getTeamsTeamActivityDetail' -Period 'D180')){$id=[string](Prop $r @('Team Id','TeamId','Team ID')); if($id){$activityById[$id]=$r}}
 $teamFilter=[uri]::EscapeDataString("resourceProvisioningOptions/Any(x:x eq 'Team')"); $teamsUri="https://graph.microsoft.com/v1.0/groups?`$filter=$teamFilter&`$select=id,displayName,description,visibility,createdDateTime,classification,assignedLabels,mail,webUrl&`$top=999"; $teams=@(Get-GraphCollection -Uri $teamsUri -Operation 'Get team groups'); if($MaxTeams-gt 0){$teams=@($teams|Select-Object -First $MaxTeams)}; WriteLog -Message ("Teams discovered: {0}" -f $teams.Count)
 $teamsPartialInventory=($MaxTeams -gt 0 -or (Test-SmartM365MaxItemsMode))
 $CurrentOperation='Prefetch Teams Graph data'; $teamBatchSeed=Get-TeamsBatchSeed -Teams $teams; WriteLog -Message ("Teams Graph batch prefetch completed: {0}/{1} teams seeded." -f $teamBatchSeed.Count,$teams.Count)
 $i=0; foreach($g in $teams){$i++; $teamId=[string]$g.id; $teamName=[string]$g.displayName; WriteLog -Message ("Processing team {0}/{1}: {2}" -f $i,$teams.Count,$teamName); $CurrentOperation="Process team $teamName"
  $seed=if($teamBatchSeed.ContainsKey($teamId)){$teamBatchSeed[$teamId]}else{$null}; $details=if($seed-and$null-ne$seed.Details){$seed.Details}else{$null}; if($null-eq$details){$details=Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/teams/{0}" -f $teamId) -Operation 'Get team details'}; $archived=if($details-and$details.PSObject.Properties['isArchived']){[bool]$details.isArchived}else{$false}
  $labels=@(); foreach($l in @($g.assignedLabels)){$labels += [string](if($l.displayName){$l.displayName}else{$l.labelId})}; $label=JoinVals $labels
  if($seed-and$null-ne$seed.Owners){$owners=@($seed.Owners)}else{$owners=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" -f $teamId) -Operation 'Get owners')}; if($seed-and$null-ne$seed.Members){$members=@($seed.Members)}else{$members=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" -f $teamId) -Operation 'Get members')}
  $ownerIds=@{}; foreach($o in $owners){$ownerIds[[string]$o.id]=$true}; $guests=@($members|Where-Object{[string]$_.userType-eq'Guest'})
  foreach($m in $members){$role=if($ownerIds.ContainsKey([string]$m.id)){'Owner'}else{'Member'}; [void]$MembersRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;UserId=[string]$m.id;DisplayName=[string]$m.displayName;UserPrincipalName=[string]$m.userPrincipalName;Mail=[string]$m.mail;UserType=[string]$m.userType;Role=$role;Status='OK';NumericValue='';TextValue=$role;Threshold='Inventory only';Details=''})}
  foreach($guest in $guests){[void]$GuestsRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;UserId=[string]$guest.id;DisplayName=[string]$guest.displayName;UserPrincipalName=[string]$guest.userPrincipalName;Mail=[string]$guest.mail;Status='Warning';NumericValue='1';TextValue='Guest';Threshold="Guests <= $GuestWarningThreshold";Details='External guest member'})}
  if($seed-and$null-ne$seed.Channels){$channels=@($seed.Channels)}else{$channels=@()}; if($null-eq$seed-or$null-eq$seed.Channels){$channels=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/teams/{0}/channels" -f $teamId) -Operation 'Get channels' -Headers $TeamsChannelRequestHeaders)}
  $unknownChannels=@($channels|Where-Object{[string]$_.membershipType-eq'unknownFutureValue'}); if($unknownChannels.Count-gt 0){WriteLog -Message ("Team {0} still returned {1} channel(s) with membershipType=unknownFutureValue despite the evolvable-enum request header; channel category totals exclude them." -f $teamName,$unknownChannels.Count) -Level WARNING}
  $standard=@($channels|Where-Object{[string]$_.membershipType-in@('','standard')}).Count; $private=@($channels|Where-Object{[string]$_.membershipType-eq'private'}).Count; $shared=@($channels|Where-Object{[string]$_.membershipType-eq'shared'}).Count
  foreach($ch in $channels){$chOwners=@(); if($IncludeChannelOwners-and [string]$ch.membershipType-in@('private','shared')){try{$cm=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/teams/{0}/channels/{1}/members?`$top=200" -f $teamId,$ch.id) -Operation 'Get channel members'); $chOwners=@($cm|Where-Object{@($_.roles)-contains'owner'}|ForEach-Object{$_.displayName})}catch{$chOwners=@('NotMeasured: ChannelMember.Read.All may be required')}}; [void]$ChannelsRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;ChannelId=[string]$ch.id;ChannelDisplayName=[string]$ch.displayName;MembershipType=[string]$ch.membershipType;CreatedDateTimeUtc=(IsoUtc $ch.createdDateTime);PrivateChannelOwners=(JoinVals $chOwners);Status='OK';NumericValue='1';TextValue=[string]$ch.membershipType;Threshold='Inventory only';Details=''})}
  $last=''; $inactive=''; $act=$activityById[$teamId]; if($act){$la=Prop $act @('Last Activity Date','LastActivityDate'); $last=IsoUtc $la; if($last){$inactive=[math]::Round(((Get-Date).ToUniversalTime()-([datetime]$la).ToUniversalTime()).TotalDays,0)}}
  $usedGb=''; $quotaGb=''; $quotaPct=''; $storageDetail='Storage not measured'; try{$drive=if($seed-and$null-ne$seed.Drive){$seed.Drive}else{Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/sites/root/drive?`$select=quota,webUrl" -f $teamId) -Operation 'Get team SharePoint quota'}; if($drive.quota-and[double]$drive.quota.total-gt 0){$usedGb=[math]::Round([double]$drive.quota.used/1GB,2); $quotaGb=[math]::Round([double]$drive.quota.total/1GB,2); $quotaPct=[math]::Round(([double]$drive.quota.used/[double]$drive.quota.total)*100,2); $storageDetail=[string]$drive.webUrl}}catch{$storageDetail='NotMeasured: '+$_.Exception.Message}
  $teamSiteId=''; $teamSiteUrl=''; $teamSiteLinkStatus='Unavailable'
  try {
   $teamSite=if($seed-and$null-ne$seed.Site){$seed.Site}else{Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/sites/root?`$select=id,webUrl" -f $teamId) -Operation 'Get team SharePoint site'}
   $teamSiteId=[string](Get-TeamsGraphPropertyValue -Object $teamSite -Name 'id')
   $teamSiteUrl=[string](Get-TeamsGraphPropertyValue -Object $teamSite -Name 'webUrl')
   if($teamSiteId){$teamSiteLinkStatus='Collected'}
  } catch { WriteLog -Message ("Team SharePoint site identity unavailable for TeamId={0}: {1}" -f $teamId,$_.Exception.Message) -Level WARNING }
  $status='OK'; $notes=New-Object 'System.Collections.Generic.List[string]'; if($owners.Count-eq 0){$status='Critical'; [void]$notes.Add('No owner'); Add-Alert $teamId $teamName Critical Owners $owners.Count 'NoOwner' "Owners >= $MinOwners" 'Team has no owner.'}elseif($owners.Count-lt$MinOwners){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Owner count below threshold'); Add-Alert $teamId $teamName Warning Owners $owners.Count 'LowOwnerCount' "Owners >= $MinOwners" 'Team has too few owners.'}
  if($inactive-ne'' -and [double]$inactive-gt$InactiveDays -and -not$archived){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Inactive team'); Add-Alert $teamId $teamName Warning Inactivity $inactive $last "<= $InactiveDays days" 'Team is a candidate for archival.'}
  if($quotaPct-ne'' -and [double]$quotaPct-gt$QuotaCriticalPercent){$status='Critical'; [void]$notes.Add('Storage quota critical'); Add-Alert $teamId $teamName Critical StorageQuotaPercent $quotaPct $storageDetail "<= $QuotaCriticalPercent percent" 'SharePoint storage quota usage is above threshold.'}
  if($guests.Count-gt$GuestWarningThreshold){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('High guest count'); Add-Alert $teamId $teamName Warning GuestCount $guests.Count Guests "<= $GuestWarningThreshold" 'Team has many external guests.'}
  if([string]$g.visibility-eq'Public' -and -not[string]::IsNullOrWhiteSpace($label)){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Public team with sensitivity label'); Add-Alert $teamId $teamName Warning PublicSensitiveLabel 1 $label 'Review public sensitive teams' 'Public team has a sensitivity/classification label.'}
  if($RequireSensitivityLabel -and [string]::IsNullOrWhiteSpace($label)){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Missing sensitivity label'); Add-Alert $teamId $teamName Warning MissingSensitivityLabel 0 NoLabel 'Sensitivity label required' 'Tenant policy expects labels.'}
  [void]$TeamsRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;Description=[string]$g.description;Visibility=[string]$g.visibility;CreatedDateTimeUtc=(IsoUtc $g.createdDateTime);Classification=[string]$g.classification;SensitivityLabel=$label;IsArchived=[string]$archived;OwnerCount=$owners.Count;MemberCount=$members.Count;GuestCount=$guests.Count;StandardChannelCount=$standard;PrivateChannelCount=$private;SharedChannelCount=$shared;LastActivityDateUtc=$last;InactiveDays=(Num $inactive);StorageUsedGB=(Num $usedGb);StorageQuotaGB=(Num $quotaGb);StorageQuotaPercent=(Num $quotaPct);Status=$status;NumericValue=(Num $members.Count);TextValue="Owners=$($owners.Count); Members=$($members.Count); Guests=$($guests.Count)";Threshold="Owners >= $MinOwners; inactive <= $InactiveDays days; storage <= $QuotaCriticalPercent percent; guests <= $GuestWarningThreshold";Details=(JoinVals $notes);SharePointSiteId=$teamSiteId;SharePointSiteUrl=$teamSiteUrl;SharePointSiteLinkStatus=$teamSiteLinkStatus;IsPartialInventory=$teamsPartialInventory})
 }
  $CurrentOperation='Resolve Teams private and shared channel sites'; Add-TeamsChannelSiteReferences -ChannelRows $ChannelsRows.ToArray()
 foreach ($teamRow in $TeamsRows) {
  $teamActivity=$activityById[[string]$teamRow.TeamId]
  $teamRow | Add-Member -NotePropertyMembers @{
   MemberCollectionStatus='Collected';OwnerCollectionStatus='Collected';ChannelCollectionStatus='Collected'
   StorageCollectionStatus=if([string]::IsNullOrWhiteSpace([string]$teamRow.StorageUsedGB)){'Unavailable'}else{'Collected'}
   ActivityEvidenceStatus=if($teamActivity){'Observed'}else{'NoReportRow'}
   ActivityReportRefreshDate=if($teamActivity){[string](Prop $teamActivity @('Report Refresh Date','ReportRefreshDate'))}else{''}
   CollectedAtUtc=[datetime]::UtcNow.ToString('o')
  }
 }
 $stamp=(Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture)
 $teamsTimestampedPath=Join-Path $OutputPath "M365_Teams_Teams_$stamp.csv"; $membersTimestampedPath=Join-Path $OutputPath "M365_Teams_Members_$stamp.csv"; $channelsTimestampedPath=Join-Path $OutputPath "M365_Teams_Channels_$stamp.csv"; $guestsTimestampedPath=Join-Path $OutputPath "M365_Teams_Guests_$stamp.csv"
 Export-InventoryCsv -Rows $TeamsRows.ToArray() -Columns $teamColumns -TimestampedPath $teamsTimestampedPath -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Teams.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Teams_History.csv')
 Export-InventoryCsv -Rows $MembersRows.ToArray() -Columns $memberColumns -TimestampedPath $membersTimestampedPath -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Members.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Members_History.csv')
 Export-InventoryCsv -Rows $ChannelsRows.ToArray() -Columns $channelColumns -TimestampedPath $channelsTimestampedPath -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Channels.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Channels_History.csv')
 Export-InventoryCsv -Rows $GuestsRows.ToArray() -Columns $guestColumns -TimestampedPath $guestsTimestampedPath -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Guests.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Guests_History.csv')
 $timestampedCsvFiles=@(
  [pscustomobject]@{Path=$teamsTimestampedPath;WorksheetName='Teams';TableName='TeamsInventory'},
  [pscustomobject]@{Path=$membersTimestampedPath;WorksheetName='Members';TableName='TeamsMembers'},
  [pscustomobject]@{Path=$channelsTimestampedPath;WorksheetName='Channels';TableName='TeamsChannels'},
  [pscustomobject]@{Path=$guestsTimestampedPath;WorksheetName='Guests';TableName='TeamsGuests'}
 )
 $workbookPath=Join-Path $OutputPath "M365_Teams_Inventory_$stamp.xlsx"; New-TeamsTimestampedWorkbook -CsvFiles $timestampedCsvFiles -Path $workbookPath|Out-Null; Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputPath -FilePattern 'M365_Teams_Inventory_*.xlsx' -RetentionDays 7 -LogFile $global:LogTextFile
 if($EnableWeeklyHistory-and-not$DryRun){Add-SmartM365WeeklyHistory -SourceCsvPaths $GeneratedCsvPaths.ToArray() -HistoryRootPath $WeeklyHistoryFolderPath -RetentionWeeks $WeeklyHistoryRetentionWeeks -HistoryLabel 'Microsoft Teams inventory' -UploadChangedFilesOnly|Out-Null}elseif($DryRun){WriteLog -Message 'DryRun enabled: WeeklyHistory skipped.' -Level INFO}
 $teamArray=$TeamsRows.ToArray(); $memberArray=$MembersRows.ToArray(); $channelArray=$ChannelsRows.ToArray(); $guestArray=$GuestsRows.ToArray(); $alertArray=$Alerts.ToArray()
 $summary=@{TotalTeams=$teamArray.Count;ActiveTeams=@($teamArray|Where-Object{$_.IsArchived-ne'True' -and ([string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-or [double]$_.InactiveDays-le$InactiveDays)}).Count;InactiveTeams=@($teamArray|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-and [double]$_.InactiveDays-gt$InactiveDays}).Count;ArchivedTeams=@($teamArray|Where-Object{$_.IsArchived-eq'True'}).Count;PublicTeams=@($teamArray|Where-Object{$_.Visibility-eq'Public'}).Count;PrivateTeams=@($teamArray|Where-Object{$_.Visibility-eq'Private'}).Count;TeamsWithGuests=@($teamArray|Where-Object{[int]$_.GuestCount-gt 0}).Count;OwnerlessTeams=@($teamArray|Where-Object{[int]$_.OwnerCount-eq 0}).Count;HighQuotaTeams=@($teamArray|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.StorageQuotaPercent)-and [double]$_.StorageQuotaPercent-gt$QuotaCriticalPercent}).Count;CriticalCount=@($alertArray|Where-Object Status -eq Critical).Count;WarningCount=@($alertArray|Where-Object Status -eq Warning).Count;MemberRows=$memberArray.Count;ChannelRows=$channelArray.Count;GuestRows=$guestArray.Count}
 $worst=Get-TeamsMailStatus -Summary $summary -PartialInventory ($MaxTeams-gt 0 -or (Test-SmartM365MaxItemsMode))
 Add-TeamsWorkbookOverview -Path $workbookPath -Summary $summary -Alerts $alertArray -Status $worst
 if(-not$DryRun){$workbookUpload=Invoke-SmartM365SharePointCsvUpload -LocalFilePath $workbookPath; if($workbookUpload){Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath $workbookPath -RetentionDays 7|Out-Null}}
 $mailFileLinks=if($ShowMailLinks){New-TeamsSharePointLinksHtml -Paths (@($timestampedCsvFiles.Path)+@($workbookPath))}else{''}
  $subject="[$($worst.ToUpperInvariant())] Microsoft Teams Inventory | $($summary.OwnerlessTeams) ownerless | $($summary.HighQuotaTeams) above storage threshold"; $html=ConvertTo-HtmlReport -AlertRows $alertArray -TeamRows $teamArray -Summary $summary -Worst $worst -Started $RunStarted -Ended (Get-Date) -FileLinksHtml $mailFileLinks
 if($DryRun){WriteLog -Message 'DryRun enabled: daily summary email skipped.' -Level INFO}else{$dailySummaryMarkerPath=Join-Path -Path (Split-Path -Path $global:LogTextFile -Parent) -ChildPath "$ScriptBaseName-DailySummary-LastSent.txt"; $dailySummarySent=Invoke-TeamsDailySummaryMail -MarkerPath $dailySummaryMarkerPath -ForceSend:$AlwaysSend -SendAction {Send-SmartM365Mail -Subject $subject -BodyHtml $html -Attachments @($workbookPath) -AllowAttachments -SuppressAttachmentLinks}; if($dailySummarySent){WriteLog -Message ("Daily Teams summary email sent: {0}" -f $subject) -Level SUCCESS}}
 Set-SmartM365CmdbSourceScope -CompleteScope ($MaxTeams -eq 0 -and $MaxItems -eq 0) -Scope 'CMDB:teams,team_members'
 $result="Teams=$($summary.TotalTeams); Critical=$($summary.CriticalCount); Warnings=$($summary.WarningCount); Members=$($memberArray.Count); Channels=$($channelArray.Count); Guests=$($guestArray.Count)"; try{Stop-SmartM365RunTranscript|Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile}catch{$null=$_}; WriteLog -Message ("Result summary: $result") -Level INFO; Write-Host "Teams inventory completed. Status=$worst; $result"; Complete-SmartM365ExecutionContext -Status $(if($worst-eq'OK'){'Success'}else{'CompletedWithWarnings'})
}catch{ $err=$_; try{WriteLog -Message ("Teams inventory failed during {0}: {1}" -f $CurrentOperation,$err.Exception.Message) -Level ERROR}catch{$null=$_}; try{Stop-SmartM365RunTranscript|Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile}catch{$null=$_}; try{Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $err -FailureStage $CurrentOperation}catch{$null=$_}; throw }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB5/kPoBTpZ23YG
# Ed3wBYUdpk3lMkrtZQ+saSjsbdS8FKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOX5b4fsjPn2rjQ+ytVD7R3jo/eYgckbJL6vSrZpaydcMA0GCSqG
# SIb3DQEBAQUABIIBgA2eIGuE3ulpAgxUzSofQQdXcFqosFf9b2pgPfS0VxZZJr7q
# eBD708q7srtXT1RDB5+yUMltH4rveb2c5vYtu+zimKAJFdBJplReDcl092FsMF/0
# ASwn4FDyHIToF3jHNP26BuZzRfJwsQz/YPL0cjxLwoVhuJO++kWB/7UN57UKZgyQ
# AvSNpbpKlIEVU5eGijXlqj+N6/rwJ4y3YYRMzTpllymr23s3TECXP8KyyRqpxzRk
# 5M2xgJw4eyRE/SK0qiyR6vRFBRtmCV4994xy/reWt2cWXTdjimTQFhxsJO4Ku6bB
# SzcXfon0TooVA3UX4ZDCqX1+OczDCgwzD2gYbQU6SCYHKsiWw+SbdubJuWHmnB+T
# 45yGfKdi5OkvfYeCi0fhmAPK36A9FlfbCT464AsEfx1wISKKwdp6I90FBqbwMEtD
# I5bRxyum74M6qVMGcK+oo0euzquMbKgMlbfRMGYnr3bFFcFKtS0kAt26JhZx1Rr3
# Diu92qWLJxRQR6r1kqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExMjAw
# MjFaMC8GCSqGSIb3DQEJBDEiBCCOqWY5FH5a9Iednp4l1+QbFxpWFI4GhGNoaifR
# AeqWHjANBgkqhkiG9w0BAQEFAASCAgCq3zAsNJTNg2BUAPNrA+zEiLUfbFTo05Yq
# GIHd+xpCwwrwuIO5hwa3A/jxhs2XlI3fWBvj5NAsd5nfRIv1DtbvudREBhyP15Kl
# Y7ZFeS1PDWBFDNR227mXgTBWmBWJHgaGz7pgy1SryII36GVBVjY1yTXVpxHpTctb
# LGXiU6HojIqaZPVfhKY9u5cO4zI3uaYzslcCdcFFqBmpBvSf601j2ETy6NwR5T3f
# QmSN30RHqQUdxT44CQ9saoouzdwdwrG/Bs5GhcTFVtpWbWBevhtGznaDFZyuTZ+p
# yKL+9UPOxSKeWest4L/oOlfuJb9FQFCLPO2fHvEHN/guQE5Sc8zH6REZs4HOFTNH
# 1KQO9RRcJ3BXdAtkx8Ozd5WXRo77K+YoV38hsCBsjpFT9mNSIOIn3sc+ocbRSUaV
# ilX7o227C2ZJYKkQgE3AnqvCdCuIb1BFCSQi3s5GLVevwywQkx++QR5fsX4Vp7jf
# QApPH3RLZAFmSsG8xKa/5KSYSAso846WI84fZewZ3y49Ji+3WfeaQDZdAiB1DqPo
# TSE7zwO7+75Bp3iApTkNLZ/GSEzXYoiSC8lhFf9Xc65PomPawMUM1r5NWj3ojlsj
# mjn0w/djTAIkI4T0wejFzjll3OcmJG76+7KP8pKotjOp1DofuJUivQz4ZrXqP1Is
# b7G1h5jSnw==
# SIG # End signature block
