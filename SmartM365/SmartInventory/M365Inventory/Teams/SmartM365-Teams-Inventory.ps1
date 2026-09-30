#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Teams tenant inventory with CSV exports and HTML alert summary.
.VERSION
0.34
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
    [switch]$AppendHistory,
    [double]$QuotaCriticalPercent=90,
    [int]$MinOwners=2,
    [int]$GuestWarningThreshold=25,
    [switch]$RequireSensitivityLabel,
    [switch]$IncludeChannelOwners,
    [string]$OutputPath,
    [int]$MaxItems = 0
)
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
$ScriptVersion="0.34"
$ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$TaskName = $ScriptBaseName
$RunStarted=Get-Date; $RunDateUtc=$RunStarted.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture); $RunId=[guid]::NewGuid().ToString(); $CurrentOperation='Initialize'
$TeamsRows=New-Object 'System.Collections.Generic.List[object]'; $MembersRows=New-Object 'System.Collections.Generic.List[object]'; $ChannelsRows=New-Object 'System.Collections.Generic.List[object]'; $GuestsRows=New-Object 'System.Collections.Generic.List[object]'; $Alerts=New-Object 'System.Collections.Generic.List[object]'; $GeneratedCsvPaths=New-Object 'System.Collections.Generic.List[string]'
if($PSVersionTable.PSVersion.Major -lt 7){throw 'This script requires PowerShell 7 or later.'}
$tenantContextPath=&{ $d=$PSScriptRoot; while($d){ foreach($c in @((Join-Path $d 'SmartM365-TenantContext.ps1'),(Join-Path $d 'Config\SmartM365-TenantContext.ps1'))){ if(Test-Path -LiteralPath $c){return $c} }; $p=Split-Path $d -Parent; if([string]::IsNullOrWhiteSpace($p)-or$p-eq$d){break}; $d=$p }; throw 'SmartM365-TenantContext.ps1 not found.' }
. $tenantContextPath
$TenantContext=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ctxDir=Split-Path $tenantContextPath -Parent; $SmartM365Root=if((Split-Path $ctxDir -Leaf)-ieq 'Config'){Split-Path $ctxDir -Parent}else{$ctxDir}
Import-Module -Name (Join-Path $SmartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.62' -Force -ErrorAction Stop
$LocalConfigPath=Join-Path $PSScriptRoot "$ScriptBaseName.local.json"; $LocalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $LocalConfigPath; $LocalTemplatePath=(Get-SmartM365JsonTemplateName -Path $LocalConfigPath)
if(-not(Test-Path -LiteralPath $LocalConfigPath)){Initialize-SmartM365LocalJsonFromTemplate -Path $LocalConfigPath -TemplatePath $LocalTemplatePath -ConfigDescription 'script local configuration'|Out-Null}
$ScriptConfig=Get-Content -LiteralPath $LocalConfigPath -Raw|ConvertFrom-Json
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
        [Parameter(Mandatory)][scriptblock]$SendAction
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
        if ($lastSentDate.Trim() -eq $today) {
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
function WorstStatus{param([object[]]$Rows) if(@($Rows|Where-Object Status -eq Critical).Count){'Critical'}elseif(@($Rows|Where-Object Status -eq Warning).Count){'Warning'}else{'OK'}}
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
            if($attempt-ge$MaxAttempts){throw}
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
    for ($offset = 0; $offset -lt $Teams.Count; $offset += 4) {
        $last = [math]::Min($offset + 3, $Teams.Count - 1)
        $requests = [System.Collections.Generic.List[object]]::new()
        $requestMap = @{}
        $requestId = 1

        foreach ($team in @($Teams[$offset..$last])) {
            $teamId = [string]$team.id
            $result[$teamId] = [pscustomobject]@{ Details=$null; Owners=$null; Members=$null; Channels=$null; Drive=$null }
            $specs = @(
                @{ Kind='Details'; Url="/teams/${teamId}?`$select=id,isArchived" },
                @{ Kind='Owners'; Url="/groups/$teamId/owners/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" },
                @{ Kind='Members'; Url="/groups/$teamId/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" },
                @{ Kind='Channels'; Headers=$TeamsChannelRequestHeaders; Url="/teams/$teamId/channels" },
                @{ Kind='Drive'; Url="/groups/$teamId/sites/root/drive?`$select=quota,webUrl" }
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
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
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
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
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
    Write-TeamsCsvAtomically -Rows $Rows -Path $LatestPath -Columns $Columns
    [void]$GeneratedCsvPaths.Add($TimestampedPath)
    [void]$GeneratedCsvPaths.Add($LatestPath)
    if (-not $global:csvGeneratedPaths) { $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase) }
    [void]$global:csvGeneratedPaths.Add($TimestampedPath)
    [void]$global:csvGeneratedPaths.Add($LatestPath)
    if ($DryRun) {
        WriteLog -Message 'DryRun enabled: SharePoint CSV upload skipped.' -Level INFO
    }
    else {
        $timestampedUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $TimestampedPath
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $LatestPath | Out-Null
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

function Ensure-TeamsImportExcelModule {
    [CmdletBinding()]
    param()

    $module = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) {
        $installCommand = 'Install-Module ImportExcel -Scope CurrentUser -Repository PSGallery -Force -AllowClobber'
        WriteLog -Message ("ImportExcel is not installed. Automatic installation starting: {0}" -f $installCommand) -Level WARNING
        if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
            throw "ImportExcel is missing and Install-Module is unavailable. Install PowerShellGet, then run: $installCommand"
        }
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            throw "Automatic ImportExcel installation failed. Run '$installCommand' with the SmartM365 execution account. $($_.Exception.Message)"
        }
        $module = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) { throw 'ImportExcel installation completed but the module is still unavailable in PSModulePath.' }
        WriteLog -Message ("ImportExcel installed automatically: version={0}; path={1}" -f $module.Version,$module.Path) -Level SUCCESS
    }
    Import-Module -Name $module.Path -Force -ErrorAction Stop
    WriteLog -Message ("ImportExcel module loaded: version={0}; path={1}" -f $module.Version,$module.Path) -Level INFO
}
function New-TeamsTimestampedWorkbook {
    param(
        [Parameter(Mandatory)][object[]]$CsvFiles,
        [Parameter(Mandatory)][string]$Path
    )

    Import-Module ImportExcel -ErrorAction Stop
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
    param([object[]]$AlertRows,[hashtable]$Summary,[string]$Worst,[datetime]$Started,[datetime]$Ended,[string]$FileLinksHtml='')
    $color=@{OK='#107c10';Warning='#ff8c00';Critical='#d13438'}
    $sb=[Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Segoe UI,Arial;background:#f5f8fb;color:#1f2937;padding:24px}.card{background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin:0 0 16px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dde7f0;padding:7px;font-size:12px;text-align:left;vertical-align:top}th{background:#eef6fc}.rowWarning{background:#fff7e6}.rowCritical{background:#fde7e9}.pill{color:#fff;border-radius:999px;padding:4px 10px;font-weight:600}.kpi td{background:#f8fafc}.kpiLabel{font-size:11px;color:#64748b;text-transform:uppercase}.kpiValue{font-size:20px;font-weight:700;color:#0f172a}</style></head><body>')
    [void]$sb.AppendLine(("<div class='card'><h1>Microsoft Teams Inventory <span class='pill' style='background:{0}'>{1}</span></h1><p>RunId: {2}<br>Machine: {3}<br>Started UTC: {4}<br>Ended UTC: {5}<br>Duration: {6}<br>Teams processed: {7}</p></div>" -f $color[$Worst],$Worst,$RunId,$env:COMPUTERNAME,(IsoUtc $Started),(IsoUtc $Ended),((New-TimeSpan -Start $Started -End $Ended).ToString()),$Summary.TotalTeams))
    $kpiKeys=@('TotalTeams','ActiveTeams','InactiveTeams','ArchivedTeams','PublicTeams','PrivateTeams','TeamsWithGuests','MemberRows','ChannelRows','GuestRows','CriticalCount','WarningCount')
    $kpiCells=@($kpiKeys|ForEach-Object{if($Summary.ContainsKey($_)){"<td><div class='kpiLabel'>{0}</div><div class='kpiValue'>{1}</div></td>" -f [Net.WebUtility]::HtmlEncode($_),[Net.WebUtility]::HtmlEncode([string]$Summary[$_])}})
    [void]$sb.AppendLine(("<div class='card'><h2>Global summary</h2><table class='kpi'><tr>{0}</tr></table></div>" -f ($kpiCells -join '')))
    [void]$sb.AppendLine(("<div class='card'><b>Summary</b>: Total={0}; Active={1}; Inactive={2}; Archived={3}; Public={4}; Private={5}; TeamsWithGuests={6}; Critical={7}; Warnings={8}</div>" -f $Summary.TotalTeams,$Summary.ActiveTeams,$Summary.InactiveTeams,$Summary.ArchivedTeams,$Summary.PublicTeams,$Summary.PrivateTeams,$Summary.TeamsWithGuests,$Summary.CriticalCount,$Summary.WarningCount))
    [void]$sb.AppendLine('<div class="card"><h2>Critical and warning findings</h2><table><tr><th>Status</th><th>Team</th><th>Check</th><th>Value</th><th>Threshold</th><th>Details</th></tr>')
    foreach($a in @($AlertRows|Sort-Object @{Expression={if($_.Status-eq'Critical'){0}else{1}}},TeamDisplayName,Check|Select-Object -First 200)){
        [void]$sb.AppendLine(("<tr class='row{0}'><td>{0}</td><td>{1}</td><td>{2}</td><td>{3} {4}</td><td>{5}</td><td>{6}</td></tr>" -f $a.Status,[Net.WebUtility]::HtmlEncode($a.TeamDisplayName),[Net.WebUtility]::HtmlEncode($a.Check),[Net.WebUtility]::HtmlEncode([string]$a.NumericValue),[Net.WebUtility]::HtmlEncode([string]$a.TextValue),[Net.WebUtility]::HtmlEncode($a.Threshold),[Net.WebUtility]::HtmlEncode($a.Details)))
    }
    [void]$sb.AppendLine('</table></div>')
    if (-not [string]::IsNullOrWhiteSpace($FileLinksHtml)) { [void]$sb.AppendLine($FileLinksHtml) }
    [void]$sb.AppendLine('</body></html>')
    $sb.ToString()
}
if([string]::IsNullOrWhiteSpace($OutputPath)){$OutputPath=[string](Get-ConfigValue 'TeamsInventoryCsvLogFolderPath' '{{DataAllRootPath}}\M365\Teams\Inventory')}
$LatestCsvFolderPath=[string](Get-ConfigValue 'LatestCsvFolderPath' $TenantContext.LatestCsvFolderPath); $WeeklyHistoryFolderPath=[string](Get-ConfigValue 'WeeklyHistoryFolderPath' (Join-Path $OutputPath 'WeeklyHistory')); $WeeklyHistoryRetentionWeeks=[int](Get-ConfigValue 'WeeklyHistoryRetentionWeeks' 52); $EnableWeeklyHistory=[bool](Get-ConfigValue 'EnableWeeklyHistory' $true)
if(-not$PSBoundParameters.ContainsKey('RequireSensitivityLabel')){$RequireSensitivityLabel=[bool](Get-ConfigValue 'RequireSensitivityLabel' $false)}; if(-not$PSBoundParameters.ContainsKey('IncludeChannelOwners')){$IncludeChannelOwners=[bool](Get-ConfigValue 'IncludeChannelOwners' $false)}
$global:RetentionMaxCSV=[int](Get-ConfigValue 'RetentionMaxCSV' 30); $global:RetentionMaxLogs=[int](Get-ConfigValue 'RetentionMaxLogs' 30)
$global:EnableSharePointUpload=[bool](Get-ConfigValue 'EnableSharePointUpload' $false); $global:SharePointSiteHostname=[string](Get-ConfigValue 'SharePointSiteHostname' ''); $global:SharePointSitePath=[string](Get-ConfigValue 'SharePointSitePath' ''); $global:SharePointLibraryDisplayName=[string](Get-ConfigValue 'SharePointLibraryDisplayName' 'Documents'); $global:SharePointTargetFolderPath=[string](Get-ConfigValue 'SharePointTargetFolderPath' '')
$AppId=[string](Get-ConfigValue 'AppId' ''); $TenantId=[string](Get-ConfigValue 'TenantId' ''); $OrgDomain=[string](Get-ConfigValue 'OrgDomain' ''); $Thumb=[string](Get-ConfigValue 'Thumbprint' (Get-ConfigValue 'Thumb' ''))
$global:AppId=$AppId; $global:TenantId=$TenantId; $global:OrgDomain=$OrgDomain; $global:Thumb=$Thumb; $global:Thumbprint=$Thumb
$teamColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','Description','Visibility','CreatedDateTimeUtc','Classification','SensitivityLabel','IsArchived','OwnerCount','MemberCount','GuestCount','StandardChannelCount','PrivateChannelCount','SharedChannelCount','LastActivityDateUtc','InactiveDays','StorageUsedGB','StorageQuotaGB','StorageQuotaPercent','Status','NumericValue','TextValue','Threshold','Details')
$memberColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','UserId','DisplayName','UserPrincipalName','Mail','UserType','Role','Status','NumericValue','TextValue','Threshold','Details')
$channelColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','ChannelId','ChannelDisplayName','MembershipType','CreatedDateTimeUtc','PrivateChannelOwners','Status','NumericValue','TextValue','Threshold','Details')
$guestColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','UserId','DisplayName','UserPrincipalName','Mail','Status','NumericValue','TextValue','Threshold','Details')
try{
 $CurrentOperation='Initialize script environment'; $OutputPath=InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $ScriptBaseName; Start-Transcript -Path $global:logTranscriptFile -Append|Out-Null; Write-SmartM365LoadedModuleVersions; WriteLog -Message "Starting $TaskName"
 $CurrentOperation='Connect Microsoft Graph'; Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true; $conn=Connect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -AppId $AppId -Thumbprint $Thumb -TenantId $TenantId -Organization $OrgDomain -GraphScopes @('Team.ReadBasic.All','TeamMember.Read.All','Channel.ReadBasic.All','Group.Read.All','Reports.Read.All'); if(-not$conn.GraphConnected){throw 'Microsoft Graph app-only connection failed.'}
 $CurrentOperation='Ensure ImportExcel module'; Ensure-TeamsImportExcelModule
 $CurrentOperation='Run preflight'; Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('Microsoft.Graph.Authentication','ImportExcel') -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('Team.ReadBasic.All','TeamMember.Read.All','Channel.ReadBasic.All','Group.Read.All','Reports.Read.All','Sites.Read.All') -GraphProbeUris @('https://graph.microsoft.com/v1.0/organization','https://graph.microsoft.com/v1.0/groups?$top=1')|Out-Null
 $CurrentOperation='Load tenant metadata'; $org=Invoke-Graph -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName' -Operation 'Get organization'; $TenantName=[string]@($org.value)[0].displayName; if([string]::IsNullOrWhiteSpace($TenantName)){$TenantName=$Tenant}
 $activityById=@{}; foreach($r in (Get-ReportRow -ReportName 'getTeamsTeamActivityDetail' -Period 'D180')){$id=[string](Prop $r @('Team Id','TeamId','Team ID')); if($id){$activityById[$id]=$r}}
 $teamFilter=[uri]::EscapeDataString("resourceProvisioningOptions/Any(x:x eq 'Team')"); $teamsUri="https://graph.microsoft.com/v1.0/groups?`$filter=$teamFilter&`$select=id,displayName,description,visibility,createdDateTime,classification,assignedLabels,mail,webUrl&`$top=999"; $teams=@(Get-GraphCollection -Uri $teamsUri -Operation 'Get team groups'); if($MaxTeams-gt 0){$teams=@($teams|Select-Object -First $MaxTeams)}; WriteLog -Message ("Teams discovered: {0}" -f $teams.Count)
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
  $status='OK'; $notes=New-Object 'System.Collections.Generic.List[string]'; if($owners.Count-eq 0){$status='Critical'; [void]$notes.Add('No owner'); Add-Alert $teamId $teamName Critical Owners $owners.Count 'NoOwner' "Owners >= $MinOwners" 'Team has no owner.'}elseif($owners.Count-lt$MinOwners){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Owner count below threshold'); Add-Alert $teamId $teamName Warning Owners $owners.Count 'LowOwnerCount' "Owners >= $MinOwners" 'Team has too few owners.'}
  if($inactive-ne'' -and [double]$inactive-gt$InactiveDays -and -not$archived){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Inactive team'); Add-Alert $teamId $teamName Warning Inactivity $inactive $last "<= $InactiveDays days" 'Team is a candidate for archival.'}
  if($quotaPct-ne'' -and [double]$quotaPct-gt$QuotaCriticalPercent){$status='Critical'; [void]$notes.Add('Storage quota critical'); Add-Alert $teamId $teamName Critical StorageQuotaPercent $quotaPct $storageDetail "<= $QuotaCriticalPercent percent" 'SharePoint storage quota usage is above threshold.'}
  if($guests.Count-gt$GuestWarningThreshold){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('High guest count'); Add-Alert $teamId $teamName Warning GuestCount $guests.Count Guests "<= $GuestWarningThreshold" 'Team has many external guests.'}
  if([string]$g.visibility-eq'Public' -and -not[string]::IsNullOrWhiteSpace($label)){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Public team with sensitivity label'); Add-Alert $teamId $teamName Warning PublicSensitiveLabel 1 $label 'Review public sensitive teams' 'Public team has a sensitivity/classification label.'}
  if($RequireSensitivityLabel -and [string]::IsNullOrWhiteSpace($label)){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Missing sensitivity label'); Add-Alert $teamId $teamName Warning MissingSensitivityLabel 0 NoLabel 'Sensitivity label required' 'Tenant policy expects labels.'}
  [void]$TeamsRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;Description=[string]$g.description;Visibility=[string]$g.visibility;CreatedDateTimeUtc=(IsoUtc $g.createdDateTime);Classification=[string]$g.classification;SensitivityLabel=$label;IsArchived=[string]$archived;OwnerCount=$owners.Count;MemberCount=$members.Count;GuestCount=$guests.Count;StandardChannelCount=$standard;PrivateChannelCount=$private;SharedChannelCount=$shared;LastActivityDateUtc=$last;InactiveDays=(Num $inactive);StorageUsedGB=(Num $usedGb);StorageQuotaGB=(Num $quotaGb);StorageQuotaPercent=(Num $quotaPct);Status=$status;NumericValue=(Num $members.Count);TextValue="Owners=$($owners.Count); Members=$($members.Count); Guests=$($guests.Count)";Threshold="Owners >= $MinOwners; inactive <= $InactiveDays days; storage <= $QuotaCriticalPercent percent; guests <= $GuestWarningThreshold";Details=(JoinVals $notes)})
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
 if(-not$DryRun){$workbookUpload=Invoke-SmartM365SharePointCsvUpload -LocalFilePath $workbookPath; if($workbookUpload){Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath $workbookPath -RetentionDays 7|Out-Null}}
 if($EnableWeeklyHistory-and-not$DryRun){Add-SmartM365WeeklyHistory -SourceCsvPaths $GeneratedCsvPaths.ToArray() -HistoryRootPath $WeeklyHistoryFolderPath -RetentionWeeks $WeeklyHistoryRetentionWeeks -HistoryLabel 'Microsoft Teams inventory' -UploadChangedFilesOnly|Out-Null}elseif($DryRun){WriteLog -Message 'DryRun enabled: WeeklyHistory skipped.' -Level INFO}
 $teamArray=$TeamsRows.ToArray(); $memberArray=$MembersRows.ToArray(); $channelArray=$ChannelsRows.ToArray(); $guestArray=$GuestsRows.ToArray(); $alertArray=$Alerts.ToArray()
 $summary=@{TotalTeams=$teamArray.Count;ActiveTeams=@($teamArray|Where-Object{$_.IsArchived-ne'True' -and ([string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-or [double]$_.InactiveDays-le$InactiveDays)}).Count;InactiveTeams=@($teamArray|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-and [double]$_.InactiveDays-gt$InactiveDays}).Count;ArchivedTeams=@($teamArray|Where-Object{$_.IsArchived-eq'True'}).Count;PublicTeams=@($teamArray|Where-Object{$_.Visibility-eq'Public'}).Count;PrivateTeams=@($teamArray|Where-Object{$_.Visibility-eq'Private'}).Count;TeamsWithGuests=@($teamArray|Where-Object{[int]$_.GuestCount-gt 0}).Count;CriticalCount=@($alertArray|Where-Object Status -eq Critical).Count;WarningCount=@($alertArray|Where-Object Status -eq Warning).Count;MemberRows=$memberArray.Count;ChannelRows=$channelArray.Count;GuestRows=$guestArray.Count}
 $mailFileLinks=New-TeamsSharePointLinksHtml -Paths (@($timestampedCsvFiles.Path)+@($workbookPath))
 $worst=WorstStatus $alertArray; $subject="[$($worst.ToUpperInvariant())] Microsoft Teams Inventory - $RunDateUtc"; $html=ConvertTo-HtmlReport -AlertRows $alertArray -Summary $summary -Worst $worst -Started $RunStarted -Ended (Get-Date) -FileLinksHtml $mailFileLinks
 if($DryRun){WriteLog -Message 'DryRun enabled: daily summary email skipped.' -Level INFO}else{$dailySummaryMarkerPath=Join-Path -Path (Split-Path -Path $global:LogTextFile -Parent) -ChildPath "$ScriptBaseName-DailySummary-LastSent.txt"; $dailySummarySent=Invoke-TeamsDailySummaryMail -MarkerPath $dailySummaryMarkerPath -SendAction {Send-SmartM365Mail -Subject $subject -BodyHtml $html}; if($dailySummarySent){WriteLog -Message ("Daily Teams summary email sent: {0}" -f $subject) -Level SUCCESS}}
 $result="Teams=$($summary.TotalTeams); Critical=$($summary.CriticalCount); Warnings=$($summary.WarningCount); Members=$($memberArray.Count); Channels=$($channelArray.Count); Guests=$($guestArray.Count)"; try{Stop-Transcript|Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile}catch{$null=$_}; WriteLog -Message ("Result summary: $result") -Level INFO; Write-Host "Teams inventory completed. Status=$worst; $result"; Complete-SmartM365ExecutionContext -Status $(if($worst-eq'OK'){'Success'}else{'CompletedWithWarnings'})
}catch{ $err=$_; try{WriteLog -Message ("Teams inventory failed during {0}: {1}" -f $CurrentOperation,$err.Exception.Message) -Level ERROR}catch{$null=$_}; try{Stop-Transcript|Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile}catch{$null=$_}; try{Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $err -FailureStage $CurrentOperation}catch{$null=$_}; throw }
