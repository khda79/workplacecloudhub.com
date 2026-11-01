<#
.SYNOPSIS
    Active Directory forest health check for PowerShell 7 and RSAT ActiveDirectory.
.VERSION
1.0.37
.DESCRIPTION
    Discovers every domain with Get-ADForest, audits domain controllers and domain health,
    exports a flat Power BI-ready CSV, and sends an HTML summary email on warnings or critical alerts.
    Minimum permissions: PowerShell 7+, RSAT ActiveDirectory module, and read access to the AD forest/domains. Remote DC admin checks require explicit T0/admin rights only when -EnableRemoteDcAdminChecks is used.

.PARAMETER NonBlockingDomainErrors
    Domains whose transient replication-query connectivity failures may remain non-blocking after retries. Their missing measurements are marked NotMeasured and the forest query coverage is marked Warning. Other domains and error types remain blocking.

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ActiveDirectory RSAT/Windows Server module.
    Minimum permissions: read access to the AD forest/domains through Get-ADForest, Get-ADDomain, Get-ADDomainController, Get-ADReplication* and related read-only AD cmdlets.
    Optional: -EnableRemoteDcAdminChecks requires explicit tier-0/admin rights for remote DC service, event, disk and AD database checks.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
#>
#requires -Version 7.0
#requires -Modules ActiveDirectory
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Parameters are consumed by script-scope helper functions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='This read-only health script does not change AD state.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='Final console status is intentional for command-line use.')]
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$OutputFolder = '',
    [switch]$AppendHistory,
    [string]$HistoryCsvPath,
    [string[]]$To = @(),
    [string]$From = '',
    [string]$SmtpServer = '',
    [int]$SmtpPort = 25,
    [switch]$UseSsl,
    [pscredential]$SmtpCredential,
    [switch]$AlwaysSend,
    [int]$RetryCount = 2,
    [int]$RetryDelaySeconds = 3,
    [int]$TcpTimeoutMs = 2000,
    [double]$TimeOffsetWarningMinutes = 5,
    [double]$DiskFreePercentCritical = 15,
    [double]$DiskFreeGbCritical = 5,
    [int]$ReplicationDelayWarningHours = 24,
    [int]$DfsrBacklogWarningCount = 100,
    [int]$DfsrBacklogCriticalCount = 1000,
    [switch]$SkipDfsrBacklog,
    [switch]$EnableRemoteDcAdminChecks,
    [string[]]$NonBlockingDomainErrors = @(),
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

if ($MaxItems -gt 0) {
    throw "-MaxItems is not supported by SmartM365-ActiveDirectory-HealthCheck because it produces a forest/domain health report where partial data would be misleading."
}
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::InvariantCulture
$RunStarted = Get-Date
$RunDateUtc = $RunStarted.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)
$RunId = [guid]::NewGuid().ToString()
$Rows = [System.Collections.ArrayList]::new()
$DomainFacts = [System.Collections.ArrayList]::new()
$script:PrivilegedUserPasswordNeverExpiresCache = @{}
$ScriptBaseName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ScriptVersion = "1.0.37"
$TaskName = "$ScriptBaseName v$ScriptVersion"
$TenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        foreach ($candidate in @((Join-Path $d 'SmartM365-TenantContext.ps1'), (Join-Path $d 'Config\SmartM365-TenantContext.ps1'))) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $TenantContextPath
$TenantContext = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$TenantContextDirectory = Split-Path -Path $TenantContextPath -Parent
$SmartM365ProjectRoot = if ((Split-Path -Path $TenantContextDirectory -Leaf) -ieq 'Config') {
    Split-Path -Path $TenantContextDirectory -Parent
}
else {
    $TenantContextDirectory
}
$CoreModulePath = Join-Path $SmartM365ProjectRoot 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $CoreModulePath -MinimumVersion '1.0.121' -Force -ErrorAction Stop
$LocalConfigPath = Join-Path $PSScriptRoot ("$ScriptBaseName.local.json"); $LocalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $LocalConfigPath
$LegacyLocalConfigPath = Join-Path $PSScriptRoot 'AD-HealthCheck.local.json'; $LegacyLocalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $LegacyLocalConfigPath
$LocalTemplatePath = (Get-SmartM365JsonTemplateName -Path $LocalConfigPath)
if (-not (Test-Path -LiteralPath $LocalConfigPath) -and (Test-Path -LiteralPath $LegacyLocalConfigPath)) {
    $LocalConfigPath = $LegacyLocalConfigPath
    $LocalTemplatePath = (Get-SmartM365JsonTemplateName -Path $LocalConfigPath)
}
$ScriptConfig = Read-SmartM365ScriptConfiguration -ConfigurationPath $LocalConfigPath -MissingPolicy CreateIfTemplate
function Resolve-ConfigToken {
    param([AllowNull()]$Value)
    Resolve-SmartM365ConfigurationTokens -Value $Value -Config $TenantContext
}
function Get-LocalConfigValue {
    param([Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    Get-SmartM365ConfigurationValue -Name $Name -LocalConfig $ScriptConfig -EffectiveConfig $TenantContext -DefaultValue $DefaultValue -EffectiveFallback Null -ResolveDefault
}
$global:AppId = [string](Get-LocalConfigValue 'AppId' '')
$global:TenantId = [string](Get-LocalConfigValue 'TenantId' '')
$global:Thumb = [string](Get-LocalConfigValue 'Thumb' '')
$global:Thumbprint = [string](Get-LocalConfigValue 'Thumbprint' $global:Thumb)
$global:EnableSharePointUpload = [bool](Get-LocalConfigValue 'EnableSharePointUpload' $false)
$global:SharePointSiteHostname = [string](Get-LocalConfigValue 'SharePointSiteHostname' '')
$global:SharePointSitePath = [string](Get-LocalConfigValue 'SharePointSitePath' '')
$global:SharePointLibraryDisplayName = [string](Get-LocalConfigValue 'SharePointLibraryDisplayName' 'Documents')
$global:SharePointTargetFolderPath = [string](Get-LocalConfigValue 'SharePointTargetFolderPath' '')
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = [string](Get-LocalConfigValue 'ADHealthCheckCsvLogFolderPath' '{{DataAllRootPath}}\ActiveDirectory\HealthCheck') }
$LatestCsvFolderPath = [string](Get-LocalConfigValue 'LatestCsvFolderPath' $TenantContext.LatestCsvFolderPath)
if (-not $PSBoundParameters.ContainsKey('AlwaysSend')) { $AlwaysSend = [bool](Get-LocalConfigValue 'AlwaysSend' $false) }
if (-not $PSBoundParameters.ContainsKey('AppendHistory')) { $AppendHistory = [bool](Get-LocalConfigValue 'AppendHistory' $false) }
if ([string]::IsNullOrWhiteSpace($HistoryCsvPath)) { $HistoryCsvPath = [string](Get-LocalConfigValue 'HistoryCsvPath' '{{DataAllRootPath}}\ActiveDirectory\HealthCheck\AD_HealthCheck_History.csv') }
if (-not $PSBoundParameters.ContainsKey('EnableRemoteDcAdminChecks')) { $EnableRemoteDcAdminChecks = [bool](Get-LocalConfigValue 'EnableRemoteDcAdminChecks' $false) }
if (-not $PSBoundParameters.ContainsKey('NonBlockingDomainErrors')) { $NonBlockingDomainErrors = @((Get-LocalConfigValue 'NonBlockingDomainErrors' @())) }
$NonBlockingDomainErrorSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($configuredDomain in @($NonBlockingDomainErrors)) {
    $domainName = ([string]$configuredDomain).Trim().TrimEnd('.')
    if ($domainName -notmatch '^(?=.{1,253}$)[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$') {
        throw "Invalid domain in NonBlockingDomainErrors: '$configuredDomain'."
    }
    [void]$NonBlockingDomainErrorSet.Add($domainName)
}
if ([string]::IsNullOrWhiteSpace($From)) { $From = [string](Get-LocalConfigValue 'From' '') }
if ([string]::IsNullOrWhiteSpace($SmtpServer)) { $SmtpServer = [string](Get-LocalConfigValue 'SmtpServer' '') }
$SendMailMode = [string](Get-LocalConfigValue 'SendMailMode' '')
$Cc = [string](Get-LocalConfigValue 'Cc' '')
if ($To.Count -eq 0) {
    $configuredTo = [string](Get-LocalConfigValue 'To' '')
    if ([string]::IsNullOrWhiteSpace($configuredTo)) { $configuredTo = [string](Get-LocalConfigValue 'ErrorMailTo' '') }
    if (-not [string]::IsNullOrWhiteSpace($configuredTo)) { $To = @($configuredTo -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
}
function ConvertTo-IsoUtc([datetime]$d){ if($null -eq $d -or $d -eq [datetime]::MinValue){''}else{$d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)} }
function Num($v){ if($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)){''}else{ try{([double]$v).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)}catch{[string]$v} } }
function Ms($s){ [int64][math]::Max(0,((Get-Date)-$s).TotalMilliseconds) }
function Add-Row{
    param(
        [AllowNull()][object]$Forest,
        [AllowNull()][object]$Domain,
        [AllowNull()][object]$DC,
        [string]$Category,
        [string]$Check,
        [ValidateSet('OK','Warning','Critical','NotMeasured')][string]$Status,
        [AllowNull()][object]$NumericValue,
        [AllowNull()][object]$TextValue,
        [AllowNull()][object]$Threshold,
        [AllowNull()][object]$Details,
        [int64]$DurationMs
    )
    $row = [ordered]@{
        RunId = $RunId
        RunDateUtc = $RunDateUtc
        Forest = [string]$Forest
        Domain = [string]$Domain
        DC = [string]$DC
        Category = $Category
        Check = $Check
        Status = $Status
        NumericValue = (Num $NumericValue)
        TextValue = [string]$TextValue
        Threshold = [string]$Threshold
        Details = [string]$Details
        DurationMs = $DurationMs
    }
    [void]$Rows.Add([pscustomobject]$row)
}
function Get-RowSnapshot{
    @($Rows.ToArray())
}
function Invoke-Retry([scriptblock]$ScriptBlock){
    for($i=1;$i -le [math]::Max(1,$RetryCount);$i++){
        try{ return & $ScriptBlock }catch{ $m=[string]$_.Exception.Message; $transient=$m -match 'timeout|temporar|server is not operational|RPC server|unavailable|busy|could not be contacted|network path|WinRM|timed out'; if(-not $transient -or $i -ge $RetryCount){throw}; Start-Sleep -Seconds $RetryDelaySeconds }
    }
}
function Test-IsTransientAdConnectivityError([Management.Automation.ErrorRecord]$ErrorRecord) {
    for ($exception = $ErrorRecord.Exception; $null -ne $exception; $exception = $exception.InnerException) {
        if ($exception.Message -match 'timeout|temporar|server is not operational|RPC server|unavailable|busy|could not be contacted|unable to contact|network path|WinRM|timed out') { return $true }
    }
    return $false
}
function Add-ForestReplicationAgeRows([string]$ForestName,[string]$RootDomain,[string[]]$Domains,[int]$TombstoneDays) {
    $oldest = 0.0
    $unavailableDomains = [Collections.Generic.List[string]]::new()
    foreach ($domainName in @($Domains | Sort-Object)) {
        $started = Get-Date
        try {
            $failures = @(Invoke-Retry { Get-ADReplicationFailure -Target $domainName -Scope Domain -ErrorAction Stop })
            foreach ($failure in $failures) {
                if ($failure.FirstFailureTime -and $failure.FirstFailureTime -ne [datetime]::MinValue) {
                    $days = ((Get-Date).ToUniversalTime() - $failure.FirstFailureTime.ToUniversalTime()).TotalDays
                    if ($days -gt $oldest) { $oldest = $days }
                }
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            if (-not $NonBlockingDomainErrorSet.Contains($domainName) -or -not (Test-IsTransientAdConnectivityError $_)) {
                throw "Replication failure query failed for domain '$domainName': $errorMessage"
            }
            [void]$unavailableDomains.Add($domainName)
            WriteLog -Message ("Non-blocking replication query error for domain '{0}': {1}. The measurement remains unavailable." -f $domainName,$errorMessage) -Level 'WARNING'
            Add-Row $ForestName $domainName '' Replication ReplicationFailureQuery NotMeasured '' Error 'Domain replication failures readable' $errorMessage (Ms $started)
        }
    }
    if ($unavailableDomains.Count -gt 0) {
        Add-Row $ForestName $RootDomain '' Replication DomainQueryCoverage Warning ($Domains.Count - $unavailableDomains.Count) ("Unavailable={0}" -f $unavailableDomains.Count) ("All {0} domains queried" -f $Domains.Count) ("Unavailable domains: {0}" -f ($unavailableDomains -join ', ')) 0
    }
    $status = if ($oldest -gt $TombstoneDays) { 'Critical' } elseif ($oldest -gt ($TombstoneDays * .8)) { 'Warning' } elseif ($unavailableDomains.Count -gt 0) { 'NotMeasured' } else { 'OK' }
    $value = if ($status -eq 'NotMeasured') { '' } else { [math]::Round($oldest,2) }
    $details = if ($unavailableDomains.Count -gt 0) { "Incomplete forest measurement; unavailable domains: $($unavailableDomains -join ', '). Observed oldest age is a lower bound." } else { 'Compared with forest tombstone lifetime' }
    Add-Row $ForestName $RootDomain '' Tombstone OldestReplicationFailureAgeDays $status $value "TombstoneLifetimeDays=$TombstoneDays" "Critical > $TombstoneDays days; Warning > 80 percent" $details 0
}
function Test-Port([string]$Computer,[int]$Port){
    $c=[Net.Sockets.TcpClient]::new(); try{ $a=$c.BeginConnect($Computer,$Port,$null,$null); if(-not $a.AsyncWaitHandle.WaitOne($TcpTimeoutMs,$false)){return $false}; $c.EndConnect($a); $true }catch{$false}finally{try{$c.Close()}catch{ $null = $_ }}
}
function Worst($r){ if($r|Where-Object Status -eq Critical){'Critical'}elseif($r|Where-Object Status -eq Warning){'Warning'}elseif($r|Where-Object Status -eq NotMeasured){'NotMeasured'}else{'OK'} }
function Rank($s){ if($s -eq 'Critical'){0}elseif($s -eq 'Warning'){1}elseif($s -eq 'NotMeasured'){2}else{3} }
function ConvertTo-HtmlSafe($s){ [Net.WebUtility]::HtmlEncode([string]$s) }
function Get-ObjectPropertyValue([object]$Object,[string[]]$Names){
    foreach($name in $Names){
        $prop = if($null -ne $Object){$Object.PSObject.Properties[$name]}else{$null}
        if($prop -and $null -ne $prop.Value){return [string]$prop.Value}
    }
    ''
}
function Get-TimeOffsetMinute([string]$DC){
    $out=& w32tm.exe /stripchart /computer:$DC /samples:3 /dataonly 2>&1
    $vals=@(); foreach($line in $out){ foreach($m in [regex]::Matches([string]$line,'([+-]?\d+(?:[\.,]\d+)?)s')){ $raw=$m.Groups[1].Value.Replace(',','.'); $d=0.0; if([double]::TryParse($raw,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){$vals += [math]::Abs($d)} } }
    if($vals.Count -eq 0){ return [pscustomobject]@{ OffsetMinutes = $null; Error = "Unable to parse w32tm output: $($out -join ' | ')" } }
    [pscustomobject]@{ OffsetMinutes = (($vals|Measure-Object -Maximum).Maximum/60); Error = '' }
}
function Get-ADGroupMembersBySid([string]$Sid,[string]$Server){
    $group = Get-ADGroup -Identity $Sid -Server $Server -ErrorAction SilentlyContinue
    if (-not $group) { return @() }
    @(Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -Server $Server -ErrorAction SilentlyContinue)
}
function Get-DomainDnsNameFromDistinguishedName([string]$DistinguishedName){
    $domainComponents = @(
        [regex]::Matches($DistinguishedName, '(?i)(?:^|,)\s*DC=([^,]+)') |
            ForEach-Object { $_.Groups[1].Value.Trim() }
    )
    if ($domainComponents.Count -eq 0) { return '' }
    $domainComponents -join '.'
}
function Get-PrivilegedUserPasswordNeverExpires([string]$DistinguishedName,[string]$DefaultServer){
    $cacheKey = $DistinguishedName.Trim().ToLowerInvariant()
    if ($script:PrivilegedUserPasswordNeverExpiresCache.ContainsKey($cacheKey)) {
        return $script:PrivilegedUserPasswordNeverExpiresCache[$cacheKey]
    }

    $owningDomain = Get-DomainDnsNameFromDistinguishedName -DistinguishedName $DistinguishedName
    $targetServer = if ([string]::IsNullOrWhiteSpace($owningDomain)) { $DefaultServer } else { $owningDomain }
    try {
        $user = Invoke-Retry { Get-ADUser -Identity $DistinguishedName -Server $targetServer -Properties PasswordNeverExpires -ErrorAction Stop }
        $result = [pscustomobject]@{
            Resolved = $true
            PasswordNeverExpires = [bool]$user.PasswordNeverExpires
            Server = $targetServer
            Error = ''
        }
    }
    catch {
        $result = [pscustomobject]@{
            Resolved = $false
            PasswordNeverExpires = $false
            Server = $targetServer
            Error = $_.Exception.Message
        }
    }
    $script:PrivilegedUserPasswordNeverExpiresCache[$cacheKey] = $result
    $result
}
function Get-ADDbVolume([string]$DC){
    try{ $v=Invoke-Command -ComputerName $DC -ScriptBlock { $o=& ntdsutil.exe 'activate instance ntds' 'files' 'info' 'quit' 'quit' 2>&1; $l=$o|Where-Object{[string]$_ -match 'Database directory|DB Path|Database path'}|Select-Object -First 1; if($l -and [string]$l -match '([A-Za-z]:)'){$matches[1]}else{$env:SystemDrive} } -ErrorAction Stop; if($v){return ([string]$v).TrimEnd('\')} }catch{ $null = $_ }
    $os=Get-CimInstance Win32_OperatingSystem -ComputerName $DC -ErrorAction Stop; ([string]$os.SystemDrive).TrimEnd('\')
}
function Get-Disk([string]$DC,[string]$DeviceId){
    $d=Get-CimInstance Win32_LogicalDisk -ComputerName $DC -Filter "DeviceID='$DeviceId'" -ErrorAction Stop
    if($null -eq $d -or [double]$d.Size -le 0){throw "Logical disk $DeviceId not found"}
    [pscustomobject]@{FreeGb=[math]::Round(([double]$d.FreeSpace/1GB),2);FreePct=[math]::Round(([double]$d.FreeSpace/[double]$d.Size*100),2);SizeGb=[math]::Round(([double]$d.Size/1GB),2)}
}
function Get-DfsrBacklog([string]$Src,[string]$Dst){
    if(-not (Get-Command dfsrdiag.exe -ErrorAction SilentlyContinue)){
        return [pscustomobject]@{Count=$null;Reason='DfsrdiagUnavailable: dfsrdiag.exe was not found on the collector host.'}
    }
    try {
        $output=@(& dfsrdiag.exe backlog /rgname:'Domain System Volume' /rfname:'SYSVOL Share' /smem:$Src /rmem:$Dst 2>&1)
        $exitCode=[int]$LASTEXITCODE
    } catch {
        return [pscustomobject]@{Count=$null;Reason="DfsrdiagInvocationFailed: $($_.Exception.Message)"}
    }
    $excerpt=(@($output | ForEach-Object { [string]$_ }) -join ' | ') -replace '\s+', ' '
    if($excerpt.Length -gt 240){$excerpt=$excerpt.Substring(0,240)+'...'}
    if($exitCode -ne 0){return [pscustomobject]@{Count=$null;Reason="DfsrdiagExitCode=$exitCode; Output=$excerpt"}}
    foreach($line in $output){
        $text=[string]$line
        if($text -match 'Backlog File Count\s*:\s*(\d+)'){return [pscustomobject]@{Count=[int]$matches[1];Reason=''}}
        if($text -match 'No Backlog'){return [pscustomobject]@{Count=0;Reason=''}}
    }
    return [pscustomobject]@{Count=$null;Reason="DfsrdiagOutputUnrecognized: $excerpt"}
}
function Send-ReportMail([string]$Subject,[string]$Body,[ValidateSet('Report','Error')][string]$MailPurpose='Report'){
    $mailParams = @{ Subject = $Subject; BodyHtml = $Body; VerboseLog = $true; MailPurpose = $MailPurpose }
    if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) { $mailParams.SmtpServer = $SmtpServer }
    if (-not [string]::IsNullOrWhiteSpace($SendMailMode)) { $mailParams.SendMailMode = $SendMailMode }
    if (-not [string]::IsNullOrWhiteSpace($From)) { $mailParams.From = $From }
    if ($To.Count -gt 0) { $mailParams.To = ($To -join ';') }
    if (-not [string]::IsNullOrWhiteSpace($Cc)) { $mailParams.Cc = $Cc }
    if ($SmtpPort -ne 25) { $mailParams.SmtpPort = $SmtpPort }
    SendEmailHtmlReport @mailParams
}
function ConvertTo-ReportHtml([object[]]$r,[string]$status,[datetime]$started,[datetime]$ended,[string]$csv,[AllowNull()][object]$forestInfo=$null,[bool]$RemoteDcAdminChecksEnabled=$false){
    $r = @($r)
    $facts = @($DomainFacts.ToArray())
    $c = @{OK='#107c10';Warning='#ff8c00';Critical='#d13438';NotMeasured='#667085'}[$status]
    $criticalRows = @($r | Where-Object Status -eq Critical)
    $warningRows = @($r | Where-Object Status -eq Warning)
    $okRows = @($r | Where-Object Status -eq OK)
    $notMeasuredRows = @($r | Where-Object Status -eq NotMeasured)
    $find = @($r | Where-Object { $_.Status -in @('Critical','Warning') } | Sort-Object @{Expression={Rank $_.Status}},Domain,DC,Category,Check | Select-Object -First 200)
    $forestName = if ($forestInfo -and $forestInfo.Name) { [string]$forestInfo.Name } else { [string]($r | Where-Object Forest | Select-Object -ExpandProperty Forest -First 1) }
    $rootDomain = if ($forestInfo -and $forestInfo.RootDomain) { [string]$forestInfo.RootDomain } else { [string]($r | Where-Object Domain | Select-Object -ExpandProperty Domain -First 1) }
    if ([string]::IsNullOrWhiteSpace($forestName)) { $forestName = 'NotAvailable' }
    if ([string]::IsNullOrWhiteSpace($rootDomain)) { $rootDomain = 'NotAvailable' }
    $forestMode = if ($forestInfo -and $forestInfo.ForestMode) { [string]$forestInfo.ForestMode } else { 'NotAvailable' }
    $domains = if ($forestInfo -and $forestInfo.Domains) { @($forestInfo.Domains | Sort-Object) } else { @($r | Where-Object Domain | Select-Object -ExpandProperty Domain -Unique | Sort-Object) }
    $domainBadges = (@($domains) | ForEach-Object { '<span class="tag">' + (ConvertTo-HtmlSafe $_) + '</span>' }) -join ' '
    $dcNames = @($r | Where-Object { $_.Category -eq 'Connectivity' -and $_.Check -eq 'Ping' -and $_.DC } | Select-Object -ExpandProperty DC -Unique | Sort-Object)
    $gcCount = if ($facts.Count -gt 0) { ($facts | Measure-Object -Property GlobalCatalogCount -Sum).Sum } else { 'NotMeasured' }
    $lockedTotal = 0
    $lockedRows = @($r | Where-Object { $_.Category -eq 'DomainStats' -and $_.Check -eq 'LockedUserAccounts' })
    foreach ($row in $lockedRows) {
        $value = 0.0
        if ([double]::TryParse([string]$row.NumericValue,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$value)) { $lockedTotal += [int]$value }
    }
    if ($lockedRows.Count -eq 0) { $lockedTotal = 'NotMeasured' }
    $domainsWithCritical = @($criticalRows | Where-Object Domain | Select-Object -ExpandProperty Domain -Unique).Count
    $dcsWithCritical = @($criticalRows | Where-Object DC | Select-Object -ExpandProperty DC -Unique).Count
    $unreachableDcs = @($criticalRows | Where-Object { $_.Category -eq 'Connectivity' -and $_.DC } | Select-Object -ExpandProperty DC -Unique).Count
    $replicationCritical = @($criticalRows | Where-Object Category -eq Replication).Count
    $sysvolCritical = @($criticalRows | Where-Object { $_.Category -eq 'SYSVOL' -and $_.Check -in @('SYSVOL','NETLOGON') }).Count
    $b = [Text.StringBuilder]::new()
    [void]$b.AppendLine('<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Segoe UI,Arial;background:#f5f8fb;color:#1f2937;padding:24px}.card{background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin:0 0 16px}.muted{color:#52657a}.small{font-size:11px}.metric{font-size:22px;font-weight:700;color:#0f172a}.tag{display:inline-block;border:1px solid #cbd8e6;background:#f8fbfe;border-radius:999px;padding:3px 8px;margin:2px;font-size:11px}.pill{color:#fff;border-radius:999px;padding:4px 10px;font-weight:600}.grid{width:100%;border-collapse:separate;border-spacing:8px}.grid td{border:1px solid #dde7f0;background:#f8fbfe;border-radius:6px;padding:10px}.riskCritical{color:#b42318;font-weight:700}.riskWarning{color:#b54708;font-weight:700}.riskNotMeasured{color:#667085;font-weight:700}.riskOK{color:#067647;font-weight:700}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dde7f0;padding:7px;font-size:12px;text-align:left;vertical-align:top}th{background:#eef6fc}.rowOK{background:#f3fbf3}.rowWarning{background:#fff7e6}.rowCritical{background:#fde7e9}.rowNotMeasured{background:#f2f4f7}</style></head><body>')
    [void]$b.AppendLine(("<div class='card'><h1>Active Directory Health Check <span class='pill' style='background:{0}'>{1}</span></h1><p class='muted'>RunId: {2}<br>Machine: {3}<br>Started UTC: {4}<br>Ended UTC: {5}<br>Duration: {6}</p></div>" -f $c,$status,(ConvertTo-HtmlSafe $RunId),(ConvertTo-HtmlSafe $env:COMPUTERNAME),(ConvertTo-IsoUtc $started),(ConvertTo-IsoUtc $ended),(ConvertTo-HtmlSafe ((New-TimeSpan -Start $started -End $ended).ToString()))))
    [void]$b.AppendLine(("<div class='card'><h2>Forest recap</h2><table><tr><th>Forest</th><td>{0}</td><th>Root domain</th><td>{1}</td></tr><tr><th>Forest mode</th><td>{2}</td><th>Remote DC admin checks</th><td>{3}</td></tr><tr><th>Domains</th><td>{4}</td><th>Domain controllers</th><td>{5}</td></tr><tr><th>Global catalogs</th><td>{6}</td><th>Domains scanned</th><td>{7}</td></tr></table></div>" -f (ConvertTo-HtmlSafe $forestName),(ConvertTo-HtmlSafe $rootDomain),(ConvertTo-HtmlSafe $forestMode),$(if($RemoteDcAdminChecksEnabled){'Enabled'}else{'Disabled'}),@($domains).Count,$dcNames.Count,$gcCount,$domainBadges))
    [void]$b.AppendLine(("<div class='card'><h2>Operational recap</h2><table class='grid'><tr><td><div class='metric'>{0}</div><div>Total checks</div></td><td><div class='metric riskCritical'>{1}</div><div>Critical</div></td><td><div class='metric riskWarning'>{2}</div><div>Warning</div></td><td><div class='metric riskOK'>{3}</div><div>OK</div></td></tr><tr><td><div class='metric'>{4}</div><div>Domains with critical findings</div></td><td><div class='metric'>{5}</div><div>DCs with critical findings</div></td><td><div class='metric'>{6}</div><div>Unreachable DCs</div></td><td><div class='metric'>{7}</div><div>Locked user accounts</div></td></tr><tr><td><div class='metric'>{8}</div><div>Replication critical rows</div></td><td><div class='metric'>{9}</div><div>SYSVOL/NETLOGON critical rows</div></td><td><div class='metric'>{10}</div><div>Total DCs</div></td><td><div class='metric'>{11}</div><div>Domains</div></td></tr></table></div>" -f $r.Count,$criticalRows.Count,$warningRows.Count,$okRows.Count,$domainsWithCritical,$dcsWithCritical,$unreachableDcs,$lockedTotal,$replicationCritical,$sysvolCritical,$dcNames.Count,@($domains).Count))
    [void]$b.AppendLine(("<div class='card'><h2>Measurement coverage</h2><p><span class='riskNotMeasured'>{0} NotMeasured</span> checks were not executed or could not be measured. They are excluded from the OK count and do not imply healthy results.</p></div>" -f $notMeasuredRows.Count))
    [void]$b.AppendLine('<div class="card"><h2>FSMO roles</h2><table><tr><th>Scope</th><th>Domain</th><th>Role</th><th>Holder</th><th>Status</th></tr>')
    foreach ($x in @($r | Where-Object Category -eq FSMO | Sort-Object Domain,Check)) { $scope = if ($x.Check -in @('SchemaMaster','DomainNamingMaster')) { 'Forest' } else { 'Domain' }; [void]$b.AppendLine(("<tr class='row{0}'><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{0}</td></tr>" -f (ConvertTo-HtmlSafe $x.Status),(ConvertTo-HtmlSafe $scope),(ConvertTo-HtmlSafe $x.Domain),(ConvertTo-HtmlSafe $x.Check),(ConvertTo-HtmlSafe $x.DC))) }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Priority findings</h2><table><tr><th>Type</th><th>Name</th><th>Critical</th><th>Warning</th></tr>')
    $priorityGroups = @()
    $priorityGroups += @($r | Where-Object DC | Group-Object DC | ForEach-Object { $it = @($_.Group); [pscustomobject]@{ Type='DC'; Name=$_.Name; Critical=@($it | Where-Object Status -eq Critical).Count; Warning=@($it | Where-Object Status -eq Warning).Count } })
    $priorityGroups += @($r | Group-Object Category | ForEach-Object { $it = @($_.Group); [pscustomobject]@{ Type='Category'; Name=$_.Name; Critical=@($it | Where-Object Status -eq Critical).Count; Warning=@($it | Where-Object Status -eq Warning).Count } })
    foreach ($g in @($priorityGroups | Where-Object { $_.Critical -gt 0 -or $_.Warning -gt 0 } | Sort-Object @{Expression='Critical';Descending=$true},@{Expression='Warning';Descending=$true},Type,Name | Select-Object -First 20)) { [void]$b.AppendLine(("<tr><td>{0}</td><td>{1}</td><td class='riskCritical'>{2}</td><td class='riskWarning'>{3}</td></tr>" -f (ConvertTo-HtmlSafe $g.Type),(ConvertTo-HtmlSafe $g.Name),$g.Critical,$g.Warning)) }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Domain status summary</h2><table><tr><th>Domain</th><th>DCs</th><th>OK</th><th>Warning</th><th>Critical</th><th>NotMeasured</th><th>Locked users</th><th>FSMO</th><th>Replication</th><th>SYSVOL</th></tr>')
    foreach ($g in ($r | Where-Object Domain | Group-Object Domain | Sort-Object Name)) {
        $it = @($g.Group)
        # A failed domain scan can leave findings without facts or account measurements.
        $fact = $facts | Where-Object Domain -eq $g.Name | Select-Object -First 1
        $locked = $it | Where-Object { $_.Category -eq 'DomainStats' -and $_.Check -eq 'LockedUserAccounts' } | Select-Object -First 1
        $lockedValue = if ($locked) { $locked.NumericValue } else { 'NotMeasured' }
        $fsmoRows = @($it | Where-Object Category -eq FSMO)
        $repRows = @($it | Where-Object Category -eq Replication)
        $sysvolRows = @($it | Where-Object Category -eq SYSVOL)
        $fsmoWorst = if ($fsmoRows.Count -gt 0) { Worst $fsmoRows } else { 'NotMeasured' }
        $repWorst = if ($repRows.Count -gt 0) { Worst $repRows } else { 'NotMeasured' }
        $sysvolWorst = if ($sysvolRows.Count -gt 0) { Worst $sysvolRows } else { 'NotMeasured' }
        $dcCount = if ($fact) { $fact.DCCount } else { 'NotMeasured' }
        [void]$b.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td class='risk{7}'>{7}</td><td class='risk{8}'>{8}</td><td class='risk{9}'>{9}</td></tr>" -f (ConvertTo-HtmlSafe $g.Name),$dcCount,@($it | Where-Object Status -eq OK).Count,@($it | Where-Object Status -eq Warning).Count,@($it | Where-Object Status -eq Critical).Count,@($it | Where-Object Status -eq NotMeasured).Count,(ConvertTo-HtmlSafe $lockedValue),$fsmoWorst,$repWorst,$sysvolWorst))
    }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Critical and warning findings</h2><table><tr><th>Status</th><th>Domain</th><th>DC</th><th>Category</th><th>Check</th><th>Value</th><th>Threshold</th><th>Details</th></tr>')
    if ($r.Count -eq 0) { [void]$b.AppendLine('<tr class="rowNotMeasured"><td colspan="8">No health measurements are available.</td></tr>') }
    elseif ($find.Count -eq 0 -and $notMeasuredRows.Count -gt 0) { [void]$b.AppendLine('<tr class="rowNotMeasured"><td colspan="8">No critical or warning findings; unmeasured checks do not imply healthy results.</td></tr>') }
    elseif ($find.Count -eq 0) { [void]$b.AppendLine('<tr class="rowOK"><td colspan="8">No critical or warning findings.</td></tr>') }
    foreach ($x in $find) { [void]$b.AppendLine(("<tr class='row{0}'><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>" -f (ConvertTo-HtmlSafe $x.Status),(ConvertTo-HtmlSafe $x.Domain),(ConvertTo-HtmlSafe $x.DC),(ConvertTo-HtmlSafe $x.Category),(ConvertTo-HtmlSafe $x.Check),(ConvertTo-HtmlSafe $x.NumericValue),(ConvertTo-HtmlSafe $x.Threshold),(ConvertTo-HtmlSafe $x.Details))) }
    [void]$b.AppendLine('</table></div><div class="card small"><h2>Technical files</h2><p>CSV: ' + (ConvertTo-HtmlSafe $csv) + '</p></div></body></html>')
    $b.ToString()
}
function Invoke-DcCheck($ForestName,$DomainName,$DC,$AllDcs){
    $dcName=if($DC.HostName){[string]$DC.HostName}else{[string]$DC.Name}
    foreach($port in 389,3268,445){$s=Get-Date;try{$ok=Invoke-Retry {Test-Port $dcName $port};$stat=if($ok -or ($port -eq 3268 -and -not $DC.IsGlobalCatalog)){'OK'}else{'Critical'};Add-Row $ForestName $DomainName $dcName Connectivity "TCP $port" $stat ([int]$ok) "Reachable=$ok" 'Open' "TCP port $port" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Connectivity "TCP $port" Critical 0 Error Open $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$ok=Invoke-Retry {Test-Connection -ComputerName $dcName -Count 1 -Quiet -ErrorAction Stop};Add-Row $ForestName $DomainName $dcName Connectivity Ping $(if($ok){'OK'}else{'Critical'}) ([int]$ok) "Reachable=$ok" Reachable 'ICMP echo test' (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Connectivity Ping Critical 0 Error Reachable $_.Exception.Message (Ms $s)}
    if ($EnableRemoteDcAdminChecks) {
        foreach($svcName in 'NTDS','DNS','Netlogon','KDC','W32Time','DFSR') {
            $s=Get-Date;try{$svc=Invoke-Retry {Get-Service -ComputerName $dcName -Name $svcName -ErrorAction Stop};$ok=$svc.Status -eq 'Running';Add-Row $ForestName $DomainName $dcName Services $svcName $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $svc.Status Running "Service $svcName" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Services $svcName Critical 0 Error Running $_.Exception.Message (Ms $s)}
        }
    }
    else {
        foreach($svcName in 'NTDS','DNS','Netlogon','KDC','W32Time','DFSR') {
            $s=Get-Date;Add-Row $ForestName $DomainName $dcName Services $svcName NotMeasured '' NotMeasured 'Requires T0 remote DC admin' 'Skipped by default to avoid remote service-control access to domain controllers. Use -EnableRemoteDcAdminChecks with a T0 account.' (Ms $s)
        }
    }
$s=Get-Date;$repErrors=$null;$f=@(Get-ADReplicationFailure -Target $dcName -Scope Server -ErrorAction SilentlyContinue -ErrorVariable repErrors);if($repErrors){Add-Row $ForestName $DomainName $dcName Replication ReplicationFailures Critical '' Error 0 ([string]$repErrors[0].Exception.Message) (Ms $s)}else{$det=($f|Select-Object -First 5|ForEach-Object{"$($_.Partner): $($_.FailureCount) since $(ConvertTo-IsoUtc $_.FirstFailureTime)"}) -join '; ';Add-Row $ForestName $DomainName $dcName Replication ReplicationFailures $(if($f.Count -eq 0){'OK'}else{'Critical'}) $f.Count "Failures=$($f.Count)" 0 $det (Ms $s)}
$s=Get-Date;$metaErrors=$null;$m=@(Get-ADReplicationPartnerMetadata -Target $dcName -Scope Server -ErrorAction SilentlyContinue -ErrorVariable metaErrors);if($metaErrors){Add-Row $ForestName $DomainName $dcName Replication MaxLastSuccessDelayHours Critical '' Error "<= $ReplicationDelayWarningHours hours" ([string]$metaErrors[0].Exception.Message) (Ms $s)}else{$old=@($m|Where-Object{$_.LastReplicationSuccess -and $_.LastReplicationSuccess -ne [datetime]::MinValue}|Sort-Object LastReplicationSuccess|Select-Object -First 1)[0];$max=0;if($old){$max=[math]::Round(((Get-Date).ToUniversalTime()-$old.LastReplicationSuccess.ToUniversalTime()).TotalHours,2)};$stat=if($max -gt $ReplicationDelayWarningHours){'Warning'}else{'OK'};Add-Row $ForestName $DomainName $dcName Replication MaxLastSuccessDelayHours $stat $max $(if($old){ConvertTo-IsoUtc $old.LastReplicationSuccess}else{''}) "<= $ReplicationDelayWarningHours hours" "Partners=$($m.Count)" (Ms $s)}
    foreach($sh in 'SYSVOL','NETLOGON'){$s=Get-Date;$unc="\\$dcName\$sh";try{$ok=Invoke-Retry {Test-Path -LiteralPath $unc -PathType Container};Add-Row $ForestName $DomainName $dcName SYSVOL $sh $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $unc Available "Share available=$ok" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName SYSVOL $sh Critical 0 $unc Available $_.Exception.Message (Ms $s)}}
    if(-not $SkipDfsrBacklog -and $AllDcs.Count -gt 1){$src=@($AllDcs|Where-Object{$_ -ine $dcName}|Select-Object -First 1)[0];$s=Get-Date;try{$backlog=Get-DfsrBacklog $src $dcName;if($null -eq $backlog.Count){Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog NotMeasured '' NotMeasured "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" "Not measurable from $src to $dcName; $($backlog.Reason)" (Ms $s)}else{$n=[int]$backlog.Count;$st=if($n -gt $DfsrBacklogCriticalCount){'Critical'}elseif($n -gt $DfsrBacklogWarningCount){'Warning'}else{'OK'};Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog $st $n "$src->$dcName" "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" 'DFSR backlog measured' (Ms $s)}}catch{Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog NotMeasured '' NotMeasured "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" "DfsrdiagUnexpectedError: $($_.Exception.Message)" (Ms $s)}}
    $s=Get-Date;try{$r=@(Resolve-DnsName -Name $dcName -ErrorAction Stop);$targets=@($r|Select-Object -First 3|ForEach-Object{$v=Get-ObjectPropertyValue $_ @('IPAddress','NameHost','NameTarget','Target','Name');if($v){$v}});Add-Row $ForestName $DomainName $dcName DNS ResolveDC $(if($r.Count -gt 0){'OK'}else{'Critical'}) $r.Count $dcName '>= 1 record' ($targets -join '; ') (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName DNS ResolveDC Critical 0 $dcName '>= 1 record' $_.Exception.Message (Ms $s)}
$s=Get-Date;$time=Get-TimeOffsetMinute $dcName;if($null -eq $time.OffsetMinutes){Add-Row $ForestName $DomainName $dcName Time W32TimeOffsetMinutes NotMeasured '' NotMeasured "<= $TimeOffsetWarningMinutes minutes" $time.Error (Ms $s)}else{Add-Row $ForestName $DomainName $dcName Time W32TimeOffsetMinutes $(if($time.OffsetMinutes -gt $TimeOffsetWarningMinutes){'Warning'}else{'OK'}) $time.OffsetMinutes 'Absolute max sample offset' "<= $TimeOffsetWarningMinutes minutes" 'w32tm /stripchart samples=3' (Ms $s)}
    $s=Get-Date
    if ($EnableRemoteDcAdminChecks) {
        try{$vol=Invoke-Retry {Get-ADDbVolume $dcName};$disk=Invoke-Retry {Get-Disk $dcName $vol};$st=if($disk.FreePct -lt $DiskFreePercentCritical -or $disk.FreeGb -lt $DiskFreeGbCritical){'Critical'}else{'OK'};Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent $st $disk.FreePct "$vol free $($disk.FreeGb) GB" ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" 'AD DB volume via ntdsutil, fallback system drive' (Ms $s);Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreeGB $st $disk.FreeGb "$vol free $($disk.FreePct) percent" ">= $DiskFreeGbCritical GB and >= $DiskFreePercentCritical percent" 'Same volume as percent check' (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent Critical '' Error ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" $_.Exception.Message (Ms $s)}
    }
    else {
        Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent NotMeasured '' NotMeasured ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" 'Skipped by default to avoid WinRM/CIM remote admin logon to domain controllers. Use -EnableRemoteDcAdminChecks with a T0 account.' (Ms $s)
        Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreeGB NotMeasured '' NotMeasured ">= $DiskFreeGbCritical GB and >= $DiskFreePercentCritical percent" 'Skipped by default to avoid WinRM/CIM remote admin logon to domain controllers. Use -EnableRemoteDcAdminChecks with a T0 account.' (Ms $s)
    }
}
function Invoke-DomainCheck($ForestName,[string]$DomainName,$ForestInfo){
    $domain=Invoke-Retry {Get-ADDomain -Identity $DomainName -Server $DomainName -ErrorAction Stop}
    $dcs=@(Invoke-Retry {Get-ADDomainController -Filter * -Server $DomainName -ErrorAction Stop}|Sort-Object HostName)
    $dcNames=@($dcs|ForEach-Object{if($_.HostName){[string]$_.HostName}else{[string]$_.Name}})
    [void]$DomainFacts.Add([pscustomobject]@{Domain=$DomainName;DomainMode=[string]$domain.DomainMode;DCCount=@($dcs).Count;GlobalCatalogCount=@($dcs|Where-Object IsGlobalCatalog).Count;PDCEmulator=[string]$domain.PDCEmulator;RIDMaster=[string]$domain.RIDMaster;InfrastructureMaster=[string]$domain.InfrastructureMaster})
    $s=Get-Date;try{$srv=@(Resolve-DnsName -Name "_ldap._tcp.$DomainName" -Type SRV -ErrorAction Stop);$targets=@($srv|Select-Object -First 5|ForEach-Object{$target=Get-ObjectPropertyValue $_ @('NameTarget','NameHost','Target','Name');$port=Get-ObjectPropertyValue $_ @('Port');if($port){"$target`:$port"}else{$target}});Add-Row $ForestName $DomainName '' DNS LDAP_SRV $(if($srv.Count -gt 0){'OK'}else{'Critical'}) $srv.Count "_ldap._tcp.$DomainName" '>= 1 SRV record' ($targets -join '; ') (Ms $s)}catch{Add-Row $ForestName $DomainName '' DNS LDAP_SRV Critical 0 "_ldap._tcp.$DomainName" '>= 1 SRV record' $_.Exception.Message (Ms $s)}
    foreach($dc in $dcs){Invoke-DcCheck $ForestName $DomainName $dc $dcNames}
    foreach($role in 'PDCEmulator','RIDMaster','InfrastructureMaster'){$s=Get-Date;try{$h=[string]$domain.$role;$ok=Test-Port $h 389;Add-Row $ForestName $DomainName $h FSMO $role $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $h 'LDAP 389 reachable' "Domain FSMO holder for $role" (Ms $s)}catch{Add-Row $ForestName $DomainName '' FSMO $role Critical 0 Error 'LDAP 389 reachable' $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$trusts=@(Get-ADTrust -Filter * -Server $DomainName -ErrorAction Stop);$trustTestCommand=Get-Command Test-ADTrustRelationship -ErrorAction SilentlyContinue;foreach($t in $trusts){$ts=Get-Date;if($trustTestCommand){try{$ok=Test-ADTrustRelationship -Identity $t.Name -Server $DomainName -ErrorAction Stop;Add-Row $ForestName $DomainName '' Trusts $t.Name $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $t.TrustType 'Trust validation succeeds' "Direction=$($t.Direction); Transitive=$($t.IsTransitive)" (Ms $ts)}catch{Add-Row $ForestName $DomainName '' Trusts $t.Name Warning 0 $t.TrustType 'Trust validation attempted' $_.Exception.Message (Ms $ts)}}else{Add-Row $ForestName $DomainName '' Trusts $t.Name OK '' $t.TrustType 'Trust validation unavailable' "Direction=$($t.Direction); Transitive=$($t.IsTransitive); Test-ADTrustRelationship not available in this PowerShell session" (Ms $ts)}};if($trusts.Count -eq 0){Add-Row $ForestName $DomainName '' Trusts TrustCount OK 0 'No trusts discovered' 'Inventory only' 'No trust returned by Get-ADTrust' (Ms $s)}}catch{Add-Row $ForestName $DomainName '' Trusts TrustEnumeration Critical '' Error 'Enumeration succeeds' $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{$locked=@(Search-ADAccount -LockedOut -UsersOnly -Server $DomainName -ErrorAction Stop);Add-Row $ForestName $DomainName '' DomainStats LockedUserAccounts $(if($locked.Count -gt 0){'Warning'}else{'OK'}) $locked.Count 'Locked user accounts' '0 preferred' 'Search-ADAccount -LockedOut -UsersOnly' (Ms $s)}catch{Add-Row $ForestName $DomainName '' DomainStats LockedUserAccounts Warning '' Error '0 preferred' $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{
        $domainAdminsSid = $domain.DomainSID.Value + '-512'
        $da=@(Get-ADGroupMembersBySid -Sid $domainAdminsSid -Server $DomainName);Add-Row $ForestName $DomainName '' DomainStats DomainAdminsMemberCount $(if($da.Count -gt 0){'OK'}else{'Warning'}) $da.Count 'Domain Admins recursive members' 'Inventory only' 'Group resolved by RID 512' (Ms $s)
        $dns=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach($m in $da){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}
        $builtinAdmins=@(Get-ADGroupMembersBySid -Sid 'S-1-5-32-544' -Server $DomainName)
        foreach($m in $builtinAdmins){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}
        if($DomainName -ieq $ForestInfo.RootDomain){foreach($rid in '518','519'){foreach($m in @(Get-ADGroupMembersBySid -Sid ($domain.DomainSID.Value+'-'+$rid) -Server $DomainName)){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}}}
        $never=0
        $unresolvedPrivilegedUsers=0
        foreach($dn in $dns){
            $passwordState = Get-PrivilegedUserPasswordNeverExpires -DistinguishedName $dn -DefaultServer $DomainName
            if(-not $passwordState.Resolved){$unresolvedPrivilegedUsers++;continue}
            if($passwordState.PasswordNeverExpires){$never++}
        }
        if($unresolvedPrivilegedUsers -gt 0){
            WriteLog -Message ("Unable to resolve {0} privileged user(s) for domain '{1}' while checking PasswordNeverExpires." -f $unresolvedPrivilegedUsers,$DomainName) -Level 'WARNING'
        }
        $privilegedStatus=if($never -gt 0 -or $unresolvedPrivilegedUsers -gt 0){'Warning'}else{'OK'}
        Add-Row $ForestName $DomainName '' DomainStats PrivilegedPasswordNeverExpires $privilegedStatus $never "PrivilegedUsers=$($dns.Count); Unresolved=$unresolvedPrivilegedUsers" '0 preferred; 0 unresolved' 'Domain Admins, Builtin Administrators when resolvable, Enterprise Admins, Schema Admins when applicable' (Ms $s)
    }catch{Add-Row $ForestName $DomainName '' DomainStats PrivilegedAccountStats Warning '' Error 'Inventory succeeds' $_.Exception.Message (Ms $s)}
}
try{
    Import-Module ActiveDirectory -ErrorAction Stop
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputFolder -LogFileName $ScriptBaseName
    . (Join-Path (Split-Path $CoreModulePath -Parent) 'SmartM365-SharePointUploadPolicy.ps1')
    Initialize-SmartM365CollectorUploadContext -LocalConfig $ScriptConfig -GlobalConfig $TenantContext -DataAllRoot $InitializeOutputPath -DataLastRoot $LatestCsvFolderPath -LogAllRoot $global:LogPath -ExternalActionsAllowed $true
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $LatestCsvFolderPath -ScopeParameters @{SkipDfsrBacklog=$SkipDfsrBacklog;EnableRemoteDcAdminChecks=$EnableRemoteDcAdminChecks}
    $OutputFolder = $InitializeOutputPath
    $transcriptPath = $null
    $transcriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
    if ($transcriptVariable -and $transcriptVariable.Value) { $transcriptPath = [string]$transcriptVariable.Value }
    else {
        $transcriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
        if ($transcriptVariable -and $transcriptVariable.Value) { $transcriptPath = [string]$transcriptVariable.Value }
    }
    if ($transcriptPath) { Start-SmartM365RunTranscript -Path $transcriptPath -Append | Out-Null }
    Write-SmartM365LoadedModuleVersions
    WriteLog -Message ("Script environment initialized at {0}" -f $OutputFolder)
    WriteLog -Message ("Starting {0}" -f $TaskName)
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputFolder) -RequiredModules @('ActiveDirectory') -RequireActiveDirectoryRead | Out-Null
    if (-not $EnableRemoteDcAdminChecks) { WriteLog -Message 'Remote DC admin checks are disabled. Service status and AD database disk checks will be marked NotMeasured; use -EnableRemoteDcAdminChecks only with a T0 account.' -Level 'INFO' }
    if ($NonBlockingDomainErrorSet.Count -gt 0) { WriteLog -Message ("NonBlockingDomainErrors configured for replication connectivity errors: {0}. Unavailable measurements will be explicit." -f (($NonBlockingDomainErrorSet | Sort-Object) -join ', ')) -Level 'INFO' }
    $forest=Invoke-Retry {Get-ADForest -ErrorAction Stop}; $forestName=[string]$forest.Name
    WriteLog -Message ("Forest discovered: {0}; Domains: {1}" -f $forestName, @($forest.Domains).Count) -Level 'INFO'
    WriteLog -Message 'Collecting forest FSMO checks.' -Level 'INFO'
    foreach($role in 'SchemaMaster','DomainNamingMaster'){$s=Get-Date;try{$h=[string]$forest.$role;$ok=Test-Port $h 389;Add-Row $forestName $forest.RootDomain $h FSMO $role $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $h 'LDAP 389 reachable' "Forest FSMO holder for $role" (Ms $s)}catch{Add-Row $forestName $forest.RootDomain '' FSMO $role Critical 0 Error 'LDAP 389 reachable' $_.Exception.Message (Ms $s)}}
    WriteLog -Message 'Collecting tombstone lifetime.' -Level 'INFO'
    $s=Get-Date;try{$cfg=(Get-ADRootDSE -Server $forest.RootDomain -ErrorAction Stop).configurationNamingContext;$ds="CN=Directory Service,CN=Windows NT,CN=Services,$cfg";$obj=Get-ADObject -Identity $ds -Properties tombstoneLifetime -Server $forest.RootDomain -ErrorAction Stop;$tomb=if($obj.tombstoneLifetime){[int]$obj.tombstoneLifetime}else{180};Add-Row $forestName $forest.RootDomain '' Tombstone TombstoneLifetimeDays OK $tomb 'Forest tombstone lifetime' 'Compare with oldest replication failure' $ds (Ms $s)}catch{$tomb=180;Add-Row $forestName $forest.RootDomain '' Tombstone TombstoneLifetimeDays Warning $tomb Defaulted 'Compare with oldest replication failure' $_.Exception.Message (Ms $s)}
    foreach($d in @($forest.Domains|Sort-Object)){WriteLog -Message ("Collecting domain checks: {0}" -f $d) -Level 'INFO';try{Invoke-DomainCheck $forestName ([string]$d) $forest}catch{Add-Row $forestName ([string]$d) '' Domain DomainScan Critical 0 Failed 'Domain scan succeeds' $_.Exception.Message 0}}
    Add-ForestReplicationAgeRows -ForestName $forestName -RootDomain $forest.RootDomain -Domains @($forest.Domains) -TombstoneDays $tomb
    $stamp=(Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture)
    $all = Get-RowSnapshot
    if(-not (Test-Path -LiteralPath $LatestCsvFolderPath)){New-Item -Path $LatestCsvFolderPath -ItemType Directory -Force|Out-Null}
    $latestCsv=Join-Path $LatestCsvFolderPath 'AD_HealthCheck.csv'
    if($AppendHistory){$csv=$HistoryCsvPath}else{$csv=Join-Path $OutputFolder "AD_HealthCheck_$stamp.csv"}
    Assert-SmartM365CsvDataCompleteness -Data $all -TimestampedPath $csv -LatestPath $latestCsv
    if($AppendHistory){if(Test-Path $csv){Repair-SmartM365CsvTenantKeySchema -Path $csv -Delimiter ',' -Encoding UTF8|Out-Null}; Add-SmartM365CsvRowsAtomically -Data $all -Path $csv -Encoding utf8BOM}else{Write-SmartM365CsvAtomically -Data $all -Path $csv -Encoding utf8BOM}
    Write-SmartM365CsvAtomically -Data $all -Path $latestCsv -Encoding utf8BOM
    Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputFolder -FilePattern '*.csv' -RetentionDays 7 -LogFile $global:LogTextFile
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csv
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestCsv
    $end=Get-Date;$worst=Worst $all;$subject="SMART365 - [$($worst.ToUpperInvariant())] Active Directory Health Check - $forestName - $RunDateUtc";$html=ConvertTo-ReportHtml -r $all -status $worst -started $RunStarted -ended $end -csv $csv -forestInfo $forest -RemoteDcAdminChecksEnabled:$EnableRemoteDcAdminChecks;if($AlwaysSend -or $worst -ne 'OK'){Send-ReportMail $subject $html}
    $summaryStatus=if($worst -eq 'OK'){'Success'}else{'CompletedWithWarnings'}
    $businessCriticalCount = @($all | Where-Object Status -eq Critical).Count
    $businessWarningCount = @($all | Where-Object Status -eq Warning).Count
    $notMeasuredCount = @($all | Where-Object Status -eq NotMeasured).Count
    WriteLog -Message ("Health findings: Critical={0}; Warning={1}; NotMeasured={2}; technical log warnings and errors are reported separately in the execution summary." -f $businessCriticalCount,$businessWarningCount,$notMeasuredCount) -Level INFO
    try { Stop-SmartM365RunTranscript | Out-Null } catch { $null = $_ }
    $global:csvGeneratedPaths = @($csv, $latestCsv)
    Write-Host "AD Health Check completed. Status=$worst; Rows=$($all.Count); Csv=$csv; Latest=$latestCsv"
    Complete-SmartM365ExecutionContext -Status $summaryStatus
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(0); exit $smartM365RunGuard.ExitCode
}catch{
    $runError = $_
    $runErrorDetail = @(
        "Exception: $($runError.Exception.GetType().FullName): $($runError.Exception.Message)"
        if ($runError.InvocationInfo -and $runError.InvocationInfo.PositionMessage) { "Position: $($runError.InvocationInfo.PositionMessage)" }
        if ($runError.ScriptStackTrace) { "Stack: $($runError.ScriptStackTrace)" }
    ) -join ' | '
    try { WriteLog -Message $runErrorDetail -Level 'ERROR' } catch { Write-Warning $runErrorDetail }
    Add-Row '' '' '' Script UnhandledError Critical 0 Failed 'Script completes' $runErrorDetail 0
    if(-not (Test-Path -LiteralPath $OutputFolder)){New-Item -Path $OutputFolder -ItemType Directory -Force|Out-Null}
    $csv=Join-Path $OutputFolder ("AD_HealthCheck_FAILED_{0}.csv" -f (Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture));$failedRows=Get-RowSnapshot;Write-SmartM365CsvAtomically -Data $failedRows -Path $csv -Encoding utf8BOM
    try{$html=ConvertTo-ReportHtml -r $failedRows -status 'Critical' -started $RunStarted -ended (Get-Date) -csv $csv;Send-ReportMail "SMART365 - [CRITICAL] Active Directory Health Check failed - $RunDateUtc" $html -MailPurpose Error}catch{Write-Warning $_.Exception.Message}
    try { Stop-SmartM365RunTranscript | Out-Null } catch { $null = $_ }
    $global:csvGeneratedPaths = @($csv)
    Write-Error $runErrorDetail
    try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'ADHealthCheck' } catch { $null = $_ };$smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(2); exit $smartM365RunGuard.ExitCode
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
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDAg4ipv+J4lzML
# 6sOmORYTOyPRsY8g4eO0DX+mnEze4aCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg4mrA6OCkujSH
# VP3DIPgN8TfzsDtEKKIsJ9izFvA3z9EwDQYJKoZIhvcNAQEBBQAEggGAqAyJhC9T
# uSf0QQ+6fyUqBCG0JCRYY/MWvAG0yd0dXVexeiSXeA+VTZMaYbEhMhqwAQ++5G+o
# 3QSGGrk3XaVHuSKeIeM23ozwk21nkMgrGQsMUXbjXcrCLTGq/NI98jSYff0JjkcK
# IcKEelfTEkuXPuiSbu76M31cSUSBJXi+XKhjW/f3m6wEZJMZPNBu1Zl/Cg93yT9o
# ZuQJId+bTTCiCouOraW/HOayHNknx8Hdy1Pn3DUOeMcC/Z+4awDJpeai9Yt3GBm1
# Ks3FF/bJ1l5qojSiNOoZvFRo56dy19wP446zLrtPelJmNJqV9IVH+7nbfQbRDypj
# 1VYkHiOjFBhHJLIZlBGRlU1NoWbb+vsKMMT5uSRghElSoPPR8AHoCwoqa/YsZTT4
# I1Y8ZRyX3rt0MXRvMBkBaw88NIphOWkYjkHaLfAwsoRtEvVmqf2JrFbsQXx+gPls
# p3PtSDBqBYbQPUY8+C+Kg3Fch/6ViYwyQg+T0vgBAai7z5cuEnXdnFsooYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTAxMTY1OTA0WjA/BgkqhkiG9w0BCQQxMgQwXIHy+nTLUtHJVeaZxaUVKgt6
# HfNjo7CXrpnQLIvGDTITZ3x8ILiVMOLpPW6lEeJ5MA0GCSqGSIb3DQEBAQUABIIC
# AImgOAsNhA2dBi3GPEDsH9/9GcZZAkzHkTNH0xgca/EMhvU4cZxaa8szSY/W8IGz
# 8JfgWuhnQBhfhALpVbT/VwinVw5DdyqxoeHZtnt/T8Gm2sp/dkgE+GoNp5+tvHvY
# TfYJkp8PQATPlvB9YM8iIYF/W/+gaBJblSk+GN2P6aMTG5CRiCYWA45k/ozpQ04p
# VH0ZpFPizIXuQ2HNh7sl1Poa+r4UKY3/xlrARHHN6QYjBfZ0nhxx9l6pKWq+A8oy
# dprenei+gW2X1Y8uUuSwhjJqUUz/z+9H+PQH7UrQgQBReUyv/e2CCPXVpNdwNbyO
# ucfvwVlTPlrIBNT9GYTNOrd7MhiKYUehGi7NTZp+JAZmeLpp6Fr4hSR6xKSgpfca
# 1m4OtInx9eLV5IBQUb+WP4bEByZNeGGjFYoDte3Kgmcfj+6hq5VsMLgZ1YhbAh1H
# aihR5IOMECOu/5ki5Vi4N6s2ip45oILAQyDhibRnvfH3yO2NnNXErH+JMeP6AZjW
# KLj+HGK7Oxy4H8gDrINKQn1tKCvpsPNehcoBM+Mnz1ASCnjwFJsCPJ5qs3+QmOkr
# snOnBkbAIZL0qAKHmOcE7EY/+dw3uojVdSKu9vwIjVc9+kbvtuO6S5/AhYlodPJD
# fgJD6vxAiHkd5CVn2HE2kh60R96mZVEaw6rwvzhrqEaB
# SIG # End signature block
