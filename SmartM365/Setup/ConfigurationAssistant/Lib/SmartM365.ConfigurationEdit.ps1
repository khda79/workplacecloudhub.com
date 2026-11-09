# SmartInventory configuration editing helpers, version 0.17.0.
# Uses the existing transport and template enrichment only when explicitly saving.
function Test-SiConfigSensitiveObject {
    param([AllowNull()]$Value,[int]$Depth=0)
    if ($Depth -gt 40) {return $true}
    if ($Value -is [pscustomobject]) {foreach ($property in $Value.PSObject.Properties) {if ($property.Name -match '(?i)password|secret|token|webhook|credential|private.?key' -or (Test-SiConfigSensitiveObject -Value $property.Value -Depth ($Depth+1))) {return $true}}}
    elseif ($Value -is [array]) {foreach ($item in $Value) {if (Test-SiConfigSensitiveObject -Value $item -Depth ($Depth+1)) {return $true}}}
    return $false
}
function Assert-SiEndpointRequirements {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Requirements)
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);$active=0
    foreach ($requirement in $Requirements) {
        if ($requirement -isnot [pscustomobject]) {throw 'Each requirement must be a JSON object.'}
        $role=[string]$requirement.Role;$product=([string]$requirement.Product).Trim();$mode=[string]$requirement.ExpectedMode;$reason=([string]$requirement.Reason).Trim()
        if ($role -cnotin @('Antivirus','EDR') -or -not $product -or $product -match '[|\x00-\x1f\x7f]' -or $mode -cnotin @('Active','Passive','NotRequired') -or ($mode -eq 'Passive' -and ($role -ne 'Antivirus' -or $product -ine 'Microsoft Defender Antivirus')) -or ($mode -eq 'NotRequired' -and -not $reason) -or -not $seen.Add($role+'|'+$product)) {throw 'Endpoint requirements contain an invalid or duplicate role, product or mode.'}
        if ($role -eq 'Antivirus' -and $mode -eq 'Active') {$active++}
    }
    if ($active -gt 1) {throw 'A scope cannot require more than one active antivirus.'}
}
function Assert-SiEndpointProtectionPolicy {
    param([Parameter(Mandatory)]$Policy)
    if ($Policy -isnot [pscustomobject] -or $Policy.Enabled -isnot [bool] -or $Policy.Requirements -isnot [array] -or $Policy.ScopeRules -isnot [array] -or (Test-SiConfigSensitiveObject $Policy)) {throw 'Endpoint protection policy has an invalid or sensitive structure.'}
    $version=([string]$Policy.PolicyVersion).Trim()
    if (($Policy.Enabled -and -not $version) -or $version -match '[|\x00-\x1f\x7f]') {throw 'An enabled policy requires a valid policy version.'}
    if ($null -ne $Policy.MaxObservationAgeHours) {
        $age=$Policy.MaxObservationAgeHours
        if (($age -isnot [int] -and $age -isnot [long] -and $age -isnot [double] -and $age -isnot [decimal]) -or [double]$age -le 0 -or [double]::IsNaN([double]$age) -or [double]::IsInfinity([double]$age)) {throw 'Maximum observation age must be null or a finite positive number.'}
    }
    Assert-SiEndpointRequirements -Requirements $Policy.Requirements
    $scopeNames=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($scope in $Policy.ScopeRules) {
        if ($scope -isnot [pscustomobject] -or $scope.DeviceSourceIds -isnot [array] -or -not $scope.DeviceSourceIds.Count -or $scope.Requirements -isnot [array]) {throw 'Each scope requires device IDs and requirements arrays.'}
        $id=([string]$scope.ScopeId).Trim()
        if (-not $id -or $id -match '[|\x00-\x1f\x7f]' -or $id -in @('Default','Ambiguous','Disabled') -or -not $scopeNames.Add($id)) {throw 'Scope IDs must be unique, nonempty and not reserved.'}
        $ids=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($device in $scope.DeviceSourceIds) {if ($device -isnot [string] -or [string]::IsNullOrWhiteSpace($device) -or $device -match '[|\x00-\x1f\x7f]' -or -not $ids.Add($device.Trim())) {throw 'Scope device IDs must be nonempty and unique.'}}
        Assert-SiEndpointRequirements -Requirements $scope.Requirements
    }
}
function Assert-SiReportMailSchedule {
    param([Parameter(Mandatory)]$Schedule)
    if ($Schedule -isnot [pscustomobject] -or $Schedule.Enabled -isnot [bool] -or $Schedule.NotBefore -isnot [string] -or $Schedule.NotBefore -notmatch '^([01]\d|2[0-3]):[0-5]\d$' -or $Schedule.TimeZone -isnot [string] -or $Schedule.QueuePath -isnot [string] -or $Schedule.TrustedSids -isnot [array] -or (Test-SiConfigSensitiveObject $Schedule)) {throw 'Report mail schedule has an invalid or sensitive structure.'}
    try {$null=[TimeZoneInfo]::FindSystemTimeZoneById($Schedule.TimeZone)} catch {throw 'Report mail schedule requires a recognized time zone.'}
    if ($Schedule.QueuePath.Length -gt 4096 -or $Schedule.QueuePath -match '[\x00-\x1f\x7f\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'Report mail queue path is invalid.'}
    $sids=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($sid in $Schedule.TrustedSids) {
        if ($sid -isnot [string] -or $sid -notmatch '^S-1-(?:\d+-)+\d+$' -or -not $sids.Add($sid)) {throw 'Trusted SIDs must be valid and unique.'}
        try {$null=[Security.Principal.SecurityIdentifier]::new($sid)} catch {throw 'Trusted SIDs must be valid and unique.'}
    }
    if ($Schedule.Enabled -and -not $sids.Count) {throw 'An enabled report mail schedule requires trusted SIDs.'}
}
function Get-SiTenantTemplateDelta {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)]$Document,[string[]]$RequiredFields=@())
    if ($Document -isnot [pscustomobject]) {throw 'Expected a tenant configuration object.'}
    $templatePath=Join-Path $Root 'Config/Tenants/tenant.local.json.template'
    $bytes=[IO.File]::ReadAllBytes($templatePath);$sha=[Security.Cryptography.SHA256]::Create()
    try {$hash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','')} finally {$sha.Dispose()}
    $template=[Text.UTF8Encoding]::new($false,$true).GetString($bytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
    if ($template -isnot [pscustomobject]) {throw 'Invalid tenant template.'}
    foreach ($field in $RequiredFields) {if ($null -eq $template.PSObject.Properties[$field]) {throw 'A required template field is missing.'}}
    $added=@(Add-SmartM365MissingJsonTemplateProperties -Target $Document -Template $template)
    $additions=@(foreach ($key in $added) {
        $node=$Document;foreach ($part in ($key -split '\.')) {$property=$node.PSObject.Properties[$part];if ($null -eq $property) {throw 'Template addition cannot be reviewed.'};$node=$property.Value}
        if ($key -match '(?i)password|secret|token|webhook|credential|private.?key') {
            if ($null -ne $node -and ($node -isnot [string] -or $node -notin @('','__USE_GLOBAL__','USE_GLOBAL'))) {throw 'A sensitive template addition requires separate review.'}
            $value='(empty or inherit)'
        } else {
            if (Test-SiConfigSensitiveObject $node) {throw 'A sensitive template addition requires separate review.'}
            $value=ConvertTo-Json -InputObject $node -Depth 100 -Compress
        }
        if ($value -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'Template addition contains unsupported direction controls.'}
        [pscustomobject]@{Field=$key;ValueJson=$value}
    })
    [pscustomobject]@{Hash=$hash;AddedKeys=$added;Additions=$additions}
}
function Format-SiTenantTemplateAdditions {
    param([Parameter(Mandatory)]$Snapshot)
    $details=if (@($Snapshot.Additions).Count) {(@($Snapshot.Additions | ForEach-Object {$_.Field+' = '+$_.ValueJson}) -join "`n")} else {'None'}
    if ($details.Length -gt 12000) {throw 'Template additions are too large for this review. Review the template separately.'}
    return $details
}
function Confirm-SiTenantTemplateDelta {
    param([Parameter(Mandatory)]$Snapshot,[Parameter(Mandatory)]$Document,[string[]]$RequiredFields=@())
    $delta=Get-SiTenantTemplateDelta -Root $Snapshot.Root -Document $Document -RequiredFields $RequiredFields
    if ($delta.Hash -cne $Snapshot.TemplateHash -or (($delta.AddedKeys -join "`n") -cne (@($Snapshot.AddedKeys) -join "`n"))) {throw 'Template additions changed; review the operation again.'}
    return $delta
}
function Get-SiEditableFields {
    param([string]$Step)
    switch ($Step) {
        'share' { @('DefaultTenant','SmartM365RootPath','WorkspaceRootPath','SmartInventoryConfigRootPath','DataAllRootPath','OrchestratorSharedDataFolderPath','LatestCsvFolderPath','LogAllRootPath','OnPremExecutionMode','RetentionMaxCSV','RetentionMaxLogs','EnableSharePointUploadDataAll','EnableSharePointUploadDataLast','EnableSharePointUploadDataPowerBI','EnableSharePointUploadLogAll','MailBrandingEnabled','MailPowerBiEnabled','MailPowerBiUrl','MailPowerBiLogoPath','MailPowerBiLinkText','MailClientName','MailClientLogoPath','MailClientLogoMaxKB','ReportMailSchedule') }
        'tenant' { @('Environment','OrgDomain','RemoteRoutingDomain','ServiceAccount','DeviceSystemMaxObservationAgeHours','DeviceSystemOffsetlessGraphTimeZoneId','AdComputerNameCountryPrefixes','AdDomainCountryMappings','ExpectedThirdPartyAV','EndpointProtectionPolicy') }
        'identity' { @('TenantId','AppId') }
        'administrators' { @('Smart365Administrators','ErrorMailTo') }
        'sharepoint' { @('EnableSharePointUpload','EnableSharePointUploadDataAll','EnableSharePointUploadDataLast','EnableSharePointUploadDataPowerBI','EnableSharePointUploadLogAll','SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath') }
        'mail' { @('From','To','Cc','SmtpServer','SendMailMode','MailSendAccessPolicyGroup','Subject','MailClientName','MailClientLogoPath','MailClientLogoMaxKB','MailBrandingEnabled','MailPowerBiEnabled','MailPowerBiUrl','MailPowerBiLogoPath','MailPowerBiLinkText') }
        default { @() }
    }
}
function Get-SiEditLocation {
    param([string]$Root,[string]$Tenant,[string]$Step)
    if ($Tenant -notmatch '^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$') { throw 'Invalid profile key.' }
    if ($Step -eq 'share') { return Join-Path $Root 'Config/SmartM365.global.local.json' }
    return Join-Path $Root ("Config/Tenants/$Tenant.local.json")
}
function Get-SiEditSnapshot {
    param([string]$Root,[string]$Tenant,[string]$Step,[string]$Field)
    if ($Field -notin @(Get-SiEditableFields $Step)) { throw 'Field is not editable in this step.' }
    $legacy = Get-SiEditLocation $Root $Tenant $Step
    $probe = Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))
    if ($probe.State -ne 'Present') { throw 'Configuration missing, invalid or conflicting; no edit is possible.' }
    $read = Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single') { throw 'Coexisting configuration files require review.' }
    $property = $read.Document.PSObject.Properties[$Field];$present=($null -ne $property);$templateHash=''
    if (-not $present) {
        $template=if ($Step -eq 'share') {Join-Path $Root 'Config/SmartM365.global.local.json.template'} else {Join-Path $Root 'Config/Tenants/tenant.local.json.template'}
        $templateBytes=[IO.File]::ReadAllBytes($template);$sha=[Security.Cryptography.SHA256]::Create()
        try {$templateHash=([BitConverter]::ToString($sha.ComputeHash($templateBytes))).Replace('-','')} finally {$sha.Dispose()}
        $model=[Text.UTF8Encoding]::new($false,$true).GetString($templateBytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
        if ($model -isnot [pscustomobject]) {throw 'Invalid configuration template.'}
        $property=$model.PSObject.Properties[$Field]
        if ($null -eq $property) {throw 'Field is absent from both configuration and template.'}
    }
    $numeric=$Field -in @('RetentionMaxCSV','RetentionMaxLogs') -or ($Step -eq 'share' -and $Field -eq 'MailClientLogoMaxKB')
    $map=$Field -in @('AdComputerNameCountryPrefixes','AdDomainCountryMappings','ExpectedThirdPartyAV','EndpointProtectionPolicy','ReportMailSchedule')
    $globalBoolean=$Step -eq 'share' -and $Field -in @('MailBrandingEnabled','MailPowerBiEnabled')
    $plainBoolean=$Field -in @('EnableSharePointUpload')
    $inheritedBoolean=($Step -ne 'share' -and $Field -in @('MailBrandingEnabled','MailPowerBiEnabled')) -or $Field -like 'EnableSharePointUpload*'
    $inheritedInteger=$Step -ne 'share' -and $Field -eq 'MailClientLogoMaxKB'
    $nullableNumber=$Field -eq 'DeviceSystemMaxObservationAgeHours'
    $numberValue=$property.Value -is [int] -or $property.Value -is [long] -or $property.Value -is [double] -or $property.Value -is [decimal]
    if (($numeric -and $property.Value -isnot [int] -and $property.Value -isnot [long]) -or ($map -and $property.Value -isnot [pscustomobject]) -or (($globalBoolean -or $plainBoolean) -and $property.Value -isnot [bool]) -or ($inheritedBoolean -and $property.Value -isnot [bool] -and ($property.Value -isnot [string] -or $property.Value -notin @('','__USE_GLOBAL__','USE_GLOBAL','true','false'))) -or ($inheritedInteger -and $property.Value -isnot [int] -and $property.Value -isnot [long] -and ($property.Value -isnot [string] -or ($property.Value -notin @('','__USE_GLOBAL__','USE_GLOBAL') -and $property.Value -notmatch '^\d+$'))) -or ($nullableNumber -and $null -ne $property.Value -and -not $numberValue -and ($property.Value -isnot [string] -or $property.Value -notin @('','__USE_GLOBAL__','USE_GLOBAL'))) -or (-not $numeric -and -not $map -and -not $plainBoolean -and -not $globalBoolean -and -not $inheritedBoolean -and -not $inheritedInteger -and -not $nullableNumber -and $property.Value -isnot [string])) { throw 'Field has an unexpected type; no edit is possible.' }
    if ($map -and (Test-SiConfigSensitiveObject $property.Value)) {throw 'This setting contains sensitive fields and cannot be displayed here.'}
    $value=if ($map) {ConvertTo-Json -InputObject $property.Value -Depth 20 -Compress} elseif ($nullableNumber -and $null -eq $property.Value) {'null'} elseif ($nullableNumber -and $numberValue) {[Convert]::ToString($property.Value,[Globalization.CultureInfo]::InvariantCulture)} else {[string]$property.Value}
    [pscustomobject]@{Root=$Root; Tenant=$Tenant; Step=$Step; Field=$Field; Path=$read.Path; Hash=$read.SHA256; Value=$value;Present=$present;TemplateHash=$templateHash}
}
function Get-SiEditTemplatePreview {
    param([Parameter(Mandatory)]$Snapshot)
    if ($Snapshot.Field -notin @(Get-SiEditableFields $Snapshot.Step)) { throw 'Field is not editable.' }
    $legacy=Get-SiEditLocation $Snapshot.Root $Snapshot.Tenant $Snapshot.Step
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Configuration changed or conflicts; refresh before saving.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -ne $Snapshot.Hash -or $read.Path -ne $Snapshot.Path) { throw 'Configuration changed; refresh before saving.' }
    $path=if ($Snapshot.Step -eq 'share') { Join-Path $Snapshot.Root 'Config/SmartM365.global.local.json.template' } else { Join-Path $Snapshot.Root 'Config/Tenants/tenant.local.json.template' }
    $bytes=[IO.File]::ReadAllBytes($path); $sha=[Security.Cryptography.SHA256]::Create()
    try { $hash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','') } finally { $sha.Dispose() }
    if (-not $Snapshot.Present -and $hash -cne $Snapshot.TemplateHash) {throw 'Template changed; reload the missing setting.'}
    $template=[Text.UTF8Encoding]::new($false,$true).GetString($bytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
    if ($template -isnot [pscustomobject]) { throw 'Invalid configuration template.' }
    $keys=@(Add-SmartM365MissingJsonTemplateProperties -Target $read.Document -Template $template)
    $items=@(foreach ($key in $keys) {
        $node=$read.Document
        foreach ($part in ($key -split '\.')) { $property=$node.PSObject.Properties[$part]; if ($null -eq $property) { throw 'Template addition cannot be displayed.' }; $node=$property.Value }
        $json=ConvertTo-Json -InputObject $node -Depth 100 -Compress
        if ($json -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') { throw 'Template addition contains unsupported direction controls.' }
        [pscustomobject]@{Field=$key;ValueJson=$json}
    })
    [pscustomobject]@{Hash=$hash;Keys=$keys;Additions=$items}
}
function ConvertTo-SiEditValue {
    param([Parameter(Mandatory)]$Snapshot,[AllowEmptyString()][string]$Value)
    if ($Snapshot.Field -eq 'ServiceAccount' -and ($Value -notmatch '^(?:[A-Za-z0-9._-]+\\[A-Za-z0-9._+-]+|[A-Za-z0-9._+-]+@[A-Za-z0-9.-]+)$' -or $Value -in @('SYSTEM','LocalSystem','NT AUTHORITY\SYSTEM'))) {throw 'Enter a standard domain account in DOMAIN\user or user@domain format.'}
    if ($Snapshot.Field -in @('TenantId','AppId')) {
        $id = [guid]::Empty
        if (-not [guid]::TryParse($Value,[ref]$id) -or $id -eq [guid]::Empty) { throw 'An actual GUID is required.' }
    }
    if ($Snapshot.Field -eq 'SendMailMode' -and $Value -notin @('Graph','SMTP','Both','__USE_GLOBAL__')) { throw 'SendMailMode must be Graph, SMTP, Both or __USE_GLOBAL__.' }
    if ($Snapshot.Field -eq 'OnPremExecutionMode' -and $Value -notin @('Auto','Direct','LocalCache')) { throw 'OnPremExecutionMode must be Auto, Direct or LocalCache.' }
    if ($Snapshot.Field -eq 'DefaultTenant' -and $Value -cnotmatch '^[a-z0-9][a-z0-9_-]{0,63}$') {throw 'DefaultTenant must be a lowercase profile key.'}
    $typedValue=$Value
    if ($Snapshot.Field -in @('EnableSharePointUpload')) {if ($Value -notin @('true','false')) {throw 'Enter true or false.'};$typedValue=($Value -ieq 'true')}
    if ($Snapshot.Field -in @('EndpointProtectionPolicy','ReportMailSchedule')) {
        if ($Value.Length -gt 65536 -or $Value -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'The policy is too large or contains unsupported direction controls.'}
        try {$typedValue=ConvertFrom-Json -InputObject $Value -NoEnumerate -ErrorAction Stop} catch {throw 'Enter a valid JSON object.'}
        if ($Snapshot.Field -eq 'EndpointProtectionPolicy') {Assert-SiEndpointProtectionPolicy -Policy $typedValue} else {Assert-SiReportMailSchedule -Schedule $typedValue}
    }
    if ($Snapshot.Field -eq 'AdComputerNameCountryPrefixes') {
        if ($Value.Length -gt 65536 -or $Value -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'The mapping is too large or contains unsupported direction controls.'}
        try {$typedValue=ConvertFrom-Json -InputObject $Value -NoEnumerate -ErrorAction Stop} catch {throw 'Enter a JSON object mapping two-letter prefixes to country labels.'}
        if ($typedValue -isnot [pscustomobject]) {throw 'Enter a JSON object mapping two-letter prefixes to country labels.'}
        $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($property in $typedValue.PSObject.Properties) {
            if ($property.Name -notmatch '^[A-Za-z]{2}$' -or -not $seen.Add($property.Name) -or $property.Value -isnot [string] -or [string]::IsNullOrWhiteSpace($property.Value) -or $property.Value.Length -gt 256 -or $property.Value -match '[\x00-\x1f\x7f\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'Every prefix must be unique, contain two letters, and have a nonempty country label.'}
        }
    }
    if ($Snapshot.Field -eq 'AdDomainCountryMappings') {
        if ($Value.Length -gt 65536 -or $Value -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'The mapping is too large or contains unsupported direction controls.'}
        try {$typedValue=ConvertFrom-Json -InputObject $Value -NoEnumerate -ErrorAction Stop} catch {throw 'Enter a JSON object mapping exact DNS domains to country labels.'}
        if ($typedValue -isnot [pscustomobject]) {throw 'Enter a JSON object mapping exact DNS domains to country labels.'}
        $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($property in $typedValue.PSObject.Properties) {
            $domain=$property.Name.Trim().TrimEnd('.').ToLowerInvariant()
            if ($domain.Length -gt 253 -or $domain -notmatch '^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$' -or -not $seen.Add($domain) -or $property.Value -isnot [string] -or [string]::IsNullOrWhiteSpace($property.Value) -or $property.Value.Length -gt 256 -or $property.Value -match '[\x00-\x1f\x7f\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'Use unique exact DNS domains and nonempty country labels.'}
        }
    }
    if ($Snapshot.Field -eq 'ExpectedThirdPartyAV') {
        if ($Value.Length -gt 65536 -or $Value -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') {throw 'The antivirus setting is too large or contains unsupported direction controls.'}
        try {$typedValue=ConvertFrom-Json -InputObject $Value -NoEnumerate -ErrorAction Stop} catch {throw 'Enter a JSON object with DisplayName and RegisteredNames.'}
        if ($typedValue -isnot [pscustomobject] -or $typedValue.DisplayName -isnot [string] -or $typedValue.RegisteredNames -isnot [array]) {throw 'Enter a JSON object with a text DisplayName and a RegisteredNames array.'}
        $display=$typedValue.DisplayName.Trim();$names=$typedValue.RegisteredNames
        if ($typedValue.DisplayName.Length -gt 256 -or $typedValue.DisplayName -match '[\x00-\x1f\x7f]' -or @($names | Where-Object {$_ -isnot [string] -or [string]::IsNullOrWhiteSpace($_) -or $_.Length -gt 256 -or $_ -match '[\x00-\x1f\x7f]'}).Count -gt 0 -or (($display.Length -eq 0) -ne ($names.Count -eq 0))) {throw 'Configure both DisplayName and registered names, or leave both empty.'}
    }
    if ($Snapshot.Field -in @('MailBrandingEnabled','MailPowerBiEnabled') -or $Snapshot.Field -like 'EnableSharePointUpload*') {
        if ($Snapshot.Step -eq 'share' -and $Snapshot.Field -in @('MailBrandingEnabled','MailPowerBiEnabled')) {if ($Value -notin @('true','false')) {throw 'Enter true or false.'};$typedValue=($Value -ieq 'true')}
        elseif ($Value -in @('','__USE_GLOBAL__','USE_GLOBAL')) {$typedValue='__USE_GLOBAL__'} elseif ($Value -in @('true','false')) {$typedValue=($Value -ieq 'true')} else {throw 'Enter true, false or __USE_GLOBAL__.'}
    }
    if ($Snapshot.Field -eq 'MailClientLogoMaxKB') {
        if ($Snapshot.Step -ne 'share' -and $Value -in @('','__USE_GLOBAL__','USE_GLOBAL')) {$typedValue='__USE_GLOBAL__'} else {
            $number=0;if ($Value -notmatch '^\d{1,7}$' -or -not [int]::TryParse($Value,[ref]$number) -or $number -lt 1 -or $number -gt 1000000) {throw 'Enter __USE_GLOBAL__ or a whole number between 1 and 1000000.'};$typedValue=$number
        }
    }
    if ($Snapshot.Field -eq 'DeviceSystemMaxObservationAgeHours') {
        if ($Value -in @('','null','__USE_GLOBAL__','USE_GLOBAL')) {$typedValue=$null}
        else {$number=[double]0;if (-not [double]::TryParse($Value,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$number) -or [double]::IsNaN($number) -or [double]::IsInfinity($number) -or $number -le 0) {throw 'Enter null or a finite positive number of hours.'};$typedValue=$number}
    }
    if ($Snapshot.Field -eq 'DeviceSystemOffsetlessGraphTimeZoneId' -and $Value -notin @('','__USE_GLOBAL__','USE_GLOBAL')) {try {$null=[TimeZoneInfo]::FindSystemTimeZoneById($Value)} catch {throw 'Enter a recognized time zone ID or leave it empty.'}}
    if ($Snapshot.Field -in @('RetentionMaxCSV','RetentionMaxLogs')) {
        $number=0
        if ($Value -notmatch '^\d{1,7}$' -or -not [int]::TryParse($Value,[ref]$number) -or $number -gt 1000000) { throw 'Retention must be a whole number between 0 and 1000000.' }
        $typedValue=$number
    }
    if ($typedValue -is [string] -and ($typedValue.Length -gt 4096 -or $typedValue -match '[\x00-\x1f\x7f\u200e\u200f\u202a-\u202e\u2066-\u2069]')) {throw 'Text contains unsupported characters or is too long.'}
    return $typedValue
}
function Set-SiEditSnapshot {
    param([Parameter(Mandatory)]$Snapshot,[AllowEmptyString()][string]$Value,[string]$ExpectedTemplateHash,[string[]]$ExpectedAddedKeys)
    if ($Snapshot.Field -notin @(Get-SiEditableFields $Snapshot.Step)) { throw 'Field is not editable.' }
    $legacy = Get-SiEditLocation $Snapshot.Root $Snapshot.Tenant $Snapshot.Step
    $probe = Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))
    if ($probe.State -ne 'Present') { throw 'Configuration changed or conflicts; refresh before saving.' }
    $read = Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -ne $Snapshot.Hash -or $read.Path -ne $Snapshot.Path) { throw 'Configuration changed; refresh before saving.' }
    $typedValue=ConvertTo-SiEditValue -Snapshot $Snapshot -Value $Value
    $template = if ($Snapshot.Step -eq 'share') { Join-Path $Snapshot.Root 'Config/SmartM365.global.local.json.template' } else { Join-Path $Snapshot.Root 'Config/Tenants/tenant.local.json.template' }
    $templateBytes=[IO.File]::ReadAllBytes($template)
    if ($ExpectedTemplateHash) {
        $sha=[Security.Cryptography.SHA256]::Create()
        try { $actualHash=([BitConverter]::ToString($sha.ComputeHash($templateBytes))).Replace('-','') } finally { $sha.Dispose() }
        if ($actualHash -cne $ExpectedTemplateHash) { throw 'Template changed; review the change again.' }
    }
    if (-not $Snapshot.Present) {
        $sha=[Security.Cryptography.SHA256]::Create();try {$actualHash=([BitConverter]::ToString($sha.ComputeHash($templateBytes))).Replace('-','')} finally {$sha.Dispose()}
        if ($actualHash -cne $Snapshot.TemplateHash) {throw 'Template changed; reload the missing setting.'}
    }
    $model = [Text.UTF8Encoding]::new($false,$true).GetString($templateBytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
    if ($model -isnot [pscustomobject]) { throw 'Invalid configuration template.' }
    $added = @(Add-SmartM365MissingJsonTemplateProperties -Target $read.Document -Template $model)
    if ($ExpectedTemplateHash -and (($added -join "`n") -cne ($ExpectedAddedKeys -join "`n"))) { throw 'Template additions changed; review the change again.' }
    if ($null -eq $read.Document.PSObject.Properties[$Snapshot.Field]) {throw 'The requested setting is absent from the template.'}
    $read.Document.($Snapshot.Field) = $typedValue
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $Snapshot.Hash -Validate { param($d) if ($d -isnot [pscustomobject]) { throw 'Expected JSON object.' } } | Out-Null
    [pscustomobject]@{Field=$Snapshot.Field; AddedKeys=$added}
}

function Get-SiGlobalCreationPlan {
    param([Parameter(Mandatory)][string]$Root)
    $base=Join-Path $Root 'Config/SmartM365.global.local'
    if ((Test-Path -LiteralPath ($base+'.json')) -or (Test-Path -LiteralPath ($base+'.json.txt'))) {throw 'A global configuration already exists or is conflicting.'}
    $path=$base+'.json.template';$bytes=[IO.File]::ReadAllBytes($path)
    $model=[Text.UTF8Encoding]::new($false,$true).GetString($bytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
    if ($model -isnot [pscustomobject] -or (Test-SiConfigSensitiveObject $model)) {throw 'The global template is invalid or contains sensitive fields; review it separately.'}
    $sha=[Security.Cryptography.SHA256]::Create();try {$hash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','')} finally {$sha.Dispose()}
    $review=ConvertTo-Json -InputObject $model -Depth 100
    if ($review.Length -gt 12000) {throw 'Global template is too large to review here.'}
    [pscustomobject]@{Root=$Root;TemplateHash=$hash;Destination=$base+'.json.txt';Review=$review}
}
function New-SiGlobalConfiguration {
    param([Parameter(Mandatory)]$Plan,[object[]]$Edits=@())
    $fresh=Get-SiGlobalCreationPlan -Root $Plan.Root
    if ($fresh.TemplateHash -cne $Plan.TemplateHash -or $fresh.Destination -cne $Plan.Destination -or $fresh.Review -cne $Plan.Review) {throw 'Global template changed; review it again.'}
    $model=$fresh.Review | ConvertFrom-Json
    foreach ($edit in $Edits) {if ($edit.Step -ne 'share' -or $edit.Field -notin @(Get-SiEditableFields share) -or $null -eq $model.PSObject.Properties[$edit.Field]) {throw 'Invalid pending global setting.'};$model.($edit.Field)=ConvertTo-SiEditValue -Snapshot $edit -Value $edit.Value}
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($model | ConvertTo-Json -Depth 100)+[Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $fresh.Destination -Bytes $bytes -ExpectedSHA256 'ABSENT' -Validate {param($d) if ($d -isnot [pscustomobject]) {throw 'Expected JSON object.'}} | Out-Null
    if (Test-Path -LiteralPath ($fresh.Destination.Substring(0,$fresh.Destination.Length-4))) {throw 'Another writer created a legacy global configuration concurrently; both files were preserved and require review.'}
    return [pscustomobject]@{Path=$fresh.Destination}
}
function Get-SiTenantCreationPlan {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$ProfileKey,[Parameter(Mandatory)][string]$OrganizationKey,[Parameter(Mandatory)][string]$EnvironmentKey)
    if ($ProfileKey -cnotmatch '^[a-z0-9][a-z0-9_-]{0,63}$') { throw 'Invalid profile key.' }
    foreach ($key in @($OrganizationKey,$EnvironmentKey)) {
        if ($key -cnotmatch '^[a-z0-9][a-z0-9-]{0,63}$') { throw 'Organization and environment keys require lowercase letters, digits or hyphens.' }
    }
    $dir = Join-Path $Root 'Config/Tenants'
    $base = Join-Path $dir ($ProfileKey + '.local')
    if ((Test-Path -LiteralPath ($base + '.json')) -or (Test-Path -LiteralPath ($base + '.json.txt'))) { throw 'A profile already exists, is unreadable or is in conflict.' }
    $template = Join-Path $dir 'tenant.local.json.template'
    $templateBytes=[IO.File]::ReadAllBytes($template);$sha=[Security.Cryptography.SHA256]::Create()
    try {$templateHash=([BitConverter]::ToString($sha.ComputeHash($templateBytes))).Replace('-','')} finally {$sha.Dispose()}
    $model=[Text.UTF8Encoding]::new($false,$true).GetString($templateBytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
    if ($model -isnot [pscustomobject]) { throw 'Invalid tenant template.' }
    foreach ($key in @('ProfileKey','OrganizationKey','EnvironmentKey','TenantKey','Environment')) {
        if ($null -eq $model.PSObject.Properties[$key]) { throw 'Tenant template is missing an identity key.' }
    }
    $model.ProfileKey=$ProfileKey
    $model.OrganizationKey=$OrganizationKey
    $model.EnvironmentKey=$EnvironmentKey
    $model.TenantKey=$OrganizationKey+'-'+$EnvironmentKey
    $model.Environment=$EnvironmentKey
    foreach ($setting in @(@('EnableTeamsNotifications',$false),@('EnableSharePointUpload',$true))) {
        $property=$model.PSObject.Properties[[string]$setting[0]]
        if ($null -eq $property -or $property.Value -isnot [bool]) {throw 'The tenant template is missing a Boolean notification or SharePoint setting.'}
        $property.Value=[bool]$setting[1]
    }
    $path = $base + '.json.txt'
    [pscustomobject]@{Root=$Root;ProfileKey=$ProfileKey;OrganizationKey=$OrganizationKey;EnvironmentKey=$EnvironmentKey;TenantKey=$model.TenantKey;Destination=$path;TemplateHash=$templateHash;Model=$model}
}
function New-SiTenantProfile {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$ProfileKey,[Parameter(Mandatory)][string]$OrganizationKey,[Parameter(Mandatory)][string]$EnvironmentKey,[string]$ExpectedTemplateHash='',[object[]]$Edits=@())
    $plan=Get-SiTenantCreationPlan -Root $Root -ProfileKey $ProfileKey -OrganizationKey $OrganizationKey -EnvironmentKey $EnvironmentKey
    if ($ExpectedTemplateHash -and $plan.TemplateHash -cne $ExpectedTemplateHash) {throw 'Tenant template changed; review the new profile again.'}
    $model=$plan.Model
    foreach ($edit in $Edits) {if ($edit.Step -eq 'share' -or $edit.Field -notin @(Get-SiEditableFields $edit.Step) -or $null -eq $model.PSObject.Properties[$edit.Field]) {throw 'Invalid pending profile setting.'};$model.($edit.Field)=ConvertTo-SiEditValue -Snapshot $edit -Value $edit.Value}
    $path=$plan.Destination;$base=$path.Substring(0,$path.Length-9)
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($model | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $path -Bytes $bytes -ExpectedSHA256 'ABSENT' -Validate { param($d) if ($d -isnot [pscustomobject]) { throw 'Expected JSON object.' } } | Out-Null
    if (Test-Path -LiteralPath ($base + '.json')) { throw 'Another writer created a legacy profile concurrently; both files were preserved and require review.' }
    return [pscustomobject]@{Path=$path; ProfileKey=$ProfileKey; TenantKey=$model.TenantKey}
}

function Get-SiBooleanSnapshot {
    param([string]$Root,[string]$Tenant,[ValidateSet('sharepoint')][string]$Step)
    $field='EnableSharePointUpload'
    $legacy=Get-SiEditLocation $Root $Tenant $Step
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Configuration missing, invalid or conflicting.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single') { throw 'Coexisting configuration files require review.' }
    $present=$null -ne $read.Document.PSObject.Properties[$field]
    $delta=Get-SiTenantTemplateDelta -Root $Root -Document $read.Document -RequiredFields @($field)
    $property=$read.Document.PSObject.Properties[$field]
    if ($null -eq $property -or $property.Value -isnot [bool]) { throw 'Expected a Boolean value or template default.' }
    [pscustomobject]@{Root=$Root;Tenant=$Tenant;Step=$Step;Field=$field;Path=$read.Path;Hash=$read.SHA256;Value=$property.Value;Present=$present;TemplateHash=$delta.Hash;AddedKeys=$delta.AddedKeys;Additions=$delta.Additions}
}
function Set-SiBooleanSnapshot {
    param([Parameter(Mandatory)]$Snapshot,[Parameter(Mandatory)][bool]$Value)
    $field=switch ($Snapshot.Step) { 'sharepoint' {'EnableSharePointUpload'} default {throw 'Invalid step.'} }
    if ($Snapshot.Field -ne $field) { throw 'Invalid field.' }
    $legacy=Get-SiEditLocation $Snapshot.Root $Snapshot.Tenant $Snapshot.Step
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Configuration changed or conflicts.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -ne $Snapshot.Hash -or $read.Path -ne $Snapshot.Path) { throw 'Configuration changed; refresh first.' }
    $delta=Confirm-SiTenantTemplateDelta -Snapshot $Snapshot -Document $read.Document -RequiredFields @($field);$added=@($delta.AddedKeys)
    if ($read.Document.$field -isnot [bool]) { throw 'Expected a Boolean.' }
    $read.Document.$field=$Value
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $Snapshot.Hash -Validate {param($d) if ($d -isnot [pscustomobject]) {throw 'Expected JSON object.'}} | Out-Null
    [pscustomobject]@{Field=$field;AddedKeys=$added}
}

function Get-SiUploadCategorySnapshot {
    param([string]$Root,[string]$Tenant,[ValidateSet('DataAll','DataLast','DataPowerBI','LogAll')][string]$Category)
    $field='EnableSharePointUpload'+$Category
    $legacy=Get-SiEditLocation $Root $Tenant 'sharepoint'
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Configuration missing, invalid or conflicting.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single') { throw 'Coexisting configuration files require review.' }
    $present=$null -ne $read.Document.PSObject.Properties[$field]
    $delta=Get-SiTenantTemplateDelta -Root $Root -Document $read.Document -RequiredFields @($field)
    $p=$read.Document.PSObject.Properties[$field]
    $value=if ($null -eq $p) { $null } else { $p.Value }
    if ($null -eq $value -or ($value -is [string] -and $value.Trim() -in @('','__USE_GLOBAL__','USE_GLOBAL'))) { $choice='Inherit' }
    elseif ($value -is [bool]) { $choice=if ($value) { 'Enabled' } else { 'Disabled' } }
    else { throw 'Unsupported category value; review the configuration.' }
    [pscustomobject]@{Root=$Root;Tenant=$Tenant;Category=$Category;Field=$field;Path=$read.Path;Hash=$read.SHA256;Choice=$choice;Present=$present;TemplateHash=$delta.Hash;AddedKeys=$delta.AddedKeys;Additions=$delta.Additions}
}
function Set-SiUploadCategorySnapshot {
    param([Parameter(Mandatory)]$Snapshot,[Parameter(Mandatory)][ValidateSet('Inherit','Enabled','Disabled')][string]$Choice)
    if ($Snapshot.Category -notin @('DataAll','DataLast','DataPowerBI','LogAll') -or $Snapshot.Field -ne ('EnableSharePointUpload'+$Snapshot.Category)) { throw 'Invalid upload category.' }
    $legacy=Get-SiEditLocation $Snapshot.Root $Snapshot.Tenant 'sharepoint'
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Configuration changed or conflicts.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -ne $Snapshot.Hash -or $read.Path -ne $Snapshot.Path) { throw 'Configuration changed; refresh first.' }
    $delta=Confirm-SiTenantTemplateDelta -Snapshot $Snapshot -Document $read.Document -RequiredFields @($Snapshot.Field);$added=@($delta.AddedKeys)
    $read.Document.($Snapshot.Field)=switch ($Choice) { 'Enabled' {$true} 'Disabled' {$false} default {'__USE_GLOBAL__'} }
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $Snapshot.Hash -Validate {param($d) if ($d -isnot [pscustomobject]) {throw 'Expected JSON object.'}} | Out-Null
    [pscustomobject]@{Field=$Snapshot.Field;AddedKeys=$added}
}

function Get-SiCertificateEditSnapshot {
    param([string]$Root,[string]$Tenant)
    $legacy=Get-SiEditLocation $Root $Tenant 'identity'
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Configuration missing, invalid or conflicting.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single') { throw 'Coexisting configuration files require review.' }
    $thumb=[string]$read.Document.Thumb;$thumbprint=[string]$read.Document.Thumbprint
    $delta=Get-SiTenantTemplateDelta -Root $Root -Document $read.Document -RequiredFields @('Thumb','Thumbprint')
    [pscustomobject]@{Root=$Root;Tenant=$Tenant;Path=$read.Path;Hash=$read.SHA256;Thumb=$thumb;Thumbprint=$thumbprint;TemplateHash=$delta.Hash;AddedKeys=$delta.AddedKeys;Additions=$delta.Additions}
}
function Set-SiCertificateEditSnapshot {
    param([Parameter(Mandatory)]$Snapshot,[Parameter(Mandatory)][string]$Thumbprint)
    $thumb=$Thumbprint.Replace(' ','').ToUpperInvariant()
    if ($thumb -notmatch '^[0-9A-F]{40}$' -or $thumb -match '^0+$') { throw 'Enter a nonzero SHA-1 certificate thumbprint of 40 hexadecimal characters.' }
    $legacy=Get-SiEditLocation $Snapshot.Root $Snapshot.Tenant 'identity'
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Configuration changed or conflicts.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -ne $Snapshot.Hash -or $read.Path -ne $Snapshot.Path) { throw 'Configuration changed; refresh first.' }
    $delta=Confirm-SiTenantTemplateDelta -Snapshot $Snapshot -Document $read.Document -RequiredFields @('Thumb','Thumbprint');$added=@($delta.AddedKeys)
    $read.Document.Thumb=$thumb; $read.Document.Thumbprint=$thumb
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $Snapshot.Hash -Validate {param($d) if ($d -isnot [pscustomobject] -or [string]$d.Thumb -cne [string]$d.Thumbprint) {throw 'Certificate fields must agree.'}} | Out-Null
    [pscustomobject]@{AddedKeys=$added}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDTDdITSkrcoHBU
# MUoozL0YTzbapQVFJNtnloXRxP3/LaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGycNVzK/Pvnx+Qk09N1zYZH8l8DOsgV0G/WpgpvHa8DMA0GCSqG
# SIb3DQEBAQUABIIBgKrar+RAlP1kAMvIf8FiwnBLhtMmB/u1EH41VuguEDSW9tll
# bdkojahr78W/tvIYaiZVeOtIdfr93KNHrcLZq+NLqkMxRmtpoLL0ZsSKmdXMmpQI
# VcpATOGAM4gbcGE4WdbrK65es/lkFp700LxbxdVQ0RSMy2Z14Rk403qBEJWeElmp
# Nm/5M0R7NNIhCJWeR5cTu5ojTWPBV3j/ALz7eXj+Kb7yWucY5ybHroX29m5VCUI9
# +p1908ShiaHedgdaZxojiL94+D00Bnvo3zgtraKWPnw7Lhr0wq0OM3xDEsfNk1rb
# BWgFV57mTj/S6je7wHG8qCV+TvHJk5lLqgVMHDvJLtQbBAXmLh3nJIMhWI5YkeEj
# BypmLkMXphV7q3Pz0Jvl9mXIaC/rxCW0hVrsbRhjVk96LWEtwSJnp7njYyJRquqP
# ri7GBVypWX2tu9eKJyaoBQ16i7fMRP4n4rE4+roS0PKpfkj/9Mb3c4HIEYQMu0k+
# CzULjsI1vk+kD5w2LKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjZaMC8GCSqGSIb3DQEJBDEiBCBTxBPQGB54r3yctGk95Hn+l5Wn2lao5FwOinUk
# 9bl58zANBgkqhkiG9w0BAQEFAASCAgA4Mwg8UtLgjQmkqwCFfhqdVWCvQ55sg4Wg
# 5RtpkUQjXE2g2O3HqOHGlWPNv6U8uOYzyS9LddV0acNeO34WmKo0Vt0MN7psnZxJ
# LreSMkTe+yDxFB9+kmplvry7xnTtoqgb6tvajblMA/Ye5I+ezbCK9fyiAR9iHV0N
# GRr3ugHQLCHDI7EN8IcQPxnnKQ2hOwhWjJT0a2SP6zdN7qSKSS0x9N8HBetp2rMC
# 7mtVIXTehvJO7X6c8OkHhHlh1b62ZjYwnGPMbusZft20VlrzYSaT19KlIsXWLBnk
# gBFyQoQq+zZg/EhntaVooulns776Yg8dM4pFLvcvWKVJqpEBsxkWDA2qOjGQNIe8
# pr+L7qIlpPnv41TYgTJPqY1IsaDqSmgB7aErx/yEECF+o9XrmdBCCLJCL0RpImpm
# Wc8zqDUgDb1T+hMqMxQqh9Dv1xGJDop9KTUr4SXkAbawr82lcHGpu0fZjDCNRZd6
# PuC8jQMcJXNweJJDyurAySmWbvFkXjGbCBroCpTTa+A3asVvWZFjwbZri1PnW1vP
# ZFwmHU0FCRKHeF1PrKXVAqojx+s9ggAy1fW0yQxWzZZtFaq4tp9unVYYI7p//pLR
# B4gm2wX+ce433So32UKE+bA2JV6SQpek0+DeH6kU3uoZ7Sxv7LIQIFpjp1k8pr2C
# nPhs4imtAw==
# SIG # End signature block
