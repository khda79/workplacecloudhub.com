# SmartM365.Core.psm1
# Core utilities: logging, initialization, cleanup, CSV export, mail, remote scheduling, cloud connections.

function Get-ModuleLocalConfig {
    [CmdletBinding()]
    param()

    $modulePath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    $configPath = Join-Path -Path (Split-Path -Parent $modulePath) -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($modulePath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        return [pscustomobject]@{}
    }

    try {
        return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw ("Failed to read module local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Initialize-SmartM365ModuleGlobalConfigFromTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$TemplatePath
    )

    if (Test-Path -LiteralPath $Path) { return $true }
    if (-not (Test-Path -LiteralPath $TemplatePath)) { return $false }

    $targetFolder = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($targetFolder) -and -not (Test-Path -LiteralPath $targetFolder)) {
        New-Item -Path $targetFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    Copy-Item -LiteralPath $TemplatePath -Destination $Path -Force -ErrorAction Stop
    Write-Host "Created missing global local configuration from template." -ForegroundColor Yellow
    Write-Host ("Local JSON : {0}" -f $Path) -ForegroundColor Yellow
    Write-Host ("Template   : {0}" -f $TemplatePath) -ForegroundColor Yellow
    Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
    return $true
}
function Get-SmartM365EffectiveModuleGlobalConfig {
    [CmdletBinding()]
    param()

    if ($null -ne $global:SmartM365GlobalConfig -and $global:SmartM365GlobalConfig.PSObject.Properties.Count -gt 0) {
        $script:SmartM365GlobalConfig = $global:SmartM365GlobalConfig
        return $script:SmartM365GlobalConfig
    }

    $requestedProfileKey = if (-not [string]::IsNullOrWhiteSpace([string]$global:SmartM365ProfileKey)) { [string]$global:SmartM365ProfileKey } else { [string]$global:SmartM365Tenant }
    if ($null -ne $script:SmartM365GlobalConfig) {
        $cachedProfileKey = ''
        $cachedProfileProperty = $script:SmartM365GlobalConfig.PSObject.Properties['ProfileKey']
        if ($null -ne $cachedProfileProperty -and $null -ne $cachedProfileProperty.Value) { $cachedProfileKey = [string]$cachedProfileProperty.Value }
        if ([string]::IsNullOrWhiteSpace($requestedProfileKey) -or [string]::IsNullOrWhiteSpace($cachedProfileKey) -or $requestedProfileKey -ieq $cachedProfileKey) {
            return $script:SmartM365GlobalConfig
        }
    }

    $script:SmartM365GlobalConfig = [pscustomobject]@{}
    $modulePath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    $searchRoot = Split-Path -Path $modulePath -Parent
    while ($searchRoot) {
        $tenantContextCandidates = @(
            (Join-Path -Path $searchRoot -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($tenantContextPath in $tenantContextCandidates) {
            if (Test-Path -LiteralPath $tenantContextPath) {
                . $tenantContextPath
                $profileKey = if ([string]::IsNullOrWhiteSpace($requestedProfileKey)) { 'test' } else { $requestedProfileKey }
                $script:SmartM365GlobalConfig = Get-SmartM365EffectiveGlobalConfig -StartPath $searchRoot -ProfileKey $profileKey
                break
            }
        }
        if ($null -ne $script:SmartM365GlobalConfig -and $script:SmartM365GlobalConfig.PSObject.Properties.Count -gt 0) { break }

        $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
        $globalTemplatePath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json.template'
        if (-not (Test-Path -LiteralPath $globalConfigPath)) { Initialize-SmartM365ModuleGlobalConfigFromTemplate -Path $globalConfigPath -TemplatePath $globalTemplatePath | Out-Null }
        if (Test-Path -LiteralPath $globalConfigPath) {
            try { $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
            catch { throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message) }
            break
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    return $script:SmartM365GlobalConfig
}
function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    if ($Value -notmatch '\{\{[^}]+\}\}') {
        return $Value
    }

    $script:SmartM365GlobalConfig = Get-SmartM365EffectiveModuleGlobalConfig

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}
function Get-ModuleLocalConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                return Resolve-SmartM365ConfigValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartM365ConfigValue -Value $property.Value
        }
    }


    $script:SmartM365GlobalConfig = Get-SmartM365EffectiveModuleGlobalConfig

    $globalProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string]) {
            $globalValue = $globalProperty.Value.Trim()
            if ([string]::IsNullOrWhiteSpace($globalValue) -or $globalValue -in @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                return $DefaultValue
            }
        }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

function Get-SmartM365CallerLocalConfig {
    [CmdletBinding()]
    param()

    foreach ($frame in Get-PSCallStack) {
        $scriptPath = $frame.ScriptName
        if ([string]::IsNullOrWhiteSpace($scriptPath) -or $scriptPath -notmatch '\.ps1$') { continue }
        $configPath = Join-Path -Path (Split-Path -Path $scriptPath -Parent) -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($scriptPath))
        if (-not (Test-Path -LiteralPath $configPath)) { continue }
        try { return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
        catch {
            WriteLog -Message ("Failed to read caller local configuration '{0}': {1}" -f $configPath, $_.Exception.Message) -Level 'WARNING'
            return [pscustomobject]@{}
        }
    }
    return [pscustomobject]@{}
}

function Get-SmartM365WeeklyHistoryConfigValue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    $callerConfig = Get-SmartM365CallerLocalConfig
    return Get-ModuleLocalConfigValue -Config $callerConfig -Name $Name -DefaultValue $DefaultValue
}

function Get-SmartM365IsoWeekName {
    [CmdletBinding()]
    param([datetime]$Date = (Get-Date))
    $calendar = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
    $dayOfWeek = $calendar.GetDayOfWeek($Date)
    if ($dayOfWeek -ge [System.DayOfWeek]::Monday -and $dayOfWeek -le [System.DayOfWeek]::Wednesday) { $Date = $Date.AddDays(3) }
    $week = $calendar.GetWeekOfYear($Date, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [System.DayOfWeek]::Monday)
    return '{0}-W{1:00}' -f $Date.Year, $week
}

function Get-SmartM365WeeklyHistoryFileName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $fileName = [System.IO.Path]::GetFileName($Path)
    return ($fileName -replace '_\d{8}[-_]\d{4,6}(?=\.csv$)', '')
}

function Save-SmartM365WeeklyInventoryHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$SourceFiles,
        [Parameter(Mandatory)][string]$HistoryRootPath,
        [int]$RetentionWeeks = 52,
        [string]$HistoryLabel = 'SmartM365 inventory'
    )
    if ([string]::IsNullOrWhiteSpace($HistoryRootPath)) { WriteLog -Message ("Weekly {0} history skipped: HistoryRootPath is empty." -f $HistoryLabel); return }
    $existingSourceFiles = @($SourceFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($existingSourceFiles.Count -eq 0) { WriteLog -Message ("Weekly {0} history skipped: no source CSV file found." -f $HistoryLabel); return }
    $weekName = Get-SmartM365IsoWeekName
    $weekFolder = Join-Path -Path $HistoryRootPath -ChildPath $weekName
    New-Item -Path $weekFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
    $copiedFiles = New-Object System.Collections.Generic.List[string]
    foreach ($sourceFile in $existingSourceFiles) {
        $destinationFileName = Get-SmartM365WeeklyHistoryFileName -Path $sourceFile
        $destinationFile = Join-Path -Path $weekFolder -ChildPath $destinationFileName
        if (Test-Path -LiteralPath $destinationFile -PathType Leaf) { continue }
        Copy-Item -LiteralPath $sourceFile -Destination $destinationFile -Force -ErrorAction Stop
        [void]$copiedFiles.Add($destinationFile)
    }
    $manifest = [pscustomobject][ordered]@{
        UpdatedAt       = (Get-Date).ToString('o')
        Week            = $weekName
        HistoryLabel    = $HistoryLabel
        HistoryRootPath = $HistoryRootPath
        Files           = @(Get-ChildItem -LiteralPath $weekFolder -Filter '*.csv' -File -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object { $_.Name })
    }
    $manifestPath = Join-Path -Path $weekFolder -ChildPath 'manifest.json'
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    if ($copiedFiles.Count -gt 0) { WriteLog -Message ("Weekly {0} history saved for {1}: {2} new file(s) in {3}" -f $HistoryLabel, $weekName, $copiedFiles.Count, $weekFolder) }
    else { WriteLog -Message ("Weekly {0} history already exists for {1}. Snapshot skipped: {2}" -f $HistoryLabel, $weekName, $weekFolder) }

    $historyUploadCandidates = @(Get-ChildItem -LiteralPath $weekFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.csv', '.json') } | Sort-Object Name)
    foreach ($historyUploadCandidate in $historyUploadCandidates) {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $historyUploadCandidate.FullName | Out-Null
    }

    if ($RetentionWeeks -gt 0) {
        $oldWeekFolders = @(Get-ChildItem -LiteralPath $HistoryRootPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}-W\d{2}$' } | Sort-Object Name -Descending | Select-Object -Skip $RetentionWeeks)
        foreach ($oldWeekFolder in $oldWeekFolders) {
            try { Remove-Item -LiteralPath $oldWeekFolder.FullName -Recurse -Force -ErrorAction Stop; WriteLog -Message ("Deleted old weekly {0} history folder: {1}" -f $HistoryLabel, $oldWeekFolder.FullName) }
            catch { WriteLog -Message ("Failed to delete old weekly {0} history folder '{1}': {2}" -f $HistoryLabel, $oldWeekFolder.FullName, $_.Exception.Message) -Level 'WARNING' }
        }
    }
}

function Add-SmartM365WeeklyHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$SourceCsvPaths,
        [string]$HistoryRootPath,
        [int]$RetentionWeeks = 52,
        [string]$HistoryLabel = 'SmartM365 inventory'
    )
    if ([string]::IsNullOrWhiteSpace($HistoryRootPath)) {
        $firstSource = @($SourceCsvPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) | Select-Object -First 1
        if ($firstSource) {
            $sourceFolder = Split-Path -Path $firstSource -Parent
            if (-not [string]::IsNullOrWhiteSpace($sourceFolder)) { $HistoryRootPath = Join-Path -Path $sourceFolder -ChildPath 'WeeklyHistory' }
        }
    }
    Save-SmartM365WeeklyInventoryHistory -SourceFiles $SourceCsvPaths -HistoryRootPath $HistoryRootPath -RetentionWeeks $RetentionWeeks -HistoryLabel $HistoryLabel
}
function Invoke-SmartM365WeeklyInventoryHistoryForCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$SourceFiles, [string]$TimestampedPath)
    $enabled = [bool](Get-SmartM365WeeklyHistoryConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true)
    if (-not $enabled) { return }
    $historyRootPath = Get-SmartM365WeeklyHistoryConfigValue -Name 'WeeklyHistoryFolderPath' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($historyRootPath) -and -not [string]::IsNullOrWhiteSpace($TimestampedPath)) {
        $sourceFolder = Split-Path -Path $TimestampedPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($sourceFolder)) { $historyRootPath = Join-Path -Path $sourceFolder -ChildPath 'WeeklyHistory' }
    }
    $retentionWeeks = [int](Get-SmartM365WeeklyHistoryConfigValue -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
    Save-SmartM365WeeklyInventoryHistory -SourceFiles $SourceFiles -HistoryRootPath $historyRootPath -RetentionWeeks $retentionWeeks
}

function Get-SmartM365MaxItemsValue {
    [CmdletBinding()]
    param()

    foreach ($name in @('SmartM365MaxItems','SmartM365TestMaxItems','MaxItems')) {
        $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
        if ($variable -and $null -ne $variable.Value) {
            try {
                $value = [int]$variable.Value
                if ($value -gt 0) { return $value }
            }
            catch { }
        }
    }

    return 0
}

function Test-SmartM365MaxItemsMode {
    [CmdletBinding()]
    param()

    return ((Get-SmartM365MaxItemsValue) -gt 0)
}

function Get-SmartM365MaxItemsSuffix {
    [CmdletBinding()]
    param()

    $maxItems = Get-SmartM365MaxItemsValue
    if ($maxItems -le 0) { return '' }
    return ('_MAXITEMS-{0}' -f $maxItems)
}

function Set-SmartM365MaxItemsMode {
    [CmdletBinding()]
    param([int]$MaxItems)

    if ($MaxItems -gt 0) {
        $global:SmartM365MaxItems = $MaxItems
        $global:SmartM365TestMaxItems = $MaxItems
        $global:SmartM365IsMaxItemsRun = $true
        WriteLog -Message ("MaxItems test mode active: only a bounded dataset should be collected/exported. Power BI DATA-LAST and WeeklyHistory standard publication must not be updated by this run. MaxItems={0}" -f $MaxItems) -Level 'WARNING'
    }
}

function Add-SmartM365MaxItemsSuffixToCsvPath {
    [CmdletBinding()]
    param([AllowNull()][string]$Path)

    $suffix = Get-SmartM365MaxItemsSuffix
    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($suffix)) { return $Path }

    $folder = Split-Path -Path $Path -Parent
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)
    if ($name -match [regex]::Escape($suffix)) { return $Path }
    if ([string]::IsNullOrWhiteSpace($folder)) { return ("{0}{1}{2}" -f $name, $suffix, $extension) }
    return (Join-Path -Path $folder -ChildPath ("{0}{1}{2}" -f $name, $suffix, $extension))
}

function Add-SmartM365MaxItemsSuffixToBaseName {
    [CmdletBinding()]
    param([string]$BaseFileName)

    $suffix = Get-SmartM365MaxItemsSuffix
    if ([string]::IsNullOrWhiteSpace($BaseFileName) -or [string]::IsNullOrWhiteSpace($suffix)) { return $BaseFileName }
    if ($BaseFileName -match [regex]::Escape($suffix)) { return $BaseFileName }
    return ("{0}{1}" -f $BaseFileName, $suffix)
}

function Add-SmartM365MaxItemsMailBanner {
    [CmdletBinding()]
    param([AllowNull()][string]$BodyHtml)

    $maxItems = Get-SmartM365MaxItemsValue
    if ($maxItems -le 0) { return $BodyHtml }
    if ($BodyHtml -match 'SmartM365MaxItemsBanner') { return $BodyHtml }

    $banner = @"
<div class="SmartM365MaxItemsBanner" style="margin:12px 0;padding:12px 14px;border:1px solid #f59e0b;border-left:6px solid #d97706;background:#fffbeb;color:#92400e;font-family:Segoe UI,Arial,sans-serif;font-size:13px;line-height:19px;">
  <strong>TEST RUN - MaxItems=$maxItems.</strong> This email and the generated CSV files come from a bounded SmartM365 test run. CSV names include <code>MAXITEMS-$maxItems</code> and must not be used by Power BI production datasets.
</div>
"@

    if ([string]::IsNullOrWhiteSpace($BodyHtml)) { return $banner }
    if ($BodyHtml -match '(?is)<body\b[^>]*>') {
        return ([regex]::Replace($BodyHtml, '(?is)(<body\b[^>]*>)', ('$1' + "`n" + $banner), 1))
    }
    return ($banner + "`n" + [string]$BodyHtml)
}

function Add-SmartM365MaxItemsSubjectPrefix {
    [CmdletBinding()]
    param([AllowNull()][string]$Subject)

    $maxItems = Get-SmartM365MaxItemsValue
    if ($maxItems -le 0) { return $Subject }
    $prefix = "[MAXITEMS-$maxItems TEST]"
    if ([string]::IsNullOrWhiteSpace($Subject)) { return $prefix }
    if ($Subject -like "*MAXITEMS-$maxItems*") { return $Subject }
    return ("{0} {1}" -f $prefix, $Subject)
}

function Limit-SmartM365RowsForMaxItems {
    [CmdletBinding()]
    param([AllowNull()][object[]]$Data)

    $maxItems = Get-SmartM365MaxItemsValue
    if ($maxItems -le 0 -or $null -eq $Data) { return @($Data) }
    $rows = @($Data)
    if ($rows.Count -le $maxItems) { return $rows }
    WriteLog -Message ("MaxItems={0}: limiting exported CSV rows from {1} to {0}." -f $maxItems, $rows.Count) -Level 'WARNING'
    return @($rows | Select-Object -First $maxItems)
}
#region Logging and file helpers

function Format-SmartM365LogLine {
    param(
        [AllowEmptyString()][string]$Message,
        [string]$Level = "INFO",
        [datetime]$Timestamp = (Get-Date)
    )

    $timestampText = $Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
    $normalizedLevel = if ([string]::IsNullOrWhiteSpace($Level)) { "INFO" } else { $Level.ToUpperInvariant() }
    $messageLines = @([regex]::Split(([string]$Message), '\r?\n'))
    foreach ($messageLine in $messageLines) {
        "{0} [{1}] {2}" -f $timestampText, $normalizedLevel, $messageLine
    }
}

function Update-SmartM365TimestampedTranscript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $timestampText = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $lines = [System.IO.File]::ReadAllLines($Path)
    $timestampPattern = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b'
    $changed = $false
    $updatedLines = foreach ($line in $lines) {
        if ($line -match $timestampPattern) {
            $line
        }
        elseif ($line -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s*(.*)$') {
            $changed = $true
            "{0} {1}" -f $Matches[1], $Matches[2]
        }
        else {
            $changed = $true
            "{0} {1}" -f $timestampText, $line
        }
    }

    if ($changed) {
        [System.IO.File]::WriteAllLines($Path, [string[]]$updatedLines, [System.Text.UTF8Encoding]::new($false))
    }
}

function WriteLog {
    param(
        [string]$Message,
        [string]$Level = ""
    )

    # Auto-detect level if not provided
    if (-not $Level -or $Level -eq "") {
        if ($Message -match "(?i)\berror\b|\bfailed\b|\bfailure\b") {
            $Level = "ERROR"
        }
        else {
            $Level = "INFO"
        }
    }
    $normalizedLevel = $Level.ToUpperInvariant()
    if ($normalizedLevel -eq 'WARNING') {
        $global:SmartM365WarningCount = [int]$global:SmartM365WarningCount + 1
    }
    elseif ($normalizedLevel -eq 'ERROR') {
        $global:SmartM365ErrorCount = [int]$global:SmartM365ErrorCount + 1
    }

    $logEntry = @(Format-SmartM365LogLine -Message $Message -Level $Level)

    switch ($normalizedLevel) {
        "ERROR"   { $logEntry | ForEach-Object { Write-Host $_ -ForegroundColor Red } }
        "WARNING" { $logEntry | ForEach-Object { Write-Host $_ -ForegroundColor Yellow } }
        "INFO"    { $logEntry | ForEach-Object { Write-Host $_ -ForegroundColor Cyan } }
        "DEBUG"   { $logEntry | ForEach-Object { Write-Host $_ -ForegroundColor Gray } }
        "SUCCESS" { $logEntry | ForEach-Object { Write-Host $_ -ForegroundColor Green } }
        default   { $logEntry | ForEach-Object { Write-Host $_ -ForegroundColor White } }
    }

    if ($global:LogTextFile) {
        Add-Content -Path $global:LogTextFile -Value $logEntry
    }

    Invoke-SmartM365TeamsNotificationFromLog -Message $Message -Level $Level
}


function Get-SmartM365ModuleDiagnosticText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [switch]$IncludeAvailable
    )

    $module = @(Get-Module -Name $Name | Sort-Object Version -Descending | Select-Object -First 1)[0]
    $state = 'loaded'
    if (-not $module -and $IncludeAvailable) {
        $module = @(Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1)[0]
        $state = 'available'
    }

    if (-not $module) {
        return ("{0}: not found" -f $Name)
    }

    $version = if ($module.Version) { $module.Version.ToString() } else { 'unknown' }
    $displayName = if (-not [string]::IsNullOrWhiteSpace([string]$module.Name)) { [string]$module.Name } else { $Name }
    $text = ("{0}: {1} ({2})" -f $displayName, $version, $state)
    if (-not [string]::IsNullOrWhiteSpace([string]$module.Path)) {
        $text = "{0}; Path={1}" -f $text, $module.Path
    }
    return $text
}

function Write-SmartM365LoadedModuleVersions {
    [CmdletBinding()]
    param(
        [string]$Header = 'Loaded module versions:',
        [string[]]$ModuleNames = @(
            'SmartM365.Core',
            'SmartM365-WindowsPowerShell5',
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.DeviceManagement',
            'Microsoft.Graph.Devices',
            'Microsoft.Graph.Reports',
            'ExchangeOnlineManagement',
            'ActiveDirectory'
        )
    )

    $names = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($name in $ModuleNames) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($seen.Add($name)) { $names.Add($name) }
    }

    foreach ($module in @(Get-Module | Where-Object { $_.Name -like 'SmartM365*' -or $_.Name -like 'Microsoft.Graph*' -or $_.Name -eq 'ExchangeOnlineManagement' -or $_.Name -eq 'ActiveDirectory' } | Sort-Object Name)) {
        if ($module -and $seen.Add([string]$module.Name)) { $names.Add([string]$module.Name) }
    }

    $diagnostics = New-Object 'System.Collections.Generic.List[string]'
    foreach ($name in $names) {
        $module = @(Get-Module -Name $name | Sort-Object Version -Descending | Select-Object -First 1)[0]
        if (-not $module) { continue }
        $diagnostics.Add((Get-SmartM365ModuleDiagnosticText -Name $name))
    }

    if ($diagnostics.Count -eq 0) { return }

    WriteLog -Message $Header -Level 'INFO'
    foreach ($diagnostic in $diagnostics) {
        WriteLog -Message ("  {0}" -f $diagnostic) -Level 'INFO'
    }
}
function Get-SmartM365ScriptVersionFromFile {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return '' }

    try {
        $content = [System.IO.File]::ReadAllText($Path)
        $match = [regex]::Match($content, '(?im)^[ \t]*\.VERSION[ \t]*\r?\n[ \t]*(?<Version>[^\r\n]+)')
        if ($match.Success) {
            return $match.Groups['Version'].Value.Trim()
        }
    }
    catch {
    }

    return ''
}

function Write-SmartM365BrandBanner {
    [CmdletBinding()]
    param()

    $bannerLines = @(
        '================================================================================',
        ' SmartM365 by WorkplaceCloudHub',
        ' Website : https://workplacecloudhub.com',
        ' GitHub  : https://github.com/khda79/workplacecloudhub.com',
        '================================================================================'
    )

    $bannerAlreadyShown = Get-Variable -Name SmartM365BrandBannerConsoleShown -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if (-not [bool]$bannerAlreadyShown) {
        Set-Variable -Name SmartM365BrandBannerConsoleShown -Scope Global -Value $true
        Microsoft.PowerShell.Utility\Write-Host $bannerLines[0] -ForegroundColor DarkCyan
        Microsoft.PowerShell.Utility\Write-Host $bannerLines[1] -ForegroundColor Cyan
        Microsoft.PowerShell.Utility\Write-Host $bannerLines[2] -ForegroundColor Yellow
        Microsoft.PowerShell.Utility\Write-Host $bannerLines[3] -ForegroundColor Yellow
        Microsoft.PowerShell.Utility\Write-Host $bannerLines[4] -ForegroundColor DarkCyan
    }

    $logTextFile = [string](Get-Variable -Name LogTextFile -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    $loggedPath = [string](Get-Variable -Name SmartM365BrandBannerLoggedPath -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    if (-not [string]::IsNullOrWhiteSpace($logTextFile) -and $loggedPath -ne $logTextFile) {
        Add-Content -Path $logTextFile -Value $bannerLines
        Set-Variable -Name SmartM365BrandBannerLoggedPath -Scope Global -Value $logTextFile
    }
}
function Write-SmartM365CompletionBanner {
    [CmdletBinding()]
    param(
        [ValidateSet('Auto', 'Success', 'Failed', 'CompletedWithWarnings')]
        [string]$Status = 'Auto',
        [string]$ScriptName = '',
        [AllowNull()][Nullable[datetime]]$StartedAt = $null,
        [AllowNull()][Nullable[datetime]]$EndedAt = $null,
        [int]$WarningCount = -1,
        [int]$ErrorCount = -1,
        [int]$GeneratedCsvFiles = -1,
        [string]$LogPath = ''
    )
    $scriptNameVariable = Get-Variable -Name SmartM365ScriptName -Scope Global -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($ScriptName) -and $scriptNameVariable) { $ScriptName = [string]$scriptNameVariable.Value }
    $startedAtVariable = Get-Variable -Name SmartM365ExecutionStartTime -Scope Global -ErrorAction SilentlyContinue
    if (($null -eq $StartedAt -or $StartedAt -eq [datetime]::MinValue) -and $startedAtVariable) { $StartedAt = [datetime]$startedAtVariable.Value }
    $warningVariable = Get-Variable -Name SmartM365WarningCount -Scope Global -ErrorAction SilentlyContinue
    if ($WarningCount -lt 0) { $WarningCount = if ($warningVariable) { [int]$warningVariable.Value } else { 0 } }
    $errorVariable = Get-Variable -Name SmartM365ErrorCount -Scope Global -ErrorAction SilentlyContinue
    if ($ErrorCount -lt 0) { $ErrorCount = if ($errorVariable) { [int]$errorVariable.Value } else { 0 } }
    $csvVariable = Get-Variable -Name csvGeneratedPaths -Scope Global -ErrorAction SilentlyContinue
    if ($GeneratedCsvFiles -lt 0) { $GeneratedCsvFiles = if ($csvVariable -and $csvVariable.Value) { @($csvVariable.Value).Count } else { 0 } }
    $logVariable = Get-Variable -Name LogTextFile -Scope Global -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($LogPath) -and $logVariable) { $LogPath = [string]$logVariable.Value }


    if ($null -eq $EndedAt -or $EndedAt -eq [datetime]::MinValue) { $EndedAt = Get-Date }
    if ($null -eq $StartedAt -or $StartedAt -eq [datetime]::MinValue) { $StartedAt = $EndedAt }
    if ([string]::IsNullOrWhiteSpace($ScriptName)) { $ScriptName = 'SmartM365' }
    if ($Status -eq 'Auto') {
        $Status = if ($ErrorCount -gt 0) { 'Failed' } elseif ($WarningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
    }

    $statusLabel = switch ($Status) {
        'Failed' { 'FAILED' }
        'CompletedWithWarnings' { 'COMPLETED WITH WARNINGS' }
        default { 'SUCCESS' }
    }
    $title = switch ($Status) {
        'Failed' { 'Execution failed' }
        'CompletedWithWarnings' { 'Execution completed with warnings' }
        default { 'Execution completed' }
    }
    $statusColor = switch ($Status) {
        'Failed' { 'Red' }
        'CompletedWithWarnings' { 'Yellow' }
        default { 'Green' }
    }
    $duration = $EndedAt - $StartedAt
    $durationText = '{0:00}:{1:00}:{2:00}' -f ([int][Math]::Floor($duration.TotalHours)), $duration.Minutes, $duration.Seconds
    $effectiveLogPath = if ([string]::IsNullOrWhiteSpace($LogPath)) { '' } else { [string]$LogPath }
    $runKey = '{0}|{1}|{2}' -f $ScriptName, $StartedAt.Ticks, $effectiveLogPath
    $existingRunKey = [string](Get-Variable -Name SmartM365CompletionBannerRunKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    if ($existingRunKey -eq $runKey) { return }

    $lines = @(
        '================================================================================',
        (' SmartM365 by WorkplaceCloudHub - {0}' -f $title),
        (' Script    : {0}' -f $ScriptName),
        (' Status    : {0}' -f $statusLabel),
        (' Duration  : {0}' -f $durationText),
        (' Warnings  : {0}' -f $WarningCount),
        (' Errors    : {0}' -f $ErrorCount),
        (' CSV files : {0}' -f $GeneratedCsvFiles)
    )
    if (-not [string]::IsNullOrWhiteSpace($effectiveLogPath)) { $lines += (' Log       : {0}' -f $effectiveLogPath) }
    $lines += '================================================================================'

    Set-Variable -Name SmartM365CompletionBannerRunKey -Scope Global -Value $runKey
    Microsoft.PowerShell.Utility\Write-Host $lines[0] -ForegroundColor DarkCyan
    Microsoft.PowerShell.Utility\Write-Host $lines[1] -ForegroundColor Cyan
    for ($i = 2; $i -lt ($lines.Count - 1); $i++) {
        Microsoft.PowerShell.Utility\Write-Host $lines[$i] -ForegroundColor $(if ($lines[$i] -like ' Status*') { $statusColor } else { 'Gray' })
    }
    Microsoft.PowerShell.Utility\Write-Host $lines[-1] -ForegroundColor DarkCyan

    if (-not [string]::IsNullOrWhiteSpace($effectiveLogPath)) {
        try {
            $logFolder = Split-Path -Path $effectiveLogPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($logFolder) -and -not (Test-Path -LiteralPath $logFolder)) {
                New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
            }
            Add-Content -LiteralPath $effectiveLogPath -Value $lines -Encoding UTF8
        }
        catch {
            Microsoft.PowerShell.Utility\Write-Warning ("SmartM365 completion banner could not be written to '{0}': {1}" -f $effectiveLogPath, $_.Exception.Message)
        }
    }
}
function Write-SmartM365ExecutionContext {
    [CmdletBinding()]
    param(
        [string]$ScriptName = $global:SmartM365ScriptName,
        [string]$ScriptVersion = '',
        [string]$ProfileKey = $global:SmartM365ProfileKey,
        [string]$TenantKey = $global:SmartM365TenantKey,
        [string]$OrganizationKey = $global:SmartM365OrganizationKey,
        [string]$EnvironmentKey = $global:SmartM365EnvironmentKey,
        [string]$TenantId = $global:SmartM365TenantId,
        [string]$OutputPath = $global:BasePath,
        [string]$ScriptPath = ''
    )

    $now = Get-Date
    $context = [ordered]@{}

    if (-not [string]::IsNullOrWhiteSpace($ScriptName)) { $context['ScriptName'] = $ScriptName }
    if (-not [string]::IsNullOrWhiteSpace($ScriptVersion)) { $context['ScriptVersion'] = $ScriptVersion }
    if (-not [string]::IsNullOrWhiteSpace($ProfileKey)) { $context['ProfileKey'] = $ProfileKey }
    if (-not [string]::IsNullOrWhiteSpace($TenantKey)) { $context['TenantKey'] = $TenantKey }
    if (-not [string]::IsNullOrWhiteSpace($OrganizationKey)) { $context['OrganizationKey'] = $OrganizationKey }
    if (-not [string]::IsNullOrWhiteSpace($EnvironmentKey)) { $context['EnvironmentKey'] = $EnvironmentKey }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $context['TenantId'] = $TenantId }

    $context['StartTimeLocal'] = $now.ToString('yyyy-MM-dd HH:mm:ss zzz')
    $context['StartTimeUtc'] = $now.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ssZ')

    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        if ($identity -and -not [string]::IsNullOrWhiteSpace($identity.Name)) {
            $context['User'] = $identity.Name
        }

        $principal = New-Object -TypeName System.Security.Principal.WindowsPrincipal -ArgumentList $identity
        $context['RunAsAdmin'] = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
            $context['User'] = ('{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME).Trim('\')
        }
        $context['RunAsAdmin'] = 'Unknown'
    }

    $context['MachineName'] = [Environment]::MachineName
    $context['PowerShellVersion'] = $PSVersionTable.PSVersion.ToString()
    $context['PowerShellEdition'] = $PSVersionTable.PSEdition
    $context['ProcessId'] = $PID
    $context['Process64Bit'] = [Environment]::Is64BitProcess
    $context['OS64Bit'] = [Environment]::Is64BitOperatingSystem

    try {
        $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $context['OperatingSystem'] = ('{0} {1}' -f $operatingSystem.Caption, $operatingSystem.Version).Trim()
    }
    catch {
        $context['OperatingSystem'] = [Environment]::OSVersion.VersionString
    }

    if (-not [string]::IsNullOrWhiteSpace($ScriptPath)) { $context['ScriptPath'] = $ScriptPath }
    try { $context['WorkingDirectory'] = (Get-Location).Path } catch {}
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $context['OutputPath'] = $OutputPath }
    if (-not [string]::IsNullOrWhiteSpace($global:LogPath)) { $context['LogPath'] = $global:LogPath }
    if (-not [string]::IsNullOrWhiteSpace($global:LogTextFile)) { $context['LogTextFile'] = $global:LogTextFile }
    if (-not [string]::IsNullOrWhiteSpace($global:logTranscriptFile)) { $context['TranscriptFile'] = $global:logTranscriptFile }

    Write-SmartM365BrandBanner
    WriteLog -Message 'Execution context:' -Level 'INFO'
    foreach ($key in $context.Keys) {
        $value = $context[$key]
        if ($null -eq $value) { continue }
        if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) { continue }
        WriteLog -Message ('  {0}: {1}' -f $key, $value) -Level 'INFO'
    }
    Write-SmartM365LoadedModuleVersions
}

function Complete-SmartM365ExecutionContext {
    [CmdletBinding()]
    param(
        [ValidateSet('Auto', 'Success', 'Failed', 'CompletedWithWarnings')]
        [string]$Status = 'Auto',
        [AllowNull()]$ErrorRecord = $null,
        [string]$FailureStage = ''
    )

    if ($global:SmartM365ExecutionSummaryWritten) { return }
    $global:SmartM365ExecutionSummaryWritten = $true

    $ended = Get-Date
    $started = if ($global:SmartM365ExecutionStartTime) { [datetime]$global:SmartM365ExecutionStartTime } else { $ended }
    $duration = $ended - $started
    $durationText = '{0:00}:{1:00}:{2:00}' -f ([int][Math]::Floor($duration.TotalHours)), $duration.Minutes, $duration.Seconds

    $warningCount = if ($null -ne $global:SmartM365WarningCount) { [int]$global:SmartM365WarningCount } else { 0 }
    $errorCount = if ($null -ne $global:SmartM365ErrorCount) { [int]$global:SmartM365ErrorCount } else { 0 }

    $hasFailure = $null -ne $ErrorRecord -or $errorCount -gt 0
    if ($Status -eq 'Auto') {
        if ($hasFailure) {
            $Status = 'Failed'
        }
        elseif ($warningCount -gt 0) {
            $Status = 'CompletedWithWarnings'
        }
        else {
            $Status = 'Success'
        }
    }
    $global:SmartM365ExecutionStatus = $Status

    $summary = [ordered]@{
        Status            = $Status
        Duration          = $durationText
        Started           = $started.ToString('yyyy-MM-dd HH:mm:ss zzz')
        Ended             = $ended.ToString('yyyy-MM-dd HH:mm:ss zzz')
        ScriptName        = $global:SmartM365ScriptName
        ProfileKey       = $global:SmartM365ProfileKey
        TenantKey         = $global:SmartM365TenantKey
        OrganizationKey   = $global:SmartM365OrganizationKey
        EnvironmentKey    = $global:SmartM365EnvironmentKey
        TenantId          = $global:SmartM365TenantId
        OutputPath        = $global:BasePath
        LogTextFile       = $global:LogTextFile
        TranscriptFile    = $global:logTranscriptFile
        GeneratedCsvFiles = if ($global:csvGeneratedPaths) { @($global:csvGeneratedPaths).Count } else { 0 }
        Warnings          = $warningCount
        Errors            = $errorCount
    }

    if (-not [string]::IsNullOrWhiteSpace($FailureStage)) {
        $summary['FailureStage'] = $FailureStage
    }

    if ($null -ne $ErrorRecord) {
        $failureMessage = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord] -and $ErrorRecord.Exception) {
            $ErrorRecord.Exception.Message
        }
        elseif ($ErrorRecord.Exception) {
            $ErrorRecord.Exception.Message
        }
        else {
            [string]$ErrorRecord
        }
        if (-not [string]::IsNullOrWhiteSpace($failureMessage)) {
            $summary['FailureMessage'] = $failureMessage
        }
    }

    if ($global:SmartM365SharePointUploadedFiles) {
        $summary['SharePointUploads'] = @($global:SmartM365SharePointUploadedFiles).Count
    }

    WriteLog -Message 'Execution summary:' -Level 'INFO'
    foreach ($key in $summary.Keys) {
        $value = $summary[$key]
        if ($null -eq $value) { continue }
        if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) { continue }
        WriteLog -Message ('  {0}: {1}' -f $key, $value) -Level 'INFO'
    }

    try {
    # Upload run log files after the execution summary is written, so SharePoint keeps the final log content.
    $logUploadCandidates = @($global:LogTextFile, $global:logTranscriptFile)
    if ($global:SmartM365MailHtmlFiles) { $logUploadCandidates += @($global:SmartM365MailHtmlFiles) }
    $logUploadCandidates = $logUploadCandidates |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
        Sort-Object -Unique
    $logUploadStartCount = if ($global:SmartM365SharePointUploadedFiles) { @($global:SmartM365SharePointUploadedFiles).Count } else { 0 }
    foreach ($logUploadCandidate in $logUploadCandidates) {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $logUploadCandidate | Out-Null
    }
    $logUploadEndCount = if ($global:SmartM365SharePointUploadedFiles) { @($global:SmartM365SharePointUploadedFiles).Count } else { $logUploadStartCount }
    if ($logUploadEndCount -gt $logUploadStartCount) {
        WriteLog -Message ('Run log SharePoint uploads: {0}' -f ($logUploadEndCount - $logUploadStartCount)) -Level 'INFO'
        if (-not [string]::IsNullOrWhiteSpace([string]$global:LogTextFile) -and (Test-Path -LiteralPath $global:LogTextFile -PathType Leaf)) {
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $global:LogTextFile | Out-Null
        }
    }
    }
    finally {
        Write-SmartM365CompletionBanner `
            -Status $Status `
            -ScriptName $global:SmartM365ScriptName `
            -StartedAt $started `
            -EndedAt $ended `
            -WarningCount $warningCount `
            -ErrorCount $errorCount `
            -GeneratedCsvFiles $summary.GeneratedCsvFiles `
            -LogPath $global:LogTextFile
    }
}

function Invoke-SmartM365TeamsNotificationFromLog {
    [CmdletBinding()]
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    if ($script:SmartM365TeamsNotificationInProgress) { return }

    $normalizedLevel = if ([string]::IsNullOrWhiteSpace($Level)) { 'INFO' } else { $Level.ToUpperInvariant() }
    $isError = $normalizedLevel -eq 'ERROR'
    $isTerminalSuccess = $normalizedLevel -eq 'SUCCESS' -and $Message -match '(?i)\b(completed|finished|complete|terminated|termine)\b' -and $Message -notmatch '(?i)\bpreflight completed\b'
    if (-not $isError -and -not $isTerminalSuccess) { return }

    $dedupeKey = '{0}|{1}' -f $normalizedLevel, $Message
    if ($null -eq $script:SmartM365TeamsNotificationLogKeys) {
        $script:SmartM365TeamsNotificationLogKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }
    if (-not $script:SmartM365TeamsNotificationLogKeys.Add($dedupeKey)) { return }

    $scriptName = if ($global:SmartM365ScriptName) {
        $global:SmartM365ScriptName
    }
    elseif ($global:LogTextFile) {
        [System.IO.Path]::GetFileNameWithoutExtension($global:LogTextFile)
    }
    elseif ($PSCommandPath) {
        [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
    }
    else {
        'SmartM365'
    }

    $facts = @{
        Script     = $scriptName
        Computer   = $env:COMPUTERNAME
        LogFile    = $global:LogTextFile
        Transcript = $global:logTranscriptFile
    }
    if (-not $isError) {
        $facts['Result summary'] = $Message
        if ($global:csvGeneratedPaths) {
            $facts['Generated CSV files'] = @($global:csvGeneratedPaths).Count
        }
        if ($global:csvFilePath1) {
            $facts['Timestamped CSV'] = $global:csvFilePath1
        }
        if ($global:csvFilePath3) {
            $facts['Latest CSV'] = $global:csvFilePath3
        }
    }

    $title = if ($isError) { "SmartM365 error - $scriptName" } else { "SmartM365 completed - $scriptName" }
    $helpUrl = ''
    if ($isError) {
        $prompt = "Help troubleshoot this SmartM365 PowerShell script error. Script: $scriptName. Computer: $env:COMPUTERNAME. Error: $Message. Log: $global:LogTextFile"
        $helpUrl = 'https://chat.openai.com/?q=' + [System.Uri]::EscapeDataString($prompt)
    }

    try {
        $script:SmartM365TeamsNotificationInProgress = $true
        Send-SmartM365TeamsNotification -Title $title -Message $Message -Level $(if ($isError) { 'ERROR' } else { 'SUCCESS' }) -Facts $facts -HelpUrl $helpUrl | Out-Null
    }
    catch {
        # Notifications must not break inventory/report execution.
    }
    finally {
        $script:SmartM365TeamsNotificationInProgress = $false
    }
}

# Backward-compatible alias
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = ""
    )
    WriteLog -Message $Message -Level $Level
}

function Test-FileLocked {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $fs.Close()
        return $false
    } catch {
        return $true
    }
}

function RemoveOldFiles {
    param (
        [Parameter(Mandatory)][Alias('FolderPath')][string]$Path,
        [Alias('FilePattern')][string]$Filter = "*.log",
        [Alias('MaxFiles')][int]$KeepCount = 10,
        [string[]]$ExcludeFiles = @(),
        [string]$LogFile
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }
    if ([string]::IsNullOrWhiteSpace($LogFile) -and -not [string]::IsNullOrWhiteSpace([string]$global:LogTextFile)) {
        $LogFile = [string]$global:LogTextFile
    }
    $writeCleanupLog = {
        param([Parameter(Mandatory)][string]$Message)

        if ([string]::IsNullOrWhiteSpace($LogFile)) { return }
        try {
            Add-Content -LiteralPath $LogFile -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -ErrorAction Stop
        }
        catch {
            Write-Verbose ("Cleanup log write skipped for '{0}': {1}" -f $LogFile, $_.Exception.Message)
        }
    }


    # Normalize exclusion list to full paths, case-insensitive
    $excludeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($ex in $ExcludeFiles) {
        try {
            $resolved = Resolve-Path -LiteralPath $ex -ErrorAction Stop
            [void]$excludeSet.Add($resolved.ProviderPath)
        } catch {
            [void]$excludeSet.Add($ex)
        }
    }

    # Add globally generated CSV files
    if ($global:csvGeneratedPaths) {
        foreach ($csvPath in $global:csvGeneratedPaths) {
            [void]$excludeSet.Add($csvPath)
        }
    }

    # Add global log files
    foreach ($varName in @('logTranscriptFile','logTextFile')) {
        $v = Get-Variable -Name $varName -Scope Global -ErrorAction SilentlyContinue
        if ($v -and $v.Value) {
            try {
                $resolved = Resolve-Path -LiteralPath $v.Value -ErrorAction Stop
                [void]$excludeSet.Add($resolved.ProviderPath)
            } catch {
                [void]$excludeSet.Add($v.Value)
            }
        }
    }

    $files = Get-ChildItem -LiteralPath $Path -Filter $Filter -File |
             Sort-Object LastWriteTime -Descending

    $filesToDelete = $files | Where-Object { -not $excludeSet.Contains($_.FullName) } |
                              Select-Object -Skip $KeepCount

    foreach ($file in $filesToDelete) {
        if (Test-FileLocked -Path $file.FullName) {
            & $writeCleanupLog "[SKIP] Locked file: $($file.Name)"
            continue
        }

        try {
            if ($file.Attributes -band [IO.FileAttributes]::ReadOnly) {
                $file.Attributes = $file.Attributes -bxor [IO.FileAttributes]::ReadOnly
            }
        } catch {
            & $writeCleanupLog "[SKIP] Cannot clear ReadOnly on $($file.Name): $($_.Exception.Message)"
            continue
        }

        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            & $writeCleanupLog "Deleted old file: $($file.Name)"
        } catch {
            & $writeCleanupLog "[SKIP] Failed to delete $($file.Name): $($_.Exception.Message)"
        }
    }
}

# Backward compatible wrapper for old scripts
function Remove-OldFiles {
    param (
        [string]$Path,
        [string]$Filter = "*.log",
        [int]$KeepCount = 10,
        [string[]]$ExcludeFiles = @(),
        [string]$LogFile
    )

    if (Get-Command -Name RemoveOldFiles -ErrorAction SilentlyContinue) {
        RemoveOldFiles @PSBoundParameters
    } else {
        try {
            $files = Get-ChildItem -Path $Path -Filter $Filter | Sort-Object LastWriteTime -Descending
            $filesToDelete = $files | Where-Object { $ExcludeFiles -notcontains $_.FullName } | Select-Object -Skip $KeepCount

            foreach ($file in $filesToDelete) {
                Remove-Item $file.FullName -Force
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Deleted old file: $($file.Name)" | Tee-Object -FilePath $LogFile -Append
            }
        }
        catch {
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error during cleanup: $($_.Exception.Message)" | Tee-Object -FilePath $LogFile -Append
        }
    }
}

function Set-SmartM365CoreContext {
    [CmdletBinding()]
    param(
        [string]$RunId,
        [string]$RunOutputRoot,
        [string]$LatestOutputRoot,
        [string]$LogPath,
        [string]$TenantKey = $global:SmartM365TenantKey,
        [string]$OrganizationKey = $global:SmartM365OrganizationKey,
        [string]$EnvironmentKey = $global:SmartM365EnvironmentKey,
        [string]$TenantId = $global:SmartM365TenantId,
        [int]$RetentionMaxCsv = 30,
        [int]$RetentionMaxLogs = 30
    )

    $script:SmartM365CoreRunId = $RunId
    $script:SmartM365CoreRunOutputRoot = $RunOutputRoot
    $script:SmartM365CoreLatestOutputRoot = $LatestOutputRoot
    $script:SmartM365CoreLogPath = $LogPath
    $script:SmartM365CoreTenantKey = $TenantKey
    $script:SmartM365CoreOrganizationKey = $OrganizationKey
    $script:SmartM365CoreEnvironmentKey = $EnvironmentKey
    $script:SmartM365CoreTenantId = $TenantId
    $script:SmartM365CoreRetentionMaxCsv = $RetentionMaxCsv
    $script:SmartM365CoreRetentionMaxLogs = $RetentionMaxLogs

    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        $global:LogTextFile = $LogPath
    }
    if ($RetentionMaxCsv -gt 0) {
        $global:RetentionMaxCSV = $RetentionMaxCsv
    }
    if ($RetentionMaxLogs -gt 0) {
        $global:RetentionMaxLogs = $RetentionMaxLogs
    }
}

function Get-SmartM365CoreContextValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    $variableName = "SmartM365Core$Name"
    $variable = Get-Variable -Name $variableName -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $variable -and $null -ne $variable.Value) {
        return $variable.Value
    }

    return $DefaultValue
}

function Get-SmartM365CsvValidationBaseName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    $baseName = ($name -replace '_\d{8}[-_]\d{6}$', '')
    $baseName = ($baseName -replace '_MAXITEMS-\d+$', '')
    return $baseName
}

function Get-SmartM365CsvValidationRule {
    [CmdletBinding()]
    param(
        [string]$BaseFileName,
        [string]$TimestampedPath,
        [string]$LatestPath
    )

    $rules = $global:SmartM365CsvValidationRules
    if ($null -eq $rules -or -not ($rules -is [hashtable])) { return $null }

    $candidateKeys = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @($BaseFileName, $LatestPath, $TimestampedPath)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        [void]$candidateKeys.Add($candidate)
        [void]$candidateKeys.Add((Split-Path -Path $candidate -Leaf))
        $base = Get-SmartM365CsvValidationBaseName -Path $candidate
        [void]$candidateKeys.Add($base)
        if (-not [string]::IsNullOrWhiteSpace($base)) {
            [void]$candidateKeys.Add(($base + '.csv'))
            foreach ($normalizedBase in @(
                ($base -replace '_PARTIAL$', ''),
                ($base -replace '_top100$', ''),
                (($base -replace '_top100$', '') -replace '_PARTIAL$', '')
            )) {
                if (-not [string]::IsNullOrWhiteSpace($normalizedBase)) {
                    [void]$candidateKeys.Add($normalizedBase)
                    [void]$candidateKeys.Add(($normalizedBase + '.csv'))
                }
            }
        }
    }

    foreach ($key in @($candidateKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        if ($rules.ContainsKey($key)) { return $rules[$key] }
    }

    return $null
}

function Assert-SmartM365CsvDataCompleteness {
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Data,
        [string[]]$Columns = @(),
        [string]$BaseFileName,
        [string]$TimestampedPath,
        [string]$LatestPath
    )

    $rule = Get-SmartM365CsvValidationRule -BaseFileName $BaseFileName -TimestampedPath $TimestampedPath -LatestPath $LatestPath
    if ($null -eq $rule) {
        if ($global:SmartM365RequireCsvValidationRules) {
            $name = if (-not [string]::IsNullOrWhiteSpace($BaseFileName)) { $BaseFileName } elseif (-not [string]::IsNullOrWhiteSpace($LatestPath)) { Get-SmartM365CsvValidationBaseName -Path $LatestPath } else { Get-SmartM365CsvValidationBaseName -Path $TimestampedPath }
            throw ("CSV validation rule is missing for '{0}'. DATA-LAST publication is blocked." -f $name)
        }
        return
    }

    $rows = @($Data)
    $displayName = if ($rule.ContainsKey('Name') -and -not [string]::IsNullOrWhiteSpace([string]$rule.Name)) { [string]$rule.Name } elseif (-not [string]::IsNullOrWhiteSpace($BaseFileName)) { $BaseFileName } elseif (-not [string]::IsNullOrWhiteSpace($LatestPath)) { Get-SmartM365CsvValidationBaseName -Path $LatestPath } else { Get-SmartM365CsvValidationBaseName -Path $TimestampedPath }
    $allowEmptyDataset = ($rule.ContainsKey('AllowEmptyDataset') -and [bool]$rule.AllowEmptyDataset)
    $criticalFields = @($rule.CriticalFields | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $requiredColumns = @($rule.RequiredColumns | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $expectedRowCount = $null
    if ($rule.ContainsKey('ExpectedRowCount') -and $null -ne $rule.ExpectedRowCount -and [string]$rule.ExpectedRowCount -ne '') {
        $expectedRowCount = [int]$rule.ExpectedRowCount
    }
    $criticalMissingFailMinRows = 50
    if ($rule.ContainsKey('CriticalMissingFailMinRows') -and $null -ne $rule.CriticalMissingFailMinRows -and [string]$rule.CriticalMissingFailMinRows -ne '') {
        $criticalMissingFailMinRows = [int]$rule.CriticalMissingFailMinRows
    }
    $criticalMissingFailPercent = 0.5
    if ($rule.ContainsKey('CriticalMissingFailPercent') -and $null -ne $rule.CriticalMissingFailPercent -and [string]$rule.CriticalMissingFailPercent -ne '') {
        $criticalMissingFailPercent = [double]$rule.CriticalMissingFailPercent
    }

    if ($rows.Count -eq 0 -and -not $allowEmptyDataset) {
        throw ("CSV '{0}' has 0 rows while this dataset is not declared as empty-capable. DATA-LAST publication is blocked." -f $displayName)
    }

    if ($null -ne $expectedRowCount -and $rows.Count -ne $expectedRowCount) {
        throw ("CSV '{0}' row count mismatch. Expected {1}, got {2}. DATA-LAST publication is blocked." -f $displayName, $expectedRowCount, $rows.Count)
    }

    $availableColumns = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($column in @($Columns)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$column)) { [void]$availableColumns.Add([string]$column) }
    }
    foreach ($row in @($rows | Select-Object -First 1)) {
        foreach ($property in @($row.PSObject.Properties.Name)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$property)) { [void]$availableColumns.Add([string]$property) }
        }
    }

    foreach ($column in @($requiredColumns + $criticalFields | Select-Object -Unique)) {
        if (-not $availableColumns.Contains([string]$column)) {
            throw ("CSV '{0}' is missing required column '{1}'. DATA-LAST publication is blocked." -f $displayName, $column)
        }
    }

    $missingExamples = New-Object System.Collections.Generic.List[string]
    $missingCriticalRows = 0
    $rowNumber = 1
    $thresholdReached = $false
    foreach ($row in $rows) {
        $rowHasMissingCriticalField = $false
        foreach ($field in $criticalFields) {
            $property = $row.PSObject.Properties[$field]
            $value = if ($null -ne $property) { $property.Value } else { $null }
            if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
                $rowHasMissingCriticalField = $true
                if ($missingExamples.Count -lt 10) { [void]$missingExamples.Add(("row {0}, field {1}" -f $rowNumber, $field)) }
            }
        }
        if ($rowHasMissingCriticalField) {
            $missingCriticalRows++
            $currentPercent = if ($rows.Count -gt 0) { [math]::Round((100.0 * $missingCriticalRows / $rows.Count), 4) } else { 0 }
            $rowThresholdHit = ($criticalMissingFailMinRows -gt 0 -and $missingCriticalRows -ge $criticalMissingFailMinRows)
            $percentThresholdHit = ($criticalMissingFailPercent -gt 0 -and $currentPercent -ge $criticalMissingFailPercent)
            if ($rowThresholdHit -or $percentThresholdHit) { $thresholdReached = $true; break }
        }
        $rowNumber++
    }

    if ($missingCriticalRows -gt 0) {
        $missingPercent = if ($rows.Count -gt 0) { [math]::Round((100.0 * $missingCriticalRows / $rows.Count), 4) } else { 0 }
        $exampleText = ($missingExamples -join '; ')
        $thresholdText = ("fail thresholds: rows>={0}, percent>={1}%" -f $criticalMissingFailMinRows, $criticalMissingFailPercent)
        if ($thresholdReached) {
            throw ("CSV '{0}' has empty critical business fields in at least {1}/{2} row(s), {3}%. {4}. Examples: {5}. DATA-LAST publication is blocked." -f $displayName, $missingCriticalRows, $rows.Count, $missingPercent, $thresholdText, $exampleText)
        }
        WriteLog -Message ("CSV validation warning for '{0}': empty critical business fields in {1}/{2} row(s), {3}%. {4}. Examples: {5}. DATA-LAST publication continues." -f $displayName, $missingCriticalRows, $rows.Count, $missingPercent, $thresholdText, $exampleText) -Level 'WARNING'
        return
    }

    WriteLog -Message ("CSV validation passed for '{0}'. Rows: {1}; CriticalFields: {2}" -f $displayName, $rows.Count, ($criticalFields -join ', '))
}
function Add-SmartM365CsvValidationRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Rules,
        [Parameter(Mandatory)][string]$BaseFileName,
        [Parameter(Mandatory)][string[]]$CriticalFields,
        [string[]]$RequiredColumns = @(),
        [switch]$AllowEmptyDataset,
        [string]$Name = '',
        [int]$CriticalMissingFailMinRows = 50,
        [double]$CriticalMissingFailPercent = 0.5
    )

    $ruleName = if ([string]::IsNullOrWhiteSpace($Name)) { $BaseFileName } else { $Name }
    $Rules[$BaseFileName] = @{
        Name                       = $ruleName
        CriticalFields             = @($CriticalFields)
        RequiredColumns            = @($RequiredColumns)
        AllowEmptyDataset          = [bool]$AllowEmptyDataset
        CriticalMissingFailMinRows = [int]$CriticalMissingFailMinRows
        CriticalMissingFailPercent = [double]$CriticalMissingFailPercent
    }
}
function Initialize-SmartM365DefaultCsvValidationRules {
    [CmdletBinding()]
    param()

    if ($global:SmartM365CsvValidationRules -isnot [hashtable]) {
        $global:SmartM365CsvValidationRules = @{}
    }

    $rules = $global:SmartM365CsvValidationRules
    $add = { param($Base,$Fields,$AllowEmpty) Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName $Base -CriticalFields $Fields -AllowEmptyDataset:$AllowEmpty }

    & $add 'AD_HealthCheck' @('Forest','Domain','Category','Check','Status') $false
    & $add 'AD_Inventory_DailySummary' @('SnapshotDate','GeneratedAt') $false
    & $add 'AD_Users_AllDomains' @('DomainName','SamAccountName','DistinguishedName','ObjectGUID') $false
    & $add 'AD_Computers_AllDomains' @('DomainName','SamAccountName','DistinguishedName','ObjectGUID') $false
    & $add 'AD_Groups_AllDomains' @('DomainName','SamAccountName','DistinguishedName','ObjectGUID') $false
    & $add 'AD_OUs_AllDomains' @('DomainName','Name','DistinguishedName') $false
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'AD_Contacts_AllDomains' -CriticalFields @('DomainName','ObjectGUID') -RequiredColumns @('ObjectType','Name','DistinguishedName','DisplayName','ProxyAddresses','Mail') -AllowEmptyDataset
    & $add 'AD_Users_DailyStats' @('Date','DomainName') $false
    & $add 'AD_Computers_DailyStats' @('Date','DomainName') $false
    & $add 'AD_Users_DuplicateUPN' @('UserPrincipalName','SamAccountName','DistinguishedName') $true
    & $add 'AD_Users_DuplicateSMTP' @('SmtpAddress','UserPrincipalName','SamAccountName','DistinguishedName') $true

    & $add 'Exchange_EXO_AcceptedDomains' @('Name','DomainName','DomainType') $false
    & $add 'Exchange_EXO_Mailboxes_AllDomains' @('UserPrincipalName','PrimarySmtpAddress','MailboxGuid') $false
    & $add 'Exchange_EXO_Mailboxes_AllDomains_Archive' @('UserPrincipalName','PrimarySmtpAddress') $false
    & $add 'Exchange_EXO_Mailboxes_AllDomains_Stats' @('UserPrincipalName','PrimarySmtpAddress','TotalItemSizeGB','ItemCount') $false
    & $add 'Exchange_EXO_Mailboxes_AllDomains_Permissions' @('UserPrincipalName','PrimarySmtpAddress') $false
    & $add 'Exchange_EXO_MailboxCalendarPermissions_AllDomains' @('Mailbox','User','AccessRights') $false
    & $add 'Exchange_EXO_MailboxCalendarPermissions_Errors' @('Index','Message') $true
    & $add 'Exchange_EXO_MigrationJobs' @('BatchName','BatchStatus') $true
    & $add 'Exchange_Mailboxes_AllSources_PermissionsByUser' @('ResolvedUser','PrimarySMTPaddress','Source') $true
    & $add 'M365_Backup_ProtectedMailboxes' @('ProtectionUnitId','Status') $true

    & $add 'Exchange_OnPrem_Mailboxes_AllDomains' @('DomainName','SamAccountName','PrimarySMTPaddress','ObjectGUID') $false
    & $add 'Exchange_OnPrem_Mailboxes_AllDomains_OnlyADPermission' @('DomainName','SamAccountName','PrimarySMTPaddress','ObjectGUID') $false
    & $add 'Exchange_OnPrem_RemoteMailboxes_AllDomains' @('DomainName','SamAccountName','PrimarySmtpAddress','ObjectGuid') $false
    & $add 'Exchange_OnPrem_Mailboxes_DailyStats' @('Date','DomainName') $false
    & $add 'Exchange_OnPrem_Mailboxes_DailyStats_Summary' @('DomainName','TotalMailboxCount') $false
    & $add 'Exchange_OnPrem_ProxyAddresses_Check' @('Identity','PrimaryAddress','ExpectedAddress','Status') $false
    & $add 'Exchange_OnPrem_ProxyAddresses_Summary' @('Summary','Count') $false
    & $add 'Exchange_OnPrem_Servers_Inventory' @('Name','Fqdn','ServerRole') $false
    & $add 'Exchange_OnPrem_Servers_Inventory_Summary' @('ScriptName','RunId','ExchangeServersCount') $false
    & $add 'Exchange_OnPrem_Servers_RemoteAccess' @('ExchangeServerName','ComputerNameTried') $false
    & $add 'Exchange_OnPrem_Servers_Compute' @('ExchangeServerName','ComputerNameTried','CollectionStatus') $false
    & $add 'Exchange_OnPrem_Servers_LogicalDisks' @('ExchangeServerName','ComputerNameTried','DeviceId','CollectionStatus') $false
    & $add 'Exchange_OnPrem_Servers_DiskDrives' @('ExchangeServerName','ComputerNameTried','Index','CollectionStatus') $false
    & $add 'Exchange_OnPrem_Infrastructure_PerServerSummary' @('ExchangeServerName','ServerRole') $false
    & $add 'Exchange_OnPrem_MigrationReadiness_Config' @('Category','ObjectName','Setting','CollectionStatus') $false

    & $add 'M365_Users_Active' @('User principal name','Object Id','AccountEnabled') $false
    & $add 'M365_Users_Activity' @('RunId','UserPrincipalName','ReportRefreshDate') $false
    & $add 'M365_Mailbox_Usage' @('RunId','User Principal Name','Report Refresh Date') $false
    & $add 'M365_OneDrive_Usage' @('RunId','Site Id','Report Refresh Date') $false
    & $add 'M365_SharePoint_SiteUsage' @('RunId','Site Id','Report Refresh Date') $false
    & $add 'M365_Apps_Activations' @('RunId','User Principal Name','Report Refresh Date') $false
    & $add 'M365_Teams_UserActivity' @('RunId','User Principal Name','Report Refresh Date') $false
    & $add 'M365_Email_Activity' @('RunId','User Principal Name','Report Refresh Date') $false
    & $add 'M365_Entra_VerifiedDomains' @('Id','IsVerified') $false
    & $add 'M365_Entra_AzureADConnect_SyncHealth' @('CheckName','Status','ExportDateTime','RunId') $false

    & $add 'M365_Licenses_Users' @('Id','User principal name','UserId','SkuId','SkuPartNumber') $false
    & $add 'M365_Licenses_ServicePlans' @('Id','User principal name','SkuId','PlanId','PlanName') $false
    & $add 'M365_Licenses_ServicePlans_Detailed' @('Id','User principal name','SkuId','PlanId','PlanName') $false
    & $add 'M365_Licenses_Tenant' @('Id','TenantSkuPartNumber') $false
    & $add 'M365_Licenses_Groups' @('Id','GroupId','GroupDisplayName') $true

    & $add 'M365_Entra_Devices' @('ObjectId','DeviceId','DisplayName') $false
    & $add 'M365_Entra_Devices_RegisteredPending' @('ObjectId','DeviceId','DisplayName') $true
    & $add 'M365_Entra_Devices_HardwareIdConflicts' @('HardwareId','DeviceCount') $true
    & $add 'M365_Entra_Devices_HardwareIdConflicts_RegisteredPending' @('HardwareId','ObjectId','DeviceId') $true
    & $add 'M365_Entra_Devices_RemovalCandidates' @('HardwareId','CandidateObjectId','PrimaryObjectId','Reason') $true

    & $add 'Intune_Devices_Inventory' @('Device ID','Device name','Azure AD Device ID') $false
    & $add 'Intune_Devices_Compliance' @('DeviceName') $false
    & $add 'Intune_Devices_BIOS' @('ManagedDeviceId','DeviceName','AzureADDeviceId') $false
    & $add 'Intune_Devices_LocalSystem' @('DeviceName','DeviceId','AzureADDeviceId') $false
    & $add 'Intune_Autopilot_Devices' @('Autopilot ID','Serial number') $false
    & $add 'Intune_Devices_UpgradeEligibility_Summary' @('ReportType','ExportDateTime') $false
    & $add 'Intune_Devices_Win11Readiness' @('DeviceName','GraphId','RunId') $false
    & $add 'Intune_WindowsUpdate_Status' @('PolicyId','DeviceId','DeviceName','AggregateState') $false
    & $add 'Intune_AutopatchAlerts_Detail' @('AlertName','Severity','SourceReport') $true
    & $add 'Intune_AutopatchAlerts_Summary' @('AlertName','Severity','SourceReport') $true
    & $add 'Intune_AutopatchAlerts_PolicySummary' @('PolicyId','PolicyName') $true
    & $add 'Intune_DiscoveredApps_Summary' @('AppId','AppName','Platform') $false
    & $add 'Intune_DiscoveredApps_DeviceDetail' @('AppId','AppName','DeviceId','DeviceName') $true
    & $add 'Intune_RBAC_GroupMembers' @('Country','IntuneRole','GroupName','GroupFound') $false

    & $add 'M365_SPO_Sites' @('RunId','SiteUrl','Status') $false
    & $add 'M365_SPO_Lists' @('RunId','SiteUrl','ListId','Status') $true
    & $add 'M365_SPO_Permissions' @('RunId','SiteUrl','Status') $true
    & $add 'M365_SPO_ExternalSharing' @('RunId','SiteUrl','SharingSignal','Status') $true
    & $add 'M365_Teams_Teams' @('RunId','TeamId','TeamDisplayName','Status') $true
    & $add 'M365_Teams_Members' @('RunId','TeamId','UserId','Role','Status') $true
    & $add 'M365_Teams_Guests' @('RunId','TeamId','UserId','Status') $true
    & $add 'M365_Teams_Channels' @('RunId','TeamId','ChannelId','ChannelDisplayName','Status') $true

    WriteLog -Message ("SmartM365 CSV validation rules loaded: {0}" -f $rules.Count) -Level 'INFO'
}
function Add-SmartM365TenantKeyToCsvData {
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Data,
        [string[]]$Columns = @(),
        [string]$TenantKey = [string](Get-SmartM365CoreContextValue -Name 'TenantKey' -DefaultValue $global:SmartM365TenantKey),
        [string]$OrganizationKey = [string](Get-SmartM365CoreContextValue -Name 'OrganizationKey' -DefaultValue $global:SmartM365OrganizationKey),
        [string]$EnvironmentKey = [string](Get-SmartM365CoreContextValue -Name 'EnvironmentKey' -DefaultValue $global:SmartM365EnvironmentKey),
        [string]$TenantId = [string](Get-SmartM365CoreContextValue -Name 'TenantId' -DefaultValue $global:SmartM365TenantId)
    )

    $identityValues = [ordered]@{
        TenantKey      = $TenantKey
        OrganizationKey = $OrganizationKey
        EnvironmentKey = $EnvironmentKey
        TenantId        = $TenantId
    }
    foreach ($identityName in @($identityValues.Keys)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$identityValues[$identityName])) { continue }
        foreach ($row in @($Data)) {
            if ($null -eq $row) { continue }
            $property = $row.PSObject.Properties[$identityName]
            if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                $identityValues[$identityName] = [string]$property.Value
                break
            }
        }
    }

    foreach ($requiredIdentityName in @('TenantKey', 'OrganizationKey', 'EnvironmentKey')) {
        if ([string]::IsNullOrWhiteSpace([string]$identityValues[$requiredIdentityName])) {
            throw "$requiredIdentityName is required for SmartM365 CSV exports. Initialize the tenant context or use -NoTenantKey for an intentionally identity-neutral export."
        }
    }

    $identityColumns = @('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
    $tenantData = @(
        foreach ($row in @($Data)) {
            if ($null -eq $row) { continue }
            $values = [ordered]@{}
            foreach ($identityName in $identityColumns) { $values[$identityName] = [string]$identityValues[$identityName] }
            if ($row -is [System.Collections.IDictionary]) {
                foreach ($key in $row.Keys) {
                    $name = [string]$key
                    if ($identityColumns -icontains $name -or $values.Contains($name)) { continue }
                    $values[$name] = $row[$key]
                }
            }
            else {
                foreach ($property in $row.PSObject.Properties) {
                    if ($identityColumns -icontains $property.Name -or $values.Contains($property.Name)) { continue }
                    $values[$property.Name] = $property.Value
                }
            }
            [pscustomobject]$values
        }
    )

    $tenantColumns = if (@($Columns).Count -gt 0 -or @($tenantData).Count -eq 0) {
        @($identityColumns) + @($Columns | Where-Object { $identityColumns -inotcontains $_ })
    }
    else { @() }

    return [pscustomobject]@{
        Data            = $tenantData
        Columns         = $tenantColumns
        TenantKey       = [string]$identityValues.TenantKey
        OrganizationKey = [string]$identityValues.OrganizationKey
        EnvironmentKey  = [string]$identityValues.EnvironmentKey
        TenantId        = [string]$identityValues.TenantId
        WasApplied      = $true
    }
}
function Add-SmartM365TenantKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)][AllowNull()]$InputObject,
        [string]$TenantKey = [string](Get-SmartM365CoreContextValue -Name 'TenantKey' -DefaultValue $global:SmartM365TenantKey),
        [string]$OrganizationKey = [string](Get-SmartM365CoreContextValue -Name 'OrganizationKey' -DefaultValue $global:SmartM365OrganizationKey),
        [string]$EnvironmentKey = [string](Get-SmartM365CoreContextValue -Name 'EnvironmentKey' -DefaultValue $global:SmartM365EnvironmentKey),
        [string]$TenantId = [string](Get-SmartM365CoreContextValue -Name 'TenantId' -DefaultValue $global:SmartM365TenantId)
    )

    process {
        if ($null -eq $InputObject) { return }
        $tenantCsv = Add-SmartM365TenantKeyToCsvData -Data @($InputObject) -TenantKey $TenantKey -OrganizationKey $OrganizationKey -EnvironmentKey $EnvironmentKey -TenantId $TenantId
        foreach ($row in @($tenantCsv.Data)) { Write-Output $row }
    }
}
function Repair-SmartM365CsvTenantKeySchema {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [char]$Delimiter = ',',
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = 'UTF8'
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $header = [string](Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop)
    if ([string]::IsNullOrWhiteSpace($header)) { return $false }
    $d = [regex]::Escape([string]$Delimiter)
    $identityHeaderPattern = '^\s*"?TenantKey"?\s*' + $d + '\s*"?OrganizationKey"?\s*' + $d + '\s*"?EnvironmentKey"?\s*' + $d + '\s*"?TenantId"?(\s*' + $d + '|\s*$)'
    if ($header -match $identityHeaderPattern) { return $false }

    foreach ($requiredIdentityName in @('TenantKey', 'OrganizationKey', 'EnvironmentKey')) {
        $value = [string](Get-SmartM365CoreContextValue -Name $requiredIdentityName -DefaultValue (Get-Variable -Name ("SmartM365{0}" -f $requiredIdentityName) -Scope Global -ValueOnly -ErrorAction SilentlyContinue))
        if ([string]::IsNullOrWhiteSpace($value)) { throw "Cannot migrate CSV '$Path': $requiredIdentityName is missing from the active context." }
    }

    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $Delimiter)
    if ($rows.Count -gt 0) {
        $columns = @($rows[0].PSObject.Properties.Name)
        Write-SmartM365CsvAtomically -Data $rows -Path $Path -Columns $columns -Encoding $Encoding -Delimiter ([string]$Delimiter)
    }
    else {
        $parent = Split-Path -Path $Path -Parent
        if ([string]::IsNullOrWhiteSpace($parent)) { $parent = (Get-Location).Path }
        $tempPath = Join-Path -Path $parent -ChildPath ((Split-Path -Path $Path -Leaf) + '.tenantkey.' + [guid]::NewGuid().ToString('N') + '.tmp')
        try {
            $separatorCount = ([regex]::Matches($header, $d)).Count
            $fakeRow = (@('x') * ($separatorCount + 1)) -join ([string]$Delimiter)
            $probe = @(@($header, $fakeRow) | ConvertFrom-Csv -Delimiter $Delimiter)
            $existingColumns = if ($probe.Count -gt 0) { @($probe[0].PSObject.Properties.Name) } else { @($header.Trim('"')) }
            $identityColumns = @('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
            $newColumns = @($identityColumns) + @($existingColumns | Where-Object { $identityColumns -inotcontains $_ })
            $newHeader = ($newColumns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ([string]$Delimiter)
            Set-Content -LiteralPath $tempPath -Value $newHeader -Encoding $Encoding -ErrorAction Stop
            Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
        }
        finally { if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue } }
    }

    Write-Verbose ("Migrated CSV to SmartM365 identity-first schema: {0}" -f $Path)
    return $true
}
function Write-SmartM365CsvAtomically {
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Data,
        [Parameter(Mandatory)][string]$Path,
        [string[]]$Columns = @(),
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",
        [string]$Delimiter = ",",
        [switch]$NoTypeInformation = $true,
        [switch]$NoTenantKey
    )

    if (-not $NoTenantKey) {
        $tenantCsv = Add-SmartM365TenantKeyToCsvData -Data $Data -Columns $Columns
        $Data = @($tenantCsv.Data)
        $Columns = @($tenantCsv.Columns)
    }

    Assert-SmartM365CsvDataCompleteness -Data $Data -Columns $Columns -TimestampedPath $Path -LatestPath $Path

    $parent = Split-Path -Path $Path -Parent
    if ([string]::IsNullOrWhiteSpace($parent)) { $parent = (Get-Location).Path }
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    $extension = [System.IO.Path]::GetExtension($Path)
    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    if ([string]::IsNullOrWhiteSpace($extension)) { $extension = '.tmp' }
    $tempPath = Join-Path -Path $parent -ChildPath ("{0}.{1}{2}" -f $nameWithoutExtension, [guid]::NewGuid().ToString('N'), $extension)

    try {
        $rows = @($Data)
        if ($rows.Count -eq 0 -and $Columns.Count -gt 0) {
            $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join $Delimiter
            Set-Content -LiteralPath $tempPath -Value $header -Encoding $Encoding -ErrorAction Stop
        }
        elseif ($Columns.Count -gt 0) {
            $exportParams = @{
                Path      = $tempPath
                Encoding  = $Encoding
                Delimiter = $Delimiter
            }
            if ($NoTypeInformation) { $exportParams.NoTypeInformation = $true }
            $rows | Select-Object -Property $Columns | Export-Csv @exportParams
        }
        else {
            $exportParams = @{
                Path      = $tempPath
                Encoding  = $Encoding
                Delimiter = $Delimiter
            }
            if ($NoTypeInformation) { $exportParams.NoTypeInformation = $true }
            $rows | Export-Csv @exportParams
        }

        Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Publish-SmartM365Csv {
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Data,
        [Parameter(Mandatory)][string]$TimestampedPath,
        [string]$LatestPath,
        [string[]]$Columns = @(),
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",
        [string]$Delimiter = ",",
        [int]$RetentionMaxCsv = -1,
        [switch]$NoSharePointUpload,
        [switch]$NoTenantKey
    )

    $Data = @(Limit-SmartM365RowsForMaxItems -Data $Data)
    if (-not $NoTenantKey) {
        $tenantCsv = Add-SmartM365TenantKeyToCsvData -Data $Data -Columns $Columns
        $Data = @($tenantCsv.Data)
        $Columns = @($tenantCsv.Columns)
    }
    if (Test-SmartM365MaxItemsMode) {
        $TimestampedPath = Add-SmartM365MaxItemsSuffixToCsvPath -Path $TimestampedPath
        if (-not [string]::IsNullOrWhiteSpace($LatestPath)) { $LatestPath = Add-SmartM365MaxItemsSuffixToCsvPath -Path $LatestPath }
        WriteLog -Message ("MaxItems test CSV publication active. CSV paths are suffixed with {0}; standard Power BI filenames are not updated." -f (Get-SmartM365MaxItemsSuffix)) -Level 'WARNING'
    }

    Assert-SmartM365CsvDataCompleteness -Data $Data -Columns $Columns -TimestampedPath $TimestampedPath -LatestPath $LatestPath

    Write-SmartM365CsvAtomically -Data $Data -Path $TimestampedPath -Columns $Columns -Encoding $Encoding -Delimiter $Delimiter -NoTenantKey:$NoTenantKey
    WriteLog -Message ("CSV exported to: {0}" -f $TimestampedPath)

    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }
    [void]$global:csvGeneratedPaths.Add($TimestampedPath)

    $publishedPath = $TimestampedPath
    if (-not [string]::IsNullOrWhiteSpace($LatestPath)) {
        Write-SmartM365CsvAtomically -Data $Data -Path $LatestPath -Columns $Columns -Encoding $Encoding -Delimiter $Delimiter -NoTenantKey:$NoTenantKey
        WriteLog -Message ("CSV latest copy written to: {0}" -f $LatestPath)
        [void]$global:csvGeneratedPaths.Add($LatestPath)
        $publishedPath = $LatestPath
    }

    if ($RetentionMaxCsv -lt 0) {
        $RetentionMaxCsv = [int](Get-SmartM365CoreContextValue -Name 'RetentionMaxCsv' -DefaultValue $global:RetentionMaxCSV)
    }
    if ($RetentionMaxCsv -gt 0) {
        $timestampedFolder = Split-Path -Path $TimestampedPath -Parent
        $timestampedName = [System.IO.Path]::GetFileNameWithoutExtension($TimestampedPath)
        $retentionPrefix = $timestampedName -replace '_\d{8}[-_]\d{6}$', ''
        if (-not [string]::IsNullOrWhiteSpace($timestampedFolder) -and -not [string]::IsNullOrWhiteSpace($retentionPrefix)) {
            RemoveOldFiles -Path $timestampedFolder -Filter "$retentionPrefix*.csv" -KeepCount $RetentionMaxCsv -LogFile $global:LogTextFile
        }
    }

    $sharePointUploads = @()
    if (-not $NoSharePointUpload) {
        $uploadCandidates = @($TimestampedPath)
        if (-not [string]::IsNullOrWhiteSpace($LatestPath)) { $uploadCandidates += $LatestPath }
        foreach ($uploadCandidate in @($uploadCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)) {
            if (-not (Test-Path -LiteralPath $uploadCandidate)) { continue }
            $uploadRecord = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $uploadCandidate
            if ($uploadRecord) { $sharePointUploads += $uploadRecord }
        }
    }

    if (Test-SmartM365MaxItemsMode) {
        WriteLog -Message 'WeeklyHistory publication skipped because MaxItems test mode is active.' -Level 'WARNING'
    }
    else {
        Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($publishedPath) -TimestampedPath $TimestampedPath
    }

    return [pscustomobject]@{
        TimestampedPath = $TimestampedPath
        LatestPath      = $LatestPath
        PublishedPath   = $publishedPath
        SharePointUploads = $sharePointUploads
    }
}

function Export-SmartM365Csv {
    [CmdletBinding(DefaultParameterSetName = 'ByBaseName')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'ByBaseName')]
        [string]$BaseFileName,

        [Parameter(Mandatory, ParameterSetName = 'ByBaseName')]
        [string]$OutputPath,

        [Parameter(Mandatory, ParameterSetName = 'ByBaseName')]
        [string]$GlobalPath,

        [Parameter(Mandatory, ParameterSetName = 'ByPath')]
        [string]$TimestampedPath,

        [Parameter(ParameterSetName = 'ByPath')]
        [string]$LatestPath,

        [Parameter(Mandatory)]
        [AllowNull()][object[]]$Data,

        [string[]]$Columns = @(),
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",
        [string]$Delimiter = ",",
        [switch]$NoTypeInformation = $true,
        [switch]$NoSharePointUpload,
        [switch]$NoTenantKey
    )

    if ($PSCmdlet.ParameterSetName -eq 'ByBaseName') {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $TimestampedPath = Join-Path $OutputPath "$BaseFileName`_$timestamp.csv"
        $LatestPath = Join-Path $GlobalPath "$BaseFileName.csv"
    }

    Publish-SmartM365Csv -Data $Data -TimestampedPath $TimestampedPath -LatestPath $LatestPath -Columns $Columns -Encoding $Encoding -Delimiter $Delimiter -NoSharePointUpload:$NoSharePointUpload -NoTenantKey:$NoTenantKey
}

function Export-SmartM365CsvFromConvert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BaseFileName,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$GlobalPath,
        [Parameter(Mandatory)][array]$Data,
        [switch]$NoTypeInformation = $true,
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",
        [string]$Delimiter = ",",
        [switch]$NoTenantKey
    )

    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName -OutputPath $OutputPath -GlobalPath $GlobalPath -Data $Data -NoTypeInformation:$NoTypeInformation -Encoding $Encoding -Delimiter $Delimiter -NoTenantKey:$NoTenantKey
}

function EnsureExchangePSSnapinLoaded {
    [CmdletBinding()]
    param (
        [string]$SnapinName = "Microsoft.Exchange.Management.PowerShell.SnapIn",
        [string[]]$RequiredCommands = @("Get-Mailbox"),
        [switch]$ViewEntireForest
    )

    # Check if the PSSnapin is registered on the server
    if (-not (Get-PSSnapin $SnapinName -Registered -ErrorAction SilentlyContinue)) {
        Write-Error "The Exchange Management PSSnapin '$SnapinName' is not registered on this server."
        Write-Error "This script must be run on an Exchange 2016 server where the Management Tools are installed."
        return $false
    }

    # Check if the PSSnapin is already loaded in the current session
    if (-not (Get-PSSnapin $SnapinName -ErrorAction SilentlyContinue)) {
        Write-Verbose "The Exchange PSSnapin '$SnapinName' is not loaded in the current session. Attempting to load it..."
        try {
            Add-PSSnapin $SnapinName -ErrorAction Stop
            Write-Verbose "The Exchange PSSnapin was loaded successfully."
        }
        catch {
            Write-Error "Failed to load the Exchange PSSnapin '$SnapinName'. Error: $($_.Exception.Message)"
            Write-Error "Ensure you are running this script on an Exchange 2016 server and have the necessary permissions."
            Write-Error "Alternatively, try running this script from the Exchange Management Shell or use a script that connects via PowerShell Remoting to localhost."
            return $false
        }
    } else {
        Write-Verbose "The Exchange PSSnapin '$SnapinName' is already loaded."
    }

    $missingCommands = @()
    foreach ($commandName in @($RequiredCommands | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            $missingCommands += $commandName
        }
    }

    if ($missingCommands.Count -gt 0) {
        Write-Error ("The following Exchange cmdlet(s) are still not available after attempting to load the snap-in: {0}" -f ($missingCommands -join ', '))
        Write-Error "This could indicate an issue with the Exchange Management Tools installation or the selected Exchange snap-in."
        return $false
    }

    if ($ViewEntireForest) {
        try {
            Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
            if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
                WriteLog -Message "Set-ADServerSettings -ViewEntireForest $true applied successfully."
            }
        }
        catch {
            Write-Error "Failed to apply 'Set-ADServerSettings -ViewEntireForest $true'. Message: $($_.Exception.Message)"
            return $false
        }
    }

    return $true
}
#endregion

#region Mail helpers and file inventory

function Get-SmartM365MailLogFolder {
    [CmdletBinding()]
    param()

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$global:LogPath)) {
        $logPathCandidate = [string]$global:LogPath
        if ([System.IO.Path]::GetExtension($logPathCandidate) -eq '.log') {
            try { $logPathCandidate = Split-Path -Path $logPathCandidate -Parent } catch {}
        }
        $candidates += $logPathCandidate
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$global:LogTextFile)) {
        try { $candidates += (Split-Path -Path ([string]$global:LogTextFile) -Parent) } catch {}
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$global:logTranscriptFile)) {
        try { $candidates += (Split-Path -Path ([string]$global:logTranscriptFile) -Parent) } catch {}
    }

    foreach ($candidate in @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        try {
            if (-not (Test-Path -LiteralPath $candidate)) {
                New-Item -Path $candidate -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            return $candidate
        }
        catch {
            try { WriteLog -Message ("Mail log folder is not writable: {0}. {1}" -f $candidate, $_.Exception.Message) -Level "WARNING" } catch {}
        }
    }

    return $null
}

function ConvertTo-SmartM365SafeFileName {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [int]$MaxLength = 90
    )

    $safe = if ([string]::IsNullOrWhiteSpace($Text)) { 'SmartM365-Mail' } else { [string]$Text }
    foreach ($invalid in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$invalid, '_')
    }
    $safe = [regex]::Replace($safe, '\s+', ' ').Trim()
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength).Trim() }
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'SmartM365-Mail' }
    return $safe
}

function Save-SmartM365MailHtmlCopy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyHtml
    )

    try {
        $folder = Get-SmartM365MailLogFolder
        if ([string]::IsNullOrWhiteSpace($folder)) { return $null }

        $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        $safeSubject = ConvertTo-SmartM365SafeFileName -Text $Subject
        $path = Join-Path -Path $folder -ChildPath ("{0}_{1}.htm" -f $timestamp, $safeSubject)
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($path, $BodyHtml, $utf8NoBom)

        if (-not $global:SmartM365MailHtmlFiles) {
            $global:SmartM365MailHtmlFiles = New-Object System.Collections.ArrayList
        }
        [void]$global:SmartM365MailHtmlFiles.Add($path)
        WriteLog -Message ("Mail HTML copy saved: {0}" -f $path) -Level "INFO"
        return $path
    }
    catch {
        WriteLog -Message ("Mail HTML copy failed: {0}" -f $_.Exception.Message) -Level "ERROR"
        throw
    }
}

function New-SmartM365MailFilesHtmlSection {
    [CmdletBinding()]
    param([string[]]$Files)

    $existingFiles = @($Files | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -Unique)
    if ($existingFiles.Count -eq 0) { return '' }

    $rows = @()
    foreach ($path in $existingFiles) {
        $item = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
        if ($null -eq $item) { continue }
        $label = ConvertTo-SmartM365EmailHtmlText $item.Name
        $locationHtml = (Get-SmartM365MailFileLocationHtml -FilePath $item.FullName).Html
        $sizeKb = [Math]::Round(($item.Length / 1KB), 1)
        $rows += "<tr><td style=`"border-top:1px solid #e2e8f0;padding:5px 7px;font-size:10px;color:#334155;`">$label</td><td style=`"border-top:1px solid #e2e8f0;padding:5px 7px;font-size:10px;color:#334155;word-break:break-all;`">$locationHtml</td><td align=`"right`" style=`"border-top:1px solid #e2e8f0;padding:5px 7px;font-size:10px;color:#334155;white-space:nowrap;`">$sizeKb KB</td></tr>"
    }

    if ($rows.Count -eq 0) { return '' }
    return @"
<div style="margin-top:18px;">
  <div style="font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px;">Technical files</div>
  <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #e2e8f0;">
    <tr><th align="left" style="background:#f8fafc;padding:5px 7px;font-size:10px;color:#475569;text-transform:uppercase;">File</th><th align="left" style="background:#f8fafc;padding:5px 7px;font-size:10px;color:#475569;text-transform:uppercase;">Location</th><th align="right" style="background:#f8fafc;padding:5px 7px;font-size:10px;color:#475569;text-transform:uppercase;">Size</th></tr>
    $($rows -join "`n")
  </table>
</div>
"@
}

function Get-SmartM365SharePointUploadRecordForLocalFile {
    [CmdletBinding()]
    param([string]$FilePath)

    if ([string]::IsNullOrWhiteSpace($FilePath) -or -not $global:SmartM365SharePointUploadedFiles) { return $null }

    $normalizedPath = [string]$FilePath
    try {
        $normalizedPath = [System.IO.Path]::GetFullPath($FilePath)
    }
    catch {
        $normalizedPath = [string]$FilePath
    }

    $matches = @($global:SmartM365SharePointUploadedFiles | Where-Object {
        $recordPath = [string]$_.LocalFilePath
        if ([string]::IsNullOrWhiteSpace($recordPath)) { return $false }

        $normalizedRecordPath = $recordPath
        try {
            $normalizedRecordPath = [System.IO.Path]::GetFullPath($recordPath)
        }
        catch {
            $normalizedRecordPath = $recordPath
        }

        $normalizedRecordPath -ieq $normalizedPath
    })

    if ($matches.Count -eq 0) { return $null }
    return $matches[-1]
}

function Get-SmartM365MailFileLocationHtml {
    [CmdletBinding()]
    param([string]$FilePath)

    $uploadRecord = Get-SmartM365SharePointUploadRecordForLocalFile -FilePath $FilePath
    if ($uploadRecord -and -not [string]::IsNullOrWhiteSpace([string]$uploadRecord.WebUrl)) {
        $escapedUrl = ConvertTo-SmartM365EmailHtmlText ([string]$uploadRecord.WebUrl)
        $displayText = [string]$uploadRecord.SharePointPath
        if ([string]::IsNullOrWhiteSpace($displayText)) { $displayText = [string]$uploadRecord.WebUrl }
        $escapedDisplayText = ConvertTo-SmartM365EmailHtmlText $displayText
        return [pscustomobject]@{
            Html = "<a href=`"$escapedUrl`" style=`"color:#075985;text-decoration:underline;`">$escapedDisplayText</a>"
            HasSharePointLocation = $true
        }
    }
    elseif ($uploadRecord -and -not [string]::IsNullOrWhiteSpace([string]$uploadRecord.SharePointPath)) {
        return [pscustomobject]@{
            Html = ConvertTo-SmartM365EmailHtmlText ([string]$uploadRecord.SharePointPath)
            HasSharePointLocation = $true
        }
    }

    return [pscustomobject]@{
        Html = ConvertTo-SmartM365EmailHtmlText ([string]$FilePath)
        HasSharePointLocation = $false
    }
}

function Convert-SmartM365MailBodyLocalPathsToSharePointLinks {
    [CmdletBinding()]
    param([AllowNull()][string]$BodyHtml)

    if ([string]::IsNullOrWhiteSpace($BodyHtml) -or -not $global:SmartM365SharePointUploadedFiles) { return $BodyHtml }

    $updatedBodyHtml = [string]$BodyHtml
    foreach ($uploadRecord in @($global:SmartM365SharePointUploadedFiles)) {
        $localPath = [string]$uploadRecord.LocalFilePath
        if ([string]::IsNullOrWhiteSpace($localPath)) { continue }

        $location = Get-SmartM365MailFileLocationHtml -FilePath $localPath
        if (-not $location.HasSharePointLocation) { continue }

        $pathVariants = @(
            $localPath,
            (ConvertTo-SmartM365EmailHtmlText $localPath)
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

        foreach ($pathVariant in $pathVariants) {
            $updatedBodyHtml = [regex]::Replace(
                $updatedBodyHtml,
                [regex]::Escape($pathVariant),
                [System.Text.RegularExpressions.MatchEvaluator]{ param($match) $location.Html }
            )
        }
    }

    return $updatedBodyHtml
}
function Add-SmartM365MailFilesSection {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$BodyHtml,
        [string[]]$Files
    )

    $section = New-SmartM365MailFilesHtmlSection -Files $Files
    if ([string]::IsNullOrWhiteSpace($section)) { return $BodyHtml }

    if ([string]::IsNullOrWhiteSpace($BodyHtml)) { return $section }
    if ($BodyHtml -match '(?is)</body>') {
        return ([regex]::Replace($BodyHtml, '(?is)</body>', ($section + "`n</body>"), 1))
    }
    return ([string]$BodyHtml + "`n" + $section)
}


function ConvertTo-SmartM365Base64Url {
    [CmdletBinding()]
    param([Parameter(Mandatory)][byte[]]$Bytes)

    return ([Convert]::ToBase64String($Bytes).TrimEnd('=') -replace '\+', '-' -replace '/', '_')
}

function Get-SmartM365CertificateByThumbprint {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Thumbprint)

    $thumb = ([string]$Thumbprint).Replace(' ', '').ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($thumb)) { throw 'Certificate thumbprint is empty.' }

    $cert = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') |
        ForEach-Object { Get-ChildItem -LiteralPath $_ -ErrorAction SilentlyContinue } |
        Where-Object { $_.Thumbprint -eq $thumb } |
        Select-Object -First 1

    if (-not $cert) { throw "Certificate not found in CurrentUser\My or LocalMachine\My ($thumb)." }
    if (-not $cert.HasPrivateKey) { throw "Certificate found but without private key ($thumb)." }
    return $cert
}

function New-SmartM365GraphClientAssertion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $now = [DateTimeOffset]::UtcNow
    $audience = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $header = [ordered]@{
        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-SmartM365Base64Url -Bytes $Certificate.GetCertHash()
    }
    $payload = [ordered]@{
        aud = $audience
        iss = $ClientId
        sub = $ClientId
        jti = [guid]::NewGuid().Guid
        nbf = [int64]$now.AddMinutes(-5).ToUnixTimeSeconds()
        exp = [int64]$now.AddMinutes(10).ToUnixTimeSeconds()
    }

    $encoding = [System.Text.Encoding]::UTF8
    $headerPart = ConvertTo-SmartM365Base64Url -Bytes $encoding.GetBytes(($header | ConvertTo-Json -Compress))
    $payloadPart = ConvertTo-SmartM365Base64Url -Bytes $encoding.GetBytes(($payload | ConvertTo-Json -Compress))
    $unsignedToken = "$headerPart.$payloadPart"
    $unsignedBytes = $encoding.GetBytes($unsignedToken)

    $rsa = $null
    try { $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate) } catch { $rsa = $null }
    if ($null -eq $rsa) { try { $rsa = $Certificate.PrivateKey } catch { $rsa = $null } }
    if ($null -eq $rsa) { throw 'Certificate private key is not an RSA key or cannot be opened.' }

    try {
        $signature = $rsa.SignData($unsignedBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    }
    catch {
        $signature = $rsa.SignData($unsignedBytes, 'SHA256')
    }

    return "$unsignedToken.$(ConvertTo-SmartM365Base64Url -Bytes $signature)"
}

function Get-SmartM365GraphAccessToken {
    [CmdletBinding()]
    param(
        [string]$AppId = $global:AppId,
        [string]$TenantId = $global:TenantId,
        [string]$Thumbprint = $(if ($global:Thumbprint) { $global:Thumbprint } else { $global:Thumb }),
        [string]$Purpose = 'Microsoft Graph REST'
    )

    $moduleLocalConfig = Get-ModuleLocalConfig
    if (-not (Test-SmartM365ConfiguredValue -Value $AppId)) { $AppId = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'AppId' -DefaultValue '' }
    if (-not (Test-SmartM365ConfiguredValue -Value $TenantId)) { $TenantId = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'TenantId' -DefaultValue '' }
    if (-not (Test-SmartM365ConfiguredValue -Value $Thumbprint)) { $Thumbprint = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'Thumbprint' -DefaultValue '' }
    if (-not (Test-SmartM365ConfiguredValue -Value $Thumbprint)) { $Thumbprint = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'Thumb' -DefaultValue '' }

    if (-not (Test-SmartM365ConfiguredValue -Value $AppId) -or -not (Test-SmartM365ConfiguredValue -Value $TenantId) -or -not (Test-SmartM365ConfiguredValue -Value $Thumbprint)) {
        throw ("{0}: Graph app-only connection values are missing (AppId, TenantId, Thumb/Thumbprint)." -f $Purpose)
    }

    if ($null -eq $script:SmartM365GraphRestTokenCache) { $script:SmartM365GraphRestTokenCache = @{} }
    $cacheKey = '{0}|{1}|{2}' -f $AppId, $TenantId, $Thumbprint
    $cached = $script:SmartM365GraphRestTokenCache[$cacheKey]
    if ($cached -and $cached.AccessToken -and $cached.ExpiresOn -gt (Get-Date).ToUniversalTime().AddMinutes(5)) { return [string]$cached.AccessToken }

    $cert = Get-SmartM365CertificateByThumbprint -Thumbprint $Thumbprint
    $assertion = New-SmartM365GraphClientAssertion -TenantId $TenantId -ClientId $AppId -Certificate $cert
    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id = $AppId
        scope = 'https://graph.microsoft.com/.default'
        grant_type = 'client_credentials'
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion = $assertion
    }

    WriteLog -Message ("Acquiring Microsoft Graph REST token for {0}." -f $Purpose) -Level 'INFO'
    $response = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    if (-not $response.access_token) { throw ("{0}: token endpoint did not return an access token." -f $Purpose) }

    $expiresIn = 3599
    try { if ($response.expires_in) { $expiresIn = [int]$response.expires_in } } catch {}
    $script:SmartM365GraphRestTokenCache[$cacheKey] = [pscustomobject]@{
        AccessToken = [string]$response.access_token
        ExpiresOn = (Get-Date).ToUniversalTime().AddSeconds([math]::Max(60, $expiresIn - 60))
    }
    return [string]$response.access_token
}

function Invoke-SmartM365GraphFileDownloadWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutputFilePath,
        [int]$MaxAttempts = 4,
        [int]$DefaultRetrySeconds = 15,
        [int]$MaximumRetrySeconds = 300,
        [string]$Operation = 'Graph file download'
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $token = Get-SmartM365GraphAccessToken -Purpose $Operation
            $params = @{ Method = 'GET'; Uri = $Uri; Headers = @{ Authorization = "Bearer $token" }; OutFile = $OutputFilePath; ErrorAction = 'Stop' }
            $invokeWebRequestCommand = Get-Command Invoke-WebRequest -ErrorAction Stop
            if ($invokeWebRequestCommand.Parameters.ContainsKey('UseBasicParsing')) { $params.UseBasicParsing = $true }
            Invoke-WebRequest @params | Out-Null
            return $true
        }
        catch {
            $statusCode = $null
            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
            $isTransient = Test-SmartM365GraphTransientError -ErrorRecord $_ -StatusCode $statusCode
            if (-not $isTransient -or $attempt -ge $MaxAttempts) { throw }
            $delay = Get-SmartM365GraphRetryAfterSeconds -ErrorRecord $_ -DefaultSeconds ($DefaultRetrySeconds * $attempt) -MaximumSeconds $MaximumRetrySeconds
            $statusText = if ($statusCode) { $statusCode } else { 'unknown' }
            WriteLog -Message ("{0} transient failure. Status={1}; attempt {2}/{3}; retrying in {4}s." -f $Operation, $statusText, $attempt, $MaxAttempts, $delay) -Level 'WARNING'
            Start-Sleep -Seconds $delay
        }
    }
}
function ConvertToRecipientArray {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Recipients
    )
    return @($Recipients) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }
}

function ConvertTo-SmartM365GraphRecipient {
    [CmdletBinding()]
    param(
        [string[]]$Recipients
    )

    @($Recipients | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        @{
            emailAddress = @{
                address = $_.Trim()
            }
        }
    })
}

function ConvertTo-SmartM365GraphFileAttachment {
    [CmdletBinding()]
    param(
        [string[]]$Attachments
    )

    $graphAttachments = @()
    foreach ($attachmentPath in @($Attachments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if (-not (Test-Path -LiteralPath $attachmentPath -PathType Leaf)) { continue }

        $file = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop
        if ($file.Length -gt 3MB) {
            throw ("Graph mail attachment '{0}' is larger than 3 MB. Large attachments require an upload session." -f $file.FullName)
        }

        $graphAttachments += @{
            '@odata.type' = '#microsoft.graph.fileAttachment'
            name          = $file.Name
            contentType   = 'application/octet-stream'
            contentBytes  = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
        }
    }

    return $graphAttachments
}

function Send-SmartM365GraphMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$From,
        [Parameter(Mandatory)][string]$To,
        [string]$Cc = "",
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyHtml,
        [string[]]$Attachments,
        [switch]$AllowAttachments,
        [switch]$SkipHtmlCopy,
        [string]$AppId = $global:AppId,
        [string]$TenantId = $global:TenantId,
        [string]$Thumbprint = $(if ($global:Thumbprint) { $global:Thumbprint } else { $global:Thumb })
    )

    $toArray = ConvertToRecipientArray -Recipients $To
    $ccArray = if ($Cc) { ConvertToRecipientArray -Recipients $Cc } else { @() }

    if (-not $toArray -or [string]::IsNullOrWhiteSpace($From)) {
        throw "Send-SmartM365GraphMail: missing required parameters (From/To)."
    }

    try {
        [void](Get-SmartM365GraphAccessToken -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'Graph mail')
    }
    catch {
        throw ("Send-SmartM365GraphMail: Microsoft Graph REST app-only connection failed. {0}" -f $_.Exception.Message)
    }

    $graphAttachmentPaths = @($Attachments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($graphAttachmentPaths.Count -gt 0) {
        $BodyHtml = Add-SmartM365MailFilesSection -BodyHtml $BodyHtml -Files $graphAttachmentPaths
        if (-not $AllowAttachments) {
            WriteLog -Message 'Graph mail attachments are disabled by default; referenced files were added to the mail body instead.' -Level 'INFO'
            $graphAttachmentPaths = @()
        }
    }

    $BodyHtml = Convert-SmartM365MailBodyLocalPathsToSharePointLinks -BodyHtml $BodyHtml
    $BodyHtml = ConvertTo-SmartM365EmailBody -BodyHtml $BodyHtml -Subject $Subject -Category 'SmartM365'
    if (-not $SkipHtmlCopy) { Save-SmartM365MailHtmlCopy -Subject $Subject -BodyHtml $BodyHtml | Out-Null }

    $message = @{
        subject      = $Subject
        body         = @{
            contentType = 'HTML'
            content     = $BodyHtml
        }
        toRecipients = @(ConvertTo-SmartM365GraphRecipient -Recipients $toArray)
    }

    if ($ccArray.Count -gt 0) {
        $message['ccRecipients'] = @(ConvertTo-SmartM365GraphRecipient -Recipients $ccArray)
    }

    $graphAttachments = @(ConvertTo-SmartM365GraphFileAttachment -Attachments $graphAttachmentPaths)
    if ($graphAttachments.Count -gt 0) {
        $message['attachments'] = $graphAttachments
    }

    $body = @{
        message         = $message
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 12

    $encodedFrom = [System.Uri]::EscapeDataString($From)
    Invoke-SmartM365GraphRestWithRetry -Method POST -Uri ("https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $encodedFrom) -Body $body -ContentType 'application/json' -Operation 'Send Graph mail' | Out-Null
    WriteLog -Message ("Graph mail sent from {0} to {1}" -f $From, ($toArray -join ';')) -Level "SUCCESS"
}

function ConvertTo-SmartM365EmailHtmlText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}


function ConvertTo-SmartM365ConfigBoolean {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [bool]$DefaultValue = $false
    )

    if ($null -eq $Value) { return $DefaultValue }
    if ($Value -is [bool]) { return [bool]$Value }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text) -or $text -in @('__USE_GLOBAL__','USE_GLOBAL')) { return $DefaultValue }

    switch -Regex ($text.ToLowerInvariant()) {
        '^(true|1|yes|y|on)$' { return $true }
        '^(false|0|no|n|off)$' { return $false }
        default { return $DefaultValue }
    }
}

function Get-SmartM365MailBrandingConfig {
    [CmdletBinding()]
    param()

    $moduleLocalConfig = Get-ModuleLocalConfig
    $callerLocalConfig = Get-SmartM365CallerLocalConfig

    function Get-MailBrandingValue {
        param([string]$Name, [AllowNull()]$DefaultValue = $null)
        $value = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name $Name -DefaultValue $DefaultValue
        return Get-ModuleLocalConfigValue -Config $callerLocalConfig -Name $Name -DefaultValue $value
    }

    $enabledValue = Get-MailBrandingValue -Name 'MailBrandingEnabled' -DefaultValue $true
    $maxLogoKbValue = Get-MailBrandingValue -Name 'MailClientLogoMaxKB' -DefaultValue 200
    $maxLogoKb = 200
    if ($null -ne $maxLogoKbValue) { [void][int]::TryParse([string]$maxLogoKbValue, [ref]$maxLogoKb) }
    if ($maxLogoKb -lt 1) { $maxLogoKb = 200 }

    return [pscustomobject]@{
        Enabled = ConvertTo-SmartM365ConfigBoolean -Value $enabledValue -DefaultValue $true
        ClientName = [string](Get-MailBrandingValue -Name 'MailClientName' -DefaultValue '')
        ClientLogoPath = [string](Get-MailBrandingValue -Name 'MailClientLogoPath' -DefaultValue '')
        ClientLogoMaxKB = $maxLogoKb
        FooterProductName = [string](Get-MailBrandingValue -Name 'MailFooterProductName' -DefaultValue 'workplacecloudhub.com')
        FooterWebsiteUrl = [string](Get-MailBrandingValue -Name 'MailFooterWebsiteUrl' -DefaultValue 'https://workplacecloudhub.com/')
        FooterGitHubLabel = [string](Get-MailBrandingValue -Name 'MailFooterGitHubLabel' -DefaultValue 'khda79/workplacecloudhub.com')
        FooterGitHubUrl = [string](Get-MailBrandingValue -Name 'MailFooterGitHubUrl' -DefaultValue 'https://github.com/khda79/workplacecloudhub.com')
    }
}

function ConvertTo-SmartM365MailLogoDataUri {
    [CmdletBinding()]
    param(
        [string]$Path,
        [int]$MaxKB = 200
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }

    $resolvedPath = Resolve-SmartM365ConfigValue -Value $Path
    if ([string]::IsNullOrWhiteSpace($resolvedPath) -or -not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        WriteLog -Message ("Mail branding logo file not found: {0}" -f $Path) -Level 'WARNING'
        return ''
    }

    try {
        $item = Get-Item -LiteralPath $resolvedPath -ErrorAction Stop
        $maxBytes = [int64]$MaxKB * 1KB
        if ($item.Length -gt $maxBytes) {
            WriteLog -Message ("Mail branding logo skipped because it is larger than {0} KB: {1}" -f $MaxKB, $item.FullName) -Level 'WARNING'
            return ''
        }

        $extension = [System.IO.Path]::GetExtension($item.FullName).TrimStart('.').ToLowerInvariant()
        $contentType = switch ($extension) {
            'png'  { 'image/png' }
            'jpg'  { 'image/jpeg' }
            'jpeg' { 'image/jpeg' }
            'gif'  { 'image/gif' }
            'svg'  { 'image/svg+xml' }
            default { 'application/octet-stream' }
        }

        $bytes = [System.IO.File]::ReadAllBytes($item.FullName)
        return ('data:{0};base64,{1}' -f $contentType, [Convert]::ToBase64String($bytes))
    }
    catch {
        WriteLog -Message ("Mail branding logo could not be read: {0}" -f $_.Exception.Message) -Level 'WARNING'
        return ''
    }
}

function Add-SmartM365MailBranding {
    [CmdletBinding()]
    param([AllowNull()][string]$BodyHtml)

    if ([string]::IsNullOrWhiteSpace($BodyHtml)) { $BodyHtml = '' }
    if ($BodyHtml -match 'SmartM365MailBranding:v1') { return $BodyHtml }

    $branding = Get-SmartM365MailBrandingConfig
    $clientName = ConvertTo-SmartM365EmailHtmlText $branding.ClientName
    $productName = ConvertTo-SmartM365EmailHtmlText $branding.FooterProductName
    $websiteUrl = ConvertTo-SmartM365EmailHtmlText $branding.FooterWebsiteUrl
    $githubLabel = ConvertTo-SmartM365EmailHtmlText $branding.FooterGitHubLabel
    $githubUrl = ConvertTo-SmartM365EmailHtmlText $branding.FooterGitHubUrl

    $logoHtml = ''
    if ($branding.Enabled) {
        $logoDataUri = ConvertTo-SmartM365MailLogoDataUri -Path $branding.ClientLogoPath -MaxKB $branding.ClientLogoMaxKB
        if (-not [string]::IsNullOrWhiteSpace($logoDataUri)) {
            $safeLogoDataUri = ConvertTo-SmartM365EmailHtmlText $logoDataUri
            $logoHtml = "<td align=`"center`" style=`"padding:0;vertical-align:middle;text-align:center;`"><img src=`"$safeLogoDataUri`" alt=`"$clientName`" width=`"92`" style=`"display:block;margin:0 auto;width:92px;max-width:92px;height:auto;border:0;outline:none;text-decoration:none;`" /></td>"
        }
    }

    $clientLabelHtml = ''
    if ($branding.Enabled -and [string]::IsNullOrWhiteSpace($logoHtml) -and -not [string]::IsNullOrWhiteSpace($clientName)) {
        $clientLabelHtml = @"
<td style="vertical-align:middle;">
  <div style="font-size:16px;line-height:22px;color:#0f172a;font-weight:700;">$clientName</div>
</td>
"@
    }
    $headerHtml = ''
    if ($branding.Enabled -and (-not [string]::IsNullOrWhiteSpace($logoHtml) -or -not [string]::IsNullOrWhiteSpace($clientLabelHtml))) {
        $headerHtml = @"
<!-- SmartM365MailBranding:v1 -->
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:12px 0 6px 0;border-collapse:collapse;">
  <tr>
    <td align="center">
      <table role="presentation" width="760" cellpadding="0" cellspacing="0" style="width:760px;max-width:760px;border-collapse:collapse;">
        <tr>
          <td align="center" style="padding:0 24px 6px 24px;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
              <tr>$logoHtml$clientLabelHtml</tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
"@
    }
    else {
        $headerHtml = '<!-- SmartM365MailBranding:v1 -->'
    }

    $footerHtml = @"
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:0 0 18px 0;border-collapse:collapse;">
  <tr>
    <td align="center">
      <table role="presentation" width="760" cellpadding="0" cellspacing="0" style="width:760px;max-width:760px;border-collapse:collapse;">
        <tr>
          <td style="padding:12px 20px;color:#64748b;font-family:Segoe UI,Arial,sans-serif;font-size:11px;line-height:16px;text-align:center;">
            <strong style="color:#334155;">$productName</strong><br />
            <a href="$websiteUrl" style="color:#2563eb;text-decoration:none;">$websiteUrl</a><br />
            <a href="$githubUrl" style="color:#2563eb;text-decoration:none;">$githubLabel</a>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
"@

    if ($BodyHtml -match '(?is)<body\b[^>]*>') {
        $branded = [regex]::Replace($BodyHtml, '(?is)(<body\b[^>]*>)', ('$1' + "`n" + $headerHtml), 1)
        if ($branded -match '(?is)</body>') {
            return [regex]::Replace($branded, '(?is)</body>', ($footerHtml + "`n" + '</body>'), 1)
        }
        return ($branded + "`n" + $footerHtml)
    }

    return ($headerHtml + "`n" + $BodyHtml + "`n" + $footerHtml)
}
function New-SmartM365EmailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [string]$Category = 'SmartM365',
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Severity = 'Info',
        [string]$Message,
        [hashtable]$SummaryData,
        [array]$SummaryRows,
        [array]$PathRows,
        [array]$Sections,
        [string]$ActionTitle,
        [string]$ActionHtml,
        [string]$BodyHtml,
        [string]$Tenant,
        [string]$HostName = $env:COMPUTERNAME,
        [string]$GeneratedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'),
        [string]$Footer = 'This automated message was generated by SmartM365.'
    )

    $palette = @{
        Info    = @{ Bg = '#eff6ff'; Border = '#bfdbfe'; Accent = '#2563eb'; Title = '#1d4ed8'; Text = '#1e3a8a' }
        Success = @{ Bg = '#ecfdf5'; Border = '#bbf7d0'; Accent = '#16a34a'; Title = '#166534'; Text = '#14532d' }
        Warning = @{ Bg = '#fff7ed'; Border = '#fed7aa'; Accent = '#f97316'; Title = '#9a3412'; Text = '#7c2d12' }
        Error   = @{ Bg = '#fef2f2'; Border = '#fecaca'; Accent = '#dc2626'; Title = '#991b1b'; Text = '#7f1d1d' }
    }
    $colors = $palette[$Severity]

    $htmlTitle = ConvertTo-SmartM365EmailHtmlText $Title
    $htmlCategory = ConvertTo-SmartM365EmailHtmlText $Category
    $htmlTenant = ConvertTo-SmartM365EmailHtmlText $Tenant
    $htmlHostName = ConvertTo-SmartM365EmailHtmlText $HostName
    $htmlGeneratedAt = ConvertTo-SmartM365EmailHtmlText $GeneratedAt
    $htmlFooter = ConvertTo-SmartM365EmailHtmlText $Footer

    $metaParts = @()
    if (-not [string]::IsNullOrWhiteSpace($Tenant)) { $metaParts += "Tenant: $htmlTenant" }
    if (-not [string]::IsNullOrWhiteSpace($HostName)) { $metaParts += "Host: $htmlHostName" }
    if (-not [string]::IsNullOrWhiteSpace($GeneratedAt)) { $metaParts += "Generated: $htmlGeneratedAt" }
    $metaHtml = $metaParts -join ' &nbsp;|&nbsp; '

    $messageHtml = ''
    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $safeMessage = (ConvertTo-SmartM365EmailHtmlText $Message) -replace "(`r`n|`n|`r)", '<br />'
        $messageHtml = "<tr><td style=`"padding:18px 24px 0 24px;font-size:14px;line-height:21px;color:#334155;`">$safeMessage</td></tr>"
    }

    $actionSectionHtml = ''
    if (-not [string]::IsNullOrWhiteSpace($ActionTitle) -or -not [string]::IsNullOrWhiteSpace($ActionHtml)) {
        $safeActionTitle = if ($ActionTitle) { ConvertTo-SmartM365EmailHtmlText $ActionTitle } else { 'Action' }
        $actionSectionHtml = @"
          <tr>
            <td style="padding:22px 24px 8px 24px;">
              <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
                <tr>
                  <td style="background:$($colors.Bg);border:1px solid $($colors.Border);border-left:5px solid $($colors.Accent);padding:12px 14px;border-radius:4px;">
                    <div style="font-size:14px;font-weight:700;color:$($colors.Title);">$safeActionTitle</div>
                    <div style="font-size:13px;line-height:19px;color:$($colors.Text);margin-top:4px;">$ActionHtml</div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
"@
    }

    $summarySectionHtml = ''
    if (($null -ne $SummaryData -and $SummaryData.Count -gt 0) -or ($null -ne $SummaryRows -and $SummaryRows.Count -gt 0)) {
        $rows = @()
        if ($null -ne $SummaryRows -and $SummaryRows.Count -gt 0) {
            foreach ($row in $SummaryRows) {
                $label = ConvertTo-SmartM365EmailHtmlText $row.Label
                $value = ConvertTo-SmartM365EmailHtmlText $row.Value
                $rows += "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$label</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:14px;font-weight:700;color:#111827;`">$value</td></tr>"
            }
        }
        else {
            foreach ($key in $SummaryData.Keys) {
                $label = ConvertTo-SmartM365EmailHtmlText $key
                $value = ConvertTo-SmartM365EmailHtmlText $SummaryData[$key]
                $rows += "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$label</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:14px;font-weight:700;color:#111827;`">$value</td></tr>"
            }
        }
        $summaryRowsHtml = $rows -join "`n"
        $summarySectionHtml = @"
          <tr>
            <td style="padding:18px 24px 0 24px;">
              <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:8px;">Summary</div>
              <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
                <tr>
                  <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Metric</th>
                  <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Value</th>
                </tr>
                $summaryRowsHtml
              </table>
            </td>
          </tr>
"@
    }

    $pathSectionHtml = ''
    if ($null -ne $PathRows -and $PathRows.Count -gt 0) {
        $pathRowsHtml = foreach ($pathRow in $PathRows) {
            $label = ConvertTo-SmartM365EmailHtmlText $pathRow.Label
            $path = ConvertTo-SmartM365EmailHtmlText $pathRow.Path
            "<tr><td style=`"width:150px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;`">$label</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$path</td></tr>"
        }
        $pathSectionHtml = @"
          <tr>
            <td style="padding:18px 24px 0 24px;">
              <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:8px;">Paths</div>
              <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
                $($pathRowsHtml -join "`n")
              </table>
            </td>
          </tr>
"@
    }

    $customSectionsHtml = ''
    if ($null -ne $Sections -and $Sections.Count -gt 0) {
        $sectionHtmlItems = foreach ($section in $Sections) {
            $sectionTitle = ConvertTo-SmartM365EmailHtmlText $section.Title
            @"
          <tr>
            <td style="padding:18px 24px 0 24px;">
              <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:8px;">$sectionTitle</div>
              $($section.Html)
            </td>
          </tr>
"@
        }
        $customSectionsHtml = $sectionHtmlItems -join "`n"
    }

    if ([string]::IsNullOrWhiteSpace($BodyHtml)) { $BodyHtml = '' }

    return @"
<!doctype html>
<!-- SmartM365EmailTemplate:v1 -->
<html>
<head>
  <meta charset="utf-8">
  <title>$htmlTitle</title>
</head>
<body style="margin:0;padding:0;background:#f4f6f8;font-family:Segoe UI,Arial,sans-serif;color:#111827;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:24px 0;">
    <tr>
      <td align="center">
        <table role="presentation" width="760" cellpadding="0" cellspacing="0" style="width:760px;max-width:760px;background:#ffffff;border:1px solid #d9e2ec;border-radius:6px;overflow:hidden;">
          <tr>
            <td style="background:#0f172a;color:#ffffff;padding:20px 24px;">
              <div style="font-size:12px;text-transform:uppercase;color:#93c5fd;font-weight:700;">$htmlCategory</div>
              <div style="font-size:24px;line-height:30px;font-weight:700;margin-top:6px;">$htmlTitle</div>
              <div style="font-size:13px;color:#cbd5e1;margin-top:8px;">$metaHtml</div>
            </td>
          </tr>
          $messageHtml
          $actionSectionHtml
          $summarySectionHtml
          $customSectionsHtml
          $pathSectionHtml
          $BodyHtml
          <tr>
            <td style="background:#f8fafc;border-top:1px solid #d9e2ec;padding:12px 24px;color:#64748b;font-size:12px;line-height:18px;">
              $htmlFooter
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"@
}

function ConvertTo-SmartM365EmailBody {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$BodyHtml,
        [string]$Subject = 'SmartM365',
        [string]$Category = 'SmartM365',
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Severity = 'Info'
    )

    $effectiveTitle = if (-not [string]::IsNullOrWhiteSpace($Subject)) { $Subject } else { 'SmartM365' }
    if (-not [string]::IsNullOrWhiteSpace($BodyHtml) -and $BodyHtml -match '(?is)<title>\s*(?<Title>.*?)\s*</title>') {
        $effectiveTitle = [System.Net.WebUtility]::HtmlDecode($matches['Title'])
    }

    if ([string]::IsNullOrWhiteSpace($BodyHtml)) {
        return Add-SmartM365MailBranding -BodyHtml (New-SmartM365EmailBody -Title $effectiveTitle -Category $Category -Severity $Severity)
    }

    if ($BodyHtml -match 'SmartM365EmailTemplate:v1') {
        return Add-SmartM365MailBranding -BodyHtml $BodyHtml
    }

    $legacyHtml = [string]$BodyHtml
    $styleBlocks = @()
    foreach ($match in [regex]::Matches($legacyHtml, '(?is)<style\b[^>]*>.*?</style>')) {
        $styleBlocks += $match.Value
    }
    $styleHtml = $styleBlocks -join "`n"

    if ($legacyHtml -match '(?is)<body\b[^>]*>(?<Body>.*?)</body>') {
        $legacyHtml = $matches['Body']
    }
    else {
        $legacyHtml = [regex]::Replace($legacyHtml, '(?is)<!doctype[^>]*>|<html\b[^>]*>|</html>|<head\b[^>]*>.*?</head>', '')
    }
    $legacyHtml = $legacyHtml.Trim()

    $wrappedBody = @"
          <tr>
            <td style="padding:18px 24px 22px 24px;font-size:13px;line-height:19px;color:#334155;">
              $styleHtml
              $legacyHtml
            </td>
          </tr>
"@

    return Add-SmartM365MailBranding -BodyHtml (New-SmartM365EmailBody -Title $effectiveTitle -Category $Category -Severity $Severity -BodyHtml $wrappedBody)
}

function NewSimpleEmailBody {
    param(
        [string]$Title,
        [string]$Message = "See attached report(s) for details."
    )

    return New-SmartM365EmailBody -Title $Title -Category 'SmartM365 Report' -Message $Message
}

function ConvertBytesToSizeString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][long]$Bytes)

    if ($Bytes -lt 1KB) { return "$Bytes B" }
    elseif ($Bytes -lt 1MB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    elseif ($Bytes -lt 1GB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    elseif ($Bytes -lt 1TB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    else { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
}

function GetFileList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Files,
        [switch]$Recurse,
        [switch]$ComputeHash
    )

    $resolved = New-Object System.Collections.Generic.List[string]

    foreach ($f in $Files) {
        if (-not $f) { continue }

        if (Test-Path -LiteralPath $f -PathType Leaf) {
            $resolved.Add((Resolve-Path -LiteralPath $f).Path) | Out-Null
            continue
        }

        if (Test-Path -LiteralPath $f -PathType Container) {
            $items = Get-ChildItem -LiteralPath $f -File -Recurse:$Recurse -ErrorAction SilentlyContinue
            foreach ($i in $items) { $resolved.Add($i.FullName) | Out-Null }
            continue
        }

        $parent = Split-Path -Path $f -Parent
        $leaf   = Split-Path -Path $f -Leaf
        if ([string]::IsNullOrWhiteSpace($parent)) { $parent = '.' }
        if (Test-Path -LiteralPath $parent) {
            $items = Get-ChildItem -Path $parent -Filter $leaf -File -Recurse:$Recurse -ErrorAction SilentlyContinue
            foreach ($i in $items) { $resolved.Add($i.FullName) | Out-Null }
        }
    }

    $resolved = $resolved | Sort-Object -Unique

    $countOK   = 0
    $countErr  = 0
    $sizeTotal = 0
    $fileList  = @()

    foreach ($path in $resolved) {
        try {
            $item = Get-Item -LiteralPath $path -ErrorAction Stop
            $hash = $null
            if ($ComputeHash) {
                $hash = (Get-FileHash -Path $path -Algorithm SHA256 -ErrorAction SilentlyContinue).Hash
            }
            $fileList += [PSCustomObject]@{
                'FileName'      = $item.Name
                'FullPath'      = $item.FullName
                'FullName'      = $item.FullName
                'Size'          = $item.Length
                'SizeString'    = ConvertBytesToSizeString -Bytes $item.Length
                'CreationTime'  = $item.CreationTime
                'LastWriteTime' = $item.LastWriteTime
                'Hash'          = $hash
            }
            $countOK++
            $sizeTotal += [int64]$item.Length
        }
        catch {
            $countErr++
        }
    }

    $summary = [PSCustomObject]@{
        'Files (count)' = $countOK
        'Total size'    = ConvertBytesToSizeString -Bytes $sizeTotal
        'Errors'        = $countErr
        'Files'         = $fileList
    }

    return $summary
}

function NewTableEmailBody {
    param(
        [string]$Title,
        [hashtable]$SummaryData,
        [string]$Message
    )

    return New-SmartM365EmailBody -Title $Title -Category 'SmartM365 Report' -Message $Message -SummaryData $SummaryData
}

function NewTableFilesEmailBody {
    param(
        [string]$Title,
        [hashtable]$SummaryData,
        [array]$Files,
        [string]$Message
    )

    foreach ($file in $Files) {
        if ($file.FileName -match '\.csv$') {
            try {
                $rowCount = (Get-Content $file.FullName | Measure-Object -Line).Lines - 1
                WriteLog "Row count for '$($file.FileName)': $rowCount" "INFO"
            }
            catch {
                $rowCount = "N/A"
                WriteLog "Failed to count rows in '$($file.FileName)': $_" "ERROR"
            }
            $file | Add-Member -NotePropertyName RowCount -NotePropertyValue $rowCount -Force
        }
        else {
            $file | Add-Member -NotePropertyName RowCount -NotePropertyValue "" -Force
        }
    }

    $fileRows = foreach ($file in $Files) {
        $fileName = ConvertTo-SmartM365EmailHtmlText $file.FileName
        $size = ConvertTo-SmartM365EmailHtmlText $file.SizeString
        $rowCount = ConvertTo-SmartM365EmailHtmlText $file.RowCount
        $creationTime = ConvertTo-SmartM365EmailHtmlText $file.CreationTime
        $lastWriteTime = ConvertTo-SmartM365EmailHtmlText $file.LastWriteTime
        $hash = ConvertTo-SmartM365EmailHtmlText $file.Hash
        "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$fileName</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$size</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$rowCount</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$creationTime</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$lastWriteTime</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$hash</td></tr>"
    }

    $fileDetailsHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">File name</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Size</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Rows</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Created</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Modified</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Hash</th>
  </tr>
  $($fileRows -join "`n")
</table>
"@

    $sections = @([pscustomobject]@{ Title = 'File details'; Html = $fileDetailsHtml })
    return New-SmartM365EmailBody -Title $Title -Category 'SmartM365 Report' -Message $Message -SummaryData $SummaryData -Sections $sections
}

function SendEmailHtmlReport {
    [CmdletBinding()]
    param(
        [string]$SmtpServer = "",
        [int]$SmtpPort = 25,
        [string]$SendMailMode = "",
        [string]$From = "",
        [string]$To = "",
        [string]$Cc = "",
        [string]$Subject = "SmartM365",
        [string]$BodyHtml,
        [string[]]$Attachments,
        [switch]$AllowAttachments,
        [switch]$VerboseLog
    )

    try {
        $moduleLocalConfig = Get-ModuleLocalConfig
        $callerLocalConfig = Get-SmartM365CallerLocalConfig
        foreach ($configName in @('SendMailMode','SmtpServer','From','To','Cc','Subject')) {
            if (-not $PSBoundParameters.ContainsKey($configName)) {
                $defaultValue = Get-Variable -Name $configName -ValueOnly
                if ($configName -eq 'To') {
                    $defaultValue = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'ErrorMailTo' -DefaultValue $defaultValue
                }
                $configValue = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name $configName -DefaultValue $defaultValue
                $configValue = Get-ModuleLocalConfigValue -Config $callerLocalConfig -Name $configName -DefaultValue $configValue
                Set-Variable -Name $configName -Value $configValue -Scope Local
            }
        }

        $Subject = Add-SmartM365MaxItemsSubjectPrefix -Subject $Subject
        $BodyHtml = Add-SmartM365MaxItemsMailBanner -BodyHtml $BodyHtml

        $toArray = ConvertToRecipientArray -Recipients $To
        $ccArray = if ($Cc) { ConvertToRecipientArray -Recipients $Cc } else { @() }

        if (-not $toArray -or [string]::IsNullOrWhiteSpace($From)) {
            throw "SendEmailHtmlReport: missing required parameters (From/To)."
        }

        if ([string]::IsNullOrWhiteSpace($Subject)) {
            if ($BodyHtml -match '<title>(.*?)</title>') {
                $Subject = $matches[1]
                if ($PSBoundParameters.ContainsKey('VerboseLog')) {
                    WriteLog -Message "Subject was empty, set from HTML title: $Subject" -Level "INFO"
                }
            } else {
                $Subject = "SmartM365"
                if ($PSBoundParameters.ContainsKey('VerboseLog')) {
                    WriteLog -Message "Subject was empty and no <title> found, set to default: $Subject" -Level "INFO"
                }
            }
        }

        $atts = @()
        foreach ($a in ($Attachments | Where-Object { $_ })) {
            if (Test-Path $a) { $atts += $a }
        }
        if ($atts.Count -gt 0) {
            $BodyHtml = Add-SmartM365MailFilesSection -BodyHtml $BodyHtml -Files $atts
            if (-not $AllowAttachments) {
                WriteLog -Message 'Email attachments are disabled by default; referenced files were added to the mail body instead.' -Level 'INFO'
                $atts = @()
            }
        }
        $BodyHtml = Convert-SmartM365MailBodyLocalPathsToSharePointLinks -BodyHtml $BodyHtml
        $BodyHtml = ConvertTo-SmartM365EmailBody -BodyHtml $BodyHtml -Subject $Subject -Category 'SmartM365'
        Save-SmartM365MailHtmlCopy -Subject $Subject -BodyHtml $BodyHtml | Out-Null

        $effectiveSendMailMode = $SendMailMode
        if ([string]::IsNullOrWhiteSpace($effectiveSendMailMode)) {
            $effectiveSendMailMode = if ([string]::IsNullOrWhiteSpace($SmtpServer)) { 'Graph' } else { 'SMTP' }
        }
        $effectiveSendMailMode = $effectiveSendMailMode.Trim()
        if ($effectiveSendMailMode -notin @('Graph','SMTP','Both')) {
            throw "SendEmailHtmlReport: unsupported SendMailMode '$effectiveSendMailMode'. Use Graph, SMTP, or Both."
        }

        $mailParams = @{
            SmtpServer  = $SmtpServer
            Port        = $SmtpPort
            From        = $From
            To          = $toArray
            Subject     = $Subject
            Body        = $BodyHtml
            BodyAsHtml  = $true
            ErrorAction = 'Stop'
        }
        if ($ccArray.Count -gt 0) { $mailParams['Cc'] = $ccArray }
        if ($atts.Count   -gt 0) { $mailParams['Attachments'] = $atts }

        if ($effectiveSendMailMode -eq 'Graph') {
            Send-SmartM365GraphMail -From $From -To ($toArray -join ';') -Cc ($ccArray -join ';') -Subject $Subject -BodyHtml $BodyHtml -Attachments $atts -AllowAttachments:$AllowAttachments -SkipHtmlCopy
            if ($PSBoundParameters.ContainsKey('VerboseLog')) {
                WriteLog -Message "Email sent to $($toArray -join ';') via Microsoft Graph" -Level "SUCCESS"
            }
            return
        }

        if ($effectiveSendMailMode -eq 'SMTP') {
            if ([string]::IsNullOrWhiteSpace($SmtpServer)) { throw 'SendEmailHtmlReport: SmtpServer is required when SendMailMode is SMTP.' }
            Send-MailMessage @mailParams
            if ($PSBoundParameters.ContainsKey('VerboseLog')) {
                WriteLog -Message "Email sent to $($toArray -join ';') via $($SmtpServer):$SmtpPort" -Level "SUCCESS"
            }
            return
        }

        try {
            Send-SmartM365GraphMail -From $From -To ($toArray -join ';') -Cc ($ccArray -join ';') -Subject $Subject -BodyHtml $BodyHtml -Attachments $atts -AllowAttachments:$AllowAttachments -SkipHtmlCopy
            if ($PSBoundParameters.ContainsKey('VerboseLog')) {
                WriteLog -Message "Email sent to $($toArray -join ';') via Microsoft Graph" -Level "SUCCESS"
            }
            return
        }
        catch {
            WriteLog -Message ("Graph mail failed; falling back to SMTP: {0}" -f $_.Exception.Message) -Level "WARNING"
            if ([string]::IsNullOrWhiteSpace($SmtpServer)) { throw 'SendEmailHtmlReport: Graph mail failed and SMTP fallback is unavailable because SmtpServer is empty.' }
            Send-MailMessage @mailParams
            if ($PSBoundParameters.ContainsKey('VerboseLog')) {
                WriteLog -Message "Email sent to $($toArray -join ';') via $($SmtpServer):$SmtpPort after Graph fallback" -Level "SUCCESS"
            }
            return
        }
    }
    catch {
        WriteLog -Message ("SendEmailHtmlReport failed: {0}" -f $_.Exception.Message) -Level "ERROR"
        throw
    }
}

function Send-SmartM365Mail {
    [CmdletBinding()]
    param(
        [string]$SmtpServer = "",
        [int]$SmtpPort = 25,
        [string]$SendMailMode = "",
        [string]$From = "",
        [string]$To = "",
        [string]$Cc = "",
        [string]$Subject = "SmartM365",
        [string]$Body,
        [string]$BodyHtml,
        [string[]]$Attachments,
        [switch]$AllowAttachments,
        [switch]$BodyAsHtml,
        [switch]$HighPriority
    )

    $moduleLocalConfig = Get-ModuleLocalConfig
    $callerLocalConfig = Get-SmartM365CallerLocalConfig
    foreach ($configName in @('SendMailMode','SmtpServer','From','To','Cc','Subject')) {
        if (-not $PSBoundParameters.ContainsKey($configName)) {
            $defaultValue = Get-Variable -Name $configName -ValueOnly
            if ($configName -eq 'To') {
                $defaultValue = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'ErrorMailTo' -DefaultValue $defaultValue
            }
            $configValue = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name $configName -DefaultValue $defaultValue
            $configValue = Get-ModuleLocalConfigValue -Config $callerLocalConfig -Name $configName -DefaultValue $configValue
            Set-Variable -Name $configName -Value $configValue -Scope Local
        }
    }

    $htmlBody = if (-not [string]::IsNullOrWhiteSpace($BodyHtml)) { $BodyHtml } else { $Body }

    $mailParams = @{
        SmtpServer = $SmtpServer
        SmtpPort   = $SmtpPort
        From       = $From
        To         = $To
        Cc         = $Cc
        Subject    = $Subject
        BodyHtml   = $htmlBody
        Attachments = $Attachments
        AllowAttachments = $AllowAttachments
    }
    if (-not [string]::IsNullOrWhiteSpace($SendMailMode)) {
        $mailParams['SendMailMode'] = $SendMailMode
    }

    SendEmailHtmlReport @mailParams
}

function SendFileListEmailReport {
    [CmdletBinding()]
    param(
        [string[]]$Files,
        [Parameter(Mandatory)][string]$Title,
        [string]$Message,
        [switch]$Recurse,
        [switch]$ComputeHash
    )

    if (-not $Files -or $Files.Count -eq 0) {
        $body = NewSimpleEmailBody -Title $Title -Message $Message
        SendEmailHtmlReport -Subject $Title -BodyHtml $body
        return
    }

    $summary = GetFileList -Files $Files -Recurse:$Recurse -ComputeHash:$ComputeHash

    $summaryData = @{
        "Files (count)" = $summary.'Files (count)'
        "Total size"    = $summary.'Total size'
        "Errors"        = $summary.Errors
    }

    $body = NewTableFilesEmailBody -Title $Title `
        -SummaryData $summaryData `
        -Files $summary.Files `
        -Message $Message

    SendEmailHtmlReport -BodyHtml $body
}

function Send-SmartM365TeamsNotification {
    <#
    .SYNOPSIS
        Sends a SmartM365 notification to a Teams workflow or incoming webhook URL.

    .DESCRIPTION
        Posts a MessageCard-compatible JSON payload to TeamsAlertsWebhookUrl or TeamsInfosWebhookUrl
        from local/global configuration, or to the explicit WebhookUrl parameter. If no URL is configured,
        the function logs and returns $false without throwing.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$WebhookUrl = "",
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level = 'INFO',
        [ValidateSet('Auto','Alerts','Infos')]
        [string]$Channel = 'Auto',
        [hashtable]$Facts,
        [string]$ResultSummary = "",
        [string]$HelpUrl = "",
        [switch]$ThrowOnError
    )

    try {
        $effectiveChannel = if ($Channel -ne 'Auto') {
            $Channel
        }
        elseif ($Level -eq 'ERROR') {
            'Alerts'
        }
        else {
            'Infos'
        }

        if (-not $PSBoundParameters.ContainsKey('WebhookUrl')) {
            $moduleLocalConfig = Get-ModuleLocalConfig
            $teamsEnabled = [bool](Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'EnableTeamsNotifications' -DefaultValue $false)
            if (-not $teamsEnabled) {
                WriteLog -Message 'Teams notification skipped: EnableTeamsNotifications is not enabled.' -Level 'INFO'
                return $false
            }

            $webhookConfigName = if ($effectiveChannel -eq 'Alerts') { 'TeamsAlertsWebhookUrl' } else { 'TeamsInfosWebhookUrl' }
            $WebhookUrl = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name $webhookConfigName -DefaultValue ''
            if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
                $WebhookUrl = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'TeamsWebhookUrl' -DefaultValue $WebhookUrl
            }
        }

        if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
            WriteLog -Message 'Teams notification skipped: TeamsAlertsWebhookUrl/TeamsInfosWebhookUrl is not configured.' -Level 'INFO'
            return $false
        }

        $themeColors = @{
            INFO    = '0078D4'
            SUCCESS = '107C10'
            WARNING = 'FFB900'
            ERROR   = 'D13438'
        }

        $factList = New-Object System.Collections.Generic.List[hashtable]
        $factList.Add(@{ name = 'Level'; value = $Level }) | Out-Null
        $factList.Add(@{ name = 'Timestamp'; value = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }) | Out-Null
        $factList.Add(@{ name = 'Computer'; value = $env:COMPUTERNAME }) | Out-Null

        $summaryFactNames = @('Result summary', 'ResultSummary', 'Summary')
        $hasSummaryFact = $false
        $summaryFromFacts = ''
        if ($null -ne $Facts) {
            foreach ($summaryFactName in $summaryFactNames) {
                if ($Facts.ContainsKey($summaryFactName)) {
                    $summaryFromFacts = [string]$Facts[$summaryFactName]
                    break
                }
            }
            $hasSummaryFact = $Facts.ContainsKey('Result summary')
        }
        if ($effectiveChannel -eq 'Infos' -and -not $hasSummaryFact) {
            $summaryText = if (-not [string]::IsNullOrWhiteSpace($ResultSummary)) {
                $ResultSummary
            }
            elseif (-not [string]::IsNullOrWhiteSpace($summaryFromFacts)) {
                $summaryFromFacts
            }
            else {
                $Message
            }
            if ($summaryText.Length -gt 600) {
                $summaryText = $summaryText.Substring(0, 597) + '...'
            }
            $factList.Add(@{ name = 'Result summary'; value = $summaryText }) | Out-Null
        }

        if ($null -ne $Facts) {
            foreach ($key in @($Facts.Keys | Sort-Object)) {
                $value = $Facts[$key]
                if ($null -eq $value) { continue }
                $factList.Add(@{
                    name  = [string]$key
                    value = [string]$value
                }) | Out-Null
            }
        }

        $actions = @()
        if (-not [string]::IsNullOrWhiteSpace($HelpUrl)) {
            $actions += @{
                '@type' = 'OpenUri'
                name    = 'Get AI help'
                targets = @(
                    @{
                        os  = 'default'
                        uri = $HelpUrl
                    }
                )
            }
        }

        $payload = @{
            '@type'    = 'MessageCard'
            '@context' = 'https://schema.org/extensions'
            summary    = $Title
            themeColor = $themeColors[$Level]
            title      = $Title
            text       = $Message
            sections   = @(
                @{
                    markdown = $true
                    facts    = @($factList)
                }
            )
        }
        if ($actions.Count -gt 0) {
            $payload['potentialAction'] = $actions
        }

        Invoke-RestMethod `
            -Method POST `
            -Uri $WebhookUrl `
            -ContentType 'application/json; charset=utf-8' `
            -Body ($payload | ConvertTo-Json -Depth 8) `
            -ErrorAction Stop | Out-Null

        WriteLog -Message ("Teams notification sent: {0}" -f $Title) -Level 'SUCCESS'
        return $true
    }
    catch {
        WriteLog -Message ("Teams notification failed: {0}" -f $_.Exception.Message) -Level 'ERROR'
        if ($ThrowOnError) {
            throw
        }
        return $false
    }
}

#endregion

#region Initialization and CSV helpers

function TestSharePath {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        WriteLog -Message "The share '$Path' is not available. Stopping the script." -Level "ERROR"
        throw "The share '$Path' is not available."
    }
}

function InitializeScriptEnvironment {
    param(
        [string]$OutputPathInit,
        [string]$OutputPath,
        [string]$LogFileName,
        [string]$CallerScriptPath
    )

    $CallerScriptPath = [string]$CallerScriptPath
    $callerScriptPath = if ([string]::IsNullOrWhiteSpace($CallerScriptPath)) { '' } else { $CallerScriptPath }
    if ([string]::IsNullOrWhiteSpace($callerScriptPath)) {
        try {
            foreach ($frame in (Get-PSCallStack)) {
                if ([string]::IsNullOrWhiteSpace([string]$frame.ScriptName)) { continue }
                if ($frame.ScriptName -eq $PSCommandPath) { continue }
                if ($frame.ScriptName -notlike '*.ps1') { continue }
                $callerScriptPath = [string]$frame.ScriptName
                break
            }
        }
        catch {
            $callerScriptPath = ''
        }
    }
    $expectedLocalConfigPath = ''
    $expectedTemplatePath = ''
    if (-not [string]::IsNullOrWhiteSpace($callerScriptPath)) {
        $callerScriptName = [System.IO.Path]::GetFileNameWithoutExtension($callerScriptPath)
        $callerScriptDirectory = Split-Path -Path $callerScriptPath -Parent
        $expectedLocalConfigPath = Join-Path -Path $callerScriptDirectory -ChildPath ("{0}.local.json" -f $callerScriptName)
        $expectedTemplatePath = "{0}.template" -f $expectedLocalConfigPath
    }

    if ([string]::IsNullOrWhiteSpace($OutputPathInit) -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPathInit = $OutputPath
    }

    if ([string]::IsNullOrWhiteSpace($OutputPathInit)) {
        $messageParts = @(
            "InitializeScriptEnvironment could not resolve an output folder.",
            "Create or update the script local JSON and set the script-specific output path key, for example '<ScriptName>CsvLogFolderPath'."
        )
        if (-not [string]::IsNullOrWhiteSpace($callerScriptPath)) {
            $messageParts += "Script: $callerScriptPath"
        }
        if (-not [string]::IsNullOrWhiteSpace($expectedLocalConfigPath)) {
            $messageParts += "Expected local JSON: $expectedLocalConfigPath"
        }
        if (-not [string]::IsNullOrWhiteSpace($expectedTemplatePath)) {
            $messageParts += "Template to copy: $expectedTemplatePath"
        }
        $messageParts += "You can also pass -OutputPath explicitly when the script supports it."
        $message = $messageParts -join ' '
        Write-Warning $message
        throw $message
    }

    try {
        if (-not (Test-Path -LiteralPath $OutputPathInit)) {
            New-Item -Path $OutputPathInit -ItemType Directory -Force -ErrorAction Stop | Out-Null
            WriteLog -Message "Created missing output directory: $OutputPathInit"
        }
    }
    catch {
        WriteLog -Message "Failed to create or access output directory '$OutputPathInit': $_" -Level "ERROR"
        throw
    }

    $moduleLocalConfig = Get-ModuleLocalConfig
    $logAllRootPath = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'LogAllRootPath' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($logAllRootPath)) {
        $message = "InitializeScriptEnvironment could not resolve LogAllRootPath. Create Config\SmartM365.global.local.json from Config\SmartM365.global.local.json.template at the SmartM365 root, then set LogAllRootPath or keep the default tokenized value."
        Write-Warning $message
        throw $message
    }
    $global:SmartM365ExecutionStartTime = Get-Date
    $global:SmartM365ExecutionSummaryWritten = $false
    $global:SmartM365WarningCount = 0
    $global:SmartM365ErrorCount = 0
    $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Initialize-SmartM365DefaultCsvValidationRules
    $global:BasePath = $OutputPathInit
    $global:SmartM365ScriptName = $LogFileName
    $global:LogPath  = Join-Path -Path $logAllRootPath -ChildPath $LogFileName

    try {
        if (-not (Test-Path -Path $global:LogPath)) {
            New-Item -Path $global:LogPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    } catch {
        Write-Error "Failed to create the log directory at path '$global:LogPath'. Error details: $_"
        throw "Script execution stopped due to failure in creating the log directory."
    }

    $global:LogTextFile       = Join-Path $global:LogPath "$LogFileName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
    $global:logTranscriptFile = Join-Path $global:LogPath "$LogFileName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_Transcript.log"

    try {
        $callerMaxItemsVariable = Get-Variable -Name MaxItems -Scope 1 -ErrorAction SilentlyContinue
        if ($callerMaxItemsVariable -and $null -ne $callerMaxItemsVariable.Value -and [int]$callerMaxItemsVariable.Value -gt 0) {
            Set-SmartM365MaxItemsMode -MaxItems ([int]$callerMaxItemsVariable.Value)
        }
    }
    catch { }
    if ($global:RetentionMaxLogs -gt 0) {
        RemoveOldFiles -FolderPath $global:LogPath -FilePattern "$LogFileName*.log" -MaxFiles $global:RetentionMaxLogs
    }
    $scriptVersion = Get-SmartM365ScriptVersionFromFile -Path $callerScriptPath
    Write-SmartM365ExecutionContext -ScriptName $LogFileName -ScriptVersion $scriptVersion -OutputPath $OutputPathInit -ScriptPath $callerScriptPath
    Write-Host "Environment initialized successfully."

    return $OutputPathInit
}

function ConvertTo-GraphDrivePath {
    param([Parameter(Mandatory)][string]$Path)

    (($Path -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) |
        ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/'
}

function Test-SmartM365ConfiguredValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string]) {
        return ($null -ne $Value)
    }

    $normalizedValue = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($normalizedValue)) { return $false }
    if ($normalizedValue -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return $false }
    if ($normalizedValue -match '^0{8}-0{4}-0{4}-0{4}-0{12}$') { return $false }
    if ($normalizedValue -match '^0{40}$') { return $false }

    return $true
}

function Get-SmartM365ExceptionDetails {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Exception)

    $messages = New-Object System.Collections.Generic.List[string]
    $current = $Exception
    while ($null -ne $current) {
        if (-not [string]::IsNullOrWhiteSpace($current.Message)) {
            $messages.Add($current.Message)
        }
        $current = $current.InnerException
    }

    if ($messages.Count -eq 0) {
        return [string]$Exception
    }

    return ($messages | Select-Object -Unique) -join " | "
}

function Connect-SmartM365GraphAppOnly {
    [CmdletBinding()]
    param(
        [string]$AppId,
        [string]$TenantId,
        [string]$Thumbprint,
        [string]$Purpose = 'Microsoft Graph'
    )

    try {
        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        }

        $moduleLocalConfig = Get-ModuleLocalConfig
        if (-not (Test-SmartM365ConfiguredValue -Value $AppId)) {
            $AppId = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'AppId' -DefaultValue ''
        }
        if (-not (Test-SmartM365ConfiguredValue -Value $TenantId)) {
            $TenantId = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'TenantId' -DefaultValue ''
        }
        if (-not (Test-SmartM365ConfiguredValue -Value $Thumbprint)) {
            $Thumbprint = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'Thumbprint' -DefaultValue ''
        }
        if (-not (Test-SmartM365ConfiguredValue -Value $Thumbprint)) {
            $Thumbprint = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'Thumb' -DefaultValue ''
        }

        if (-not (Test-SmartM365ConfiguredValue -Value $AppId) -or
            -not (Test-SmartM365ConfiguredValue -Value $TenantId) -or
            -not (Test-SmartM365ConfiguredValue -Value $Thumbprint)) {
            WriteLog -Message ("{0} skipped: Graph app-only connection values are missing (AppId, TenantId, Thumb/Thumbprint)." -f $Purpose) -Level "WARNING"
            return $false
        }

        WriteLog -Message ("Connecting to Microsoft Graph for {0}." -f $Purpose) -Level "INFO"
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context) {
            WriteLog -Message ("{0} skipped: Microsoft Graph connection did not return a context." -f $Purpose) -Level "WARNING"
            return $false
        }

        WriteLog -Message ("Microsoft Graph connected for {0}." -f $Purpose) -Level "SUCCESS"
        return $true
    }
    catch {
        WriteLog -Message ("{0} skipped: failed to connect Microsoft Graph: {1}" -f $Purpose, (Get-SmartM365ExceptionDetails -Exception $_.Exception)) -Level "WARNING"
        return $false
    }
}

function Connect-SmartM365GraphForSharePointUpload {
    [CmdletBinding()]
    param(
        [string]$AppId,
        [string]$TenantId,
        [string]$Thumbprint
    )

    $connectionKey = '{0}|{1}|{2}' -f $AppId, $TenantId, $Thumbprint
    if ($script:SmartM365SharePointUploadConnectionKey -eq $connectionKey) {
        return $true
    }

    try {
        [void](Get-SmartM365GraphAccessToken -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'SharePoint upload')
        $script:SmartM365SharePointUploadConnectionKey = $connectionKey
        return $true
    }
    catch {
        WriteLog -Message ("SharePoint upload skipped: failed to acquire Microsoft Graph REST token: {0}" -f $_.Exception.Message) -Level "WARNING"
        return $false
    }
}
function Get-SmartM365GraphRetryAfterSeconds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [int]$DefaultSeconds = 15,
        [int]$MaximumSeconds = 300
    )

    $retryAfter = $null
    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.Headers) {
            $headerValues = $ErrorRecord.Exception.Response.Headers.GetValues('Retry-After')
            $retryAfter = @($headerValues | Select-Object -First 1)[0]
        }
    }
    catch {}

    if ([string]::IsNullOrWhiteSpace([string]$retryAfter)) {
        try { $retryAfter = $ErrorRecord.Exception.Data['Retry-After'] } catch {}
    }

    $seconds = 0
    if ($retryAfter -and [int]::TryParse([string]$retryAfter, [ref]$seconds) -and $seconds -gt 0) {
        return [math]::Min($seconds, $MaximumSeconds)
    }

    return [math]::Min([math]::Max(1, $DefaultSeconds), $MaximumSeconds)
}

function Test-SmartM365GraphTransientError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [AllowNull()][Nullable[int]]$StatusCode
    )

    if ($StatusCode -in @(429, 500, 502, 503, 504)) { return $true }

    $details = [string]$ErrorRecord.Exception.Message
    try {
        if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ErrorDetails.Message)) {
            $details = $details + ' ' + [string]$ErrorRecord.ErrorDetails.Message
        }
    }
    catch {}

    try {
        $response = $ErrorRecord.Exception.Response
        if ($response -is [System.Net.Http.HttpResponseMessage] -and $null -ne $response.Content) {
            $responseBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                $details = $details + ' ' + $responseBody
            }
        }
    }
    catch {}

    if ($StatusCode -eq 409 -and $details -match '(?i)resourceModified|eTag[^\r\n]*(?:mismatch|changed)|resource[^\r\n]*changed since the caller last read it') {
        return $true
    }

    return ($details -match '(?i)TooManyRequests|throttl|timeout|temporarily unavailable')
}
function Invoke-SmartM365GraphRestWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PUT','PATCH','DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [AllowNull()]$Body = $null,
        [string]$ContentType = 'application/json',
        [int]$MaxAttempts = 4,
        [int]$DefaultRetrySeconds = 15,
        [int]$MaximumRetrySeconds = 300,
        [string]$Operation = 'Graph request'
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $token = Get-SmartM365GraphAccessToken -Purpose $Operation
            $params = @{
                Method      = $Method
                Uri         = $Uri
                Headers     = @{ Authorization = "Bearer $token" }
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) { $params.Body = $Body }
            if (-not [string]::IsNullOrWhiteSpace($ContentType)) { $params.ContentType = $ContentType }
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
            $message = [string]$_.Exception.Message
            $isTransient = Test-SmartM365GraphTransientError -ErrorRecord $_ -StatusCode $statusCode
            if (-not $isTransient -or $attempt -ge $MaxAttempts) {
                $responseBody = $null
                try {
                    if ($_.Exception.Response) {
                        $errorResponse = $_.Exception.Response
                        if ($errorResponse -is [System.Net.Http.HttpResponseMessage]) {
                            if ($errorResponse.Content) { $responseBody = $errorResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult() }
                        }
                        else {
                            $responseStream = $errorResponse.GetResponseStream()
                            if ($responseStream) {
                                $reader = New-Object System.IO.StreamReader($responseStream)
                                try { $responseBody = $reader.ReadToEnd() } finally { $reader.Dispose() }
                            }
                        }
                    }
                } catch {}
                try {
                    if ([string]::IsNullOrWhiteSpace($responseBody) -and $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)) {
                        $responseBody = [string]$_.ErrorDetails.Message
                    }
                } catch {}
                $statusText = if ($statusCode) { $statusCode } else { 'unknown' }
                if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                    throw ("{0} failed. Method={1}; Status={2}; Error={3}; Body={4}" -f $Operation, $Method, $statusText, $message, $responseBody)
                }
                throw ("{0} failed. Method={1}; Status={2}; Error={3}. No response body was returned." -f $Operation, $Method, $statusText, $message)
            }

            $delay = Get-SmartM365GraphRetryAfterSeconds -ErrorRecord $_ -DefaultSeconds ($DefaultRetrySeconds * $attempt) -MaximumSeconds $MaximumRetrySeconds
            $statusText = if ($statusCode) { $statusCode } else { 'unknown' }
            WriteLog -Message ("{0} transient failure. Status={1}; attempt {2}/{3}; retrying in {4}s." -f $Operation, $statusText, $attempt, $MaxAttempts, $delay) -Level "WARNING"
            Start-Sleep -Seconds $delay
        }
    }
}
function Invoke-SmartM365SharePointLargeFileUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [Parameter(Mandatory)][string]$DriveId,
        [Parameter(Mandatory)][string]$TargetPath,
        [int]$ChunkSizeBytes = 10485760
    )

    $fileInfo = Get-Item -LiteralPath $LocalFilePath -ErrorAction Stop
    $sessionBody = @{
        item = @{
            '@microsoft.graph.conflictBehavior' = 'replace'
        }
    } | ConvertTo-Json -Depth 5
    $sessionUri = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/createUploadSession" -f $DriveId, $TargetPath
    $session = Invoke-SmartM365GraphRestWithRetry -Method POST -Uri $sessionUri -Body $sessionBody -ContentType 'application/json' -Operation 'Create SharePoint upload session'
    if ($null -eq $session -or [string]::IsNullOrWhiteSpace([string]$session.uploadUrl)) {
        throw "SharePoint upload session did not return an uploadUrl."
    }

    $chunkMultiple = 327680
    if (($ChunkSizeBytes % $chunkMultiple) -ne 0) {
        $ChunkSizeBytes = [int]([math]::Floor($ChunkSizeBytes / $chunkMultiple) * $chunkMultiple)
        if ($ChunkSizeBytes -lt $chunkMultiple) { $ChunkSizeBytes = $chunkMultiple }
    }

    $uploadUrl = [string]$session.uploadUrl
    $lastUploadResponse = $null
    $stream = [System.IO.File]::OpenRead($LocalFilePath)
    try {
        $fileSize = [int64]$stream.Length
        $buffer = New-Object byte[] $ChunkSizeBytes
        $offset = [int64]0
        while ($offset -lt $fileSize) {
            $remaining = $fileSize - $offset
            $readSize = [int][math]::Min($ChunkSizeBytes, $remaining)
            $bytesRead = $stream.Read($buffer, 0, $readSize)
            if ($bytesRead -le 0) { break }

            $chunk = New-Object byte[] $bytesRead
            [System.Array]::Copy($buffer, 0, $chunk, 0, $bytesRead)
            $rangeEnd = $offset + $bytesRead - 1
            $uploaded = $false
            for ($attempt = 1; -not $uploaded -and $attempt -le 4; $attempt++) {
                $response = $null
                $responseBody = $null
                $statusCode = $null
                try {
                    $headers = @{ 'Content-Range' = ("bytes {0}-{1}/{2}" -f $offset, $rangeEnd, $fileSize) }
                    $putParams = @{
                        Method      = 'PUT'
                        Uri         = $uploadUrl
                        Headers     = $headers
                        ContentType = 'application/octet-stream'
                        Body        = $chunk
                        ErrorAction = 'Stop'
                    }
                    $invokeWebRequestCommand = Get-Command Invoke-WebRequest -ErrorAction Stop
                    if ($invokeWebRequestCommand.Parameters.ContainsKey('UseBasicParsing')) {
                        $putParams.UseBasicParsing = $true
                    }
                    if ($invokeWebRequestCommand.Parameters.ContainsKey('SkipHttpErrorCheck')) {
                        $putParams.SkipHttpErrorCheck = $true
                    }

                    $response = Invoke-WebRequest @putParams
                    $statusCode = [int]$response.StatusCode
                    $responseBody = [string]$response.Content

                    if ($statusCode -lt 200 -or $statusCode -gt 299) {
                        throw ("SharePoint chunk upload failed. Status={0}. Body: {1}" -f $statusCode, $responseBody)
                    }

                    if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                        try { $lastUploadResponse = $responseBody | ConvertFrom-Json -ErrorAction Stop } catch { $lastUploadResponse = $responseBody }
                    }
                    $uploaded = $true
                }
                catch {
                    try {
                        if ($_.Exception.Response) {
                            $errorResponse = $_.Exception.Response
                            if ($errorResponse -is [System.Net.Http.HttpResponseMessage]) {
                                $statusCode = [int]$errorResponse.StatusCode
                                if ($errorResponse.Content) {
                                    $responseBody = $errorResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                                }
                            }
                            else {
                                $statusCode = [int]$errorResponse.StatusCode
                                $responseStream = $errorResponse.GetResponseStream()
                                if ($responseStream) {
                                    $reader = New-Object System.IO.StreamReader($responseStream)
                                    try { $responseBody = $reader.ReadToEnd() } finally { $reader.Dispose() }
                                }
                            }
                        }
                    } catch {}
                    try {
                        if ([string]::IsNullOrWhiteSpace($responseBody) -and $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)) {
                            $responseBody = [string]$_.ErrorDetails.Message
                        }
                    } catch {}
                    $isTransient = Test-SmartM365GraphTransientError -ErrorRecord $_ -StatusCode $statusCode
                    if (-not $isTransient -or $attempt -ge 4) {
                        $statusText = if ($statusCode) { $statusCode } else { 'unknown' }
                        $rangeText = "bytes {0}-{1}/{2}" -f $offset, $rangeEnd, $fileSize
                        if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                            throw ("SharePoint chunk upload failed. Status={0}. Range={1}. Error={2}. Body={3}" -f $statusText, $rangeText, $_.Exception.Message, $responseBody)
                        }
                        throw ("SharePoint chunk upload failed. Status={0}. Range={1}. Error={2}. No response body was returned." -f $statusText, $rangeText, $_.Exception.Message)
                    }
                    $delay = Get-SmartM365GraphRetryAfterSeconds -ErrorRecord $_ -DefaultSeconds (10 * $attempt) -MaximumSeconds 300
                    $statusText = if ($statusCode) { $statusCode } else { 'unknown' }
                    WriteLog -Message ("SharePoint chunk upload transient failure. Status={0}; attempt {1}/4; retrying in {2}s." -f $statusText, $attempt, $delay) -Level "WARNING"
                    Start-Sleep -Seconds $delay
                }
                finally {
                    if ($null -ne $response -and $response -is [System.IDisposable]) { $response.Dispose() }
                }
            }

            $offset += $bytesRead
        }
    }
    finally {
        $stream.Dispose()
    }
    return $lastUploadResponse
}
function ConvertTo-SmartM365SharePointDataRootPath {
    [CmdletBinding()]
    param([AllowNull()][string]$TargetFolderPath)

    if ([string]::IsNullOrWhiteSpace($TargetFolderPath)) { return $TargetFolderPath }

    $normalized = ([string]$TargetFolderPath -replace '\\', '/').Trim('/')
    if ($normalized -match '^(?<Parent>.*?)(?:/)?CSV$') {
        $parent = [string]$matches['Parent']
        if ([string]::IsNullOrWhiteSpace($parent)) { return 'DATA' }
        return ($parent.TrimEnd('/') + '/DATA')
    }

    return $normalized
}

function Get-SmartM365SharePointRelativeFilePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LocalFilePath)

    try { $normalizedPath = [System.IO.Path]::GetFullPath($LocalFilePath) }
    catch { $normalizedPath = [string]$LocalFilePath }

    $normalizedPath = $normalizedPath -replace '/', '\'
    foreach ($rootName in @('DATA-LAST','DATA-ALL','LOG-ALL')) {
        $match = [regex]::Match($normalizedPath, ('\\' + [regex]::Escape($rootName) + '\\'), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($match.Success) {
            $relativePath = $normalizedPath.Substring($match.Index + 1)
            return (($relativePath -replace '\\', '/').Trim('/'))
        }
    }

    return [System.IO.Path]::GetFileName($LocalFilePath)
}

function Add-SmartM365SharePointUploadRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [Parameter(Mandatory)][string]$SharePointPath,
        [Parameter(Mandatory)][string]$DriveId,
        [AllowNull()]$DriveItem
    )

    $fileInfo = Get-Item -LiteralPath $LocalFilePath -ErrorAction Stop
    $record = [pscustomobject][ordered]@{
        LocalFilePath  = $fileInfo.FullName
        FileName       = $fileInfo.Name
        SharePointPath = $SharePointPath
        WebUrl         = if ($DriveItem -and $DriveItem.webUrl) { [string]$DriveItem.webUrl } else { '' }
        DriveId        = $DriveId
        ItemId         = if ($DriveItem -and $DriveItem.id) { [string]$DriveItem.id } else { '' }
        Size           = $fileInfo.Length
        UploadedAt     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
    }

    if (-not $global:SmartM365SharePointUploadedFiles) {
        $global:SmartM365SharePointUploadedFiles = New-Object System.Collections.ArrayList
    }
    [void]$global:SmartM365SharePointUploadedFiles.Add($record)
    return $record
}

function Invoke-SmartM365SharePointCsvUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [bool]$Enabled = [bool]$global:EnableSharePointUpload,
        [string]$SiteHostname = $global:SharePointSiteHostname,
        [string]$SitePath = $global:SharePointSitePath,
        [string]$LibraryDisplayName = $global:SharePointLibraryDisplayName,
        [string]$TargetFolderPath = $global:SharePointTargetFolderPath,
        [string]$AppId = $global:AppId,
        [string]$TenantId = $global:TenantId,
        [string]$Thumbprint = $(if ($global:Thumbprint) { $global:Thumbprint } else { $global:Thumb })
    )

    if (-not $Enabled) { return }
    if (-not (Test-Path -LiteralPath $LocalFilePath)) {
        WriteLog -Message "SharePoint upload skipped: local file not found: $LocalFilePath" -Level "WARNING"
        return
    }
    if ([string]::IsNullOrWhiteSpace($SiteHostname) -or [string]::IsNullOrWhiteSpace($SitePath) -or [string]::IsNullOrWhiteSpace($LibraryDisplayName) -or [string]::IsNullOrWhiteSpace($TargetFolderPath)) {
        WriteLog -Message "SharePoint upload skipped: SharePointSiteHostname, SharePointSitePath, SharePointLibraryDisplayName or SharePointTargetFolderPath is missing." -Level "WARNING"
        return
    }
    if (-not (Connect-SmartM365GraphForSharePointUpload -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint)) {
        return
    }

    try {
        if ($null -eq $script:SmartM365SharePointDriveIdCache) { $script:SmartM365SharePointDriveIdCache = @{} }

        $driveCacheKey = '{0}|{1}|{2}' -f $SiteHostname, $SitePath, $LibraryDisplayName
        if ($script:SmartM365SharePointDriveIdCache.ContainsKey($driveCacheKey)) {
            $driveId = $script:SmartM365SharePointDriveIdCache[$driveCacheKey]
        }
        else {
            $site = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $SiteHostname, $SitePath) -Operation 'Resolve SharePoint site'
            $drives = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}/drives" -f $site.id) -Operation 'Resolve SharePoint document libraries'
            $driveList = @($drives.value)
            $normalize = { param($Text) if ($null -eq $Text) { '' } else { ([string]$Text).Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{M}', '' } }
            $drive = @($driveList | Where-Object { $_.name -ieq $LibraryDisplayName } | Select-Object -First 1)[0]
            if (-not $drive) {
                $targetNorm = & $normalize $LibraryDisplayName
                $drive = @($driveList | Where-Object { (& $normalize $_.name) -ieq $targetNorm } | Select-Object -First 1)[0]
            }
            if (-not $drive) {
                $available = ($driveList | ForEach-Object { $_.name }) -join ' | '
                WriteLog -Message "SharePoint upload skipped: document library '$LibraryDisplayName' not found. Available drives: $available" -Level "WARNING"
                return
            }
            $driveId = $drive.id
            $script:SmartM365SharePointDriveIdCache[$driveCacheKey] = $driveId
        }

        $fileInfo = Get-Item -LiteralPath $LocalFilePath -ErrorAction Stop
        $targetRootPath = ConvertTo-SmartM365SharePointDataRootPath -TargetFolderPath $TargetFolderPath
        $relativeFilePath = Get-SmartM365SharePointRelativeFilePath -LocalFilePath $fileInfo.FullName
        $sharePointPath = (($targetRootPath.TrimEnd('/')) + '/' + $relativeFilePath.TrimStart('/'))
        $targetPath = ConvertTo-GraphDrivePath $sharePointPath
        $largeUploadThresholdBytes = 250MB
        $uploadedItem = $null

        if ($fileInfo.Length -gt $largeUploadThresholdBytes) {
            WriteLog -Message ("SharePoint large file upload started: {0} ({1:N1} MB)" -f $sharePointPath, ($fileInfo.Length / 1MB))
            $uploadedItem = Invoke-SmartM365SharePointLargeFileUpload -LocalFilePath $fileInfo.FullName -DriveId $driveId -TargetPath $targetPath
        }
        else {
            $bytes = [System.IO.File]::ReadAllBytes($fileInfo.FullName)
            $uri = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/content" -f $driveId, $targetPath
            $uploadedItem = Invoke-SmartM365GraphRestWithRetry -Method PUT -Uri $uri -Body $bytes -ContentType 'application/octet-stream' -Operation 'Upload SharePoint file'
        }

        if (-not $uploadedItem -or (-not $uploadedItem.webUrl -and -not $uploadedItem.id)) {
            $itemUri = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}" -f $driveId, $targetPath
            $uploadedItem = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri $itemUri -Operation 'Resolve uploaded SharePoint file'
        }

        $record = Add-SmartM365SharePointUploadRecord -LocalFilePath $fileInfo.FullName -SharePointPath $sharePointPath -DriveId $driveId -DriveItem $uploadedItem
        if (-not [string]::IsNullOrWhiteSpace($record.WebUrl)) {
            WriteLog -Message ("SharePoint file uploaded: {0} ({1})" -f $record.SharePointPath, $record.WebUrl) -Level "INFO"
        }
        else {
            WriteLog -Message ("SharePoint file uploaded: {0}" -f $record.SharePointPath) -Level "INFO"
        }
        return $record
    }
    catch {
        WriteLog -Message ("SharePoint upload failed but script continues: {0}" -f $_.Exception.Message) -Level "WARNING"
    }
}
function Invoke-SmartM365SharePointFileDownload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [string]$SharePointRelativePath = '',
        [bool]$Enabled = [bool]$global:EnableSharePointUpload,
        [string]$SiteHostname = $global:SharePointSiteHostname,
        [string]$SitePath = $global:SharePointSitePath,
        [string]$LibraryDisplayName = $global:SharePointLibraryDisplayName,
        [string]$TargetFolderPath = $global:SharePointTargetFolderPath,
        [string]$AppId = $global:AppId,
        [string]$TenantId = $global:TenantId,
        [string]$Thumbprint = $(if ($global:Thumbprint) { $global:Thumbprint } else { $global:Thumb }),
        [switch]$Force
    )

    if (-not $Force -and (Test-Path -LiteralPath $LocalFilePath -PathType Leaf)) {
        return (Get-Item -LiteralPath $LocalFilePath -ErrorAction SilentlyContinue)
    }
    if (-not $Enabled) { return $null }
    if ([string]::IsNullOrWhiteSpace($SiteHostname) -or [string]::IsNullOrWhiteSpace($SitePath) -or [string]::IsNullOrWhiteSpace($LibraryDisplayName) -or [string]::IsNullOrWhiteSpace($TargetFolderPath)) {
        WriteLog -Message "SharePoint download skipped: SharePointSiteHostname, SharePointSitePath, SharePointLibraryDisplayName or SharePointTargetFolderPath is missing." -Level "WARNING"
        return $null
    }
    if (-not (Connect-SmartM365GraphForSharePointUpload -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint)) {
        return $null
    }

    try {
        if ($null -eq $script:SmartM365SharePointDriveIdCache) { $script:SmartM365SharePointDriveIdCache = @{} }

        $driveCacheKey = '{0}|{1}|{2}' -f $SiteHostname, $SitePath, $LibraryDisplayName
        if ($script:SmartM365SharePointDriveIdCache.ContainsKey($driveCacheKey)) {
            $driveId = $script:SmartM365SharePointDriveIdCache[$driveCacheKey]
        }
        else {
            $site = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $SiteHostname, $SitePath) -Operation 'Resolve SharePoint site'
            $drives = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}/drives" -f $site.id) -Operation 'Resolve SharePoint document libraries'
            $driveList = @($drives.value)
            $normalize = { param($Text) if ($null -eq $Text) { '' } else { ([string]$Text).Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{M}', '' } }
            $drive = @($driveList | Where-Object { $_.name -ieq $LibraryDisplayName } | Select-Object -First 1)[0]
            if (-not $drive) {
                $targetNorm = & $normalize $LibraryDisplayName
                $drive = @($driveList | Where-Object { (& $normalize $_.name) -ieq $targetNorm } | Select-Object -First 1)[0]
            }
            if (-not $drive) {
                $available = ($driveList | ForEach-Object { $_.name }) -join ' | '
                WriteLog -Message "SharePoint download skipped: document library '$LibraryDisplayName' not found. Available drives: $available" -Level "WARNING"
                return $null
            }
            $driveId = $drive.id
            $script:SmartM365SharePointDriveIdCache[$driveCacheKey] = $driveId
        }

        $targetRootPath = ConvertTo-SmartM365SharePointDataRootPath -TargetFolderPath $TargetFolderPath
        $relativeFilePath = if ([string]::IsNullOrWhiteSpace($SharePointRelativePath)) {
            Get-SmartM365SharePointRelativeFilePath -LocalFilePath $LocalFilePath
        }
        else {
            (([string]$SharePointRelativePath -replace '\\', '/').Trim('/'))
        }
        if ([string]::IsNullOrWhiteSpace($relativeFilePath) -or $relativeFilePath -match '(^|/)\.\.(/|$)') {
            throw ("Invalid SharePoint relative file path: {0}" -f $SharePointRelativePath)
        }
        $sharePointPath = (($targetRootPath.TrimEnd('/')) + '/' + $relativeFilePath.TrimStart('/'))
        $targetPath = ConvertTo-GraphDrivePath $sharePointPath
        $destinationFolder = Split-Path -Path $LocalFilePath -Parent
        if (-not [string]::IsNullOrWhiteSpace($destinationFolder) -and -not (Test-Path -LiteralPath $destinationFolder)) {
            New-Item -Path $destinationFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $uri = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/content" -f $driveId, $targetPath
        Invoke-SmartM365GraphFileDownloadWithRetry -Uri $uri -OutputFilePath $LocalFilePath -Operation 'Download SharePoint file' | Out-Null
        if (Test-Path -LiteralPath $LocalFilePath -PathType Leaf) {
            $item = Get-Item -LiteralPath $LocalFilePath -ErrorAction Stop
            WriteLog -Message ("SharePoint file downloaded: {0} -> {1}" -f $sharePointPath, $item.FullName) -Level "INFO"
            return $item
        }
    }
    catch {
        WriteLog -Message ("SharePoint recovery download unavailable; script continues: {0}" -f $_.Exception.Message) -Level "INFO"
    }

    return $null
}

function Resolve-SmartM365CsvPathWithSharePointFallback {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Description = 'CSV',
        [switch]$Required
    )

    if (Test-Path -LiteralPath $Path -PathType Leaf) { return $Path }

    WriteLog -Message ("{0} not found locally, attempting SharePoint recovery: {1}" -f $Description, $Path) -Level "INFO"
    $downloaded = Invoke-SmartM365SharePointFileDownload -LocalFilePath $Path
    if ($downloaded -and (Test-Path -LiteralPath $downloaded.FullName -PathType Leaf)) { return $downloaded.FullName }

    if ($Required) { throw ("Required {0} not found locally or in SharePoint: {1}" -f $Description, $Path) }
    return $null
}

function Import-SmartM365CsvWithSharePointFallback {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Description = 'CSV',
        [string]$Delimiter = '',
        [string]$Encoding = '',
        [switch]$Required
    )

    $resolvedPath = Resolve-SmartM365CsvPathWithSharePointFallback -Path $Path -Description $Description -Required:$Required
    if ([string]::IsNullOrWhiteSpace($resolvedPath)) { return @() }

    $params = @{ LiteralPath = $resolvedPath; ErrorAction = 'Stop' }
    if (-not [string]::IsNullOrWhiteSpace($Delimiter)) { $params.Delimiter = [char]$Delimiter[0] }
    if (-not [string]::IsNullOrWhiteSpace($Encoding)) { $params.Encoding = $Encoding }
    return @(Import-Csv @params)
}
function ExportAndCopyCsv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$BaseFileName,

        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [string]$GlobalPath,

        [Parameter(Mandatory)]
        [array]$Data,

        [Parameter()]
        [switch]$NoTypeInformation = $true,

        [Parameter()]
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",

        [Parameter()]
        [char]$Delimiter,

        [Parameter()]
        [switch]$NoTenantKey
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $runBaseFileName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName $BaseFileName
    $global:csvFilePath1 = Join-Path $OutputPath "$runBaseFileName`_$timestamp.csv"
    $global:csvFilePath2 = Join-Path $OutputPath "$runBaseFileName.csv"
    $global:csvFilePath3 = Join-Path $GlobalPath "$runBaseFileName.csv"
    $Data = @(Limit-SmartM365RowsForMaxItems -Data $Data)
    if (-not $NoTenantKey) {
        $tenantCsv = Add-SmartM365TenantKeyToCsvData -Data $Data
        $Data = @($tenantCsv.Data)
    }
    if (Test-SmartM365MaxItemsMode) {
        WriteLog -Message ("MaxItems test CSV publication active. CSV paths are suffixed with {0}; standard Power BI filenames are not updated." -f (Get-SmartM365MaxItemsSuffix)) -Level 'WARNING'
    }

    Assert-SmartM365CsvDataCompleteness -Data $Data -BaseFileName $BaseFileName -TimestampedPath $global:csvFilePath1 -LatestPath $global:csvFilePath3

    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }

    try {
        $exportParams = @{
            Path     = $csvFilePath1
            Encoding = $Encoding
        }

        if ($NoTypeInformation) {
            $exportParams.NoTypeInformation = $true
        }

        if ($PSBoundParameters.ContainsKey('Delimiter')) {
            $exportParams.Delimiter = $Delimiter
        }

        $Data | Export-Csv @exportParams
        WriteLog -Message "CSV export to: $csvFilePath1"
        [void]$global:csvGeneratedPaths.Add($csvFilePath1)
    } catch {
        WriteLog -Message "Failed to export to: $csvFilePath1 - $_" -Level Error
    }

    try {
        Copy-Item -Path $csvFilePath1 -Destination $csvFilePath2 -Force
        WriteLog -Message "CSV copied to: $csvFilePath2"
        [void]$global:csvGeneratedPaths.Add($csvFilePath2)
    } catch {
        WriteLog -Message "Failed to copy to: $csvFilePath2 - $_" -Level Error
    }

    if (-not (Test-Path -Path $GlobalPath)) {
        try {
            New-Item -Path $GlobalPath -ItemType Directory -Force | Out-Null
            WriteLog -Message "Created missing directory: $GlobalPath"
        } catch {
            WriteLog -Message "Failed to create directory: $GlobalPath - $_" -Level Error
            return
        }
    }

    $maxRetries = 3
    $retryDelaySec = 10
    $attempt = 0
    $globalCopyDone = $false

    while (-not $globalCopyDone -and $attempt -lt $maxRetries) {
        $attempt++
        try {
            Copy-Item -Path $csvFilePath2 -Destination $csvFilePath3 -Force -ErrorAction Stop
            WriteLog -Message "CSV copied to global path: $csvFilePath3"
            [void]$global:csvGeneratedPaths.Add($csvFilePath3)
            $globalCopyDone = $true
        } catch {
            if ($attempt -lt $maxRetries) {
                WriteLog -Message "Copy to global path failed (attempt $attempt/$maxRetries), retrying in $retryDelaySec s... - $_" -Level Warning
                Start-Sleep -Seconds $retryDelaySec
            } else {
                WriteLog -Message "Failed to copy to global path after $maxRetries attempts: $csvFilePath3 - $_" -Level Error
            }
        }
    }

    if ($global:RetentionMaxCSV -gt 0) {
        RemoveOldFiles -FolderPath $OutputPath -FilePattern "$BaseFileName`_*.csv" -MaxFiles $global:RetentionMaxCSV
    }

    if ($globalCopyDone) {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath3
    }
    else {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath2
    }

    $historySourcePath = if ($globalCopyDone) { $csvFilePath3 } else { $csvFilePath2 }
    if (Test-SmartM365MaxItemsMode) {
        WriteLog -Message 'WeeklyHistory publication skipped because MaxItems test mode is active.' -Level 'WARNING'
    }
    else {
        Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($historySourcePath) -TimestampedPath $csvFilePath1
    }

    WriteLog -Message "CSV export completed: $csvFilePath1"
}

function ExportAndCopyCsvFromConvert {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$BaseFileName,

        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [string]$GlobalPath,

        [Parameter(Mandatory)]
        [array]$Data,

        [Parameter()]
        [switch]$NoTypeInformation = $true,

        [Parameter()]
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",

        [Parameter()]
        [string]$Delimiter = ",",

        [Parameter()]
        [switch]$NoTenantKey
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $runBaseFileName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName $BaseFileName
    $global:csvFilePath1 = Join-Path $OutputPath "$runBaseFileName`_$timestamp.csv"
    $global:csvFilePath2 = Join-Path $OutputPath "$runBaseFileName.csv"
    $global:csvFilePath3 = Join-Path $GlobalPath "$runBaseFileName.csv"
    $Data = @(Limit-SmartM365RowsForMaxItems -Data $Data)
    if (-not $NoTenantKey) {
        $tenantCsv = Add-SmartM365TenantKeyToCsvData -Data $Data
        $Data = @($tenantCsv.Data)
    }
    if (Test-SmartM365MaxItemsMode) {
        WriteLog -Message ("MaxItems test CSV publication active. CSV paths are suffixed with {0}; standard Power BI filenames are not updated." -f (Get-SmartM365MaxItemsSuffix)) -Level 'WARNING'
    }

    Assert-SmartM365CsvDataCompleteness -Data $Data -BaseFileName $BaseFileName -TimestampedPath $global:csvFilePath1 -LatestPath $global:csvFilePath3

    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }

    try {
        $csvContent = if ($NoTypeInformation) {
            $Data | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter
        } else {
            $Data | ConvertTo-Csv -Delimiter $Delimiter
        }

        try {
            $csvContent | Out-File -FilePath $csvFilePath1 -Encoding $Encoding
            WriteLog -Message "CSV exported to: $csvFilePath1"
            [void]$global:csvGeneratedPaths.Add($csvFilePath1)
        } catch {
            WriteLog -Message "Failed to export CSV to: $csvFilePath1 - $_" -Level Error
            return
        }

        try {
            Copy-Item -Path $csvFilePath1 -Destination $csvFilePath2 -Force
            WriteLog -Message "CSV copied to: $csvFilePath2"
            [void]$global:csvGeneratedPaths.Add($csvFilePath2)
        } catch {
            WriteLog -Message "Failed to copy CSV to: $csvFilePath2 - $_" -Level Error
        }

        if (-not (Test-Path -Path $GlobalPath)) {
            try {
                New-Item -Path $GlobalPath -ItemType Directory -Force | Out-Null
                WriteLog -Message "Created missing directory: $GlobalPath"
            } catch {
                WriteLog -Message "Failed to create directory: $GlobalPath - $_" -Level Error
                return
            }
        }

        $maxRetries = 3
        $retryDelaySec = 10
        $attempt = 0
        $globalCopyDone = $false

        while (-not $globalCopyDone -and $attempt -lt $maxRetries) {
            $attempt++
            try {
                Copy-Item -Path $csvFilePath1 -Destination $csvFilePath3 -Force -ErrorAction Stop
                WriteLog -Message "CSV copied to global path: $csvFilePath3"
                [void]$global:csvGeneratedPaths.Add($csvFilePath3)
                $globalCopyDone = $true
            } catch {
                if ($attempt -lt $maxRetries) {
                    WriteLog -Message "Copy to global path failed (attempt $attempt/$maxRetries), retrying in $retryDelaySec s... - $_" -Level Warning
                    Start-Sleep -Seconds $retryDelaySec
                } else {
                    WriteLog -Message "Failed to copy to global path after $maxRetries attempts: $csvFilePath3 - $_" -Level Error
                }
            }
        }

        if ($global:RetentionMaxCSV -gt 0) {
            RemoveOldFiles -FolderPath $OutputPath -FilePattern "$BaseFileName`_*.csv" -MaxFiles $global:RetentionMaxCSV
        }

        if ($globalCopyDone) {
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath3
        }
        else {
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath2
        }

        $historySourcePath = if ($globalCopyDone) { $csvFilePath3 } else { $csvFilePath2 }
        if (Test-SmartM365MaxItemsMode) {
        WriteLog -Message 'WeeklyHistory publication skipped because MaxItems test mode is active.' -Level 'WARNING'
    }
    else {
        Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($historySourcePath) -TimestampedPath $csvFilePath1
    }

    } catch {
        WriteLog -Message "Unexpected error during CSV export process: $_" -Level Error
    }
}

#endregion

#region Remote scheduled task helper

function NewRemoteScheduledTaskAndWait {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$RemoteComputerName,

        [Parameter(Mandatory)]
        [string]$TaskName,

        [Parameter(Mandatory)]
        [string]$ScriptPathOnShare,

        [Parameter()]
        [string]$TriggerTime,

        [Parameter()]
        [string]$KeyFileName = "",

        [Parameter()]
        [string]$CredentialUserName = "DOMAIN\svc_account"
    )

    if (-not $TriggerTime) {
        $TriggerTime = (Get-Date).AddMinutes(1).ToString("HH:mm")
        Write-Log "TriggerTime not provided. Defaulting to $TriggerTime"
    }

    try {
        if (-not $PSBoundParameters.ContainsKey('CredentialUserName')) {
            $moduleLocalConfig = Get-ModuleLocalConfig
            $CredentialUserName = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'CredentialUserName' -DefaultValue $CredentialUserName
        }

        Write-Log "Loading credentials from key file '$KeyFileName'"

        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        $keyFilePath = Join-Path $scriptDir $KeyFileName

        if (!(Test-Path $keyFilePath)) {
            throw "Key file '$keyFilePath' not found."
        }

        $securePassword = Get-Content $keyFilePath | ConvertTo-SecureString
        $credential = New-Object System.Management.Automation.PSCredential ($CredentialUserName, $securePassword)

        Write-Log "Starting task creation on $RemoteComputerName"

        $username = $credential.UserName
        $password = $credential.GetNetworkCredential().Password
        $escapedScriptPath = $ScriptPathOnShare.Replace('"', '""')
        $command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$escapedScriptPath`""
        $taskNameEscaped = "`"$TaskName`""

        $createArgs = "/Create /S $RemoteComputerName /TN $taskNameEscaped /TR `"$command`" /SC ONCE /ST $TriggerTime /RU `"$username`" /RP `"$password`" /RL HIGHEST /F"
        $runArgs    = "/Run /S $RemoteComputerName /TN $taskNameEscaped"
        $queryArgs  = "/Query /S $RemoteComputerName /TN $taskNameEscaped /FO LIST /V"

        Write-Log "Creating task with schtasks.exe"
        $create = Start-Process -FilePath "schtasks.exe" -ArgumentList $createArgs -NoNewWindow -Wait -PassThru
        if ($create.ExitCode -ne 0) {
            throw "Failed to create task. ExitCode: $($create.ExitCode)"
        }

        Write-Log "Running task"
        $run = Start-Process -FilePath "schtasks.exe" -ArgumentList $runArgs -NoNewWindow -Wait -PassThru
        if ($run.ExitCode -ne 0) {
            throw "Failed to start task. ExitCode: $($run.ExitCode)"
        }

        Write-Log "Waiting for task to complete..."
        $maxWait = 600
        $elapsed = 0
        do {
            Start-Sleep -Seconds 5
            $elapsed += 5
            $output = & schtasks.exe $queryArgs
            $lastRunResult = ($output | Where-Object { $_ -like "Last Run Result*" }) -replace "Last Run Result:\s*", ""
        } while ($lastRunResult -eq "0x41301" -and $elapsed -lt $maxWait)

        Write-Log "Task completed with result: $lastRunResult"

        $success = $lastRunResult -eq "0x0"
        if ($success) {
            Write-Log "Task executed successfully on $RemoteComputerName"
        } else {
            Write-Log "Task failed with code $lastRunResult" "ERROR"
        }

        return @{
            Computer  = $RemoteComputerName
            TaskName  = $TaskName
            ResultCode = $lastRunResult
            Success   = $success
            LogFile   = $global:LogTextFile
        }
    }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
        return @{
            Computer = $RemoteComputerName
            TaskName = $TaskName
            Success  = $false
            Error    = $_.Exception.Message
            LogFile  = $global:LogTextFile
        }
    }
}

#endregion

#region Cloud services connections (EXO + Graph)

function ConvertFrom-SmartM365JwtPayload {
    [CmdletBinding()]
    param([string]$AccessToken)

    if ([string]::IsNullOrWhiteSpace($AccessToken)) { return $null }
    $parts = $AccessToken.Split('.')
    if ($parts.Count -lt 2) { return $null }
    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
        1 { return $null }
    }
    try {
        $json = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
        return $json | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Get-SmartM365GraphGrantedPermissionValues {
    [CmdletBinding()]
    param([string]$GraphAccessToken)

    $granted = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    if (-not [string]::IsNullOrWhiteSpace($GraphAccessToken)) {
        $claims = ConvertFrom-SmartM365JwtPayload -AccessToken $GraphAccessToken
        foreach ($role in @($claims.roles)) { if (-not [string]::IsNullOrWhiteSpace([string]$role)) { [void]$granted.Add([string]$role) } }
        foreach ($scope in @(([string]$claims.scp) -split '\s+')) { if (-not [string]::IsNullOrWhiteSpace($scope)) { [void]$granted.Add($scope) } }
        return @($granted)
    }

    $context = $null
    try { $context = Get-MgContext -ErrorAction SilentlyContinue } catch { $context = $null }
    if ($null -eq $context) { return @($granted) }

    foreach ($scope in @($context.Scopes)) { if (-not [string]::IsNullOrWhiteSpace([string]$scope)) { [void]$granted.Add([string]$scope) } }

    $clientId = [string]$context.ClientId
    if ([string]::IsNullOrWhiteSpace($clientId)) { return @($granted) }

    try {
        $escapedClientId = $clientId.Replace("'", "''")
        $appResponse = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '{0}'&`$select=id,appId,displayName" -f $escapedClientId) -OutputType PSObject -ErrorAction Stop
        $appServicePrincipal = @($appResponse.value) | Select-Object -First 1
        if ($null -eq $appServicePrincipal -or [string]::IsNullOrWhiteSpace([string]$appServicePrincipal.id)) { return @($granted) }

        $graphResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles" -OutputType PSObject -ErrorAction Stop
        $graphServicePrincipal = @($graphResponse.value) | Select-Object -First 1
        if ($null -eq $graphServicePrincipal -or [string]::IsNullOrWhiteSpace([string]$graphServicePrincipal.id)) { return @($granted) }

        $roleLookup = @{}
        foreach ($role in @($graphServicePrincipal.appRoles)) {
            if ($null -ne $role.id -and -not [string]::IsNullOrWhiteSpace([string]$role.value)) {
                $roleLookup[[string]$role.id] = [string]$role.value
            }
        }

        $assignments = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/servicePrincipals/{0}/appRoleAssignments?`$top=999" -f $appServicePrincipal.id) -OutputType PSObject -ErrorAction Stop
        foreach ($assignment in @($assignments.value)) {
            if ([string]$assignment.resourceId -ne [string]$graphServicePrincipal.id) { continue }
            $roleId = [string]$assignment.appRoleId
            if ($roleLookup.ContainsKey($roleId)) { [void]$granted.Add($roleLookup[$roleId]) }
        }
    }
    catch {
        WriteLog -Message ("Preflight Graph app role lookup could not be completed: {0}" -f $_.Exception.Message) -Level "WARNING"
    }

    return @($granted)
}
function Invoke-SmartM365Preflight {
    [CmdletBinding()]
    param(
        [string]$ScriptName = $MyInvocation.MyCommand.Name,
        [string[]]$RequiredModules = @(),
        [switch]$SkipMissingModuleInstall,
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$InstallModuleScope = 'CurrentUser',
        [string]$InstallModuleRepository = 'PSGallery',
        [string[]]$RequiredCommands = @(),
        [string[]]$OutputPaths = @(),
        [string[]]$GraphProbeUris = @(),
        [string[]]$RequiredGraphApplicationPermissions = @(),
        [string]$GraphAccessToken,
        [string[]]$ExchangeOnlineProbeCommands = @(),
        [switch]$RequireActiveDirectoryRead,
        [switch]$RequireExchangeOnPrem,
        [switch]$SkipOutputPathCreation
    )

    $failures = New-Object System.Collections.Generic.List[string]
    WriteLog -Message ("Preflight started for {0}" -f $ScriptName) -Level "INFO"

    foreach ($moduleName in @($RequiredModules | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $availableModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        if ($null -eq $availableModule) {
            $nonInstallableModules = @('ActiveDirectory')
            if ($SkipMissingModuleInstall -or $moduleName -in $nonInstallableModules) {
                $failures.Add("Required PowerShell module not found: $moduleName")
                continue
            }
            if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) {
                $failures.Add("Required PowerShell module not found and Install-Module is unavailable: $moduleName")
                continue
            }
            try {
                WriteLog -Message ("Preflight module missing; installing {0} from {1} with Scope={2}." -f $moduleName, $InstallModuleRepository, $InstallModuleScope) -Level "WARNING"
                Install-Module -Name $moduleName -Scope $InstallModuleScope -Repository $InstallModuleRepository -Force -AllowClobber -ErrorAction Stop
                $availableModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
                if ($null -eq $availableModule) { throw "Install-Module completed but module was not found in module paths." }
                WriteLog -Message ("Preflight module installed: {0} {1}; Path={2}" -f $availableModule.Name, $availableModule.Version, $availableModule.Path) -Level "SUCCESS"
            }
            catch {
                $failures.Add("Required PowerShell module '$moduleName' could not be installed: $($_.Exception.Message)")
                continue
            }
        }

        try {
            Import-Module -Name $moduleName -ErrorAction Stop
            WriteLog -Message ("Preflight module OK: {0}" -f (Get-SmartM365ModuleDiagnosticText -Name $moduleName -IncludeAvailable)) -Level "INFO"
        }
        catch {
            $failures.Add("Required PowerShell module '$moduleName' could not be imported: $($_.Exception.Message)")
        }
    }
    foreach ($commandName in @($RequiredCommands | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            $failures.Add("Required PowerShell command not found: $commandName")
        }
        else {
            WriteLog -Message ("Preflight command OK: {0}" -f $commandName) -Level "INFO"
        }
    }

    foreach ($path in @($OutputPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        try {
            if (-not (Test-Path -LiteralPath $path)) {
                if ($SkipOutputPathCreation) {
                    $failures.Add("Output path does not exist: $path")
                    continue
                }
                $null = New-Item -Path $path -ItemType Directory -Force -ErrorAction Stop
            }

            $probePath = Join-Path -Path $path -ChildPath (".smartm365-preflight-{0}.tmp" -f ([guid]::NewGuid().Guid))
            Set-Content -LiteralPath $probePath -Value "SmartM365 preflight" -Encoding UTF8 -ErrorAction Stop
            Remove-Item -LiteralPath $probePath -Force -ErrorAction Stop
            WriteLog -Message ("Preflight output path writable: {0}" -f $path) -Level "INFO"
        }
        catch {
            $failures.Add("Output path is not writable: $path. $($_.Exception.Message)")
        }
    }

    if ($RequireActiveDirectoryRead) {
        try {
            if (-not (Get-Command -Name Get-ADForest -ErrorAction SilentlyContinue)) {
                throw "Get-ADForest is not available."
            }
            $null = Get-ADForest -ErrorAction Stop
            WriteLog -Message "Preflight Active Directory read OK: Get-ADForest" -Level "INFO"
        }
        catch {
            $failures.Add("Active Directory read preflight failed: $($_.Exception.Message)")
        }
    }

    if ($RequireExchangeOnPrem) {
        try {
            if (Get-Command -Name EnsureExchangePSSnapinLoaded -ErrorAction SilentlyContinue) {
                if (-not (EnsureExchangePSSnapinLoaded)) {
                    throw "Exchange on-premises PowerShell snap-in is not ready."
                }
            }
            if (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue) {
                $null = Get-ExchangeServer -ErrorAction Stop | Select-Object -First 1
            }
            else {
                throw "Get-ExchangeServer is not available."
            }
            WriteLog -Message "Preflight Exchange on-premises read OK: Get-ExchangeServer" -Level "INFO"
        }
        catch {
            $failures.Add("Exchange on-premises preflight failed: $($_.Exception.Message)")
        }
    }

    $requiredGraphPermissions = @($RequiredGraphApplicationPermissions | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($requiredGraphPermissions.Count -gt 0) {
        WriteLog -Message ("Preflight minimum Graph application permissions: {0}" -f ($requiredGraphPermissions -join ', ')) -Level "INFO"
        try {
            $grantedGraphPermissions = @(Get-SmartM365GraphGrantedPermissionValues -GraphAccessToken $GraphAccessToken)
            foreach ($permissionName in $requiredGraphPermissions) {
                if ($grantedGraphPermissions -notcontains $permissionName) {
                    $failures.Add("Missing Microsoft Graph application permission: $permissionName")
                }
                else {
                    WriteLog -Message ("Preflight Graph permission declared and granted: {0}" -f $permissionName) -Level "INFO"
                }
            }
        }
        catch {
            $failures.Add("Graph application permission validation failed: $($_.Exception.Message)")
        }
    }
    if ($GraphProbeUris.Count -gt 0) {
        $useAccessToken = -not [string]::IsNullOrWhiteSpace($GraphAccessToken)
        if (-not $useAccessToken) {
            try {
                if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
                    throw "Invoke-MgGraphRequest is not available."
                }
                $context = Get-MgContext -ErrorAction SilentlyContinue
                if ($null -eq $context) {
                    throw "Microsoft Graph is not connected."
                }
            }
            catch {
                $failures.Add("Microsoft Graph connection preflight failed: $($_.Exception.Message)")
            }
        }

        if ($useAccessToken -or $null -ne (Get-MgContext -ErrorAction SilentlyContinue)) {
            foreach ($uri in $GraphProbeUris) {
                if ([string]::IsNullOrWhiteSpace($uri)) { continue }
                try {
                    if ($useAccessToken) {
                        $headers = @{ Authorization = "Bearer $GraphAccessToken" }
                        $null = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ErrorAction Stop
                    }
                    else {
                        $null = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                    }
                    WriteLog -Message ("Preflight Graph permission OK: {0}" -f $uri) -Level "INFO"
                }
                catch {
                    $failures.Add("Graph permission probe failed for '$uri': $($_.Exception.Message)")
                }
            }
        }
    }

    foreach ($commandName in @($ExchangeOnlineProbeCommands | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        try {
            if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
                throw "Command is not available."
            }
            switch ($commandName) {
                "Get-Mailbox" { $null = & $commandName -ResultSize 1 -ErrorAction Stop }
                "Get-EXOMailbox" { $null = & $commandName -ResultSize 1 -ErrorAction Stop }
                "Get-MigrationBatch" { $null = & $commandName -ResultSize 1 -ErrorAction Stop }
                default { $null = & $commandName -ErrorAction Stop | Select-Object -First 1 }
            }
            WriteLog -Message ("Preflight Exchange Online RBAC OK: {0}" -f $commandName) -Level "INFO"
        }
        catch {
            $failures.Add("Exchange Online RBAC probe failed for '$commandName': $($_.Exception.Message)")
        }
    }

    if ($failures.Count -gt 0) {
        foreach ($failure in $failures) {
            WriteLog -Message ("Preflight failed: {0}" -f $failure) -Level "ERROR"
        }
        throw ("SmartM365 preflight failed for {0}: {1}" -f $ScriptName, ($failures -join " | "))
    }

    WriteLog -Message ("Preflight completed for {0}" -f $ScriptName) -Level "SUCCESS"
    return $true
}

function Connect-SmartM365CloudSession {
    param (
        [string]$AppId,
        [string]$Thumbprint,
        [string]$TenantId,
        [string]$Organization,
        [bool]$ExchangeOnline = $true,
        [bool]$Graph = $true,
        [string[]]$GraphScopes = @("User.Read.All", "Directory.Read.All")
    )

    $exchangeSuccess = $false
    $graphSuccess = $false

    $useCertAuth = $AppId -and $Thumbprint -and $TenantId

    if (-not $Organization -and $TenantId -and $TenantId.Contains(".")) {
        $Organization = $TenantId
    }

    if ($useCertAuth) {
        try {
            $thumbUpper = $Thumbprint.ToUpper()
            $cert = @("Cert:\CurrentUser\My","Cert:\LocalMachine\My") |
                ForEach-Object { Get-ChildItem $_ -ErrorAction SilentlyContinue } |
                Where-Object { $_.Thumbprint -eq $thumbUpper } |
                Select-Object -First 1

            if (-not $cert) {
                WriteLog "Certificate not found in CurrentUser\My or LocalMachine\My ($Thumbprint)." "ERROR"
                return @{
                    ExchangeOnlineConnected = $false
                    GraphConnected          = $false
                }
            }

            if (-not $cert.HasPrivateKey) {
                WriteLog "Certificate found but WITHOUT private key. Thumbprint: $Thumbprint" "ERROR"
                return @{
                    ExchangeOnlineConnected = $false
                    GraphConnected          = $false
                }
            }

            WriteLog ("Certificate resolved: {0} | {1}" -f $cert.Subject, $cert.Thumbprint) "SUCCESS"
        } catch {
            WriteLog "Error while resolving certificate: $($_.Exception.Message)" "ERROR"
            return @{
                ExchangeOnlineConnected = $false
                GraphConnected          = $false
            }
        }
    }

    if ($Graph) {
        try {
            if (-not (Get-Module -Name Microsoft.Graph.Authentication)) {
                if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
                    throw "Required module 'Microsoft.Graph.Authentication' is not installed."
                }
                WriteLog "Loading Graph Authentication module: Microsoft.Graph.Authentication..." "INFO"
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            }

            $scopeToModuleMap = @{
                "User.Read.All"         = "Microsoft.Graph.Users"
                "Device.Read.All"       = "Microsoft.Graph.Devices"
                "Group.Read.All"        = "Microsoft.Graph.Groups"
                "Reports.Read.All"      = "Microsoft.Graph.Reports"
                "Policy.Read.All"       = "Microsoft.Graph.Policies"
                "Organization.Read.All" = "Microsoft.Graph.Organization"
            }

            $modulesToImport = $GraphScopes | ForEach-Object {
                if ($scopeToModuleMap.ContainsKey($_)) {
                    $scopeToModuleMap[$_]
                }
            } | Select-Object -Unique

            foreach ($module in $modulesToImport) {
                if (-not (Get-Module -Name $module)) {
                    $availableModule = Get-Module -ListAvailable -Name $module | Sort-Object Version -Descending | Select-Object -First 1
                    if (-not $availableModule) {
                        throw "Required Graph submodule '$module' is not installed."
                    }
                    WriteLog "Loading Graph submodule: $module..." "INFO"
                    Import-Module $availableModule.Path -ErrorAction Stop
                    WriteLog -Message ("Loaded module version: {0}" -f (Get-SmartM365ModuleDiagnosticText -Name $module)) -Level "INFO"
                }
            }
        } catch {
            WriteLog "Failed to load Microsoft.Graph modules: $($_.Exception.Message)" "ERROR"
            return @{
                ExchangeOnlineConnected = $false
                GraphConnected          = $false
            }
        }
    }

    if ($ExchangeOnline) {
        try {
            if (-not (Get-Module -Name ExchangeOnlineManagement)) {
                if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                    throw "Required module 'ExchangeOnlineManagement' is not installed."
                }
                WriteLog "Loading ExchangeOnlineManagement module..." "INFO"
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                WriteLog -Message ("Loaded module version: {0}" -f (Get-SmartM365ModuleDiagnosticText -Name 'ExchangeOnlineManagement')) -Level "INFO"
            }
        } catch {
            WriteLog "Failed to load ExchangeOnlineManagement module: $($_.Exception.Message)" "ERROR"
            return @{
                ExchangeOnlineConnected = $false
                GraphConnected          = $graphSuccess
            }
        }
    }

    if ($Graph) {
        try {
            WriteLog "Connecting to Microsoft Graph..." "INFO"
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

            if ($useCertAuth) {
                Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop
            } else {
                Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
            }

            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -eq $context) {
                throw "Microsoft Graph connection did not return a context."
            }

            WriteLog "Microsoft Graph connection successful." "SUCCESS"
            $graphSuccess = $true
        } catch {
            WriteLog "Failed to connect to Microsoft Graph: $(Get-SmartM365ExceptionDetails -Exception $_.Exception)" "ERROR"
        }
    }

    if ($ExchangeOnline) {
        try {
            WriteLog "Connecting to Exchange Online..." "INFO"
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}

            if ($useCertAuth) {
                if (-not $Organization) {
                    WriteLog "No Organization specified for Exchange Online app-only connection." "ERROR"
                    return @{
                        ExchangeOnlineConnected = $false
                        GraphConnected          = $graphSuccess
                    }
                }

                Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $Organization -ShowBanner:$false
            } else {
                Connect-ExchangeOnline -ShowProgress:$true -ShowBanner:$false
            }

            WriteLog "Exchange Online connection successful." "SUCCESS"
            $exchangeSuccess = $true
        } catch {
            WriteLog "Failed to connect to Exchange Online: $($_.Exception.Message)" "ERROR"
        }
    }

    return @{
        ExchangeOnlineConnected = $exchangeSuccess
        GraphConnected          = $graphSuccess
    }
}

function Disconnect-SmartM365CloudSession {
    param (
        [bool]$ExchangeOnline = $true,
        [bool]$Graph = $true,
        [bool]$VerboseDisconnect = $false
    )

    if ($ExchangeOnline) {
        try {
            if ($VerboseDisconnect) { WriteLog "Disconnecting from Exchange Online..." "INFO" }
            Disconnect-ExchangeOnline -Confirm:$false | Out-Null
            if ($VerboseDisconnect) { WriteLog "Disconnected from Exchange Online." "SUCCESS" }
        } catch {
            if ($VerboseDisconnect) { WriteLog "Failed to disconnect from Exchange Online: $($_.Exception.Message)" "WARNING" }
        }
    }

    if ($Graph) {
        try {
            if ($VerboseDisconnect) { WriteLog "Disconnecting from Microsoft Graph..." "INFO" }
            if ($null -eq (Get-MgContext -ErrorAction SilentlyContinue)) {
                if ($VerboseDisconnect) { WriteLog "No active Microsoft Graph session to disconnect." "INFO" }
                return
            }
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            if ($VerboseDisconnect) { WriteLog "Disconnected from Microsoft Graph." "SUCCESS" }
        } catch {
            if ($VerboseDisconnect) { WriteLog "Failed to disconnect from Microsoft Graph: $($_.Exception.Message)" "WARNING" }
        }
    }
}

#endregion

Export-ModuleMember -Function `
    Format-SmartM365LogLine, Update-SmartM365TimestampedTranscript, WriteLog, Write-Log, Get-SmartM365ModuleDiagnosticText, Write-SmartM365LoadedModuleVersions, Write-SmartM365ExecutionContext, Write-SmartM365CompletionBanner, Complete-SmartM365ExecutionContext, Test-FileLocked, RemoveOldFiles, Remove-OldFiles, EnsureExchangePSSnapinLoaded, `
    Set-SmartM365CoreContext, Get-SmartM365MaxItemsValue, Test-SmartM365MaxItemsMode, Get-SmartM365MaxItemsSuffix, Set-SmartM365MaxItemsMode, Add-SmartM365MaxItemsSuffixToCsvPath, Add-SmartM365MaxItemsSuffixToBaseName, Add-SmartM365MaxItemsMailBanner, Add-SmartM365MaxItemsSubjectPrefix, Limit-SmartM365RowsForMaxItems, Get-SmartM365CsvValidationBaseName, Get-SmartM365CsvValidationRule, Assert-SmartM365CsvDataCompleteness, Add-SmartM365CsvValidationRule, Initialize-SmartM365DefaultCsvValidationRules, Add-SmartM365TenantKey, Repair-SmartM365CsvTenantKeySchema, Write-SmartM365CsvAtomically, Publish-SmartM365Csv, Export-SmartM365Csv, Export-SmartM365CsvFromConvert, `
    ConvertTo-SmartM365ConfigBoolean, Get-SmartM365MailBrandingConfig, ConvertTo-SmartM365MailLogoDataUri, Add-SmartM365MailBranding, ConvertToRecipientArray, ConvertTo-SmartM365EmailHtmlText, New-SmartM365EmailBody, ConvertTo-SmartM365EmailBody, Convert-SmartM365MailBodyLocalPathsToSharePointLinks, NewSimpleEmailBody, ConvertBytesToSizeString, GetFileList, `
    NewTableEmailBody, NewTableFilesEmailBody, SendEmailHtmlReport, Send-SmartM365Mail, Send-SmartM365GraphMail, SendFileListEmailReport, Send-SmartM365TeamsNotification, `
    TestSharePath, InitializeScriptEnvironment, Connect-SmartM365GraphAppOnly, ConvertTo-SmartM365SharePointDataRootPath, Get-SmartM365SharePointRelativeFilePath, Invoke-SmartM365SharePointCsvUpload, Invoke-SmartM365SharePointFileDownload, Resolve-SmartM365CsvPathWithSharePointFallback, Import-SmartM365CsvWithSharePointFallback, `
    ExportAndCopyCsv, ExportAndCopyCsvFromConvert, Save-SmartM365WeeklyInventoryHistory, Add-SmartM365WeeklyHistory, `
    NewRemoteScheduledTaskAndWait, `
    Invoke-SmartM365Preflight, Connect-SmartM365CloudSession, Disconnect-SmartM365CloudSession


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAvW8NL8aJ7c3o8
# HW707hVKrwBAbiI8JNL+Og7G6KNuxaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIMNu6+zknppHB9oH108TW5ToNZnVBy6xzZqeq/9+kMrXMA0GCSqG
# SIb3DQEBAQUABIIBgCE2SnaJGigCTGVCFxHIawWnDyoi4y5wM1O5VsMsDNtojmJR
# WajnaKBx+111cuglwH+/dXun3t8U3Tmjfe4+Eii66xOUBagSBO51QFqJieQ6bp5X
# dAT0SVTxczD8o2y8swRpFuWs4t6boPM0sFqqLdqZ/yWph0SFN7Na2/1ugAO+2ykN
# 4UJpHNBMW6j1Wv9p1A+zQzM8/lpAGII6I74z6gq+gRll28+8am//Ptjr+1PQcu8h
# jHFLIXnuyoB73MtXqhsqUByPz2mybBCBBwDl7rHDHHD64tiZZot6J7oi2r0xjxsQ
# J8lBK1E1zLs6RxlIGS1x+Z9KqZmG+PbP55UjTBgZdxj6/h9CizFx5mSmk4rLFa/O
# Tb+6XAi7kD2w8dLCuCtJcU60tFt0Os1gQGULySwLuqjmZTwmRvnfrPORScMPQepS
# w0+Or48X5CFTzaXH5sCxrA6G562Ful8FMp9XMNm2f+BWA3T54QOFVGdofAD6UV/b
# My8zwY5TBOj76YwpDaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQyMTQx
# MjRaMC8GCSqGSIb3DQEJBDEiBCCvRfDaSm85tbQLhdBOq1uoSHQPkWN5UY6ldqq3
# 3NeajTANBgkqhkiG9w0BAQEFAASCAgA8Nb2Ako+NJrC1XZNMcIymSOh7NvK6S4/i
# gsxIQ3whImuU3XB2EnupsKMscedNj+P1yRdz2aDMXTUbnyZr0UU9F/+LSw8k1wss
# RgekQArwTWB+Lirpx8Yz6z7XLUqVIAByz4AGCtUhKEYhT3ciyu+tkirf2eZVJo25
# nbIaqesiSC8pz09w+aG4Ur/E2hcM0AfM4CNBzXksOQkexnCNj5aqHdlYHdE0I6Rq
# 9YeyS00Mfw2dK4U0Uv8qadEg/6y1GpSgvRJuIw2sAyij46cbk8U9fpN05QmGzwvD
# 0OmSLgObBDv3+0Ne/t9maUti+wvun/UO9cmW2zjp0cPkiJ0DnYdxdo2IwIYyxMJH
# q9RitrNJUarpW4SRNt/L/hw6yI+nV+RTvnA/+tqtEikNwSzpzWVZCXpCjo8i7X+/
# pZ8iKT6xyXpzKh4BzdjU1y3uBp1vHQquieOBKLzULil3atuMVMUpp3Zw2Ioj3d6m
# 0ded5MOs94SGviC9+mzyEVLwauuS61774V90IUlpKitHZRDI2hP1EI7uKg1EpD2K
# kzoTrbIU3KFpi5crvv3X/NALGt2cShdrjiAdd5tBOuMkhDRlc7PRY4dvjf8umYri
# acbvl8yCOvOcA7Iulk11MY+5/oPh3OFzoGaUAam/l0fFqioP0uza9ScTaz1Un6yT
# b4fbCnQBFA==
# SIG # End signature block
