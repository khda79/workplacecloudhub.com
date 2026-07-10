
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
    Write-Host 'Edit the local JSON now if needed. Press Enter to continue with the current file values.' -ForegroundColor Yellow
    Read-Host 'Press Enter to continue' | Out-Null
    return $true
}
function Get-SmartM365EffectiveModuleGlobalConfig {
    [CmdletBinding()]
    param()

    if ($null -ne $script:SmartM365GlobalConfig) {
        return $script:SmartM365GlobalConfig
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
                $tenantKey = if ([string]::IsNullOrWhiteSpace([string]$global:SmartM365Tenant)) { 'test' } else { [string]$global:SmartM365Tenant }
                $script:SmartM365GlobalConfig = Get-SmartM365EffectiveGlobalConfig -StartPath $searchRoot -TenantKey $tenantKey
                break
            }
        }
        if ($null -ne $script:SmartM365GlobalConfig -and $script:SmartM365GlobalConfig.PSObject.Properties.Count -gt 0) {
            break
        }

        $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'

        $globalTemplatePath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json.template'

        if (-not (Test-Path -LiteralPath $globalConfigPath)) {

            Initialize-SmartM365ModuleGlobalConfigFromTemplate -Path $globalConfigPath -TemplatePath $globalTemplatePath | Out-Null

        }

        if (Test-Path -LiteralPath $globalConfigPath) {
            try {
                $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            }
            catch {
                throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
            }
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
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) {
            return $DefaultValue
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
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllLines($Path, [string[]]$updatedLines, $utf8NoBom)
    }
}

function WriteLog {
    param(
        [string]$Message,
        [string]$Level = ""
    )

    # Automatically detect error keywords if no level is provided
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

    Write-Host $bannerLines[0] -ForegroundColor DarkCyan
    Write-Host $bannerLines[1] -ForegroundColor Cyan
    Write-Host $bannerLines[2] -ForegroundColor Yellow
    Write-Host $bannerLines[3] -ForegroundColor Yellow
    Write-Host $bannerLines[4] -ForegroundColor DarkCyan

    if ($global:LogTextFile) {
        Add-Content -Path $global:LogTextFile -Value $bannerLines
    }
}
function Write-SmartM365ExecutionContext {
    [CmdletBinding()]
    param(
        [string]$ScriptName = $global:SmartM365ScriptName,
        [string]$ScriptVersion = '',
        [string]$TenantKey = $global:SmartM365Tenant,
        [string]$OutputPath = $global:BasePath,
        [string]$ScriptPath = ''
    )

    $now = Get-Date
    $context = [ordered]@{}

    if (-not [string]::IsNullOrWhiteSpace($ScriptName)) { $context['ScriptName'] = $ScriptName }
    if (-not [string]::IsNullOrWhiteSpace($ScriptVersion)) { $context['ScriptVersion'] = $ScriptVersion }
    if (-not [string]::IsNullOrWhiteSpace($TenantKey)) { $context['TenantKey'] = $TenantKey }

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

    $summary = [ordered]@{
        Status            = $Status
        Duration          = $durationText
        Started           = $started.ToString('yyyy-MM-dd HH:mm:ss zzz')
        Ended             = $ended.ToString('yyyy-MM-dd HH:mm:ss zzz')
        ScriptName        = $global:SmartM365ScriptName
        TenantKey         = $global:SmartM365Tenant
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

    # Upload run log files after the execution summary is written, so SharePoint keeps the final log content.
    $logUploadCandidates = @($global:LogTextFile, $global:logTranscriptFile) |
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
    }
    finally {
        $script:SmartM365TeamsNotificationInProgress = $false
    }
}

function Send-SmartM365TeamsNotification {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$WebhookUrl = "",
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message,
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
                $factList.Add(@{ name = [string]$key; value = [string]$value }) | Out-Null
            }
        }

        $payload = @{
            '@type'    = 'MessageCard'
            '@context' = 'https://schema.org/extensions'
            summary    = $Title
            themeColor = $themeColors[$Level]
            title      = $Title
            text       = $Message
            sections   = @(@{ markdown = $true; facts = @($factList) })
        }

        if (-not [string]::IsNullOrWhiteSpace($HelpUrl)) {
            $payload['potentialAction'] = @(
                @{
                    '@type' = 'OpenUri'
                    name    = 'Get AI help'
                    targets = @(@{ os = 'default'; uri = $HelpUrl })
                }
            )
        }

        $json = $payload | ConvertTo-Json -Depth 10
        Invoke-RestMethod -Method POST -Uri $WebhookUrl -Body $json -ContentType 'application/json' -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        if ($ThrowOnError) { throw }
        WriteLog -Message ("Teams notification failed: {0}" -f $_.Exception.Message) -Level 'WARNING'
        return $false
    }
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

    # Normalize the exclusion list as full paths, case-insensitive
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

    # Ajouter les fichiers de log globaux
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
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SKIP] Locked file: $($file.Name)" | Tee-Object -FilePath $LogFile -Append | Out-Null
            continue
        }

        try {
            if ($file.Attributes -band [IO.FileAttributes]::ReadOnly) {
                $file.Attributes = $file.Attributes -bxor [IO.FileAttributes]::ReadOnly
            }
        } catch {
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SKIP] Cannot clear ReadOnly on $($file.Name): $($_.Exception.Message)" | Tee-Object -FilePath $LogFile -Append | Out-Null
            continue
        }

        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Deleted old file: $($file.Name)" | Tee-Object -FilePath $LogFile -Append | Out-Null
        } catch {
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SKIP] Failed to delete $($file.Name): $($_.Exception.Message)" | Tee-Object -FilePath $LogFile -Append | Out-Null
        }
    }
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
function ConvertToRecipientArray {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Recipients  # Recipients string, supports ';' or ',' separators
    )
    return @($Recipients) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }
}

function Test-SmartM365ConfiguredValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string]) { return ($null -ne $Value) }

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
        return ($null -ne (Get-MgContext -ErrorAction SilentlyContinue))
    }
    catch {
        WriteLog -Message ("{0} skipped: failed to connect Microsoft Graph: {1}" -f $Purpose, (Get-SmartM365ExceptionDetails -Exception $_.Exception)) -Level "WARNING"
        return $false
    }
}

function ConvertTo-SmartM365GraphRecipient {
    [CmdletBinding()]
    param([string[]]$Recipients)

    @($Recipients | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        @{ emailAddress = @{ address = $_.Trim() } }
    })
}

function ConvertTo-SmartM365GraphFileAttachment {
    [CmdletBinding()]
    param([string[]]$Attachments)

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
        [string]$AppId = $global:AppId,
        [string]$TenantId = $global:TenantId,
        [string]$Thumbprint = $global:Thumbprint
    )

    if (-not (Test-SmartM365ConfiguredValue -Value $Thumbprint)) { $Thumbprint = $global:Thumb }

    $toArray = ConvertToRecipientArray -Recipients $To
    $ccArray = if ($Cc) { ConvertToRecipientArray -Recipients $Cc } else { @() }
    if (-not $toArray -or [string]::IsNullOrWhiteSpace($From)) {
        throw "Send-SmartM365GraphMail: missing required parameters (From/To)."
    }

    if (-not (Connect-SmartM365GraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'Graph mail')) {
        throw "Send-SmartM365GraphMail: Microsoft Graph app-only connection failed."
    }

    $BodyHtml = ConvertTo-SmartM365EmailBody -BodyHtml $BodyHtml -Subject $Subject -Category 'SmartM365'

    $message = @{
        subject      = $Subject
        body         = @{ contentType = 'HTML'; content = $BodyHtml }
        toRecipients = ConvertTo-SmartM365GraphRecipient -Recipients $toArray
    }
    if ($ccArray.Count -gt 0) { $message['ccRecipients'] = ConvertTo-SmartM365GraphRecipient -Recipients $ccArray }
    $graphAttachments = ConvertTo-SmartM365GraphFileAttachment -Attachments $Attachments
    if ($graphAttachments.Count -gt 0) { $message['attachments'] = $graphAttachments }

    $body = @{ message = $message; saveToSentItems = $false } | ConvertTo-Json -Depth 12
    $encodedFrom = [System.Uri]::EscapeDataString($From)
    Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $encodedFrom) -Body $body -ContentType 'application/json' | Out-Null
    WriteLog -Message ("Graph mail sent from {0} to {1}" -f $From, ($toArray -join ';')) -Level "SUCCESS"
}

function ConvertTo-SmartM365EmailHtmlText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
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
        return New-SmartM365EmailBody -Title $effectiveTitle -Category $Category -Severity $Severity
    }

    if ($BodyHtml -match 'SmartM365EmailTemplate:v1') {
        return $BodyHtml
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

    return New-SmartM365EmailBody -Title $effectiveTitle -Category $Category -Severity $Severity -BodyHtml $wrappedBody
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
    <#
        Returns a summary of files (counts, total size, errors) and details for each file.
        Each file is analyzed and the summary + file list is returned.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Files,  # File paths, folders, or wildcard patterns
        [switch]$Recurse,                        # If a folder is passed, recurse into it
        [switch]$ComputeHash                     # (optional) compute SHA256 hashes (slower)
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
                'FullName'      = $item.FullName   # <-- This property is required for CSV row count
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
        [int]$SmtpPort = 25,                                 # Default SMTP port
        [ValidateSet("Graph","SMTP","Both")]
        [string]$SendMailMode = "",
        [string]$From = "",
        [string]$To = "",
        [string]$Cc = "",                                    # Default CC (empty) team@example.com
        [string]$Subject = "SmartM365 Report",                         # Default subject
        [string]$BodyHtml,                                   # HTML body (required)
        [string[]]$Attachments,                              # Attachments
        [switch]$VerboseLog                                  # Enable verbose logging
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

        # Recipients
        $toArray = ConvertToRecipientArray -Recipients $To
        $ccArray = if ($Cc) { ConvertToRecipientArray -Recipients $Cc } else { @() }

        if (-not $toArray -or [string]::IsNullOrWhiteSpace($From)) {
            throw "SendEmailHtmlReport: missing required parameters (From/To)."
        }

        # Dynamic subject: from <title> if Subject is empty, else default
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

        # Filter attachments that exist
        $BodyHtml = ConvertTo-SmartM365EmailBody -BodyHtml $BodyHtml -Subject $Subject -Category 'SmartM365'
        $atts = @()
        foreach ($a in ($Attachments | Where-Object { $_ })) {
            if (Test-Path $a) { $atts += $a }
        }

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
            Send-SmartM365GraphMail -From $From -To ($toArray -join ';') -Cc ($ccArray -join ';') -Subject $Subject -BodyHtml $BodyHtml -Attachments $atts
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
            Send-SmartM365GraphMail -From $From -To ($toArray -join ';') -Cc ($ccArray -join ';') -Subject $Subject -BodyHtml $BodyHtml -Attachments $atts
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

function SendFileListEmailReport {
    <#
        If -Files is provided and not empty, calls GetFileList, builds the HTML report, and sends it.
        If -Files is null or empty, sends only the message as the email body.
        Parameters:
            -Files      : array of files/folders/patterns (optional)
            -Title      : email subject/title
            -Message    : optional message at the bottom
            -Recurse    : optional, recursive search
            -ComputeHash: optional, SHA256 hash
    #>
    [CmdletBinding()]
    param(
        [string[]]$Files,
        [Parameter(Mandatory)][string]$Title,
        [string]$Message,
        [switch]$Recurse,
        [switch]$ComputeHash
    )

    if (-not $Files -or $Files.Count -eq 0) {
        # No files: send only the message
        $body = NewSimpleEmailBody -Title $Title -Message $Message
        SendEmailHtmlReport -Subject $Title -BodyHtml $body
        return
    }

    # Inventory
    $summary = GetFileList -Files $Files -Recurse:$Recurse -ComputeHash:$ComputeHash

    # Prepare summary hashtable (excluding the Files property)
    $summaryData = @{
        "Files (count)" = $summary.'Files (count)'
        "Total size"    = $summary.'Total size'
        "Errors"        = $summary.Errors
    }

    # Generate HTML body
    $body = NewTableFilesEmailBody -Title $Title `
        -SummaryData $summaryData `
        -Files $summary.Files `
        -Message $Message

    # Send the email
    SendEmailHtmlReport -BodyHtml $body
}

function TestSharePath {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        WriteLog -Message "The share '$Path' is not available. Stopping the script." -Level "ERROR"
        throw "The share '$Path' is not available."
    }
}

function Invoke-SmartM365Preflight {
    [CmdletBinding()]
    param(
        [string]$ScriptName = $MyInvocation.MyCommand.Name,
        [string[]]$RequiredModules = @(),
        [string[]]$RequiredCommands = @(),
        [string[]]$OutputPaths = @(),
        [string[]]$GraphProbeUris = @(),
        [string]$GraphAccessToken,
        [string[]]$ExchangeOnlineProbeCommands = @(),
        [switch]$RequireActiveDirectoryRead,
        [switch]$RequireExchangeOnPrem,
        [switch]$SkipOutputPathCreation
    )

    $failures = New-Object System.Collections.Generic.List[string]
    WriteLog -Message ("Preflight started for {0}" -f $ScriptName) -Level "INFO"

    foreach ($moduleName in @($RequiredModules | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $failures.Add("Required PowerShell module not found: $moduleName")
        }
        else {
            WriteLog -Message ("Preflight module OK: {0}" -f (Get-SmartM365ModuleDiagnosticText -Name $moduleName -IncludeAvailable)) -Level "INFO"
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

function InitializeScriptEnvironment {
    param(
        [string]$OutputPathInit,
        [string]$OutputPath,
		[string]$LogFileName
    )

    $callerScriptPath = ''
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

    if ([string]::IsNullOrWhiteSpace($OutputPathInit) -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPathInit = $OutputPath
    }

    if ([string]::IsNullOrWhiteSpace($OutputPathInit)) {
        WriteLog -Message "InitializeScriptEnvironment requires OutputPathInit/OutputPath from the script local.json." -Level "ERROR"
        throw "InitializeScriptEnvironment requires OutputPathInit/OutputPath from the script local.json."
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
        WriteLog -Message "InitializeScriptEnvironment requires LogAllRootPath from Config\SmartM365.global.local.json." -Level "ERROR"
        throw "InitializeScriptEnvironment requires LogAllRootPath from Config\SmartM365.global.local.json."
    }
    $global:SmartM365ExecutionStartTime = Get-Date
    $global:SmartM365ExecutionSummaryWritten = $false
    $global:SmartM365WarningCount = 0
    $global:SmartM365ErrorCount = 0
    $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $global:BasePath = $OutputPathInit
	$global:SmartM365ScriptName = $LogFileName
	$global:LogPath = Join-Path -Path $logAllRootPath -ChildPath $LogFileName

	try {
		if (-not (Test-Path -Path $global:LogPath)) {
			New-Item -Path $global:LogPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
		}
	} catch {
		Write-Error "Failed to create the log directory at path '$global:LogPath'. Error details: $_"
		throw "Script execution stopped due to failure in creating the log directory."
	}

    $global:LogTextFile = Join-Path $global:LogPath "$LogFileName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
    $global:logTranscriptFile = Join-Path $global:LogPath "$LogFileName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_Transcript.log"
    if ($global:RetentionMaxLogs -gt 0) {
        RemoveOldFiles -FolderPath $global:LogPath -FilePattern "$LogFileName*.log" -MaxFiles $global:RetentionMaxLogs
    }
    $scriptVersion = Get-SmartM365ScriptVersionFromFile -Path $callerScriptPath
    Write-SmartM365ExecutionContext -ScriptName $LogFileName -ScriptVersion $scriptVersion -OutputPath $OutputPathInit -ScriptPath $callerScriptPath
    Write-host  "Environment initialized successfully."

    return $OutputPathInit
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
        [char]$Delimiter
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $global:csvFilePath1 = Join-Path $OutputPath "$BaseFileName`_$timestamp.csv"
    $global:csvFilePath2 = Join-Path $OutputPath "$BaseFileName.csv"
    $global:csvFilePath3 = Join-Path $GlobalPath "$BaseFileName.csv"

    # Init global exclusion set
    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }

    # Export with timestamp
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

    # Copy to fixed name in output path
    try {
        Copy-Item -Path $csvFilePath1 -Destination $csvFilePath2 -Force
        WriteLog -Message "CSV copied to: $csvFilePath2"
        [void]$global:csvGeneratedPaths.Add($csvFilePath2)
    } catch {
        WriteLog -Message "Failed to copy to: $csvFilePath2 - $_" -Level Error
    }

    # Create global directory if it doesn't exist
    if (-not (Test-Path -Path $GlobalPath)) {
        try {
            New-Item -Path $GlobalPath -ItemType Directory -Force | Out-Null
            WriteLog -Message "Created missing directory: $GlobalPath"
        } catch {
            WriteLog -Message "Failed to create directory: $GlobalPath - $_" -Level Error
            return
        }
    }

    # Copy to global path
    try {
        Copy-Item -Path $csvFilePath2 -Destination $csvFilePath3 -Force
        WriteLog -Message "CSV copied to global path: $csvFilePath3"
        [void]$global:csvGeneratedPaths.Add($csvFilePath3)
    } catch {
        WriteLog -Message "Failed to copy to global path: $csvFilePath3 - $_" -Level Error
    }

    if ($global:RetentionMaxCSV -gt 0) {
        RemoveOldFiles -FolderPath $OutputPath -FilePattern "$BaseFileName`_*.csv" -MaxFiles $global:RetentionMaxCSV
    }

    if (Test-Path -LiteralPath $csvFilePath3) {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath3
    }
    else {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath2
    }

    $historySourcePath = if (Test-Path -LiteralPath $csvFilePath3 -PathType Leaf) { $csvFilePath3 } else { $csvFilePath2 }
    Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($historySourcePath) -TimestampedPath $csvFilePath1

    # Final log
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
        [string]$Delimiter = ","
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $global:csvFilePath1 = Join-Path $OutputPath "$BaseFileName`_$timestamp.csv"
    $global:csvFilePath2 = Join-Path $OutputPath "$BaseFileName.csv"
    $global:csvFilePath3 = Join-Path $GlobalPath "$BaseFileName.csv"

    # Init global exclusion set
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

        try {
            Copy-Item -Path $csvFilePath1 -Destination $csvFilePath3 -Force
            WriteLog -Message "CSV copied to global path: $csvFilePath3"
            [void]$global:csvGeneratedPaths.Add($csvFilePath3)
        } catch {
            WriteLog -Message "Failed to copy CSV to global path: $csvFilePath3 - $_" -Level Error
        }

        if ($global:RetentionMaxCSV -gt 0) {
            RemoveOldFiles -FolderPath $OutputPath -FilePattern "$BaseFileName`_*.csv" -MaxFiles $global:RetentionMaxCSV
        }

        if (Test-Path -LiteralPath $csvFilePath3) {
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath3
        }
        else {
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvFilePath2
        }

        $historySourcePath = if (Test-Path -LiteralPath $csvFilePath3 -PathType Leaf) { $csvFilePath3 } else { $csvFilePath2 }
        Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($historySourcePath) -TimestampedPath $csvFilePath1

    } catch {
        WriteLog -Message "Unexpected error during CSV export process: $_" -Level Error
    }
}
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

    # Default TriggerTime to current time + 1 minute if not provided
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
        $runArgs = "/Run /S $RemoteComputerName /TN $taskNameEscaped"
        $queryArgs = "/Query /S $RemoteComputerName /TN $taskNameEscaped /FO LIST /V"

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
            Computer = $RemoteComputerName
            TaskName = $TaskName
            ResultCode = $lastRunResult
            Success = $success
            LogFile = $global:LogTextFile
        }
    }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
        return @{
            Computer = $RemoteComputerName
            TaskName = $TaskName
            Success = $false
            Error = $_.Exception.Message
            LogFile = $global:LogTextFile
        }
    }
}



#region SmartM365 cloud session compatibility
function Connect-SmartM365CloudSession {
    param (
        [string]$AppId,
        [string]$Thumbprint,
        [string]$TenantId,
        [bool]$ExchangeOnline = $true,
        [bool]$Graph = $true,
        [string[]]$GraphScopes = @("User.Read.All", "Directory.Read.All")
    )
    $exchangeSuccess = $false
    $graphSuccess = $false

	#$AppId     = "00000000-0000-0000-0000-000000000000"
	#$TenantId  = "00000000-0000-0000-0000-000000000000"
	#$Thumb     = "0000000000000000000000000000000000000000"
	#$OrgDomain = "contoso.onmicrosoft.com"
	#Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumb -NoWelcome
	#Connect-ExchangeOnline -AppId $AppId -Organization $OrgDomain -CertificateThumbprint $Thumb -ShowBanner:$false
	
    $useCertAuth = $AppId -and $Thumbprint -and $TenantId

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
                }
            }
        } catch {
            WriteLog "Failed to load Microsoft.Graph submodules: $($_.Exception.Message)" "ERROR"
            return
        }
    }
	
	if ($ExchangeOnline) {
		try {
			if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
				WriteLog "Loading ExchangeOnlineManagement module..." "INFO"
				Import-Module ExchangeOnlineManagement -ErrorAction Stop
			}
		} catch {
			WriteLog "Failed to load ExchangeOnlineManagement module: $($_.Exception.Message)" "ERROR"
			return
		}
    }

    if ($Graph) {
        try {
            WriteLog "Connecting to Microsoft Graph..." "INFO"
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
            if ($useCertAuth) {
                Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
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
                Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $TenantId
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
        GraphConnected = $graphSuccess
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
            Disconnect-ExchangeOnline -Confirm:$false
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

#region SmartM365 SharePoint CSV upload compatibility

function ConvertTo-GraphDrivePath {
    param([Parameter(Mandatory = $true)][string]$Path)

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

function Connect-SmartM365GraphForSharePointUpload {
    [CmdletBinding()]
    param(
        [string]$AppId,
        [string]$TenantId,
        [string]$Thumbprint
    )

    try {
        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        }

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $context) {
            return $true
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
            WriteLog -Message "SharePoint upload skipped: Graph app-only connection values are missing (AppId, TenantId, Thumb/Thumbprint)." -Level "WARNING"
            return $false
        }

        WriteLog -Message "Connecting to Microsoft Graph for SharePoint upload." -Level "INFO"
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context) {
            WriteLog -Message "SharePoint upload skipped: Microsoft Graph connection did not return a context." -Level "WARNING"
            return $false
        }

        WriteLog -Message "Microsoft Graph connected for SharePoint upload." -Level "SUCCESS"
        return $true
    }
    catch {
        WriteLog -Message ("SharePoint upload skipped: failed to connect Microsoft Graph: {0}" -f $_.Exception.Message) -Level "WARNING"
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
            $params = @{
                Method      = $Method
                Uri         = $Uri
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) { $params.Body = $Body }
            if (-not [string]::IsNullOrWhiteSpace($ContentType)) { $params.ContentType = $ContentType }
            return Invoke-MgGraphRequest @params
        }
        catch {
            $statusCode = $null
            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
            $message = [string]$_.Exception.Message
            $isTransient = $statusCode -in @(429, 500, 502, 503, 504) -or $message -match 'TooManyRequests|throttl|timeout|temporarily unavailable'
            if (-not $isTransient -or $attempt -ge $MaxAttempts) {
                $responseBody = $null
                try {
                    if ($_.Exception.Response) {
                        $errorResponse = $_.Exception.Response
                        if ($errorResponse -is [System.Net.Http.HttpResponseMessage]) {
                            if ($errorResponse.Content) {
                                $responseBody = $errorResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                            }
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
                    $isTransient = $statusCode -in @(429, 500, 502, 503, 504) -or $_.Exception.Message -match 'timeout|temporarily unavailable|throttl'
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
            WriteLog -Message ("SharePoint file uploaded: {0} ({1})" -f $record.SharePointPath, $record.WebUrl)
        }
        else {
            WriteLog -Message ("SharePoint file uploaded: {0}" -f $record.SharePointPath)
        }
        return $record
    }
    catch {
        WriteLog -Message ("SharePoint upload failed but script continues: {0}" -f $_.Exception.Message) -Level "WARNING"
    }
}
#endregion

#region SmartM365 cleanup compatibility
function Remove-OldFiles {
    param (
        [string]$Path,
        [string]$Filter = "*.log",
        [int]$KeepCount = 10,
        [string[]]$ExcludeFiles = @(),
        [string]$LogFile
    )

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

function Set-SmartM365CoreContext {
    [CmdletBinding()]
    param(
        [string]$RunId,
        [string]$RunOutputRoot,
        [string]$LatestOutputRoot,
        [string]$LogPath,
        [int]$RetentionMaxCsv = 30,
        [int]$RetentionMaxLogs = 30
    )

    $script:SmartM365CoreRunId = $RunId
    $script:SmartM365CoreRunOutputRoot = $RunOutputRoot
    $script:SmartM365CoreLatestOutputRoot = $LatestOutputRoot
    $script:SmartM365CoreLogPath = $LogPath
    $script:SmartM365CoreRetentionMaxCsv = $RetentionMaxCsv
    $script:SmartM365CoreRetentionMaxLogs = $RetentionMaxLogs
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) { $global:LogTextFile = $LogPath }
}

function Write-SmartM365CsvAtomically {
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Data,
        [Parameter(Mandatory)][string]$Path,
        [string[]]$Columns = @(),
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",
        [string]$Delimiter = ","
    )

    $parent = Split-Path -Path $Path -Parent
    if ([string]::IsNullOrWhiteSpace($parent)) { $parent = (Get-Location).Path }
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    $leaf = Split-Path -Path $Path -Leaf
    $tempPath = Join-Path -Path $parent -ChildPath ("{0}.tmp.{1}" -f $leaf, [guid]::NewGuid().ToString("N"))

    try {
        $rows = @($Data)
        if ($rows.Count -eq 0 -and $Columns.Count -gt 0) {
            $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join $Delimiter
            Set-Content -LiteralPath $tempPath -Value $header -Encoding $Encoding -ErrorAction Stop
        }
        elseif ($Columns.Count -gt 0) {
            $rows | Select-Object -Property $Columns | Export-Csv -Path $tempPath -NoTypeInformation -Encoding $Encoding -Delimiter $Delimiter -ErrorAction Stop
        }
        else {
            $rows | Export-Csv -Path $tempPath -NoTypeInformation -Encoding $Encoding -Delimiter $Delimiter -ErrorAction Stop
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
        [switch]$NoSharePointUpload
    )

    Write-SmartM365CsvAtomically -Data $Data -Path $TimestampedPath -Columns $Columns -Encoding $Encoding -Delimiter $Delimiter
    WriteLog -Message ("CSV exported to: {0}" -f $TimestampedPath)

    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }
    [void]$global:csvGeneratedPaths.Add($TimestampedPath)

    $publishedPath = $TimestampedPath
    if (-not [string]::IsNullOrWhiteSpace($LatestPath)) {
        Write-SmartM365CsvAtomically -Data $Data -Path $LatestPath -Columns $Columns -Encoding $Encoding -Delimiter $Delimiter
        WriteLog -Message ("CSV latest copy written to: {0}" -f $LatestPath)
        [void]$global:csvGeneratedPaths.Add($LatestPath)
        $publishedPath = $LatestPath
    }

    if ($RetentionMaxCsv -lt 0 -and $global:RetentionMaxCSV) {
        $RetentionMaxCsv = [int]$global:RetentionMaxCSV
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

    Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($publishedPath) -TimestampedPath $TimestampedPath

    return New-Object psobject -Property @{
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
        [switch]$NoSharePointUpload
    )

    if ($PSCmdlet.ParameterSetName -eq 'ByBaseName') {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $TimestampedPath = Join-Path $OutputPath "$BaseFileName`_$timestamp.csv"
        $LatestPath = Join-Path $GlobalPath "$BaseFileName.csv"
    }

    Publish-SmartM365Csv -Data $Data -TimestampedPath $TimestampedPath -LatestPath $LatestPath -Columns $Columns -Encoding $Encoding -Delimiter $Delimiter -NoSharePointUpload:$NoSharePointUpload
}

#endregion
