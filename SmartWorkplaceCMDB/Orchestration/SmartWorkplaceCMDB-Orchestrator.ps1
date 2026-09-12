<#
.SYNOPSIS
Runs autonomous SmartWorkplaceCMDB collection and curation pipelines.

.DESCRIPTION
Coordinates source collection, normalization, relationship consolidation,
data-quality publication, tenant/date dimensions, contract build, and the local
HTML report. The default mode is read-only validation. Live collection requires
the explicit -Collect switch. Offline fixture runs never connect to a tenant.

.VERSION
1.1.5
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
    [string]$OrganizationKey,
    [string]$EnvironmentKey,
    [string]$TenantKey,
    [string]$TenantId,
    [string]$DataRootPath,
    [string]$DataAllRootPath,
    [string]$LatestOutputRootPath,
    [string]$LogRootPath,
    [string]$GlobalConfigPath,
    [string]$TenantConfigPath,
    [ValidateSet(
        'Full',
        'EntraUsers',
        'EntraGroups',
        'EntraDevices',
        'TenantIdentityHealth',
        'IntuneDevices',
        'M365SubscribedSkus',
        'M365UserLicenses',
        'ExchangeOnlineMailboxes',
        'ActiveDirectory',
        'CuratedOnly'
    )]
    [string]$Pipeline = 'Full',
    [switch]$Collect,
    [switch]$ValidateOnly,
    [switch]$ValidateExistingOutputs,
    [string]$FixtureRootPath,
    [ValidateRange(0, 2147483647)]
    [int]$MaxItems = 0,
    [switch]$NoConfigWrite,
    [switch]$DisableSharePointUpload
)

$ScriptVersion = '1.1.5'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Add-SmartWorkplaceCMDBBannerLog {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Path,
        [Parameter(Mandatory)][string[]]$Lines
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $encoding = New-Object System.Text.UTF8Encoding($false)
    foreach ($line in $Lines) {
        [IO.File]::AppendAllText(
            $Path,
            ([string]$line + [Environment]::NewLine),
            $encoding
        )
    }
}

function Write-SmartWorkplaceCMDBStartupBanner {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$LogPath = '',
        [switch]$NoConsole
    )

    $lines = @(
        ('=' * 80),
        ' SmartWorkplaceCMDB by WorkplaceCloudHub',
        ' Website : https://workplacecloudhub.com',
        ' GitHub  : https://github.com/khda79/workplacecloudhub.com',
        ('=' * 80)
    )
    if (-not $NoConsole) {
        Microsoft.PowerShell.Utility\Write-Host $lines[0] -ForegroundColor DarkCyan
        Microsoft.PowerShell.Utility\Write-Host $lines[1] -ForegroundColor Cyan
        Microsoft.PowerShell.Utility\Write-Host $lines[2] -ForegroundColor Yellow
        Microsoft.PowerShell.Utility\Write-Host $lines[3] -ForegroundColor Yellow
        Microsoft.PowerShell.Utility\Write-Host $lines[4] -ForegroundColor DarkCyan
    }
    Add-SmartWorkplaceCMDBBannerLog -Path $LogPath -Lines $lines
}

function Write-SmartWorkplaceCMDBConsole {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'OUTPUT')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'DEBUG' { 'DarkGray' }
        'OUTPUT' { 'Gray' }
        default { 'Gray' }
    }
    $lines = @([string]$Message -split "`r?`n")
    if ($lines.Count -eq 0) { $lines = @('') }
    foreach ($line in $lines) {
        Microsoft.PowerShell.Utility\Write-Host (
            '[{0}] {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $line
        ) -ForegroundColor $color
    }
}

function Write-SmartWorkplaceCMDBCompletionBanner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][datetimeoffset]$StartedDateTime,
        [AllowEmptyString()][string]$Pipeline = '',
        [AllowEmptyString()][string]$Mode = '',
        [int]$StepCount = 0,
        [int]$WarningCount = 0,
        [int]$ErrorCount = 0,
        [AllowEmptyString()][string]$LogPath = '',
        [switch]$NoConsole
    )

    $endedDateTime = [datetimeoffset]::Now
    $duration = $endedDateTime - $StartedDateTime
    $lines = @(
        ('=' * 80),
        ' SmartWorkplaceCMDB execution summary',
        (' Status   : {0}' -f $Status),
        (' Pipeline : {0}' -f $Pipeline),
        (' Mode     : {0}' -f $Mode),
        (' Duration : {0}' -f $duration.ToString('hh\:mm\:ss')),
        (' Steps    : {0}' -f $StepCount),
        (' Warnings : {0}' -f $WarningCount),
        (' Errors   : {0}' -f $ErrorCount),
        (' Log      : {0}' -f $(if ([string]::IsNullOrWhiteSpace($LogPath)) { 'not created' } else { $LogPath })),
        ('=' * 80)
    )
    if (-not $NoConsole) {
        $statusColor = if ($ErrorCount -gt 0 -or $Status -eq 'Failed') {
            'Red'
        }
        elseif ($WarningCount -gt 0 -or $Status -eq 'CompletedWithWarnings') {
            'Yellow'
        }
        else { 'Green' }
        Microsoft.PowerShell.Utility\Write-Host $lines[0] -ForegroundColor DarkCyan
        Microsoft.PowerShell.Utility\Write-Host $lines[1] -ForegroundColor Cyan
        foreach ($line in $lines[2..9]) {
            Microsoft.PowerShell.Utility\Write-Host $line -ForegroundColor $statusColor
        }
        Microsoft.PowerShell.Utility\Write-Host $lines[10] -ForegroundColor DarkCyan
    }
    Add-SmartWorkplaceCMDBBannerLog -Path $LogPath -Lines $lines
}

function Add-SmartWorkplaceCMDBOrchestratorStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$List,
        [Parameter(Mandatory)][hashtable]$Catalog,
        [Parameter(Mandatory)][string[]]$Names
    )
    foreach ($name in $Names) {
        if (-not $Catalog.ContainsKey($name)) {
            throw "Unknown orchestrator step '$name'."
        }
        $List.Add($Catalog[$name])
    }
}

function Get-SmartWorkplaceCMDBStepParameter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Step,
        [Parameter(Mandatory)][hashtable]$CommonParameters,
        [Parameter(Mandatory)][string]$Mode,
        [string]$FixtureRootPath,
        [int]$MaxItems
    )
    $parameters = @{} + $CommonParameters
    if ($Mode -eq 'Validate') {
        $parameters['ValidateOnly'] = $true
    }
    if ($Step.Kind -eq 'Collect') {
        if ($Mode -eq 'Collect' -and
            $Step.PSObject.Properties['RequiresExplicitCollect'] -and
            [bool]$Step.RequiresExplicitCollect) {
            $parameters['Collect'] = $true
        }
        if ($Mode -eq 'Fixture' -or
            ($Mode -eq 'Validate' -and
                -not [string]::IsNullOrWhiteSpace($FixtureRootPath))) {
            $parameters['InputJsonPath'] = Join-Path `
                $FixtureRootPath `
                $Step.FixtureName
        }
        if ($Mode -ne 'Validate' -and $MaxItems -gt 0) {
            $parameters['MaxItems'] = $MaxItems
        }
    }
    return $parameters
}

function Get-SmartWorkplaceCMDBOrchestratorSetting {
    [CmdletBinding()]
    param(
        [AllowNull()]$Configuration,
        [Parameter(Mandatory)][string]$Name,
        $DefaultValue
    )

    if ($null -eq $Configuration) {
        return $DefaultValue
    }
    if ($Configuration -is [System.Collections.IDictionary] -and
        $Configuration.Contains($Name)) {
        return $Configuration[$Name]
    }
    $property = $Configuration.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }
    return $DefaultValue
}

function Get-SmartWorkplaceCMDBCsvSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$RootPath)

    $snapshot = @{}
    foreach ($root in $RootPath) {
        if ([string]::IsNullOrWhiteSpace($root) -or
            -not (Test-Path -LiteralPath $root -PathType Container)) {
            continue
        }
        foreach ($file in @(Get-ChildItem -LiteralPath $root -Filter '*.csv' -File -Recurse)) {
            $snapshot[$file.FullName] = '{0}|{1}' -f
                $file.Length,
                $file.LastWriteTimeUtc.Ticks
        }
    }
    return $snapshot
}

function Write-SmartWorkplaceCMDBOrchestratorLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Results,
        [Parameter(Mandatory)][string]$Path
    )
    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    @($Results.ToArray()) |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function ConvertTo-SmartWorkplaceCMDBLogName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)

    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $escaped = [regex]::Escape((-join $invalid))
    $name = [regex]::Replace($Value, "[$escaped]", '-')
    $name = [regex]::Replace($name, '[^A-Za-z0-9._-]', '-')
    return $name.Trim('-')
}

function Write-SmartWorkplaceCMDBTextLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowEmptyString()][string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'OUTPUT')]
        [string]$Level = 'INFO'
    )

    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $lines = @([string]$Message -split "`r?`n")
    if ($lines.Count -eq 0) { $lines = @('') }
    foreach ($line in $lines) {
        $entry = '[{0}] [{1}] {2}{3}' -f
            (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'),
            $Level,
            $line,
            [Environment]::NewLine
        [IO.File]::AppendAllText($Path, $entry, $encoding)
    }
}

function Invoke-SmartWorkplaceCMDBLogRetention {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$Filter,
        [ValidateRange(0, 36500)][int]$RetentionDays,
        [ValidateRange(0, 100000)][int]$MaxFiles,
        [switch]$Recurse,
        [string[]]$ExcludePath = @()
    )

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        return @()
    }
    $excluded = New-Object 'System.Collections.Generic.HashSet[string]' `
        ([StringComparer]::OrdinalIgnoreCase)
    foreach ($path in @($ExcludePath)) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            [void]$excluded.Add([IO.Path]::GetFullPath($path))
        }
    }
    $removed = New-Object System.Collections.Generic.List[string]
    $files = @(Get-ChildItem -LiteralPath $FolderPath -Filter $Filter -File `
            -Recurse:$Recurse -ErrorAction SilentlyContinue)
    if ($RetentionDays -gt 0) {
        $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
        foreach ($file in @($files | Where-Object {
                    $_.LastWriteTime -lt $cutoff -and
                    -not $excluded.Contains($_.FullName)
                })) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                $removed.Add("age:$($file.FullName)")
            }
            catch {
                $removed.Add("failed:$($file.FullName):$($_.Exception.Message)")
            }
        }
    }
    if ($MaxFiles -gt 0) {
        $remaining = @(Get-ChildItem -LiteralPath $FolderPath -Filter $Filter -File `
                -Recurse:$Recurse -ErrorAction SilentlyContinue)
        foreach ($group in @($remaining | Group-Object DirectoryName)) {
            $overflow = @($group.Group |
                Sort-Object LastWriteTimeUtc, Name -Descending |
                Select-Object -Skip $MaxFiles)
            foreach ($file in $overflow) {
                if ($excluded.Contains($file.FullName)) { continue }
                try {
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    $removed.Add("count:$($file.FullName)")
                }
                catch {
                    $removed.Add("failed:$($file.FullName):$($_.Exception.Message)")
                }
            }
        }
    }
    return @($removed.ToArray())
}

function Invoke-SmartWorkplaceCMDBLoggedStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Step,
        [Parameter(Mandatory)][hashtable]$Parameters,
        [AllowEmptyString()][string]$LogPath = '',
        [AllowEmptyString()][string]$TranscriptPath = ''
    )

    $transcriptStarted = $false
    if (-not [string]::IsNullOrWhiteSpace($TranscriptPath)) {
        $transcriptFolder = Split-Path -Parent $TranscriptPath
        if (-not (Test-Path -LiteralPath $transcriptFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $transcriptFolder -Force | Out-Null
        }
        Start-Transcript -LiteralPath $TranscriptPath -Force -ErrorAction Stop | Out-Null
        $transcriptStarted = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        Write-SmartWorkplaceCMDBTextLog -Path $LogPath -Message (
            "Started step '{0}'. Script='{1}'." -f $Step.Name, $Step.ScriptPath)
    }
    $previousParentRunId = [string]$env:SMARTWORKPLACECMDB_PARENT_RUN_ID
    $env:SMARTWORKPLACECMDB_PARENT_RUN_ID = [string]$script:SmartWorkplaceCMDBCurrentRunId
    try {
        if ($transcriptStarted) {
            Write-SmartWorkplaceCMDBConsole -Message (
                "Started step '{0}'. Script='{1}'." -f $Step.Name, $Step.ScriptPath)
        }
        & $Step.ScriptPath @Parameters *>&1 | ForEach-Object {
            $record = $_
            $level = if ($record -is [System.Management.Automation.ErrorRecord]) {
                'ERROR'
            }
            elseif ($record -is [System.Management.Automation.WarningRecord]) {
                'WARN'
            }
            elseif ($record -is [System.Management.Automation.DebugRecord]) {
                'DEBUG'
            }
            elseif ($record -is [System.Management.Automation.VerboseRecord]) {
                'DEBUG'
            }
            elseif ($record -is [System.Management.Automation.InformationRecord]) {
                'INFO'
            }
            else {
                'OUTPUT'
            }
            $text = if ($record -is [System.Management.Automation.InformationRecord]) {
                [string]$record.MessageData
            }
            else {
                [string]$record
            }
            if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
                Write-SmartWorkplaceCMDBTextLog -Path $LogPath `
                    -Message $text -Level $level
            }
            $displayRecord = $record -is [System.Management.Automation.ErrorRecord] -or
                $record -is [System.Management.Automation.WarningRecord] -or
                $record -is [System.Management.Automation.DebugRecord] -or
                $record -is [System.Management.Automation.VerboseRecord] -or
                $record -is [System.Management.Automation.InformationRecord] -or
                $record -is [string]
            if ($displayRecord -and -not [string]::IsNullOrWhiteSpace($text)) {
                Write-SmartWorkplaceCMDBConsole -Message $text -Level $level
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
            Write-SmartWorkplaceCMDBTextLog -Path $LogPath -Message (
                "Completed step '{0}'." -f $Step.Name)
        }
        if ($transcriptStarted) {
            Write-SmartWorkplaceCMDBConsole -Message (
                "Completed step '{0}'." -f $Step.Name)
        }
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
            Write-SmartWorkplaceCMDBTextLog -Path $LogPath -Message (
                "Failed step '{0}': {1}" -f $Step.Name, $_.Exception.Message) `
                -Level ERROR
        }
        if ($transcriptStarted) {
            Write-SmartWorkplaceCMDBConsole -Message (
                "Failed step '{0}': {1}" -f $Step.Name, $_.Exception.Message) `
                -Level ERROR
        }
        throw
    }
    finally {
        $env:SMARTWORKPLACECMDB_PARENT_RUN_ID = $previousParentRunId
        if ($transcriptStarted) {
            Stop-Transcript -ErrorAction Stop | Out-Null
        }
    }
}

$script:SmartWorkplaceCMDBConsoleStarted = [datetimeoffset]::Now
$script:SmartWorkplaceCMDBCompletionWritten = $false
$script:SmartWorkplaceCMDBCurrentLogPath = ''
$script:SmartWorkplaceCMDBCurrentPipeline = $Pipeline
$script:SmartWorkplaceCMDBCurrentMode = ''
$script:SmartWorkplaceCMDBCurrentStepCount = 0
$script:SmartWorkplaceCMDBCurrentRunId = ''
$script:SmartWorkplaceCMDBCurrentRunGuard = $null
$script:SmartWorkplaceCMDBCurrentStepName = ''
$script:SmartWorkplaceCMDBCurrentTranscriptPath = ''
Write-SmartWorkplaceCMDBStartupBanner

try {
if ($Collect -and $ValidateOnly) {
    throw '-Collect and -ValidateOnly cannot be used together.'
}
if ($Collect -and -not [string]::IsNullOrWhiteSpace($FixtureRootPath)) {
    throw '-Collect cannot be combined with -FixtureRootPath.'
}
if ($MaxItems -gt 0 -and $Pipeline -in @('Full', 'CuratedOnly')) {
    throw '-MaxItems requires an individual source pipeline.'
}

$mode = if (-not [string]::IsNullOrWhiteSpace($FixtureRootPath) -and
    -not $ValidateOnly) {
    'Fixture'
}
elseif ($Collect) {
    'Collect'
}
else {
    'Validate'
}
$script:SmartWorkplaceCMDBCurrentMode = $mode

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
Import-Module $coreModulePath -Force
$sharePointModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.SharePoint\SmartWorkplaceCMDB.SharePoint.psd1'
Import-Module $sharePointModulePath -Force

if (-not [string]::IsNullOrWhiteSpace($FixtureRootPath)) {
    $FixtureRootPath = [IO.Path]::GetFullPath($FixtureRootPath)
    if (-not (Test-Path -LiteralPath $FixtureRootPath -PathType Container)) {
        throw "FixtureRootPath does not exist: '$FixtureRootPath'."
    }
}

if ($mode -ne 'Validate' -and $MaxItems -gt 0) {
    $boundedName = 'MAXITEMS-{0}_{1}_{2}' -f
        $MaxItems,
        ([datetime]::UtcNow.ToString('yyyyMMdd-HHmmssfff')),
        ([guid]::NewGuid().ToString('N'))
    $boundedBase = if ([string]::IsNullOrWhiteSpace($DataRootPath)) {
        Join-Path $projectRoot 'Data'
    } else { [IO.Path]::GetFullPath($DataRootPath) }
    $DataRootPath = Join-Path $boundedBase (Join-Path 'TestRuns' $boundedName)
    # Pin every child root: absolute config/CLI paths must not bypass isolation.
    $DataAllRootPath = Join-Path $DataRootPath 'DATA-ALL'
    $LatestOutputRootPath = Join-Path $DataRootPath 'DATA-LAST'
    $LogRootPath = Join-Path $DataRootPath 'LOG-ALL'
    Write-SmartWorkplaceCMDBConsole -Message (
        "Bounded output is isolated from canonical data: '{0}'." -f
        $DataRootPath
    )
}

$boundParameterCopy = @{}
foreach ($key in @(
        'Tenant',
        'OrganizationKey',
        'EnvironmentKey',
        'TenantKey',
        'TenantId',
        'DataRootPath',
        'DataAllRootPath',
        'LatestOutputRootPath',
        'LogRootPath'
    )) {
    $value = Get-Variable -Name $key -ValueOnly
    if ($key -eq 'Tenant' -or
        -not [string]::IsNullOrWhiteSpace([string]$value)) {
        $boundParameterCopy[$key] = $value
    }
}
$context = Resolve-SmartWorkplaceCMDBContext `
    -BoundParameters $boundParameterCopy `
    -GlobalConfigPath $GlobalConfigPath `
    -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($mode -ne 'Collect' -or $NoConfigWrite)
$paths = Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths -Fixture:($mode -eq 'Fixture') -MaxItems $MaxItems -ExplicitDataRoot:([bool]$DataRootPath) -NoWrite:($mode -eq 'Validate')
$DataRootPath = $paths.DataRootPath
$DataAllRootPath = $paths.DataAllRootPath
$LatestOutputRootPath = $paths.LatestOutputRootPath
$LogRootPath = $paths.LogRootPath
$sharePointConfiguration = Get-SmartWorkplaceCMDBOrchestratorSetting `
    -Configuration $context.Configuration `
    -Name 'SharePoint' `
    -DefaultValue $null
$graphConfiguration = Get-SmartWorkplaceCMDBOrchestratorSetting `
    -Configuration $context.Configuration `
    -Name 'MicrosoftGraph' `
    -DefaultValue $null
$notificationConfiguration = Get-SmartWorkplaceCMDBOrchestratorSetting `
    -Configuration $context.Configuration `
    -Name 'Notifications' `
    -DefaultValue $null
$loggingConfiguration = Get-SmartWorkplaceCMDBOrchestratorSetting `
    -Configuration $context.Configuration `
    -Name 'Logging' `
    -DefaultValue $null
$activeDirectoryConfiguration = Get-SmartWorkplaceCMDBOrchestratorSetting `
    -Configuration $context.Configuration `
    -Name 'ActiveDirectory' `
    -DefaultValue $null
$activeDirectoryTargetDomains = [string](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        -Configuration $activeDirectoryConfiguration `
        -Name 'TargetDomains' `
        -DefaultValue ''
)
$activeDirectoryScoped = $Pipeline -in @('Full', 'ActiveDirectory') -and (
    -not [bool](Get-SmartWorkplaceCMDBOrchestratorSetting `
            -Configuration $activeDirectoryConfiguration `
            -Name 'ForestWide' `
            -DefaultValue $true) -or
    -not [string]::IsNullOrWhiteSpace([string](
            Get-SmartWorkplaceCMDBOrchestratorSetting `
                -Configuration $activeDirectoryConfiguration `
                -Name 'SearchBase' `
                -DefaultValue '')) -or
    -not [string]::IsNullOrWhiteSpace($activeDirectoryTargetDomains)
)
$sharePointEnabled = [bool](Get-SmartWorkplaceCMDBOrchestratorSetting `
        -Configuration $sharePointConfiguration `
        -Name 'Enabled' `
        -DefaultValue $false)
$sharePointEligible = $mode -eq 'Collect' -and
    $MaxItems -eq 0 -and
    -not $DisableSharePointUpload -and
    -not $activeDirectoryScoped -and
    $sharePointEnabled
$summaryEnabled = [bool](Get-SmartWorkplaceCMDBOrchestratorSetting `
        -Configuration $notificationConfiguration `
        -Name 'Enabled' `
        -DefaultValue $false)
if ($summaryEnabled -and $Pipeline -eq 'Full') {
    $summaryMode = [string](Get-SmartWorkplaceCMDBOrchestratorSetting `
        $notificationConfiguration 'SendMailMode' 'Graph')
    $summaryFrom = [string](Get-SmartWorkplaceCMDBOrchestratorSetting `
        $notificationConfiguration 'From' '')
    $summaryTo = [string](Get-SmartWorkplaceCMDBOrchestratorSetting `
        $notificationConfiguration 'To' '')
    if ($summaryMode.ToUpperInvariant() -notin @('GRAPH','SMTP','BOTH')) {
        throw "Unsupported Notifications.SendMailMode '$summaryMode'. Use Graph, SMTP, or Both."
    }
    if ([string]::IsNullOrWhiteSpace($summaryFrom) -or
        [string]::IsNullOrWhiteSpace($summaryTo)) {
        throw 'Notifications.From and Notifications.To are required when full-collection summaries are enabled.'
    }
    $graphMailConfigured = -not [string]::IsNullOrWhiteSpace([string]$paths.TenantId) -and
        -not [string]::IsNullOrWhiteSpace([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                $graphConfiguration 'ClientId' '')) -and
        -not [string]::IsNullOrWhiteSpace([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                $graphConfiguration 'CertificateThumbprint' ''))
    $smtpMailConfigured = -not [string]::IsNullOrWhiteSpace([string](
        Get-SmartWorkplaceCMDBOrchestratorSetting $notificationConfiguration 'SmtpServer' ''))
    if ($summaryMode.ToUpperInvariant() -eq 'GRAPH' -and -not $graphMailConfigured) {
        throw 'MicrosoftGraph TenantId, ClientId, and CertificateThumbprint are required for Graph summary mail.'
    }
    if ($summaryMode.ToUpperInvariant() -eq 'SMTP' -and -not $smtpMailConfigured) {
        throw 'Notifications.SmtpServer is required for SMTP summary mail.'
    }
    if ($summaryMode.ToUpperInvariant() -eq 'BOTH' -and
        -not ($graphMailConfigured -or $smtpMailConfigured)) {
        throw 'Graph app-only settings or Notifications.SmtpServer are required for summary mail fallback.'
    }
}
$summaryEligible = $mode -eq 'Collect' -and
    $Pipeline -eq 'Full' -and
    $MaxItems -eq 0 -and
    -not $activeDirectoryScoped -and
    $summaryEnabled
$loggingEnabled = $mode -ne 'Validate' -and [bool](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        $loggingConfiguration 'Enabled' $true)
$orchestratorLogRetentionDays = [math]::Max(0, [int](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        $loggingConfiguration 'OrchestratorLogRetentionDays' 30))
$stepLogRetentionDays = [math]::Max(0, [int](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        $loggingConfiguration 'StepLogRetentionDays' 30))
$runCsvRetentionDays = [math]::Max(0, [int](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        $loggingConfiguration 'RunCsvRetentionDays' 90))
$maxOrchestratorLogs = [math]::Max(0, [int](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        $loggingConfiguration 'MaxOrchestratorLogs' 30))
$maxStepLogsPerScript = [math]::Max(0, [int](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        $loggingConfiguration 'MaxStepLogsPerScript' 30))
$maxRunCsvFiles = [math]::Max(0, [int](
    Get-SmartWorkplaceCMDBOrchestratorSetting `
        $loggingConfiguration 'MaxRunCsvFiles' 90))
$sharePointBeforeSnapshot = if ($sharePointEligible) {
    Get-SmartWorkplaceCMDBCsvSnapshot -RootPath @(
        $paths.DataAllRootPath,
        $paths.LatestOutputRootPath,
        $paths.LogRootPath
    )
}
else {
    @{}
}

$commonParameters = @{
    Tenant = $Tenant
}
foreach ($entry in @(
        @{ Name = 'OrganizationKey'; Value = $OrganizationKey },
        @{ Name = 'EnvironmentKey'; Value = $EnvironmentKey },
        @{ Name = 'TenantKey'; Value = $TenantKey },
        @{ Name = 'TenantId'; Value = $TenantId },
        @{ Name = 'DataRootPath'; Value = $DataRootPath },
        @{ Name = 'DataAllRootPath'; Value = $DataAllRootPath },
        @{ Name = 'LatestOutputRootPath'; Value = $LatestOutputRootPath },
        @{ Name = 'LogRootPath'; Value = $LogRootPath },
        @{ Name = 'GlobalConfigPath'; Value = $GlobalConfigPath },
        @{ Name = 'TenantConfigPath'; Value = $TenantConfigPath }
    )) {
    if (-not [string]::IsNullOrWhiteSpace([string]$entry.Value)) {
        $commonParameters[$entry.Name] = $entry.Value
    }
}
if ($mode -ne 'Collect' -or $NoConfigWrite) {
    $commonParameters['NoConfigWrite'] = $true
}

$catalog = @{}
$catalog['EntraUsersCollect'] = [pscustomobject]@{
    Name = 'Entra users collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1'
    FixtureName = 'EntraUsers.sample.json'
}
$catalog['EntraUsersNormalize'] = [pscustomobject]@{
    Name = 'Entra users normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Normalize.ps1'
    FixtureName = ''
}
$catalog['EntraGroupsCollect'] = [pscustomobject]@{
    Name = 'Entra groups collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Collect.ps1'
    FixtureName = 'EntraGroups.sample.json'
}
$catalog['EntraGroupsNormalize'] = [pscustomobject]@{
    Name = 'Entra groups normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Normalize.ps1'
    FixtureName = ''
}
$catalog['EntraDevicesCollect'] = [pscustomobject]@{
    Name = 'Entra devices collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Collect.ps1'
    FixtureName = 'EntraDevices.sample.json'
}
$catalog['EntraDevicesNormalize'] = [pscustomobject]@{
    Name = 'Entra devices normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Normalize.ps1'
    FixtureName = ''
}
$catalog['VerifiedDomainsCollect'] = [pscustomobject]@{
    Name = 'Entra verified domains collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraVerifiedDomains-Collect.ps1'
    FixtureName = 'EntraVerifiedDomains.sample.json'
}
$catalog['TenantIdentityHealthNormalize'] = [pscustomobject]@{
    Name = 'Verified domains and hybrid identity coverage'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraVerifiedDomains-Normalize.ps1'
    FixtureName = ''
}
$catalog['IntuneDevicesCollect'] = [pscustomobject]@{
    Name = 'Intune managed devices collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1'
    FixtureName = 'IntuneManagedDevices.sample.json'
}
$catalog['IntuneDevicesNormalize'] = [pscustomobject]@{
    Name = 'Intune device enrichment'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1'
    FixtureName = ''
}
$catalog['IntuneHardwareCollect'] = [pscustomobject]@{
    Name = 'Intune device hardware collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneHardware-Collect.ps1'
    FixtureName = 'IntuneManagedDevices.sample.json'
    RequiresExplicitCollect = $true
}
$catalog['UserDeviceRelationships'] = [pscustomobject]@{
    Name = 'Primary user-device relationships'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneUserDeviceRelationships-Normalize.ps1'
    FixtureName = ''
}
$catalog['M365SkusCollect'] = [pscustomobject]@{
    Name = 'Microsoft 365 subscribed SKUs collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Collect.ps1'
    FixtureName = 'M365SubscribedSkus.sample.json'
}
$catalog['M365SkusNormalize'] = [pscustomobject]@{
    Name = 'Microsoft 365 subscribed SKUs normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Normalize.ps1'
    FixtureName = ''
}
$catalog['M365LicensesCollect'] = [pscustomobject]@{
    Name = 'Microsoft 365 user licenses collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Collect.ps1'
    FixtureName = 'M365UserLicenseAssignments.sample.json'
}
$catalog['M365LicensesNormalize'] = [pscustomobject]@{
    Name = 'Microsoft 365 user licenses normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Normalize.ps1'
    FixtureName = ''
}
$catalog['ExchangeMailboxesCollect'] = [pscustomobject]@{
    Name = 'Exchange Online mailboxes collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Collect.ps1'
    FixtureName = 'ExchangeOnlineMailboxes.sample.json'
}
$catalog['ExchangeMailboxesNormalize'] = [pscustomobject]@{
    Name = 'Exchange Online mailboxes normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Normalize.ps1'
    FixtureName = ''
}
$catalog['ActiveDirectoryCollect'] = [pscustomobject]@{
    Name = 'Active Directory collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Collect.ps1'
    FixtureName = 'ActiveDirectory.sample.json'
}
$catalog['ActiveDirectoryNormalize'] = [pscustomobject]@{
    Name = 'Active Directory normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Normalize.ps1'
    FixtureName = ''
}
$catalog['Relationships'] = [pscustomobject]@{
    Name = 'General relationship consolidation'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-Relationships-Normalize.ps1'
    FixtureName = ''
}
$catalog['DataQuality'] = [pscustomobject]@{
    Name = 'Data-quality normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-DataQuality-Normalize.ps1'
    FixtureName = ''
}
$catalog['Dimensions'] = [pscustomobject]@{
    Name = 'Tenant and date dimensions'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-Dimensions-Normalize.ps1'
    FixtureName = ''
}
$catalog['Build'] = [pscustomobject]@{
    Name = 'Contract build and manifest'
    Kind = 'Build'
    ScriptPath = Join-Path $projectRoot 'Build\SmartWorkplaceCMDB-Build.ps1'
    FixtureName = ''
}
$catalog['Report'] = [pscustomobject]@{
    Name = 'Local HTML overview report'
    Kind = 'Report'
    ScriptPath = Join-Path $projectRoot 'Reports\SmartWorkplaceCMDB-Report.ps1'
    FixtureName = ''
}

$selectedSteps = New-Object System.Collections.Generic.List[object]
switch ($Pipeline) {
    'Full' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraUsersCollect', 'EntraUsersNormalize',
            'EntraGroupsCollect', 'EntraGroupsNormalize',
            'EntraDevicesCollect', 'EntraDevicesNormalize',
            'ActiveDirectoryCollect', 'ActiveDirectoryNormalize',
            'VerifiedDomainsCollect', 'TenantIdentityHealthNormalize',
            'IntuneDevicesCollect', 'IntuneDevicesNormalize',
            'IntuneHardwareCollect',
            'UserDeviceRelationships',
            'M365SkusCollect', 'M365SkusNormalize',
            'M365LicensesCollect', 'M365LicensesNormalize',
            'ExchangeMailboxesCollect', 'ExchangeMailboxesNormalize',
            'Relationships', 'DataQuality', 'Dimensions', 'Build', 'Report'
        )
    }
    'EntraUsers' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraUsersCollect', 'EntraUsersNormalize'
        )
    }
    'EntraGroups' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraGroupsCollect', 'EntraGroupsNormalize'
        )
    }
    'EntraDevices' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraDevicesCollect', 'EntraDevicesNormalize'
        )
    }
    'TenantIdentityHealth' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'VerifiedDomainsCollect', 'TenantIdentityHealthNormalize'
        )
    }
    'IntuneDevices' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'IntuneDevicesCollect', 'IntuneDevicesNormalize',
            'IntuneHardwareCollect',
            'UserDeviceRelationships'
        )
    }
    'M365SubscribedSkus' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'M365SkusCollect', 'M365SkusNormalize'
        )
    }
    'M365UserLicenses' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'M365LicensesCollect', 'M365LicensesNormalize'
        )
    }
    'ExchangeOnlineMailboxes' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'ExchangeMailboxesCollect', 'ExchangeMailboxesNormalize'
        )
    }
    'ActiveDirectory' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'ActiveDirectoryCollect', 'ActiveDirectoryNormalize'
        )
    }
    'CuratedOnly' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'TenantIdentityHealthNormalize',
            'Relationships', 'DataQuality', 'Dimensions', 'Build', 'Report'
        )
    }
}

foreach ($step in @($selectedSteps.ToArray())) {
    if (-not (Test-Path -LiteralPath $step.ScriptPath -PathType Leaf)) {
        throw "Orchestrator step script is missing: '$($step.ScriptPath)'."
    }
    if (($mode -eq 'Fixture' -or
            ($mode -eq 'Validate' -and
                -not [string]::IsNullOrWhiteSpace($FixtureRootPath))) -and
        $step.Kind -eq 'Collect') {
        $fixturePath = Join-Path $FixtureRootPath $step.FixtureName
        if (-not (Test-Path -LiteralPath $fixturePath -PathType Leaf)) {
            throw "Required fixture is missing: '$fixturePath'."
        }
    }
}

$executionSteps = @($selectedSteps.ToArray())
if ($mode -eq 'Validate' -and
    -not $ValidateExistingOutputs -and
    $Pipeline -ne 'CuratedOnly') {
    $executionSteps = @($executionSteps |
        Where-Object Kind -in @('Collect', 'Build', 'Report'))
}

$runId = [guid]::NewGuid().ToString('N')
$runStarted = [datetimeoffset]::UtcNow
$script:SmartWorkplaceCMDBCurrentRunId = $runId
$computerName = if ([string]::IsNullOrWhiteSpace([string]$env:COMPUTERNAME)) {
    'unknown-host'
}
else {
    ConvertTo-SmartWorkplaceCMDBLogName -Value $env:COMPUTERNAME
}
$runStamp = $runStarted.ToString('yyyyMMdd-HHmmssfff')
$orchestratorLogFolderPath = Join-Path $paths.LogRootPath `
    'Orchestration\Logs'
$runCsvFolderPath = Join-Path $paths.LogRootPath 'Orchestration\Runs'
$stepLogRootPath = Join-Path $paths.LogRootPath 'Jobs'
$orchestratorLogPath = if ($loggingEnabled) {
    Join-Path $orchestratorLogFolderPath (
        'SmartWorkplaceCMDB-Orchestrator_{0}_{1}.log' -f
        $computerName,
        $runStamp)
}
else { '' }
$script:SmartWorkplaceCMDBCurrentLogPath = $orchestratorLogPath
$summaryScriptPath = Join-Path $projectRoot 'Reports\SmartWorkplaceCMDB-CollectionSummary.ps1'
$summaryParameters = @{
    Tenant = $paths.ProfileKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantKey = $paths.TenantKey
    TenantId = $paths.TenantId
    DataRootPath = $paths.DataRootPath
    DataAllRootPath = $paths.DataAllRootPath
    LatestOutputRootPath = $paths.LatestOutputRootPath
    LogRootPath = $paths.LogRootPath
    RunId = $runId
    NoConfigWrite = $true
}
if (-not [string]::IsNullOrWhiteSpace($GlobalConfigPath)) {
    $summaryParameters['GlobalConfigPath'] = $GlobalConfigPath
}
if (-not [string]::IsNullOrWhiteSpace($TenantConfigPath)) {
    $summaryParameters['TenantConfigPath'] = $TenantConfigPath
}
if ($summaryEligible) {
    try {
        @(& $summaryScriptPath @summaryParameters -CaptureBaselineOnly `
                -SnapshotDateTime $runStarted) | Out-Null
    }
    catch {
        Write-SmartWorkplaceCMDBConsole -Level WARN -Message (
            'The pre-collection summary baseline could not be captured. ' +
            'The collection will continue and unavailable comparisons will be shown as n/a. ' +
            $_.Exception.Message)
    }
}
$results = New-Object System.Collections.Generic.List[object]
$logPath = if (-not $loggingEnabled) {
    ''
}
else {
    Join-Path $runCsvFolderPath (
        'SmartWorkplaceCMDB-Orchestrator_{0}_{1}.csv' -f
        $computerName,
        $runStamp)
}

$preflight = Test-SmartWorkplaceCMDBPreflight `
    -Context $context `
    -ProjectRoot $projectRoot `
    -Pipeline $Pipeline `
    -Mode $mode `
    -ScriptPath @($executionSteps | ForEach-Object { [string]$_.ScriptPath }) `
    -ThrowOnFailure:($mode -eq 'Collect')
foreach ($check in @($preflight.Checks | Where-Object Status -in @('Failed','Warning'))) {
    Write-SmartWorkplaceCMDBConsole `
        -Level $(if ($check.Status -eq 'Failed') {'ERROR'} else {'WARN'}) `
        -Message ("Preflight {0}: {1}" -f $check.Name, $check.Details)
}

$script:SmartWorkplaceCMDBCurrentRunGuard = Enter-SmartWorkplaceCMDBRunGuard `
    -Paths $paths `
    -Pipeline $Pipeline `
    -RunId $runId `
    -StartedDateTime $runStarted `
    -NoWrite:($mode -ne 'Collect')
Update-SmartWorkplaceCMDBRunState `
    -RunGuard $script:SmartWorkplaceCMDBCurrentRunGuard `
    -Status 'Running' `
    -LogPath $orchestratorLogPath

if ($loggingEnabled) {
    Write-SmartWorkplaceCMDBStartupBanner `
        -LogPath $orchestratorLogPath -NoConsole
    Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath -Message (
        'Started orchestration. RunId={0}; Tenant={1}; Pipeline={2}; Mode={3}; Version={4}.' -f
        $runId,
        $paths.TenantKey,
        $Pipeline,
        $mode,
        $ScriptVersion)
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $orchestratorLogFolderPath `
            -Filter 'SmartWorkplaceCMDB-Orchestrator_*.log' `
            -RetentionDays $orchestratorLogRetentionDays `
            -MaxFiles $maxOrchestratorLogs `
            -ExcludePath $orchestratorLogPath)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $runCsvFolderPath `
            -Filter 'SmartWorkplaceCMDB-Orchestrator_*.csv' `
            -RetentionDays $runCsvRetentionDays `
            -MaxFiles $maxRunCsvFiles `
            -ExcludePath $logPath)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $stepLogRootPath `
            -Filter '*.log' `
            -RetentionDays $stepLogRetentionDays `
            -MaxFiles $maxStepLogsPerScript `
            -Recurse)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $stepLogRootPath `
            -Filter '*.transcript.txt' `
            -RetentionDays $stepLogRetentionDays `
            -MaxFiles $maxStepLogsPerScript `
            -Recurse)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
}

$failedMessage = ''
try {
    for ($index = 0; $index -lt $executionSteps.Count; $index++) {
        $step = $executionSteps[$index]
        $sequence = $index + 1
        $elapsed = [datetimeoffset]::UtcNow - $runStarted
        $etaText = 'estimating'
        if ($index -gt 0) {
            $averageSeconds = $elapsed.TotalSeconds / $index
            $remainingSeconds = $averageSeconds * (
                $executionSteps.Count - $index
            )
            $etaText = [timespan]::FromSeconds(
                [math]::Max(0, $remainingSeconds)
            ).ToString('hh\:mm\:ss')
        }
        Write-SmartWorkplaceCMDBConsole -Message (
            "[{0}/{1}] {2} (elapsed {3}; ETA {4})" -f
            $sequence,
            $executionSteps.Count,
            $step.Name,
            $elapsed.ToString('hh\:mm\:ss'),
            $etaText
        )

        $stepStarted = [datetimeoffset]::UtcNow
        $stepScriptName = ConvertTo-SmartWorkplaceCMDBLogName -Value (
            [IO.Path]::GetFileNameWithoutExtension([string]$step.ScriptPath))
        $stepLogPath = if ($loggingEnabled) {
            Join-Path (Join-Path $stepLogRootPath $stepScriptName) (
                '{0}_{1}_{2}_{3:D2}.log' -f
                $stepScriptName,
                $computerName,
                $runStamp,
                $sequence)
        }
        else { '' }
        $stepTranscriptPath = if ($loggingEnabled) {
            Join-Path (Join-Path $stepLogRootPath $stepScriptName) (
                '{0}_{1}_{2}_{3:D2}.transcript.txt' -f
                $stepScriptName,
                $computerName,
                $runStamp,
                $sequence)
        }
        else { '' }
        if ($loggingEnabled) {
            Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath -Message (
                '[{0}/{1}] Started {2}. StepLog={3}; Transcript={4}' -f
                $sequence,
                $executionSteps.Count,
                $step.Name,
                $stepLogPath,
                $stepTranscriptPath)
        }
        $status = 'Completed'
        $errorText = ''
        try {
            $script:SmartWorkplaceCMDBCurrentStepName = $step.Name
            $script:SmartWorkplaceCMDBCurrentTranscriptPath = $stepTranscriptPath
            Update-SmartWorkplaceCMDBRunState `
                -RunGuard $script:SmartWorkplaceCMDBCurrentRunGuard `
                -Status 'Running' `
                -CurrentStep $step.Name `
                -CompletedStepCount $results.Count `
                -LogPath $stepLogPath `
                -TranscriptPath $stepTranscriptPath
            $parameters = Get-SmartWorkplaceCMDBStepParameter `
                -Step $step `
                -CommonParameters $commonParameters `
                -Mode $mode `
                -FixtureRootPath $FixtureRootPath `
                -MaxItems $MaxItems
            Invoke-SmartWorkplaceCMDBLoggedStep `
                -Step $step `
                -Parameters $parameters `
                -LogPath $stepLogPath `
                -TranscriptPath $stepTranscriptPath
            if ($mode -eq 'Validate') {
                $status = 'Validated'
            }
        }
        catch {
            $status = 'Failed'
            $errorText = $_.Exception.Message
            if ($loggingEnabled) {
                Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
                    -Message ("Failed {0}: {1}" -f $step.Name, $errorText) `
                    -Level ERROR
            }
            throw
        }
        finally {
            $stepEnded = [datetimeoffset]::UtcNow
            if ($loggingEnabled -and $status -ne 'Failed') {
                Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath -Message (
                    'Completed {0}. DurationSeconds={1}.' -f
                    $step.Name,
                    [math]::Round(($stepEnded - $stepStarted).TotalSeconds, 3))
            }
            if ($status -ne 'Failed') {
                Write-SmartWorkplaceCMDBConsole -Message (
                    '[{0}/{1}] {2} {3} (duration {4})' -f
                    $sequence,
                    $executionSteps.Count,
                    $step.Name,
                    $status.ToLowerInvariant(),
                    ($stepEnded - $stepStarted).ToString('hh\:mm\:ss')
                )
            }
            $results.Add([pscustomobject][ordered]@{
                RunId = $runId
                Sequence = $sequence
                Pipeline = $Pipeline
                Mode = $mode
                Step = $step.Name
                Status = $status
                StartedDateTime = $stepStarted.ToString('o')
                EndedDateTime = $stepEnded.ToString('o')
                DurationSeconds = [math]::Round(
                    ($stepEnded - $stepStarted).TotalSeconds,
                    3
                )
                LogPath = $stepLogPath
                TranscriptPath = $stepTranscriptPath
                Error = $errorText
            })
            Update-SmartWorkplaceCMDBRunState `
                -RunGuard $script:SmartWorkplaceCMDBCurrentRunGuard `
                -Status $(if ($status -eq 'Failed') {'Failed'} else {'Running'}) `
                -CurrentStep $step.Name `
                -CompletedStepCount $results.Count `
                -LogPath $stepLogPath `
                -TranscriptPath $stepTranscriptPath `
                -Error $errorText
        }
    }
    $script:SmartWorkplaceCMDBCurrentStepCount = $results.Count
}
catch {
    $failedMessage = $_.Exception.Message
    $script:SmartWorkplaceCMDBCurrentStepCount = $results.Count
}
finally {
    if ($loggingEnabled) {
        Write-SmartWorkplaceCMDBOrchestratorLog `
            -Results $results `
            -Path $logPath
    }
}

if (-not [string]::IsNullOrWhiteSpace($failedMessage)) {
    if ($loggingEnabled) {
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Orchestration failed: $failedMessage" -Level ERROR
    }
    $failureEmailError = ''
    if ($summaryEnabled -and $mode -eq 'Collect' -and $Pipeline -eq 'Full') {
        try {
            @(& $summaryScriptPath @summaryParameters `
                    -RunStatus 'Failed' `
                    -SnapshotDateTime ([datetimeoffset]::UtcNow) `
                    -OperationalError $failedMessage `
                    -FailedStep $script:SmartWorkplaceCMDBCurrentStepName `
                    -FailureLogPath $orchestratorLogPath `
                    -FailureTranscriptPath $script:SmartWorkplaceCMDBCurrentTranscriptPath) | Out-Null
        }
        catch {
            $failureEmailError = $_.Exception.Message
            if ($loggingEnabled) {
                Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
                    -Message "Failure notification email failed: $failureEmailError" -Level WARN
            }
        }
    }
    Exit-SmartWorkplaceCMDBRunGuard `
        -RunGuard $script:SmartWorkplaceCMDBCurrentRunGuard `
        -Status 'Failed' `
        -Error $failedMessage
    throw "SmartWorkplaceCMDB orchestration failed: $failedMessage"
}

$sharePointRecords = @()
$sharePointError = ''
if ($sharePointEligible) {
    try {
        $afterSnapshot = Get-SmartWorkplaceCMDBCsvSnapshot -RootPath @(
            $paths.DataAllRootPath,
            $paths.LatestOutputRootPath,
            $paths.LogRootPath
        )
        $changedFiles = @($afterSnapshot.Keys | Where-Object {
                -not $sharePointBeforeSnapshot.ContainsKey($_) -or
                $sharePointBeforeSnapshot[$_] -ne $afterSnapshot[$_]
            } | Sort-Object)
        if ($changedFiles.Count -gt 0) {
            $sharePointRecords = @(Publish-SmartWorkplaceCMDBSharePointFile `
                    -LocalFilePath $changedFiles `
                    -DataAllRootPath $paths.DataAllRootPath `
                    -LatestOutputRootPath $paths.LatestOutputRootPath `
                    -LogRootPath $paths.LogRootPath `
                    -TenantId ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $graphConfiguration 'TenantId' '')) `
                    -ClientId ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $graphConfiguration 'ClientId' '')) `
                    -CertificateThumbprint ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $graphConfiguration 'CertificateThumbprint' '')) `
                    -SiteHostname ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'SiteHostname' '')) `
                    -SitePath ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'SitePath' '')) `
                    -LibraryDisplayName ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'LibraryDisplayName' 'Documents')) `
                    -TargetFolderPath ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'TargetFolderPath' 'SMART-CMDB/DATA')))
        }
    }
    catch {
        $sharePointError = $_.Exception.Message
        if ($loggingEnabled) {
            Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
                -Message "SharePoint publication failed: $sharePointError" `
                -Level WARN
        }
        Write-SmartWorkplaceCMDBConsole -Level WARN -Message (
            "SmartWorkplaceCMDB SharePoint publication failed but collection outputs are preserved: $sharePointError")
    }
}
elseif ($mode -eq 'Collect' -and $MaxItems -gt 0 -and $sharePointEnabled) {
    Write-SmartWorkplaceCMDBConsole -Message `
        'SharePoint publication skipped for the bounded MaxItems run.'
}
elseif ($mode -eq 'Collect' -and $activeDirectoryScoped -and $sharePointEnabled) {
    Write-SmartWorkplaceCMDBConsole -Message `
        'SharePoint publication skipped for the scoped Active Directory run.'
}

$sharePointFailureCount = @($sharePointRecords |
    Where-Object Status -eq 'Failed').Count
$sharePointUploadCount = @($sharePointRecords |
    Where-Object Status -eq 'Uploaded').Count
if ($sharePointEligible) {
    Write-SmartWorkplaceCMDBConsole -Message (
        'SmartWorkplaceCMDB SharePoint publication completed. Uploaded={0}; Failed={1}; Target={2}.' -f
        $sharePointUploadCount,
        ($sharePointFailureCount + [int](-not [string]::IsNullOrWhiteSpace($sharePointError))),
        ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                $sharePointConfiguration 'TargetFolderPath' 'SMART-CMDB/DATA'))
    )
}

$runEnded = [datetimeoffset]::UtcNow
$runStatus = if ($mode -eq 'Validate') {
    'Validated'
}
elseif ($sharePointFailureCount -gt 0 -or
    -not [string]::IsNullOrWhiteSpace($sharePointError)) {
    'CompletedWithWarnings'
}
else {
    'Completed'
}
$summaryResult = $null
$summaryError = ''
if ($summaryEligible) {
    try {
        $summaryResult = & $summaryScriptPath @summaryParameters `
            -RunStatus $runStatus `
            -SnapshotDateTime $runEnded
        Write-SmartWorkplaceCMDBConsole -Message (
            "SmartWorkplaceCMDB full-collection summary email sent. HTML='{0}'." -f
            $summaryResult.HtmlPath
        )
        if ($loggingEnabled) {
            Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath -Message (
                "Full-collection summary email sent. HTML='{0}'." -f
                $summaryResult.HtmlPath)
        }
    }
    catch {
        $summaryError = $_.Exception.Message
        $runStatus = 'CompletedWithWarnings'
        if ($loggingEnabled) {
            Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
                -Message "Full-collection summary email failed: $summaryError" `
                -Level WARN
        }
        Write-SmartWorkplaceCMDBConsole -Level WARN -Message (
            "SmartWorkplaceCMDB summary email failed but collection outputs are preserved: $summaryError")
    }
}
elseif ($mode -eq 'Collect' -and $Pipeline -eq 'Full' -and
    $MaxItems -gt 0 -and $summaryEnabled) {
    Write-SmartWorkplaceCMDBConsole -Message `
        'Full-collection summary email skipped for the bounded MaxItems run.'
}
elseif ($mode -eq 'Collect' -and $Pipeline -eq 'Full' -and
    $activeDirectoryScoped -and $summaryEnabled) {
    Write-SmartWorkplaceCMDBConsole -Message `
        'Full-collection summary email skipped because Active Directory collection is scoped.'
}
Write-SmartWorkplaceCMDBConsole -Message (
    "SmartWorkplaceCMDB orchestration {0}. Pipeline={1}; Mode={2}; Steps={3}; Duration={4}." -f
    $runStatus.ToLowerInvariant(),
    $Pipeline,
    $mode,
    $results.Count,
    ($runEnded - $runStarted).ToString('hh\:mm\:ss')
)

if ($loggingEnabled) {
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $orchestratorLogFolderPath `
            -Filter 'SmartWorkplaceCMDB-Orchestrator_*.log' `
            -RetentionDays $orchestratorLogRetentionDays `
            -MaxFiles $maxOrchestratorLogs `
            -ExcludePath $orchestratorLogPath)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $runCsvFolderPath `
            -Filter 'SmartWorkplaceCMDB-Orchestrator_*.csv' `
            -RetentionDays $runCsvRetentionDays `
            -MaxFiles $maxRunCsvFiles `
            -ExcludePath $logPath)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $stepLogRootPath `
            -Filter '*.log' `
            -RetentionDays $stepLogRetentionDays `
            -MaxFiles $maxStepLogsPerScript `
            -Recurse)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
    foreach ($retentionEntry in @(Invoke-SmartWorkplaceCMDBLogRetention `
            -FolderPath $stepLogRootPath `
            -Filter '*.transcript.txt' `
            -RetentionDays $stepLogRetentionDays `
            -MaxFiles $maxStepLogsPerScript `
            -Recurse)) {
        $level = if ($retentionEntry.StartsWith('failed:')) { 'WARN' } else { 'INFO' }
        Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath `
            -Message "Retention: $retentionEntry" -Level $level
    }
    Write-SmartWorkplaceCMDBTextLog -Path $orchestratorLogPath -Message (
        'Completed orchestration. Status={0}; Steps={1}; DurationSeconds={2}.' -f
        $runStatus,
        $results.Count,
        [math]::Round(($runEnded - $runStarted).TotalSeconds, 3))
}

$completionWarningCount = $sharePointFailureCount +
    [int](-not [string]::IsNullOrWhiteSpace($sharePointError)) +
    [int](-not [string]::IsNullOrWhiteSpace($summaryError))
Write-SmartWorkplaceCMDBCompletionBanner `
    -LogPath $orchestratorLogPath `
    -Status $runStatus `
    -Pipeline $Pipeline `
    -Mode $mode `
    -StartedDateTime $runStarted `
    -StepCount $results.Count `
    -WarningCount $completionWarningCount `
    -ErrorCount 0
$script:SmartWorkplaceCMDBCompletionWritten = $true

Exit-SmartWorkplaceCMDBRunGuard `
    -RunGuard $script:SmartWorkplaceCMDBCurrentRunGuard `
    -Status $runStatus

[pscustomobject]@{
    Status = $runStatus
    ScriptVersion = $ScriptVersion
    RunId = $runId
    Pipeline = $Pipeline
    Mode = $mode
    StepCount = $results.Count
    FailedStepCount = @($results |
        Where-Object Status -eq 'Failed').Count
    StartedDateTime = $runStarted.ToString('o')
    EndedDateTime = $runEnded.ToString('o')
    DurationSeconds = [math]::Round(
        ($runEnded - $runStarted).TotalSeconds,
        3
    )
    DataRootPath = $paths.DataRootPath
    LatestOutputRootPath = $paths.LatestOutputRootPath
    LogPath = $logPath
    OrchestratorLogPath = $orchestratorLogPath
    StepLogRootPath = if ($loggingEnabled) { $stepLogRootPath } else { '' }
    StepTranscriptRootPath = if ($loggingEnabled) { $stepLogRootPath } else { '' }
    LoggingEnabled = $loggingEnabled
    PreflightStatus = $preflight.Status
    PreflightFailedCount = $preflight.FailedCount
    PreflightWarningCount = $preflight.WarningCount
    RunStatePath = if ($script:SmartWorkplaceCMDBCurrentRunGuard) { $script:SmartWorkplaceCMDBCurrentRunGuard.StatePath } else { '' }
    OrchestratorLogRetentionDays = $orchestratorLogRetentionDays
    StepLogRetentionDays = $stepLogRetentionDays
    RunCsvRetentionDays = $runCsvRetentionDays
    MaxOrchestratorLogs = $maxOrchestratorLogs
    MaxStepLogsPerScript = $maxStepLogsPerScript
    MaxRunCsvFiles = $maxRunCsvFiles
    SharePointEnabled = $sharePointEnabled
    SharePointEligible = $sharePointEligible
    SharePointUploadCount = $sharePointUploadCount
    SharePointFailureCount = $sharePointFailureCount +
        [int](-not [string]::IsNullOrWhiteSpace($sharePointError))
    SharePointTargetFolderPath = [string](
        Get-SmartWorkplaceCMDBOrchestratorSetting `
            $sharePointConfiguration `
            'TargetFolderPath' `
            'SMART-CMDB/DATA'
    )
    SharePointError = $sharePointError
    SummaryEmailEnabled = $summaryEnabled
    SummaryEmailEligible = $summaryEligible
    SummaryEmailStatus = if ($summaryResult) { [string]$summaryResult.Status } elseif ($summaryEligible) { 'Failed' } else { 'NotApplicable' }
    SummaryEmailHtmlPath = if ($summaryResult) { [string]$summaryResult.HtmlPath } else { '' }
    SummaryEmailError = $summaryError
}
}
catch {
    Exit-SmartWorkplaceCMDBRunGuard `
        -RunGuard $script:SmartWorkplaceCMDBCurrentRunGuard `
        -Status 'Failed' `
        -Error $_.Exception.Message
    if (-not $script:SmartWorkplaceCMDBCompletionWritten) {
        $message = $_.Exception.Message
        Write-SmartWorkplaceCMDBConsole -Level ERROR -Message (
            "SmartWorkplaceCMDB failed: $message")
        Write-SmartWorkplaceCMDBCompletionBanner `
            -LogPath $script:SmartWorkplaceCMDBCurrentLogPath `
            -Status 'Failed' `
            -Pipeline $script:SmartWorkplaceCMDBCurrentPipeline `
            -Mode $script:SmartWorkplaceCMDBCurrentMode `
            -StartedDateTime $script:SmartWorkplaceCMDBConsoleStarted `
            -StepCount $script:SmartWorkplaceCMDBCurrentStepCount `
            -WarningCount 0 `
            -ErrorCount 1
        $script:SmartWorkplaceCMDBCompletionWritten = $true
    }
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCm77/kCzUqY4OG
# vQlkbzhreEIh4Io/olqGw+1e/JFBgaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFPlgNT+HKot1oJY2l1PDq4cRY93k+BwnwusR7H2/WbYMA0GCSqG
# SIb3DQEBAQUABIIBgHmskcyaCEcEbYjLVuK6xq7vtWtVwM4A1pIdM/LaByl21Q0f
# 1c7Tyf7Izy5UGzaXRFAPL4XohBek67k1/tu8nb20GxjIsiGKgIq8QNvzDK7Va49M
# /T9AHDXgty1gByPPMjkldfcf+mRvRJTxysz2RJXqX4jpo4P4I0VaWU+HidJTOrZD
# Khs35MjQ5/onnz3MS+TsOZ/0wlKIPZITlnxYQzk8ZbWWJqjXaaKfkLyKauiN9Hy6
# OvTB1ZJq38pAKzqc+g691ZJ+I7eKLTdCCakAy2UvNcDysBflllg56vSpi5uhoC12
# wKO3sBCe4EN80n2NZyjypLWwI7rW1qLs7Bbnr5wJPOtaOa5CcJ/jnYWad/xU2jug
# DHqfEVG4Vxg5ftpkqNJqcIJGtUObSDiScIGNmm9L+Xd4m6xVJTdZmwDn9tZ2O4Jj
# xVMqSuwM48vBwMJ/QUwB3RDXqsrC1x3lpGj+97rxMLjlPitns/3wdVvqFgHFYJ3m
# LcesO7iAUiBRV+OFNKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxODA2
# NDZaMC8GCSqGSIb3DQEJBDEiBCCNNEcTafBG4hEuB5+Hh27vKPXApRW4fEFVumt2
# 1kG9XzANBgkqhkiG9w0BAQEFAASCAgADXVrRSGeNN9zKNV4/nafX6mbtGH8EcjNd
# vO1YqlAFqaNlMO2Gq8+snoCMpBRtsMS+mCX9tW1bJy3+87ePy3QqcENUW5M+YRC8
# nmKaBw64dYAsEHlt0vr+wLc+vyECuXyAwi0yMxyc5gg7+3Lk/QiFlJp2DPLHvMSm
# IpfLH76j/TT0dUtaETJyUljIk7RkfU6GjS8e+KAUEN0ssDG7NhdAvH8/FGBeux/1
# 5S+u11fMoFiBsjY1GkN7HvePHqPi9iwxWv3jDFM20SSoEhjnmJSZEGJL4DcFY6pc
# D3UjBSFtYUV2gI2O7ew2G4d7wxBs5KrpO3y1CF0kXlmAT6sdHYeCoY2f82/9NKni
# KMkhucJ9pomw7mivj1AtWBp6qKoH1/iSwGfNuKkZ1qHkEdpmtYxX7D4FN/v58acc
# RghsvjnBBrgZVIaojkoQ4RlwAvIFpjcCY5zcbggXtBjovbyiPFcPNMo4vmNbdDlh
# 2d769AOXeJ1dJ+e7MruILwA/5TXlXki37/ggKqZp4wSffNE4WYvKrRvM/NEC9op4
# ZWtCFpi5Nf1mVE1l0xjACnbPS/Wdu1BIXSRwJN/wwhDaEoFtjwQbDUQsrOW7ZyVS
# +0Mj4+2BBvzlKoxlsUt+0+gm1GnCJQtETh/BxgdZPj9Kel0woz/UorT7PYX8k7bs
# nrTm81BJrQ==
# SIG # End signature block
