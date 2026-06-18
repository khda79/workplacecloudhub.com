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

        $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'SmartM365.global.local.json'
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

    $logEntry = @(Format-SmartM365LogLine -Message $Message -Level $Level)

    switch ($Level.ToUpper()) {
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

function Invoke-SmartM365TeamsNotificationFromLog {
    [CmdletBinding()]
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    if ($script:SmartM365TeamsNotificationInProgress) { return }

    $normalizedLevel = if ([string]::IsNullOrWhiteSpace($Level)) { 'INFO' } else { $Level.ToUpperInvariant() }
    $isError = $normalizedLevel -eq 'ERROR'
    $isTerminalSuccess = $normalizedLevel -eq 'SUCCESS' -and $Message -match '(?i)\b(completed|finished|termin[eé]|complete)\b' -and $Message -notmatch '(?i)\bpreflight completed\b'
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
        [int]$RetentionMaxCsv = 30,
        [int]$RetentionMaxLogs = 30
    )

    $script:SmartM365CoreRunId = $RunId
    $script:SmartM365CoreRunOutputRoot = $RunOutputRoot
    $script:SmartM365CoreLatestOutputRoot = $LatestOutputRoot
    $script:SmartM365CoreLogPath = $LogPath
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

function Write-SmartM365CsvAtomically {
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Data,
        [Parameter(Mandatory)][string]$Path,
        [string[]]$Columns = @(),
        [ValidateSet("ASCII", "BigEndianUnicode", "Default", "OEM", "Unicode", "UTF7", "UTF8", "UTF32")]
        [string]$Encoding = "UTF8",
        [string]$Delimiter = ",",
        [switch]$NoTypeInformation = $true
    )

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

    if (-not $NoSharePointUpload) {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $publishedPath
    }

    return [pscustomobject]@{
        TimestampedPath = $TimestampedPath
        LatestPath      = $LatestPath
        PublishedPath   = $publishedPath
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
        [switch]$NoSharePointUpload
    )

    if ($PSCmdlet.ParameterSetName -eq 'ByBaseName') {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $TimestampedPath = Join-Path $OutputPath "$BaseFileName`_$timestamp.csv"
        $LatestPath = Join-Path $GlobalPath "$BaseFileName.csv"
    }

    Publish-SmartM365Csv -Data $Data -TimestampedPath $TimestampedPath -LatestPath $LatestPath -Columns $Columns -Encoding $Encoding -Delimiter $Delimiter -NoSharePointUpload:$NoSharePointUpload
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
        [string]$Delimiter = ","
    )

    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName -OutputPath $OutputPath -GlobalPath $GlobalPath -Data $Data -NoTypeInformation:$NoTypeInformation -Encoding $Encoding -Delimiter $Delimiter
}

function EnsureExchangePSSnapinLoaded {
    [CmdletBinding()]
    param (
        [string]$SnapinName = "Microsoft.Exchange.Management.PowerShell.SnapIn"
    )

    if (-not (Get-PSSnapin $SnapinName -Registered -ErrorAction SilentlyContinue)) {
        Write-Error "The Exchange Management PSSnapin '$SnapinName' is not registered on this server."
        Write-Error "This script must be run on an Exchange 2016 server where the Management Tools are installed."
        return $false
    }

    if (-not (Get-PSSnapin $SnapinName -ErrorAction SilentlyContinue)) {
        Write-Verbose "The Exchange PSSnapin '$SnapinName' is not loaded in the current session. Attempting to load it..."
        try {
            Add-PSSnapin $SnapinName -ErrorAction Stop
            Write-Verbose "The Exchange PSSnapin was loaded successfully."
        }
        catch {
            Write-Error "Failed to load the Exchange PSSnapin '$SnapinName'. Error: $($_.Exception.Message)"
            Write-Error "Ensure you are running this script on an Exchange 2016 server and have the necessary permissions."
            return $false
        }
    } else {
        Write-Verbose "The Exchange PSSnapin '$SnapinName' is already loaded."
    }

    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Error "The Get-Mailbox cmdlet is still not available after attempting to load the snap-in."
        return $false
    }

    return $true
}

#endregion

#region Mail helpers and file inventory

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
        [string]$AppId = $global:AppId,
        [string]$TenantId = $global:TenantId,
        [string]$Thumbprint = $(if ($global:Thumbprint) { $global:Thumbprint } else { $global:Thumb })
    )

    $toArray = ConvertToRecipientArray -Recipients $To
    $ccArray = if ($Cc) { ConvertToRecipientArray -Recipients $Cc } else { @() }

    if (-not $toArray -or [string]::IsNullOrWhiteSpace($From)) {
        throw "Send-SmartM365GraphMail: missing required parameters (From/To)."
    }

    if (-not (Connect-SmartM365GraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'Graph mail')) {
        throw "Send-SmartM365GraphMail: Microsoft Graph app-only connection failed."
    }

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

    $graphAttachments = @(ConvertTo-SmartM365GraphFileAttachment -Attachments $Attachments)
    if ($graphAttachments.Count -gt 0) {
        $message['attachments'] = $graphAttachments
    }

    $body = @{
        message         = $message
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 12

    $encodedFrom = [System.Uri]::EscapeDataString($From)
    Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $encodedFrom) -Body $body -ContentType 'application/json' | Out-Null
    WriteLog -Message ("Graph mail sent from {0} to {1}" -f $From, ($toArray -join ';')) -Level "SUCCESS"
}

function NewSimpleEmailBody {
    param(
        [string]$Title,
        [string]$Message = "See attached report(s) for details."
    )

    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $body = @"
<html>
  <body style="font-family:Segoe UI,Arial,sans-serif; font-size:13px; color:#222;">
    <h2 style="margin:0 0 10px 0;">$Title</h2>
    <p>Date: $now</p>
    <p>$Message</p>
  </body>
</html>
"@
    return $body
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

    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $rows = foreach ($key in $SummaryData.Keys) {
        "<tr>
            <td style='padding:6px;border:1px solid #ddd;'>$key</td>
            <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$($SummaryData[$key])</td>
         </tr>"
    }

    $body = @"
<html>
  <body style="font-family:Segoe UI,Arial,sans-serif; font-size:13px; color:#222;">
    <h2 style="margin:0 0 10px 0;">$Title</h2>
    <p>Date: $now</p>

    $(if ($Message) { "<p style='margin-top:16px;'>$Message</p>" })

    <h3 style="margin:16px 0 8px 0;">Summary</h3>
    <table style="border-collapse:collapse;border:1px solid #ddd;">
      <thead>
        <tr style="background:#f5f5f5;">
          <th style="padding:6px;border:1px solid #ddd;text-align:left;">Metric</th>
          <th style="padding:6px;border:1px solid #ddd;text-align:right;">Value</th>
        </tr>
      </thead>
      <tbody>
        $($rows -join "`n")
      </tbody>
    </table>
  </body>
</html>
"@
    return $body
}

function NewTableFilesEmailBody {
    param(
        [string]$Title,
        [hashtable]$SummaryData,
        [array]$Files,
        [string]$Message
    )

    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    foreach ($file in $Files) {
        if ($file.FileName -match '\.csv$') {
            try {
                $rowCount = (Get-Content $file.FullName | Measure-Object -Line).Lines - 1
                WriteLog "Row count for '$($file.FileName)': $rowCount" "INFO"
            } catch {
                $rowCount = "N/A"
                WriteLog "Failed to count rows in '$($file.FileName)': $_" "ERROR"
            }
            $file | Add-Member -NotePropertyName RowCount -NotePropertyValue $rowCount -Force
        } else {
            $file | Add-Member -NotePropertyName RowCount -NotePropertyValue "" -Force
        }
    }

    $rows = foreach ($key in $SummaryData.Keys) {
        "<tr>
            <td style='padding:6px;border:1px solid #ddd;'>$key</td>
            <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$($SummaryData[$key])</td>
         </tr>"
    }

    $fileRows = foreach ($file in $Files) {
        "<tr>
            <td style='padding:6px;border:1px solid #ddd;'>$($file.FileName)</td>
            <td style='padding:6px;border:1px solid #ddd;'>$($file.SizeString)</td>
            <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$($file.RowCount)</td>
            <td style='padding:6px;border:1px solid #ddd;'>$($file.CreationTime)</td>
            <td style='padding:6px;border:1px solid #ddd;'>$($file.LastWriteTime)</td>
            <td style='padding:6px;border:1px solid #ddd;'>$($file.Hash)</td>
        </tr>"
    }

    $body = @"
<html>
  <body style="font-family:Segoe UI,Arial,sans-serif; font-size:13px; color:#222;">
    <h2 style="margin:0 0 10px 0;">$Title</h2>
    <p>Date: $now</p>

    $(if ($Message) { "<p style='margin-top:16px;'>$Message</p>" })

    <h3 style="margin:16px 0 8px 0;">Summary</h3>
    <table style="border-collapse:collapse;border:1px solid #ddd;">
      <thead>
        <tr style="background:#f5f5f5;">
          <th style="padding:6px;border:1px solid #ddd;text-align:left;">Metric</th>
          <th style="padding:6px;border:1px solid #ddd;text-align:right;">Value</th>
        </tr>
      </thead>
      <tbody>
        $($rows -join "`n")
      </tbody>
    </table>

    <h3 style="margin:16px 0 8px 0;">File Details</h3>
    <table style="border-collapse:collapse;border:1px solid #ddd;">
      <thead>
        <tr style="background:#f5f5f5;">
          <th style="padding:6px;border:1px solid #ddd;">File Name</th>
          <th style="padding:6px;border:1px solid #ddd;">Size</th>
          <th style="padding:6px;border:1px solid #ddd;text-align:right;">Rows</th>
          <th style="padding:6px;border:1px solid #ddd;">Creation Time</th>
          <th style="padding:6px;border:1px solid #ddd;">Last Write Time</th>
          <th style="padding:6px;border:1px solid #ddd;">Hash</th>
        </tr>
      </thead>
      <tbody>
        $($fileRows -join "`n")
      </tbody>
    </table>
  </body>
</html>
"@
    return $body
}

function SendEmailHtmlReport {
    [CmdletBinding()]
    param(
        [string]$SmtpServer = "",
        [int]$SmtpPort = 25,
        [string]$From = "",
        [string]$To = "",
        [string]$Cc = "",
        [string]$Subject = "SmartM365",
        [string]$BodyHtml,
        [string[]]$Attachments,
        [switch]$VerboseLog
    )

    try {
        $moduleLocalConfig = Get-ModuleLocalConfig
        foreach ($configName in @('SmtpServer','From','To','Cc','Subject')) {
            if (-not $PSBoundParameters.ContainsKey($configName)) {
                $defaultValue = Get-Variable -Name $configName -ValueOnly
                if ($configName -eq 'To') {
                    $defaultValue = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'ErrorMailTo' -DefaultValue $defaultValue
                }
                Set-Variable -Name $configName -Value (Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name $configName -DefaultValue $defaultValue) -Scope Local
            }
        }

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

        if ([string]::IsNullOrWhiteSpace($SmtpServer)) {
            Send-SmartM365GraphMail -From $From -To ($toArray -join ';') -Cc ($ccArray -join ';') -Subject $Subject -BodyHtml $BodyHtml -Attachments $atts
            if ($PSBoundParameters.ContainsKey('VerboseLog')) {
                WriteLog -Message "Email sent to $($toArray -join ';') via Microsoft Graph" -Level "SUCCESS"
            }
            return
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

        Send-MailMessage @mailParams

        if ($PSBoundParameters.ContainsKey('VerboseLog')) {
            WriteLog -Message "Email sent to $($toArray -join ';') via $($SmtpServer):$SmtpPort" -Level "SUCCESS"
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
        [string]$From = "",
        [string]$To = "",
        [string]$Cc = "",
        [string]$Subject = "SmartM365",
        [string]$Body,
        [string]$BodyHtml,
        [string[]]$Attachments,
        [switch]$BodyAsHtml,
        [switch]$HighPriority
    )

    $moduleLocalConfig = Get-ModuleLocalConfig
    foreach ($configName in @('SmtpServer','From','To','Cc','Subject')) {
        if (-not $PSBoundParameters.ContainsKey($configName)) {
            $defaultValue = Get-Variable -Name $configName -ValueOnly
            if ($configName -eq 'To') {
                $defaultValue = Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name 'ErrorMailTo' -DefaultValue $defaultValue
            }
            Set-Variable -Name $configName -Value (Get-ModuleLocalConfigValue -Config $moduleLocalConfig -Name $configName -DefaultValue $defaultValue) -Scope Local
        }
    }

    $htmlBody = if (-not [string]::IsNullOrWhiteSpace($BodyHtml)) { $BodyHtml } else { $Body }

    SendEmailHtmlReport `
        -SmtpServer $SmtpServer `
        -SmtpPort $SmtpPort `
        -From $From `
        -To $To `
        -Cc $Cc `
        -Subject $Subject `
        -BodyHtml $htmlBody `
        -Attachments $Attachments
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
        $body = @"
<html>
  <body style='font-family:Segoe UI,Arial,sans-serif; font-size:13px; color:#222;'>
    <h2 style='margin:0 0 10px 0;'>$Title</h2>
    <p>$Message</p>
  </body>
</html>
"@
        SendEmailHtmlReport -BodyHtml $body
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
        [string]$LogFileName
    )

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
        WriteLog -Message "InitializeScriptEnvironment requires LogAllRootPath from SmartM365.global.local.json." -Level "ERROR"
        throw "InitializeScriptEnvironment requires LogAllRootPath from SmartM365.global.local.json."
    }

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
    if ($global:RetentionMaxLogs -gt 0) {
        RemoveOldFiles -FolderPath $global:LogPath -FilePattern "$LogFileName*.log" -MaxFiles $global:RetentionMaxLogs
    }
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

    Connect-SmartM365GraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'SharePoint upload'
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
        WriteLog -Message "SharePoint upload skipped: local CSV not found: $LocalFilePath" -Level "WARNING"
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
        $site = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $SiteHostname, $SitePath)
        $drives = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}/drives" -f $site.id)
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

        $fileName = [System.IO.Path]::GetFileName($LocalFilePath)
        $targetPath = ConvertTo-GraphDrivePath (Join-Path -Path $TargetFolderPath -ChildPath $fileName)
        $bytes = [System.IO.File]::ReadAllBytes($LocalFilePath)
        $uri = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/content" -f $drive.id, $targetPath
        Invoke-MgGraphRequest -Method PUT -Uri $uri -Body $bytes -ContentType 'application/octet-stream' | Out-Null
        WriteLog -Message "SharePoint CSV uploaded: $TargetFolderPath/$fileName"
    }
    catch {
        WriteLog -Message ("SharePoint upload failed but script continues: {0}" -f $_.Exception.Message) -Level "WARNING"
    }
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
                WriteLog -Message "Copy to global path failed (attempt $attempt/$maxRetries), retrying in $retryDelaySec s... — $_" -Level Warning
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
                    WriteLog -Message "Copy to global path failed (attempt $attempt/$maxRetries), retrying in $retryDelaySec s... — $_" -Level Warning
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
            WriteLog -Message ("Preflight module OK: {0}" -f $moduleName) -Level "INFO"
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
    Format-SmartM365LogLine, Update-SmartM365TimestampedTranscript, WriteLog, Write-Log, Test-FileLocked, RemoveOldFiles, Remove-OldFiles, EnsureExchangePSSnapinLoaded, `
    Set-SmartM365CoreContext, Write-SmartM365CsvAtomically, Publish-SmartM365Csv, Export-SmartM365Csv, Export-SmartM365CsvFromConvert, `
    ConvertToRecipientArray, NewSimpleEmailBody, ConvertBytesToSizeString, GetFileList, `
    NewTableEmailBody, NewTableFilesEmailBody, SendEmailHtmlReport, Send-SmartM365Mail, Send-SmartM365GraphMail, SendFileListEmailReport, Send-SmartM365TeamsNotification, `
    TestSharePath, InitializeScriptEnvironment, Connect-SmartM365GraphAppOnly, Invoke-SmartM365SharePointCsvUpload, `
    ExportAndCopyCsv, ExportAndCopyCsvFromConvert, `
    NewRemoteScheduledTaskAndWait, `
    Invoke-SmartM365Preflight, Connect-SmartM365CloudSession, Disconnect-SmartM365CloudSession


