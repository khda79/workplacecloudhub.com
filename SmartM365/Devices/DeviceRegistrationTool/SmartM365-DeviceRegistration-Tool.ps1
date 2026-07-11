<#
.SYNOPSIS
SmartM365 local device registration diagnostic and repair tool.

.DESCRIPTION
Runs locally on a Windows device to inspect Microsoft Entra device registration
state through dsregcmd /status, Active Directory domain join state, domain
controller reachability, and local Intune enrollment signals.

The tool opens the GUI by default. Use -Cli for command-line mode.

By default the tool is diagnostic-only and does not run dsregcmd /leave.
Repair actions require explicit switches and safety criteria.

.PARAMETER Gui
Opens the WPF interface. This is the default behavior and the switch is kept
for explicit launchers.

.PARAMETER Cli
Runs in command-line mode instead of opening the GUI.

.PARAMETER Mode
User mode is diagnostic-only and suitable for end users. Admin mode unlocks
guarded repair options for support operators.

.PARAMETER RepairDisabledDeletedDevice
Allows dsregcmd /leave only when dsregcmd reports AzureAdJoined=YES,
DeviceAuthStatus=FAILED, and the device appears disabled or deleted.

.PARAMETER TriggerJoin
Triggers the Windows Automatic-Device-Join scheduled task after diagnostics.

.PARAMETER AllowIntuneEnrolledAction
Allows repair actions even when a local Intune enrollment is detected.

.PARAMETER AllowDsregLeave
Allows the broader Hybrid Join repair dsregcmd /leave guard used by SmartM365-Invoke-IntuneHybridJoinRepair.ps1.

.PARAMETER AllowRemoveStaleIntuneEnrollment
Allows removal of stale local Intune enrollment traces after validation.

.PARAMETER AllowRemoveNonIntuneMdmEnrollment
Allows removal of non-Intune MDM enrollment traces after validation.

.PARAMETER TriggerIntuneAutoEnrollment
Triggers deviceenroller.exe /c /AutoEnrollMDM when Hybrid Join is healthy and MDM auto-enrollment policy is configured.

.PARAMETER AuditOnly
Runs diagnostics and reports what would be done without executing repair actions.

.PARAMETER RetryCount
Number of post-action status retries. Default is 0.

.PARAMETER RetrySleepMinutes
Delay between post-action status retries. Default is 5.

.PARAMETER OutputRoot
Root folder for local logs, CSV summaries, transcripts, and dsregcmd snapshots.

.PARAMETER LogRetentionCount
Number of latest local log and diagnostic artifact files to keep per type.

.PARAMETER ConfigPath
Optional path to a JSON configuration file. CLI parameters override JSON values.

.PARAMETER RequireDomainConnectivity
Requires the device to be AD domain joined and connected to a domain controller
with a domain user session before continuing Hybrid Join checks or repair actions.

.PARAMETER DeviceProfile
Diagnostic profile. IntuneOnly skips AD domain controller validation. HybridJoin
requires AD domain join, domain controller reachability, and an AD domain user.

.PARAMETER SupportBundle
Creates a support ZIP bundle with current run logs, diagnostic outputs, CSV, and summary.

.PARAMETER SupportEmail
Support mailbox used when emailing support summaries and support bundles.

.PARAMETER SupportEmailSendMode
Email behavior for support summaries and bundles. Draft displays the Outlook draft. Send sends directly.

.PARAMETER LogoPath
Optional logo path used as the window icon and header logo. Defaults to WorkplaceCloudHub-lockup-WPF.png next to the script.

.PARAMETER DefaultLanguage
GUI language. Use auto, en, fr, or a language available in the language catalog.

.PARAMETER ForceLanguage
Optional GUI language override from configuration or command line. When set, it takes precedence over DefaultLanguage.

.PARAMETER LanguageCatalogPath
Optional path to a PowerShell data file containing localized GUI strings.

.PARAMETER JsonOutput
Writes the CLI result object as JSON instead of the human-readable summary.

.PARAMETER NoTranscript
Disables PowerShell transcript creation.

.EXITCODES
0 = Healthy / success
1 = Error
2 = Not AD domain joined
3 = Attention required / no repair performed / join still pending
4 = Domain controller not reachable
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Gui,
    [switch]$Cli,
    [ValidateSet("User", "Admin")]
    [string]$Mode = "User",
    [switch]$RepairDisabledDeletedDevice,
    [switch]$TriggerJoin,
    [switch]$AllowIntuneEnrolledAction,
    [switch]$AllowDsregLeave,
    [switch]$AllowRemoveStaleIntuneEnrollment,
    [switch]$AllowRemoveNonIntuneMdmEnrollment,
    [switch]$TriggerIntuneAutoEnrollment,
    [switch]$AuditOnly,
    [ValidateRange(0, 12)]
    [int]$RetryCount = 0,
    [ValidateRange(1, 120)]
    [int]$RetrySleepMinutes = 5,
    [string]$OutputRoot = (Join-Path $env:ProgramData "SmartM365\DeviceRegistrationTool"),
    [ValidateRange(0, 365)]
    [int]$LogRetentionCount = 10,
    [string]$ConfigPath,
    [object]$RequireDomainConnectivity = $false,
    [ValidateSet("Auto", "IntuneOnly", "HybridJoin")]
    [string]$DeviceProfile = "Auto",
    [switch]$SupportBundle,
    [string]$SupportEmail = "",
    [ValidateSet("Draft", "Send")]
    [string]$SupportEmailSendMode = "Draft",
    [string]$LogoPath = "",
    [string]$DefaultLanguage = "auto",
    [string]$ForceLanguage = "",
    [string]$LanguageCatalogPath = "",
    [switch]$JsonOutput,
    [switch]$NoTranscript
)

$Script:ToolName = "Smart DeviceRegistration Tool"
$Script:Version = "1.0"
$Script:RunId = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$Script:ComputerName = $env:COMPUTERNAME
$Script:RunLogPath = $null
$Script:ExecutionMode = $Mode
$Script:PolicyState = $null
$Script:UserAdminState = $null
$Script:IntuneEnrollmentState = $null
$Script:MdmEnrollmentState = $null
function ConvertTo-DeviceRegistrationBoolean {
    param(
        [object]$Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value) { return $Default }

    if ($Value -is [bool]) {
        return [bool]$Value
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $Default }

    switch -Regex ($text) {
        '^(1|true|yes|y)$' { return $true }
        '^(0|false|no|n)$' { return $false }
        default { return $Default }
    }
}

$Script:RequireDomainConnectivity = ConvertTo-DeviceRegistrationBoolean -Value $RequireDomainConnectivity -Default $false
$Script:DeviceProfile = $DeviceProfile
$Script:SupportEmail = $SupportEmail
$Script:SupportEmailSendMode = $SupportEmailSendMode
$Script:LogoPath = $LogoPath
$Script:DefaultLanguage = $DefaultLanguage
$Script:ForceLanguage = $ForceLanguage
$Script:LanguageCatalogPath = $LanguageCatalogPath
$Script:StringsCatalog = $null
$Script:Strings = $null
$Script:SelectedLanguage = "en"
$script:CliBoundParameters = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $script:CliBoundParameters[$key] = $true
}

function ConvertTo-SafeArgument {
    param([Parameter(Mandatory = $true)][string]$Value)

    if ($Value -match '[\s"]') {
        return '"' + ($Value -replace '"', '\"') + '"'
    }

    return $Value
}

function Get-FirstExistingPath {
    param([string[]]$Paths)

    foreach ($path in $Paths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }

    return ""
}

function Resolve-DeviceRegistrationToolPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    try {
        if ([IO.Path]::IsPathRooted($Path)) {
            return $Path
        }

        $basePath = $PSScriptRoot
        if ($script:ToolConfig -and $script:ToolConfig.PSObject.Properties.Name -contains "ConfigPath") {
            $configDir = Split-Path -Path $script:ToolConfig.ConfigPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($configDir)) {
                $basePath = $configDir
            }
        }

        return (Join-Path $basePath $Path)
    }
    catch {
        return $Path
    }
}

function Get-EffectiveLogoPath {
    $candidate = Resolve-DeviceRegistrationToolPath -Path $script:LogoPath
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $defaultLogo = Join-Path $PSScriptRoot "WorkplaceCloudHub-lockup-WPF.png"
    if (Test-Path -LiteralPath $defaultLogo) {
        return $defaultLogo
    }

    return ""
}

function Get-DeviceRegistrationLogoImage {
    param([Parameter(Mandatory = $true)][string]$Path)

    $extension = [IO.Path]::GetExtension($Path)
    $uri = New-Object System.Uri($Path)

    if ($extension -ieq ".ico") {
        $decoder = New-Object System.Windows.Media.Imaging.IconBitmapDecoder(
            $uri,
            [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        )

        $frame = $decoder.Frames |
            Sort-Object -Property @{ Expression = { $_.PixelWidth * $_.PixelHeight }; Descending = $true } |
            Select-Object -First 1

        if ($null -ne $frame) {
            if ($frame.CanFreeze) { $frame.Freeze() }
            return $frame
        }
    }

    $image = New-Object System.Windows.Media.Imaging.BitmapImage
    $image.BeginInit()
    $image.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $image.UriSource = $uri
    $image.EndInit()
    if ($image.CanFreeze) { $image.Freeze() }
    return $image
}

function Get-DeviceRegistrationStringsCatalogPath {
    $defaultPath = Join-Path $PSScriptRoot "SmartM365-DeviceRegistration-Tool.strings.psd1"
    if (-not [string]::IsNullOrWhiteSpace($script:LanguageCatalogPath)) {
        return (Resolve-DeviceRegistrationToolPath -Path $script:LanguageCatalogPath)
    }

    return $defaultPath
}

function Import-DeviceRegistrationStringsCatalog {
    $path = Get-DeviceRegistrationStringsCatalogPath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Strings resource file not found: $path"
    }

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw ("Strings resource file could not be parsed: {0}" -f (($parseErrors | ForEach-Object { $_.Message }) -join '; '))
    }

    $statement = $ast.EndBlock.Statements | Select-Object -First 1
    if ($null -eq $statement -or $null -eq $statement.PipelineElements -or $statement.PipelineElements.Count -eq 0) {
        throw "Strings resource file is empty: $path"
    }

    $catalog = $statement.PipelineElements[0].Expression.SafeGetValue()
    if (-not $catalog.ContainsKey("Strings") -or -not $catalog.Strings.ContainsKey("en")) {
        throw "Strings resource file is invalid or missing the English fallback: $path"
    }

    return $catalog
}

function Resolve-DeviceRegistrationLanguage {
    param([Parameter(Mandatory = $true)][hashtable]$Catalog)

    $candidates = New-Object System.Collections.Generic.List[string]
    $configuredLanguage = if (-not [string]::IsNullOrWhiteSpace($script:ForceLanguage)) { [string]$script:ForceLanguage } else { [string]$script:DefaultLanguage }
    if ([string]::IsNullOrWhiteSpace($configuredLanguage)) { $configuredLanguage = "auto" }

    if ($configuredLanguage -eq "auto") {
        try {
            $culture = [Globalization.CultureInfo]::CurrentUICulture
            $candidates.Add($culture.Name)
            $candidates.Add($culture.TwoLetterISOLanguageName)
        }
        catch {
            $candidates.Add("en")
        }
    }
    else {
        $candidates.Add($configuredLanguage)
    }

    $aliases = if ($Catalog.ContainsKey("Aliases")) { $Catalog.Aliases } else { @{} }
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $normalized = $candidate.Trim()
        if ($aliases.ContainsKey($normalized)) { $normalized = [string]$aliases[$normalized] }
        elseif ($aliases.ContainsKey($normalized.ToLowerInvariant())) { $normalized = [string]$aliases[$normalized.ToLowerInvariant()] }

        if ($Catalog.Strings.ContainsKey($normalized)) { return $normalized }

        $baseName = ($normalized -split "-")[0]
        if ($aliases.ContainsKey($baseName)) { $baseName = [string]$aliases[$baseName] }
        if ($Catalog.Strings.ContainsKey($baseName)) { return $baseName }
    }

    if ($Catalog.ContainsKey("DefaultLanguage")) { return [string]$Catalog.DefaultLanguage }
    return "en"
}

function Get-DeviceRegistrationStrings {
    try {
        $catalog = Import-DeviceRegistrationStringsCatalog
        $script:StringsCatalog = $catalog
        $selected = Resolve-DeviceRegistrationLanguage -Catalog $catalog
        $script:SelectedLanguage = $selected

        $english = @{}
        foreach ($key in $catalog.Strings.en.Keys) {
            $english[$key] = $catalog.Strings.en[$key]
        }

        $localized = $catalog.Strings[$selected]
        foreach ($key in $localized.Keys) {
            $english[$key] = $localized[$key]
        }

        if (-not $english.ContainsKey("FlowDirection")) { $english.FlowDirection = "LeftToRight" }
        if (-not $english.ContainsKey("LanguageLabel")) { $english.LanguageLabel = "Language" }
        if (-not $english.ContainsKey("LanguageAuto")) { $english.LanguageAuto = "Automatic" }
        return [pscustomobject]$english
    }
    catch {
        Write-Warning ("Unable to load localized strings: {0}" -f $_.Exception.Message)
        $script:SelectedLanguage = "en"
        return [pscustomobject]@{
            FlowDirection = "LeftToRight"
            WindowTitle = $Script:ToolName
            Eyebrow = "LOCAL DEVICE REGISTRATION"
            Title = $Script:ToolName
            Subtitle = "Hybrid Join, Entra device registration, policy checks, and support-ready diagnostics."
            UserModeDiagnosticOnly = "User mode. Diagnostic only."
            AdminModeAvailable = "Admin mode. Guarded repair options are available."
            RunDiagnostic = "Run diagnostic"
            RunRepair = "Run repair"
            RefreshPrt = "Refresh Azure AD PRT"
            RefreshingPrt = "Refreshing Azure AD PRT..."
            RefreshPrtCompletedFormat = "Azure AD PRT refresh completed. ExitCode={0}; Output={1}"
            TriggerIntuneAutoEnrollment = "Trigger Intune auto-enrollment"
            RemoveStaleIntuneEnrollment = "Remove stale local Intune enrollment trace"
            RemoveNonIntuneMdmEnrollment = "Remove non-Intune MDM enrollment"
            AuditOnly = "Audit only"
            LanguageLabel = "Language"
            LanguageAuto = "Automatic"
        }
    }
}

function Initialize-DeviceRegistrationConfigFromTemplate {
    param(
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [Parameter(Mandatory = $true)][string]$TemplatePath
    )

    if (Test-Path -LiteralPath $ConfigPath) { return $false }
    if (-not (Test-Path -LiteralPath $TemplatePath)) { return $false }

    $parent = Split-Path -Path $ConfigPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    Copy-Item -LiteralPath $TemplatePath -Destination $ConfigPath -ErrorAction Stop
    Write-Host ((@(
        'Created Device Registration Tool local JSON from template.',
        "Local JSON: $ConfigPath",
        "Template: $TemplatePath",
        'Review the generated local JSON values; continuing with default template values unless edited before next run.'
    )) -join [Environment]::NewLine) -ForegroundColor Yellow
    return $true
}

function Get-DeviceRegistrationToolConfig {
    param([string]$Path)

    $programConfig = Join-Path (Join-Path $env:ProgramData "SmartM365\DeviceRegistrationTool") "SmartM365-DeviceRegistration-Tool.config.json"
    $scriptConfig = Join-Path $PSScriptRoot "SmartM365-DeviceRegistration-Tool.config.json"
    $templateConfig = Join-Path $PSScriptRoot "SmartM365-DeviceRegistration-Tool.config.template.json"

    if (-not [string]::IsNullOrWhiteSpace($Path) -and -not (Test-Path -LiteralPath $Path)) {
        Initialize-DeviceRegistrationConfigFromTemplate -ConfigPath $Path -TemplatePath $templateConfig | Out-Null
    }
    elseif (-not (Test-Path -LiteralPath $scriptConfig) -and -not (Test-Path -LiteralPath $programConfig) -and (Test-Path -LiteralPath $templateConfig)) {
        Initialize-DeviceRegistrationConfigFromTemplate -ConfigPath $scriptConfig -TemplatePath $templateConfig | Out-Null
    }

    $effectivePath = Get-FirstExistingPath -Paths @($Path, $scriptConfig, $programConfig, $templateConfig)

    if ([string]::IsNullOrWhiteSpace($effectivePath)) {
        return $null
    }

    try {
        $config = Get-Content -LiteralPath $effectivePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $config | Add-Member -NotePropertyName ConfigPath -NotePropertyValue $effectivePath -Force
        return $config
    }
    catch {
        Write-Warning ("Unable to read configuration file '{0}': {1}" -f $effectivePath, $_.Exception.Message)
        return $null
    }
}
function Get-ConfigValue {
    param(
        [object]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        [object]$CurrentValue,
        [ValidateSet("bool", "int", "string")]
        [string]$Type = "string"
    )

    if ($null -eq $Config) { return $CurrentValue }
    if ($script:CliBoundParameters.ContainsKey($Name)) { return $CurrentValue }

    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $CurrentValue }

    switch ($Type) {
        "bool" { return [System.Convert]::ToBoolean($property.Value) }
        "int" { return [int]$property.Value }
        default { return [string]$property.Value }
    }
}

function Apply-DeviceRegistrationToolConfig {
    param([object]$Config)

    if ($null -eq $Config) { return }

    $script:RetryCount = Get-ConfigValue -Config $Config -Name "RetryCount" -CurrentValue $script:RetryCount -Type int
    $script:RetrySleepMinutes = Get-ConfigValue -Config $Config -Name "RetrySleepMinutes" -CurrentValue $script:RetrySleepMinutes -Type int
    $script:OutputRoot = Get-ConfigValue -Config $Config -Name "OutputRoot" -CurrentValue $script:OutputRoot -Type string
    $script:LogRetentionCount = Get-ConfigValue -Config $Config -Name "LogRetentionCount" -CurrentValue $script:LogRetentionCount -Type int
    $script:RequireDomainConnectivity = Get-ConfigValue -Config $Config -Name "RequireDomainConnectivity" -CurrentValue $script:RequireDomainConnectivity -Type bool
    $script:DeviceProfile = Get-ConfigValue -Config $Config -Name "DeviceProfile" -CurrentValue $script:DeviceProfile -Type string
    $script:SupportEmail = Get-ConfigValue -Config $Config -Name "SupportEmail" -CurrentValue $script:SupportEmail -Type string
    $script:SupportEmailSendMode = Get-ConfigValue -Config $Config -Name "SupportEmailSendMode" -CurrentValue $script:SupportEmailSendMode -Type string
    $script:LogoPath = Get-ConfigValue -Config $Config -Name "LogoPath" -CurrentValue $script:LogoPath -Type string
    $script:DefaultLanguage = Get-ConfigValue -Config $Config -Name "DefaultLanguage" -CurrentValue $script:DefaultLanguage -Type string
    $script:ForceLanguage = Get-ConfigValue -Config $Config -Name "ForceLanguage" -CurrentValue $script:ForceLanguage -Type string
    $script:LanguageCatalogPath = Get-ConfigValue -Config $Config -Name "LanguageCatalogPath" -CurrentValue $script:LanguageCatalogPath -Type string
    if (@("Auto", "IntuneOnly", "HybridJoin") -notcontains $script:DeviceProfile) {
        $script:DeviceProfile = "Auto"
    }
    if (@("Draft", "Send") -notcontains $script:SupportEmailSendMode) {
        $script:SupportEmailSendMode = "Draft"
    }

    if (-not $script:CliBoundParameters.ContainsKey("RequireDomainConnectivity")) {
        if ($script:DeviceProfile -eq "HybridJoin") {
            $script:RequireDomainConnectivity = $true
        }
        elseif ($script:DeviceProfile -eq "IntuneOnly") {
            $script:RequireDomainConnectivity = $false
        }
    }

    $Script:RequireDomainConnectivity = [bool]$script:RequireDomainConnectivity
    $Script:DeviceProfile = [string]$script:DeviceProfile
    $Script:SupportEmail = [string]$script:SupportEmail
    $Script:SupportEmailSendMode = [string]$script:SupportEmailSendMode
    $Script:LogoPath = [string]$script:LogoPath
    $Script:DefaultLanguage = [string]$script:DefaultLanguage
    $Script:ForceLanguage = [string]$script:ForceLanguage
    $Script:LanguageCatalogPath = [string]$script:LanguageCatalogPath
}

function Invoke-Relaunch64BitIfNeeded {
    try {
        if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
            if ($env:SMARTM365_DEVICE_REGISTRATION_RELAUNCHED64 -eq "1") {
                return
            }

            $env:SMARTM365_DEVICE_REGISTRATION_RELAUNCHED64 = "1"
            $powershell64 = Join-Path $env:WINDIR "SysNative\WindowsPowerShell\v1.0\powershell.exe"

            if (-not (Test-Path -LiteralPath $powershell64)) {
                $powershell64 = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
            }

            $arguments = @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass"
            )
            if (-not $Cli) { $arguments += @("-WindowStyle", "Hidden") }
            $arguments += @("-File", (ConvertTo-SafeArgument -Value $PSCommandPath))

            if ($Gui) { $arguments += "-Gui" }
            if ($Cli) { $arguments += "-Cli" }
            $arguments += @("-Mode", $Mode)
            if ($RepairDisabledDeletedDevice) { $arguments += "-RepairDisabledDeletedDevice" }
            if ($TriggerJoin) { $arguments += "-TriggerJoin" }
            if ($AllowIntuneEnrolledAction) { $arguments += "-AllowIntuneEnrolledAction" }
            if ($AllowDsregLeave) { $arguments += "-AllowDsregLeave" }
            if ($AllowRemoveStaleIntuneEnrollment) { $arguments += "-AllowRemoveStaleIntuneEnrollment" }
            if ($AllowRemoveNonIntuneMdmEnrollment) { $arguments += "-AllowRemoveNonIntuneMdmEnrollment" }
            if ($TriggerIntuneAutoEnrollment) { $arguments += "-TriggerIntuneAutoEnrollment" }
            if ($AuditOnly) { $arguments += "-AuditOnly" }
            if ($NoTranscript) { $arguments += "-NoTranscript" }
            $arguments += @("-RetryCount", $RetryCount, "-RetrySleepMinutes", $RetrySleepMinutes)
            $arguments += @("-LogRetentionCount", $LogRetentionCount)
            $arguments += @("-RequireDomainConnectivity", ([string]$Script:RequireDomainConnectivity).ToLowerInvariant())
            $arguments += @("-DeviceProfile", $DeviceProfile)
            if ($SupportBundle) { $arguments += "-SupportBundle" }
            if ($SupportEmail) { $arguments += @("-SupportEmail", (ConvertTo-SafeArgument -Value $SupportEmail)) }
            $arguments += @("-SupportEmailSendMode", $SupportEmailSendMode)
            if ($LogoPath) { $arguments += @("-LogoPath", (ConvertTo-SafeArgument -Value $LogoPath)) }
            if ($Script:DefaultLanguage) { $arguments += @("-DefaultLanguage", (ConvertTo-SafeArgument -Value $Script:DefaultLanguage)) }
            if ($Script:ForceLanguage) { $arguments += @("-ForceLanguage", (ConvertTo-SafeArgument -Value $Script:ForceLanguage)) }
            if ($Script:LanguageCatalogPath) { $arguments += @("-LanguageCatalogPath", (ConvertTo-SafeArgument -Value $Script:LanguageCatalogPath)) }
            if ($JsonOutput) { $arguments += "-JsonOutput" }
            if ($OutputRoot) { $arguments += @("-OutputRoot", (ConvertTo-SafeArgument -Value $OutputRoot)) }
            if ($ConfigPath) { $arguments += @("-ConfigPath", (ConvertTo-SafeArgument -Value $ConfigPath)) }

            $process = Start-Process -FilePath $powershell64 -ArgumentList $arguments -Wait -PassThru
            exit $process.ExitCode
        }
    }
    catch {
        Write-Warning ("64-bit relaunch check failed: {0}" -f $_.Exception.Message)
    }
}

function Invoke-RelaunchStaIfNeeded {
    if ($Cli) {
        return
    }

    try {
        if ([Threading.Thread]::CurrentThread.GetApartmentState() -eq "STA") {
            return
        }

        if ($env:SMARTM365_DEVICE_REGISTRATION_RELAUNCHED_STA -eq "1") {
            return
        }

        $env:SMARTM365_DEVICE_REGISTRATION_RELAUNCHED_STA = "1"
        $hostExe = (Get-Process -Id $PID).Path

        if ([string]::IsNullOrWhiteSpace($hostExe)) {
            $hostExe = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
        }

        $arguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-STA"
        )
        if (-not $Cli) { $arguments += @("-WindowStyle", "Hidden") }
        $arguments += @("-File", (ConvertTo-SafeArgument -Value $PSCommandPath))

        if ($Gui) { $arguments += "-Gui" }
        $arguments += @("-Mode", $Mode)
        if ($RepairDisabledDeletedDevice) { $arguments += "-RepairDisabledDeletedDevice" }
        if ($TriggerJoin) { $arguments += "-TriggerJoin" }
        if ($AllowIntuneEnrolledAction) { $arguments += "-AllowIntuneEnrolledAction" }
        if ($AllowDsregLeave) { $arguments += "-AllowDsregLeave" }
        if ($AllowRemoveStaleIntuneEnrollment) { $arguments += "-AllowRemoveStaleIntuneEnrollment" }
        if ($AllowRemoveNonIntuneMdmEnrollment) { $arguments += "-AllowRemoveNonIntuneMdmEnrollment" }
        if ($TriggerIntuneAutoEnrollment) { $arguments += "-TriggerIntuneAutoEnrollment" }
        if ($AuditOnly) { $arguments += "-AuditOnly" }
        if ($NoTranscript) { $arguments += "-NoTranscript" }
        $arguments += @("-RetryCount", $RetryCount, "-RetrySleepMinutes", $RetrySleepMinutes)
        $arguments += @("-LogRetentionCount", $LogRetentionCount)
        $arguments += @("-RequireDomainConnectivity", ([string]$Script:RequireDomainConnectivity).ToLowerInvariant())
        $arguments += @("-DeviceProfile", $DeviceProfile)
        if ($SupportBundle) { $arguments += "-SupportBundle" }
        if ($SupportEmail) { $arguments += @("-SupportEmail", (ConvertTo-SafeArgument -Value $SupportEmail)) }
        $arguments += @("-SupportEmailSendMode", $SupportEmailSendMode)
        if ($LogoPath) { $arguments += @("-LogoPath", (ConvertTo-SafeArgument -Value $LogoPath)) }
        if ($Script:DefaultLanguage) { $arguments += @("-DefaultLanguage", (ConvertTo-SafeArgument -Value $Script:DefaultLanguage)) }
        if ($Script:ForceLanguage) { $arguments += @("-ForceLanguage", (ConvertTo-SafeArgument -Value $Script:ForceLanguage)) }
        if ($Script:LanguageCatalogPath) { $arguments += @("-LanguageCatalogPath", (ConvertTo-SafeArgument -Value $Script:LanguageCatalogPath)) }
        if ($JsonOutput) { $arguments += "-JsonOutput" }
        if ($OutputRoot) { $arguments += @("-OutputRoot", (ConvertTo-SafeArgument -Value $OutputRoot)) }
        if ($ConfigPath) { $arguments += @("-ConfigPath", (ConvertTo-SafeArgument -Value $ConfigPath)) }

        $process = Start-Process -FilePath $hostExe -ArgumentList $arguments -Wait -PassThru
        exit $process.ExitCode
    }
    catch {
        Write-Warning ("STA relaunch check failed: {0}" -f $_.Exception.Message)
    }
}

function Invoke-RelaunchElevatedIfNeeded {
    if ($Mode -ne "Admin") {
        return
    }

    if (Test-ProcessElevated) {
        return
    }

    if ($env:SMARTM365_DEVICE_REGISTRATION_ELEVATION_REQUESTED -eq "1") {
        return
    }

    $env:SMARTM365_DEVICE_REGISTRATION_ELEVATION_REQUESTED = "1"

    $hostExe = (Get-Process -Id $PID).Path
    if ([string]::IsNullOrWhiteSpace($hostExe)) {
        $hostExe = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
    }

    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-STA"
    )
    if (-not $Cli) { $arguments += @("-WindowStyle", "Hidden") }
    $arguments += @(
        "-File", (ConvertTo-SafeArgument -Value $PSCommandPath),
        "-Mode", "Admin"
    )

    if ($Gui) { $arguments += "-Gui" }
    if ($Cli) { $arguments += "-Cli" }
    if ($RepairDisabledDeletedDevice) { $arguments += "-RepairDisabledDeletedDevice" }
    if ($TriggerJoin) { $arguments += "-TriggerJoin" }
    if ($AllowIntuneEnrolledAction) { $arguments += "-AllowIntuneEnrolledAction" }
    if ($AllowDsregLeave) { $arguments += "-AllowDsregLeave" }
    if ($AllowRemoveStaleIntuneEnrollment) { $arguments += "-AllowRemoveStaleIntuneEnrollment" }
    if ($AllowRemoveNonIntuneMdmEnrollment) { $arguments += "-AllowRemoveNonIntuneMdmEnrollment" }
    if ($TriggerIntuneAutoEnrollment) { $arguments += "-TriggerIntuneAutoEnrollment" }
    if ($AuditOnly) { $arguments += "-AuditOnly" }
    if ($NoTranscript) { $arguments += "-NoTranscript" }
    $arguments += @("-RetryCount", $RetryCount, "-RetrySleepMinutes", $RetrySleepMinutes)
    $arguments += @("-LogRetentionCount", $LogRetentionCount)
    $arguments += @("-RequireDomainConnectivity", ([string]$Script:RequireDomainConnectivity).ToLowerInvariant())
    $arguments += @("-DeviceProfile", $DeviceProfile)
    if ($SupportBundle) { $arguments += "-SupportBundle" }
    if ($SupportEmail) { $arguments += @("-SupportEmail", (ConvertTo-SafeArgument -Value $SupportEmail)) }
    $arguments += @("-SupportEmailSendMode", $SupportEmailSendMode)
    if ($LogoPath) { $arguments += @("-LogoPath", (ConvertTo-SafeArgument -Value $LogoPath)) }
    if ($Script:DefaultLanguage) { $arguments += @("-DefaultLanguage", (ConvertTo-SafeArgument -Value $Script:DefaultLanguage)) }
    if ($Script:ForceLanguage) { $arguments += @("-ForceLanguage", (ConvertTo-SafeArgument -Value $Script:ForceLanguage)) }
    if ($Script:LanguageCatalogPath) { $arguments += @("-LanguageCatalogPath", (ConvertTo-SafeArgument -Value $Script:LanguageCatalogPath)) }
    if ($JsonOutput) { $arguments += "-JsonOutput" }
    if ($OutputRoot) { $arguments += @("-OutputRoot", (ConvertTo-SafeArgument -Value $OutputRoot)) }
    if ($ConfigPath) { $arguments += @("-ConfigPath", (ConvertTo-SafeArgument -Value $ConfigPath)) }

    try {
        Start-Process -FilePath $hostExe -ArgumentList $arguments -Verb RunAs | Out-Null
        exit 0
    }
    catch {
        Write-Warning ("Admin mode requires elevation, but elevation was not completed: {0}" -f $_.Exception.Message)
    }
}

function Initialize-SmartM365DeviceRegistrationOutput {
    param([Parameter(Mandatory = $true)][string]$Root)

    $effectiveRoot = $Root
    $paths = [PSCustomObject]@{
        Root          = $effectiveRoot
        Logs          = Join-Path $effectiveRoot "Logs"
        Output        = Join-Path $effectiveRoot "Output"
        Transcripts   = Join-Path $effectiveRoot "Transcripts"
        CsvPath       = Join-Path $effectiveRoot ("SmartM365_DeviceRegistration_{0}.csv" -f $Script:ComputerName)
        Transcript    = Join-Path (Join-Path $effectiveRoot "Transcripts") ("Transcript_{0}_{1}.txt" -f $Script:ComputerName, $Script:RunId)
        RunLog        = Join-Path (Join-Path $effectiveRoot "Logs") ("SmartM365-DeviceRegistration-Tool_{0}_{1}.log" -f $Script:ComputerName, $Script:RunId)
    }

    try {
        foreach ($path in @($paths.Root, $paths.Logs, $paths.Output, $paths.Transcripts)) {
            if (-not (Test-Path -LiteralPath $path)) {
                New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop | Out-Null
            }
        }

        $writeProbe = Join-Path $paths.Logs (".write-test_{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
        Set-Content -LiteralPath $writeProbe -Value "" -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $writeProbe -Force -ErrorAction SilentlyContinue
    }
    catch {
        $localAppData = [Environment]::GetFolderPath("LocalApplicationData")
        if ([string]::IsNullOrWhiteSpace($localAppData)) {
            $localAppData = $env:TEMP
        }

        $effectiveRoot = Join-Path $localAppData "SmartM365\DeviceRegistrationTool"
        $paths = [PSCustomObject]@{
            Root          = $effectiveRoot
            Logs          = Join-Path $effectiveRoot "Logs"
            Output        = Join-Path $effectiveRoot "Output"
            Transcripts   = Join-Path $effectiveRoot "Transcripts"
            CsvPath       = Join-Path $effectiveRoot ("SmartM365_DeviceRegistration_{0}.csv" -f $Script:ComputerName)
            Transcript    = Join-Path (Join-Path $effectiveRoot "Transcripts") ("Transcript_{0}_{1}.txt" -f $Script:ComputerName, $Script:RunId)
            RunLog        = Join-Path (Join-Path $effectiveRoot "Logs") ("SmartM365-DeviceRegistration-Tool_{0}_{1}.log" -f $Script:ComputerName, $Script:RunId)
        }

        foreach ($path in @($paths.Root, $paths.Logs, $paths.Output, $paths.Transcripts)) {
            if (-not (Test-Path -LiteralPath $path)) {
                New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop | Out-Null
            }
        }

        $writeProbe = Join-Path $paths.Logs (".write-test_{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
        Set-Content -LiteralPath $writeProbe -Value "" -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $writeProbe -Force -ErrorAction SilentlyContinue
    }

    $Script:RunLogPath = $paths.RunLog
    return $paths
}

function Write-SmartM365DeviceRegistrationLog {
    param([Parameter(Mandatory = $true)][string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "{0} [{1}] {2}" -f $timestamp, $Script:ComputerName, $_ })
    $line | ForEach-Object { Write-Verbose $_ }

    if ($Script:RunLogPath) {
        try {
            Add-Content -LiteralPath $Script:RunLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
        }
        catch {
        }
    }
}

function Update-SmartM365DeviceRegistrationTranscript {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $updatedLines = [System.IO.File]::ReadAllLines($Path) | ForEach-Object {
        if ($_ -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b') {
            $_
        }
        elseif ($_ -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s*(.*)$') {
            "{0} {1}" -f $Matches[1], $Matches[2]
        }
        else {
            "{0} {1}" -f $timestamp, $_
        }
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($Path, [string[]]$updatedLines, $utf8NoBom)
}

function Invoke-DeviceRegistrationFileRetention {
    param(
        [Parameter(Mandatory = $true)][psobject]$Paths,
        [ValidateRange(0, 365)][int]$KeepLatest = 10
    )

    if ($KeepLatest -lt 1) {
        Write-SmartM365DeviceRegistrationLog "File retention disabled because LogRetentionCount is 0."
        return
    }

    $sets = @(
        [PSCustomObject]@{ Path = $Paths.Logs; Pattern = "SmartM365-DeviceRegistration-Tool_*.log"; Label = "run logs" },
        [PSCustomObject]@{ Path = $Paths.Transcripts; Pattern = "Transcript_*.txt"; Label = "transcripts" },
        [PSCustomObject]@{ Path = $Paths.Output; Pattern = "*_dsreg_status_*.txt"; Label = "dsreg status snapshots" },
        [PSCustomObject]@{ Path = $Paths.Output; Pattern = "*_dsreg_refreshprt_*.txt"; Label = "dsreg refreshprt outputs" },
        [PSCustomObject]@{ Path = $Paths.Output; Pattern = "*_dsreg_leave_*.txt"; Label = "dsreg leave outputs" },
        [PSCustomObject]@{ Path = $Paths.Output; Pattern = "*_events_*.txt"; Label = "event log exports" },
        [PSCustomObject]@{ Path = $Paths.Output; Pattern = "*_support_summary_*.txt"; Label = "support summaries" },
        [PSCustomObject]@{ Path = $Paths.Output; Pattern = "*_result_*.json"; Label = "JSON results" },
        [PSCustomObject]@{ Path = $Paths.Root; Pattern = "SmartM365_DeviceRegistration_SupportBundle_*.zip"; Label = "support bundles" }
    )

    foreach ($set in $sets) {
        if (-not (Test-Path -LiteralPath $set.Path)) {
            continue
        }

        try {
            $files = @(Get-ChildItem -LiteralPath $set.Path -Filter $set.Pattern -File -ErrorAction Stop | Sort-Object LastWriteTime -Descending)
            if ($files.Count -le $KeepLatest) {
                continue
            }

            $currentRunFiles = @($files | Where-Object { $_.Name -like "*$Script:RunId*" })
            $eligibleFiles = @($files | Where-Object { $_.Name -notlike "*$Script:RunId*" })
            $removeAfter = [Math]::Max(0, $KeepLatest - $currentRunFiles.Count)
            $filesToRemove = @($eligibleFiles | Select-Object -Skip $removeAfter)
            foreach ($file in $filesToRemove) {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            }

            Write-SmartM365DeviceRegistrationLog ("Retention removed {0} old {1}; kept latest {2}." -f $filesToRemove.Count, $set.Label, $KeepLatest)
        }
        catch {
            Write-SmartM365DeviceRegistrationLog ("Retention failed for {0}: {1}" -f $set.Label, $_.Exception.Message)
        }
    }
}

function Export-DeviceRegistrationEventLogs {
    param(
        [Parameter(Mandatory = $true)][psobject]$Paths,
        [ValidateRange(1, 168)][int]$LookbackHours = 24
    )

    $eventLogs = @(
        "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin",
        "Microsoft-Windows-User Device Registration/Admin"
    )
    $startTime = (Get-Date).AddHours(-1 * $LookbackHours)
    $exported = New-Object System.Collections.Generic.List[string]

    foreach ($logName in $eventLogs) {
        $safeName = ($logName -replace "[\\\/: ]", "_")
        $eventPath = Join-Path $Paths.Output ("{0}_events_{1}_{2}.txt" -f $Script:ComputerName, $safeName, $Script:RunId)

        try {
            $events = @(Get-WinEvent -FilterHashtable @{ LogName = $logName; StartTime = $startTime } -ErrorAction Stop | Select-Object -First 80)
            if ($events.Count -eq 0) {
                "No events found since $startTime." | Out-File -LiteralPath $eventPath -Encoding UTF8 -Force
            }
            else {
                $events |
                    Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message |
                    Format-List |
                    Out-String -Width 220 |
                    Out-File -LiteralPath $eventPath -Encoding UTF8 -Force
            }

            $exported.Add($eventPath)
            Write-SmartM365DeviceRegistrationLog ("Event log exported: {0}" -f $eventPath)
        }
        catch {
            Write-SmartM365DeviceRegistrationLog ("Event log export skipped for {0}: {1}" -f $logName, $_.Exception.Message)
        }
    }

    return @($exported)
}

function Format-DeviceRegistrationSupportSummary {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    @(
        ("{0} v{1}" -f $Result.ToolName, $Result.Version),
        ("RunId: {0}" -f $Result.RunId),
        ("Computer: {0}" -f $Result.ComputerName),
        ("Status: {0}" -f $Result.Status),
        ("Message: {0}" -f $Result.Message),
        ("Mode/Profile: {0}/{1}" -f $Result.Mode, $Result.DeviceProfile),
        ("User: {0}; DomainUser={1}; Elevated={2}" -f $Result.CurrentUser, $Result.CurrentUserIsDomainUser, $Result.CurrentProcessElevated),
        ("Domain: Joined={0}; Name={1}; DCReachable={2}; Check={3}" -f $Result.DomainJoined, $Result.DomainName, $Result.DcReachable, $Result.DcCheck),
        ("Intune: Enrolled={0}; Confidence={1}; {2}" -f $Result.IntuneEnrolled, $Result.EnrollmentConfidence, $Result.EnrollmentStatusMessage),
        ("MDM policy: {0}" -f $Result.MdmPolicyStatusMessage),
        ("Dsreg: AzureAdJoined={0}; DeviceAuthStatus={1}; ExitCode={2}" -f $Result.AzureAdJoined, $Result.DeviceAuthStatus, $Result.DsregExitCode),
        ("Errors: Client={0}; ServerSubCode={1}; Phase={2}" -f $Result.ClientErrorCode, $Result.ServerErrorSubCode, $Result.ErrorPhase),
        "",
        "Issues / warnings:",
        (Format-DeviceRegistrationFindingsText -Result $Result),
        "",
        "Information:",
        (Format-DeviceRegistrationInfoText -Result $Result),
        "",
        ("Run log: {0}" -f $Result.LogPath),
        ("CSV: {0}" -f $Result.CsvPath),
        ("Snapshot: {0}" -f $Result.DsregSnapshotPath)
    ) -join [Environment]::NewLine
}

function New-DeviceRegistrationSupportBundle {
    param(
        [Parameter(Mandatory = $true)][psobject]$Paths,
        [Parameter(Mandatory = $true)][psobject]$Result,
        [string[]]$EventLogPaths = @()
    )

    $summaryPath = Join-Path $Paths.Output ("{0}_support_summary_{1}.txt" -f $Script:ComputerName, $Script:RunId)
    Format-DeviceRegistrationSupportSummary -Result $Result | Out-File -LiteralPath $summaryPath -Encoding UTF8 -Force

    $bundlePath = Join-Path $Paths.Root ("SmartM365_DeviceRegistration_SupportBundle_{0}_{1}.zip" -f $Script:ComputerName, $Script:RunId)
    $items = New-Object System.Collections.Generic.List[string]

    foreach ($path in @($Result.LogPath, $Result.CsvPath, $Result.DsregSnapshotPath, $summaryPath)) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            $items.Add($path)
        }
    }

    foreach ($path in $EventLogPaths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            $items.Add($path)
        }
    }

    $runArtifacts = @(Get-ChildItem -LiteralPath $Paths.Output -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$Script:RunId*" })
    foreach ($artifact in $runArtifacts) {
        if (-not ($items -contains $artifact.FullName)) {
            $items.Add($artifact.FullName)
        }
    }

    if ($items.Count -eq 0) {
        return ""
    }

    if (Test-Path -LiteralPath $bundlePath) {
        Remove-Item -LiteralPath $bundlePath -Force -ErrorAction SilentlyContinue
    }

    Compress-Archive -LiteralPath @($items) -DestinationPath $bundlePath -Force
    Write-SmartM365DeviceRegistrationLog ("Support bundle created: {0}" -f $bundlePath)
    return $bundlePath
}

function Send-DeviceRegistrationSupportBundleEmail {
    param(
        [Parameter(Mandatory = $true)][string]$BundlePath,
        [Parameter(Mandatory = $true)][psobject]$Result,
        [Parameter(Mandatory = $true)][string]$To,
        [ValidateSet("Draft", "Send")][string]$SendMode = "Draft"
    )

    if ([string]::IsNullOrWhiteSpace($To)) {
        return [PSCustomObject]@{
            Status  = "NotConfigured"
            Message = "SupportEmail is not configured."
            EmlPath = ""
        }
    }

    if (-not (Test-Path -LiteralPath $BundlePath)) {
        return [PSCustomObject]@{
            Status  = "Failed"
            Message = "Support bundle was not found: $BundlePath"
            EmlPath = ""
        }
    }

    $subject = "[SmartM365 DRT] $($Result.ComputerName) - $($Result.Status) - $($Result.RunId)"
    $body = @(
        "Hello,",
        "",
        "Please find attached the Smart DeviceRegistration Tool support bundle.",
        "",
        (Format-DeviceRegistrationSupportSummary -Result $Result),
        "",
        "Regards"
    ) -join [Environment]::NewLine

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = $To
        $mail.Subject = $subject
        $mail.Body = $body
        [void]$mail.Attachments.Add($BundlePath)

        if ($SendMode -eq "Send") {
            $mail.Send()
            Write-SmartM365DeviceRegistrationLog ("Support bundle email sent to {0}." -f $To)
            return [PSCustomObject]@{
                Status  = "Sent"
                Message = "Support bundle email sent to $To."
                EmlPath = ""
            }
        }

        $mail.Display($false)
        Write-SmartM365DeviceRegistrationLog ("Support bundle email draft opened for {0}." -f $To)
        return [PSCustomObject]@{
            Status  = "DraftOpened"
            Message = "Support bundle email draft opened for $To."
            EmlPath = ""
        }
    }
    catch {
        Write-SmartM365DeviceRegistrationLog ("Outlook email failed: {0}" -f $_.Exception.Message)
        $emlPath = Join-Path (Split-Path -Path $BundlePath -Parent) ("SmartM365_DeviceRegistration_EmailDraft_{0}_{1}.eml" -f $Result.ComputerName, $Result.RunId)
        try {
            $boundary = "----SmartM365DeviceRegistrationBoundary$($Result.RunId)"
            $attachmentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($BundlePath))
            $attachmentName = Split-Path -Path $BundlePath -Leaf
            $encodedAttachment = ($attachmentBytes -split "(.{1,76})" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`r`n"
            $eml = @(
                "To: $To",
                "Subject: $subject",
                "MIME-Version: 1.0",
                "Content-Type: multipart/mixed; boundary=`"$boundary`"",
                "",
                "--$boundary",
                "Content-Type: text/plain; charset=utf-8",
                "Content-Transfer-Encoding: 8bit",
                "",
                $body,
                "",
                "--$boundary",
                "Content-Type: application/zip; name=`"$attachmentName`"",
                "Content-Transfer-Encoding: base64",
                "Content-Disposition: attachment; filename=`"$attachmentName`"",
                "",
                $encodedAttachment,
                "--$boundary--"
            ) -join "`r`n"

            $eml | Out-File -LiteralPath $emlPath -Encoding UTF8 -Force
            return [PSCustomObject]@{
                Status  = "DraftFileCreated"
                Message = "Outlook was not available. Email draft file created: $emlPath"
                EmlPath = $emlPath
            }
        }
        catch {
            return [PSCustomObject]@{
                Status  = "Failed"
                Message = "Support email failed and EML fallback failed: $($_.Exception.Message)"
                EmlPath = ""
            }
        }
    }
}

function Send-DeviceRegistrationSupportSummaryEmail {
    param(
        [Parameter(Mandatory = $true)][psobject]$Result,
        [Parameter(Mandatory = $true)][string]$To,
        [ValidateSet("Draft", "Send")][string]$SendMode = "Draft"
    )

    if ([string]::IsNullOrWhiteSpace($To)) {
        return [PSCustomObject]@{
            Status  = "NotConfigured"
            Message = "SupportEmail is not configured."
            EmlPath = ""
        }
    }

    $subject = "[SmartM365 DRT Summary] $($Result.ComputerName) - $($Result.Status) - $($Result.RunId)"
    $body = @(
        "Hello,",
        "",
        "Please find below the Smart DeviceRegistration Tool support summary.",
        "",
        (Format-DeviceRegistrationSupportSummary -Result $Result),
        "",
        "Regards"
    ) -join [Environment]::NewLine

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = $To
        $mail.Subject = $subject
        $mail.Body = $body

        if ($SendMode -eq "Send") {
            $mail.Send()
            Write-SmartM365DeviceRegistrationLog ("Support summary email sent to {0}." -f $To)
            return [PSCustomObject]@{
                Status  = "Sent"
                Message = "Support summary email sent to $To."
                EmlPath = ""
            }
        }

        $mail.Display($false)
        Write-SmartM365DeviceRegistrationLog ("Support summary email draft opened for {0}." -f $To)
        return [PSCustomObject]@{
            Status  = "DraftOpened"
            Message = "Support summary email draft opened for $To."
            EmlPath = ""
        }
    }
    catch {
        Write-SmartM365DeviceRegistrationLog ("Outlook summary email failed: {0}" -f $_.Exception.Message)
        $root = if (-not [string]::IsNullOrWhiteSpace($Result.CsvPath)) { Split-Path -Path $Result.CsvPath -Parent } else { $env:TEMP }
        $emlPath = Join-Path $root ("SmartM365_DeviceRegistration_SummaryEmailDraft_{0}_{1}.eml" -f $Result.ComputerName, $Result.RunId)
        try {
            $eml = @(
                "To: $To",
                "Subject: $subject",
                "MIME-Version: 1.0",
                "Content-Type: text/plain; charset=utf-8",
                "Content-Transfer-Encoding: 8bit",
                "",
                $body
            ) -join "`r`n"

            $eml | Out-File -LiteralPath $emlPath -Encoding UTF8 -Force
            return [PSCustomObject]@{
                Status  = "DraftFileCreated"
                Message = "Outlook was not available. Summary email draft file created: $emlPath"
                EmlPath = $emlPath
            }
        }
        catch {
            return [PSCustomObject]@{
                Status  = "Failed"
                Message = "Support summary email failed and EML fallback failed: $($_.Exception.Message)"
                EmlPath = ""
            }
        }
    }
}

function Write-AtomicCsvAppend {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][psobject]$RowObject,
        [Parameter(Mandatory = $true)][string]$RunIdValue
    )

    $tmpRow = "{0}.{1}.row.tmp" -f $Path, $RunIdValue
    $tmpNew = "{0}.{1}.new.tmp" -f $Path, $RunIdValue

    try {
        $RowObject | Export-Csv -LiteralPath $tmpRow -NoTypeInformation -Encoding UTF8 -Force

        if (-not (Test-Path -LiteralPath $Path)) {
            Move-Item -LiteralPath $tmpRow -Destination $Path -Force
            return
        }

        $existing = @()
        try {
            $existing = Import-Csv -LiteralPath $Path -ErrorAction Stop
        }
        catch {
            $backup = "{0}.{1}.corrupt.bak" -f $Path, $RunIdValue
            Copy-Item -LiteralPath $Path -Destination $backup -Force
        }

        $newRow = Import-Csv -LiteralPath $tmpRow
        $combined = @()
        if ($existing) { $combined += $existing }
        if ($newRow) { $combined += $newRow }

        $combined | Export-Csv -LiteralPath $tmpNew -NoTypeInformation -Encoding UTF8 -Force
        Move-Item -LiteralPath $tmpNew -Destination $Path -Force
    }
    finally {
        Remove-Item -LiteralPath $tmpRow -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tmpNew -ErrorAction SilentlyContinue
    }
}

function Get-DsregValue {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string[]]$Lines,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    foreach ($name in $Names) {
        $pattern = "^\s*{0}\s*:\s*(.*)\s*$" -f [regex]::Escape($name)
        $line = $Lines | Where-Object { $_ -match $pattern } | Select-Object -First 1

        if ($line -and $line -match $pattern) {
            $value = $matches[1].Trim()
            while ($value.StartsWith(":")) {
                $value = $value.TrimStart(":").Trim()
            }
            return $value
        }
    }

    return ""
}

function ConvertFrom-DsregStatus {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string[]]$Lines)

    [PSCustomObject]@{
        AzureAdJoined       = Get-DsregValue -Lines $Lines -Names @("AzureAdJoined")
        EnterpriseJoined    = Get-DsregValue -Lines $Lines -Names @("EnterpriseJoined")
        DomainJoined        = Get-DsregValue -Lines $Lines -Names @("DomainJoined")
        DeviceId            = Get-DsregValue -Lines $Lines -Names @("DeviceId")
        TenantName          = Get-DsregValue -Lines $Lines -Names @("TenantName")
        TenantId            = Get-DsregValue -Lines $Lines -Names @("TenantId")
        DeviceAuthStatus    = Get-DsregValue -Lines $Lines -Names @("DeviceAuthStatus")
        ClientErrorCode     = Get-DsregValue -Lines $Lines -Names @("Client ErrorCode", "ClientErrorCode")
        ServerErrorCode     = Get-DsregValue -Lines $Lines -Names @("Server ErrorCode", "ServerErrorCode")
        ServerErrorSubCode  = Get-DsregValue -Lines $Lines -Names @("Server ErrorSubCode", "ServerErrorSubCode")
        ServerOperation     = Get-DsregValue -Lines $Lines -Names @("Server Operation", "ServerOperation")
        ServerMessage       = Get-DsregValue -Lines $Lines -Names @("Server Message", "ServerMessage")
        HttpsStatus         = Get-DsregValue -Lines $Lines -Names @("Https Status", "HttpsStatus")
        RequestId           = Get-DsregValue -Lines $Lines -Names @("Request Id", "RequestId")
        ErrorPhase          = Get-DsregValue -Lines $Lines -Names @("Error Phase", "ErrorPhase")
        MdmUrl              = Get-DsregValue -Lines $Lines -Names @("MdmUrl", "MDMUrl")
        MdmTouUrl           = Get-DsregValue -Lines $Lines -Names @("MdmTouUrl", "MDMTouUrl")
        MdmComplianceUrl    = Get-DsregValue -Lines $Lines -Names @("MdmComplianceUrl", "MDMComplianceUrl")
        KeySignTest         = Get-DsregValue -Lines $Lines -Names @("KeySignTest")
        AzureAdPrt          = Get-DsregValue -Lines $Lines -Names @("AzureAdPrt")
    }
}

function ConvertTo-CleanNativeOutput {
    param([object[]]$Output)

    $lines = @($Output | ForEach-Object { [string]$_ })
    $clean = @()
    foreach ($line in $lines) {
        if ($line -match "^\s*At\s+.*:\d+\s+char:\d+") { continue }
        if ($line -match "^\s*\+\s+") { continue }
        if ($line -match "^\s*CategoryInfo\s*:") { continue }
        if ($line -match "^\s*FullyQualifiedErrorId\s*:") { continue }
        if ($line -match "^\s*NativeCommandError\s*$") { continue }
        $clean += $line
    }

    return (($clean -join " ") -replace "\s+", " ").Trim()
}

function Get-DsregStatusSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$DsregcmdPath,
        [Parameter(Mandatory = $true)][string]$OutputDir,
        [Parameter(Mandatory = $true)][string]$Phase
    )

    $snapshotPath = Join-Path $OutputDir ("{0}_dsreg_status_{1}_{2}.txt" -f $Script:ComputerName, $Phase, $Script:RunId)
    Write-SmartM365DeviceRegistrationLog ("Running dsregcmd /status. Phase={0}" -f $Phase)

    $output = & $DsregcmdPath /status 2>&1
    $dsregExitCode = $LASTEXITCODE
    $lines = @($output | ForEach-Object { [string]$_ })
    $lines | Out-File -LiteralPath $snapshotPath -Encoding UTF8 -Force

    if (-not $lines -or $lines.Count -eq 0) {
        Start-Sleep -Seconds 2
        $lines = Get-Content -LiteralPath $snapshotPath -ErrorAction Stop
    }

    if (-not $lines -or $lines.Count -eq 0) {
        throw "dsregcmd /status returned no output. Snapshot=$snapshotPath"
    }

    $parsed = ConvertFrom-DsregStatus -Lines $lines
    $parsed | Add-Member -NotePropertyName SnapshotPath -NotePropertyValue $snapshotPath -Force
    $parsed | Add-Member -NotePropertyName ExitCode -NotePropertyValue $dsregExitCode -Force

    return $parsed
}

function Invoke-DeviceRegistrationPrtRefresh {
    param([Parameter(Mandatory = $true)][string]$OutputRoot)

    $paths = Initialize-SmartM365DeviceRegistrationOutput -Root $OutputRoot
    $dsregcmdPath = Join-Path $env:WINDIR "System32\dsregcmd.exe"
    if (-not (Test-Path -LiteralPath $dsregcmdPath)) {
        $dsregcmdPath = "dsregcmd.exe"
    }

    $outputPath = Join-Path $paths.Output ("{0}_dsreg_refreshprt_{1}.txt" -f $Script:ComputerName, $Script:RunId)
    Write-SmartM365DeviceRegistrationLog "Running dsregcmd /refreshprt for current user session."

    $output = & $dsregcmdPath /refreshprt 2>&1
    $exitCode = $LASTEXITCODE
    $lines = @(
        "Command: dsregcmd /refreshprt",
        "ExitCode: $exitCode",
        "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "",
        ($output | ForEach-Object { [string]$_ })
    )
    $lines | Out-File -LiteralPath $outputPath -Encoding UTF8 -Force
    Write-SmartM365DeviceRegistrationLog ("dsregcmd /refreshprt exit code: {0}; Output={1}" -f $exitCode, $outputPath)

    [PSCustomObject]@{
        ExitCode   = $exitCode
        OutputPath = $outputPath
        Message    = if ($exitCode -eq 0) { "Azure AD PRT refresh completed." } else { "Azure AD PRT refresh returned exit code $exitCode." }
    }
}

function Get-IntuneEnrollmentState {
    $strongEvidence = New-Object System.Collections.Generic.List[string]
    $weakEvidence = New-Object System.Collections.Generic.List[string]
    $inspectionIssues = New-Object System.Collections.Generic.List[string]

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Registry64
        )

        $enrollKey = $base.OpenSubKey("SOFTWARE\Microsoft\Enrollments", $false)
        if ($enrollKey) {
            foreach ($subName in $enrollKey.GetSubKeyNames()) {
                $subKey = $enrollKey.OpenSubKey($subName, $false)
                if (-not $subKey) {
                    continue
                }

                $providerId = [string]$subKey.GetValue("ProviderID", "")
                $discoveryUrl = [string]$subKey.GetValue("DiscoveryServiceFullURL", "")
                $enrollmentState = [string]$subKey.GetValue("EnrollmentState", "")
                $upn = [string]$subKey.GetValue("UPN", "")

                if ($providerId -eq "MS DM Server") {
                    $strongEvidence.Add("Enrollments\$subName ProviderID=MS DM Server")
                }
                elseif ($discoveryUrl -match "enrollment|manage\.microsoft|manage\.microsoft\.com") {
                    $strongEvidence.Add("Enrollments\$subName DiscoveryServiceFullURL=$discoveryUrl")
                }
                elseif (-not [string]::IsNullOrWhiteSpace($upn) -and -not [string]::IsNullOrWhiteSpace($enrollmentState)) {
                    $weakEvidence.Add("Enrollments\$subName UPN present; EnrollmentState=$enrollmentState")
                }
            }
        }

        $omadmKey = $base.OpenSubKey("SOFTWARE\Microsoft\Provisioning\OMADM\Accounts", $false)
        if ($omadmKey) {
            $accountNames = @($omadmKey.GetSubKeyNames())
            if ($accountNames.Count -gt 0) {
                $weakEvidence.Add(("OMADM accounts present: {0}" -f ($accountNames -join ",")))
            }
        }
    }
    catch {
        $inspectionIssues.Add(("Registry inspection failed: {0}" -f $_.Exception.Message))
    }

    try {
        $enterpriseTasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskPath -like "\Microsoft\Windows\EnterpriseMgmt\*" })
        if ($enterpriseTasks.Count -gt 0) {
            $weakEvidence.Add(("EnterpriseMgmt scheduled tasks present: {0}" -f $enterpriseTasks.Count))
        }
    }
    catch {
        $inspectionIssues.Add(("EnterpriseMgmt task inspection failed: {0}" -f $_.Exception.Message))
    }

    try {
        $imeService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue
        if ($imeService) {
            $weakEvidence.Add(("IntuneManagementExtension service present: {0}" -f $imeService.Status))
        }
    }
    catch {
        $inspectionIssues.Add(("IME service inspection failed: {0}" -f $_.Exception.Message))
    }

    $confidence = if ($strongEvidence.Count -gt 0) {
        "Strong"
    }
    elseif ($weakEvidence.Count -gt 0) {
        "Weak"
    }
    else {
        "None"
    }

    $allEvidence = @()
    $allEvidence += @($strongEvidence)
    $allEvidence += @($weakEvidence)
    $allEvidence += @($inspectionIssues)

    [PSCustomObject]@{
        Enrolled       = ($strongEvidence.Count -gt 0)
        Confidence     = $confidence
        StrongEvidence = (@($strongEvidence) -join " | ")
        WeakEvidence   = (@($weakEvidence) -join " | ")
        InspectionIssues = (@($inspectionIssues) -join " | ")
        Evidence       = ($allEvidence -join " | ")
    }
}

function Get-IntuneEnrollmentDetected {
    return [bool](Get-IntuneEnrollmentState).Enrolled
}

function Test-RegistrySubKeyExists {
    param(
        [Parameter(Mandatory = $true)]$BaseKey,
        [Parameter(Mandatory = $true)][string]$SubKeyPath
    )

    try {
        $key = $BaseKey.OpenSubKey($SubKeyPath, $false)
        if ($key) {
            $key.Close()
            return $true
        }
    }
    catch { }

    return $false
}

function Test-EnterpriseMgmtTaskFolderExists {
    param([Parameter(Mandatory = $true)][string]$EnrollmentId)

    try {
        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        [void]$service.GetFolder("\Microsoft\Windows\EnterpriseMgmt\$EnrollmentId")
        return $true
    }
    catch {
        return $false
    }
}

function Get-MdmEnrollmentState {
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Registry64
        )
        $enrollKey = $base.OpenSubKey("SOFTWARE\Microsoft\Enrollments", $false)
        if (-not $enrollKey) {
            return [PSCustomObject]@{
                AnyMdmEnrollmentDetected = $false
                IntuneEnrollmentDetected = $false
                NonIntuneMdmEnrollmentDetected = $false
                EnrollmentCount = 0
                IntuneEnrollmentIds = ""
                NonIntuneEnrollmentIds = ""
                UnconfirmedIntuneEnrollmentIds = ""
                ProviderIds = ""
                EnrollmentDetails = ""
                IgnoredEnrollmentDetails = ""
            }
        }

        $internalProviderIds = @("Deploy Authority", "Cloud Authority", "Local Authority")
        $entries = @()
        $ignoredEntries = @()
        $unconfirmedIntuneEntries = @()

        foreach ($subName in $enrollKey.GetSubKeyNames()) {
            try {
                $sub = $enrollKey.OpenSubKey($subName, $false)
                if (-not $sub) { continue }

                $providerId = [string]$sub.GetValue("ProviderID", "")
                $discoveryServiceFullUrl = [string]$sub.GetValue("DiscoveryServiceFullURL", "")
                $enrollmentType = [string]$sub.GetValue("EnrollmentType", "")
                $upn = [string]$sub.GetValue("UPN", "")
                $aadResourceId = [string]$sub.GetValue("AADResourceID", "")

                $hasProviderId = -not [string]::IsNullOrWhiteSpace($providerId)
                $hasDiscoveryUrl = -not [string]::IsNullOrWhiteSpace($discoveryServiceFullUrl)
                $isInternalProvider = $hasProviderId -and ($internalProviderIds -contains $providerId)
                $isIntuneProvider = ($providerId -eq "MS DM Server")
                $isIntuneDiscovery = ($discoveryServiceFullUrl -match "(?i)enrollment\.manage\.microsoft\.com")
                $isExternalProvider = $hasProviderId -and -not $isInternalProvider -and -not $isIntuneProvider
                $isExternalDiscovery = $hasDiscoveryUrl -and -not $isIntuneDiscovery

                $statusKeyPresent = Test-RegistrySubKeyExists -BaseKey $base -SubKeyPath ("SOFTWARE\Microsoft\Enrollments\Status\{0}" -f $subName)
                $omadmAccountPresent = Test-RegistrySubKeyExists -BaseKey $base -SubKeyPath ("SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\{0}" -f $subName)
                $policyProviderPresent = Test-RegistrySubKeyExists -BaseKey $base -SubKeyPath ("SOFTWARE\Microsoft\PolicyManager\Providers\{0}" -f $subName)
                $enterpriseMgmtTaskPresent = Test-EnterpriseMgmtTaskFolderExists -EnrollmentId $subName

                $evidence = @()
                if ($statusKeyPresent) { $evidence += "StatusKey" }
                if ($omadmAccountPresent) { $evidence += "OMADMAccount" }
                if ($policyProviderPresent) { $evidence += "PolicyProvider" }
                if ($enterpriseMgmtTaskPresent) { $evidence += "EnterpriseMgmtTasks" }
                $evidenceText = ($evidence -join ",")

                $isIntuneCandidate = $isIntuneProvider -or $isIntuneDiscovery
                $isIntuneEnrollment = $isIntuneProvider
                $isNonIntuneMdmEnrollment = $isExternalProvider -or ($isExternalDiscovery -and ($evidence.Count -ge 1))
                $isMdm = $isIntuneEnrollment -or $isNonIntuneMdmEnrollment

                $entry = [PSCustomObject]@{
                    EnrollmentId = $subName
                    ProviderID = $providerId
                    DiscoveryServiceFullURL = $discoveryServiceFullUrl
                    EnrollmentType = $enrollmentType
                    UPN = $upn
                    AADResourceID = $aadResourceId
                    Evidence = $evidenceText
                    IsIntune = $isIntuneEnrollment
                }

                if ($isMdm) {
                    $entries += $entry
                }
                elseif ($isIntuneCandidate) {
                    $unconfirmedIntuneEntries += $entry
                    $ignoredEntries += $entry
                }
                elseif ($hasProviderId -or $hasDiscoveryUrl -or (-not [string]::IsNullOrWhiteSpace($enrollmentType))) {
                    $ignoredEntries += $entry
                }
            }
            catch {
                Write-SmartM365DeviceRegistrationLog ("MDM enrollment subkey inspection failed. EnrollmentId={0}; Error={1}" -f $subName, $_.Exception.Message)
            }
        }

        $intuneEntries = @($entries | Where-Object { $_.IsIntune })
        $nonIntuneEntries = @($entries | Where-Object { -not $_.IsIntune })
        $providerIds = @($entries | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_.ProviderID)) { "<empty>" } else { $_.ProviderID } } | Select-Object -Unique)
        $details = @($entries | ForEach-Object {
            "EnrollmentId={0},ProviderID={1},DiscoveryURL={2},EnrollmentType={3},UPN={4},Evidence={5}" -f $_.EnrollmentId, $_.ProviderID, $_.DiscoveryServiceFullURL, $_.EnrollmentType, $_.UPN, $_.Evidence
        })
        $ignoredDetails = @($ignoredEntries | ForEach-Object {
            "EnrollmentId={0},ProviderID={1},DiscoveryURL={2},EnrollmentType={3},UPN={4},Evidence={5}" -f $_.EnrollmentId, $_.ProviderID, $_.DiscoveryServiceFullURL, $_.EnrollmentType, $_.UPN, $_.Evidence
        })

        return [PSCustomObject]@{
            AnyMdmEnrollmentDetected = ($entries.Count -gt 0)
            IntuneEnrollmentDetected = ($intuneEntries.Count -gt 0)
            NonIntuneMdmEnrollmentDetected = ($nonIntuneEntries.Count -gt 0)
            EnrollmentCount = $entries.Count
            IntuneEnrollmentIds = (($intuneEntries | ForEach-Object { $_.EnrollmentId }) -join ";")
            NonIntuneEnrollmentIds = (($nonIntuneEntries | ForEach-Object { $_.EnrollmentId }) -join ";")
            UnconfirmedIntuneEnrollmentIds = (($unconfirmedIntuneEntries | ForEach-Object { $_.EnrollmentId }) -join ";")
            ProviderIds = ($providerIds -join ";")
            EnrollmentDetails = ($details -join " | ")
            IgnoredEnrollmentDetails = ($ignoredDetails -join " | ")
        }
    }
    catch {
        Write-SmartM365DeviceRegistrationLog ("MDM enrollment inspection failed: {0}" -f $_.Exception.Message)
        return [PSCustomObject]@{
            AnyMdmEnrollmentDetected = $false
            IntuneEnrollmentDetected = $false
            NonIntuneMdmEnrollmentDetected = $false
            EnrollmentCount = 0
            IntuneEnrollmentIds = ""
            NonIntuneEnrollmentIds = ""
            UnconfirmedIntuneEnrollmentIds = ""
            ProviderIds = ""
            EnrollmentDetails = ""
            IgnoredEnrollmentDetails = $_.Exception.Message
        }
    }
}

function Remove-DeviceRegistrationMdmEnrollment {
    param(
        [Parameter(Mandatory = $true)][string]$EnrollmentIds,
        [Parameter(Mandatory = $true)][string]$OutputDirPath,
        [Parameter(Mandatory = $true)][string]$RemovalLabel
    )

    $ids = @($EnrollmentIds -split ";" | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($ids.Count -eq 0) {
        return [PSCustomObject]@{ Success = $false; EnrollmentIds = ""; BackupDir = ""; RemovedItems = ""; Detail = ("No enrollment id was provided for {0}." -f $RemovalLabel) }
    }

    $safeRemovalLabel = ($RemovalLabel.ToLowerInvariant() -replace '[^a-z0-9]+', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($safeRemovalLabel)) { $safeRemovalLabel = "mdm_enrollment" }
    $backupDir = Join-Path $OutputDirPath ("{0}_{1}_backup_{2}" -f $Script:ComputerName, $safeRemovalLabel, $Script:RunId)
    New-Item -ItemType Directory -Path $backupDir -Force -ErrorAction Stop | Out-Null

    $removedItems = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]

    foreach ($id in $ids) {
        $safeId = $id -replace '[\\/:*?"<>|{}]', '_'
        $registryPaths = @(
            "HKLM\SOFTWARE\Microsoft\Enrollments\$id",
            "HKLM\SOFTWARE\Microsoft\Enrollments\Status\$id",
            "HKLM\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$id",
            "HKLM\SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\$id",
            "HKLM\SOFTWARE\Microsoft\PolicyManager\Providers\$id",
            "HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$id",
            "HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Logger\$id",
            "HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\$id",
            "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MDM\Enrollments\$id"
        )

        foreach ($regPath in $registryPaths) {
            $psPath = "Registry::$regPath"
            if (-not (Test-Path -LiteralPath $psPath)) { continue }

            $safeRegName = ($regPath -replace '[\\/:*?"<>|{} ]', '_')
            $backupFile = Join-Path $backupDir ("{0}_{1}.reg" -f $safeId, $safeRegName)
            try {
                $exportOutput = & reg.exe export $regPath $backupFile /y 2>&1
                if ($LASTEXITCODE -ne 0) {
                    $warnings.Add(("Registry export failed before removal. Path={0}; ExitCode={1}; Output={2}" -f $regPath, $LASTEXITCODE, (ConvertTo-CleanNativeOutput -Output $exportOutput)))
                }

                Remove-Item -LiteralPath $psPath -Recurse -Force -ErrorAction Stop
                $removedItems.Add($regPath)
            }
            catch {
                $errors.Add(("Registry removal failed. Path={0}; Error={1}" -f $regPath, $_.Exception.Message))
            }
        }

        $taskPath = "\Microsoft\Windows\EnterpriseMgmt\$id\"
        try {
            $tasks = @(Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue)
            foreach ($task in $tasks) {
                try {
                    Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
                    $removedItems.Add(("ScheduledTask={0}{1}" -f $task.TaskPath, $task.TaskName))
                }
                catch {
                    $errors.Add(("Scheduled task removal failed. Task={0}{1}; Error={2}" -f $task.TaskPath, $task.TaskName, $_.Exception.Message))
                }
            }
        }
        catch {
            Write-SmartM365DeviceRegistrationLog ("Scheduled task enumeration skipped/failed. TaskPath={0}; Error={1}" -f $taskPath, $_.Exception.Message)
        }

        try {
            $schedule = New-Object -ComObject Schedule.Service
            $schedule.Connect()
            $enterpriseMgmtFolder = $schedule.GetFolder("\Microsoft\Windows\EnterpriseMgmt")
            $enterpriseMgmtFolder.DeleteFolder($id, 0)
            $removedItems.Add(("ScheduledTaskFolder=\Microsoft\Windows\EnterpriseMgmt\{0}" -f $id))
        }
        catch {
            Write-SmartM365DeviceRegistrationLog ("Scheduled task folder removal skipped/failed. Folder=\Microsoft\Windows\EnterpriseMgmt\{0}; Error={1}" -f $id, $_.Exception.Message)
        }
    }

    $detail = if ($errors.Count -eq 0) {
        ("{0} registry keys and EnterpriseMgmt tasks removed. Certificates were not removed." -f $RemovalLabel)
    }
    else {
        ($errors -join " | ")
    }
    if ($warnings.Count -gt 0) {
        $detail = "{0} Warnings: {1}" -f $detail, ($warnings -join " | ")
    }

    return [PSCustomObject]@{
        Success = ($errors.Count -eq 0)
        EnrollmentIds = ($ids -join ";")
        BackupDir = $backupDir
        RemovedItems = ($removedItems -join " | ")
        Detail = $detail
    }
}

function Get-RegistryPolicyValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubKeyPath,

        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Registry64
        )

        $key = $base.OpenSubKey($SubKeyPath, $false)
        if (-not $key) {
            return $null
        }

        return $key.GetValue($ValueName, $null)
    }
    catch {
        return $null
    }
}

function Get-ScheduledTaskStateSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName
    )

    try {
        $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
        return [PSCustomObject]@{
            Present = $true
            State   = [string]$task.State
            Detail  = "Task found."
        }
    }
    catch {
        return [PSCustomObject]@{
            Present = $false
            State   = ""
            Detail  = $_.Exception.Message
        }
    }
}

function Get-DeviceRegistrationPolicyState {
    param(
        [bool]$InformationalOnly = $false,
        [string]$InformationalReason = ""
    )

    $mdmSubKey = "SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM"
    $workplaceJoinSubKey = "SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin"

    $autoEnrollMdm = Get-RegistryPolicyValue -SubKeyPath $mdmSubKey -ValueName "AutoEnrollMDM"
    $useAadCredentialType = Get-RegistryPolicyValue -SubKeyPath $mdmSubKey -ValueName "UseAADCredentialType"
    $autoWorkplaceJoin = Get-RegistryPolicyValue -SubKeyPath $workplaceJoinSubKey -ValueName "autoWorkplaceJoin"
    $joinTask = Get-ScheduledTaskStateSafe -TaskPath "\Microsoft\Windows\Workplace Join\" -TaskName "Automatic-Device-Join"

    $credentialType = switch ([string]$useAadCredentialType) {
        "1" { "UserCredential" }
        "2" { "DeviceCredential" }
        default {
            if ($null -eq $useAadCredentialType) { "NotConfigured" } else { "Unknown($useAadCredentialType)" }
        }
    }

    $issues = New-Object System.Collections.Generic.List[string]

    if ([string]$autoEnrollMdm -ne "1") {
        $issues.Add("MDM auto-enrollment policy is missing or disabled.")
    }

    if ($null -eq $useAadCredentialType) {
        $issues.Add("MDM AAD credential type policy is missing.")
    }

    if ([string]$autoWorkplaceJoin -ne "1") {
        $issues.Add("Workplace Join auto-join policy is missing or disabled.")
    }

    if (-not $joinTask.Present) {
        $issues.Add("Automatic-Device-Join scheduled task is missing.")
    }
    elseif ($joinTask.State -eq "Disabled") {
        $issues.Add("Automatic-Device-Join scheduled task is disabled.")
    }

    $mdmPolicyStatus = if ([string]$autoEnrollMdm -eq "1") { "Enabled" } else { "MissingOrDisabled" }
    $mdmPolicyStatusMessage = if ($mdmPolicyStatus -eq "Enabled") {
        "MDM auto-enrollment policy is enabled."
    }
    else {
        "MDM auto-enrollment policy is missing or disabled."
    }

    if ($InformationalOnly -and $mdmPolicyStatus -ne "Enabled" -and -not [string]::IsNullOrWhiteSpace($InformationalReason)) {
        $mdmPolicyStatusMessage = "{0} {1}" -f $mdmPolicyStatusMessage, $InformationalReason
    }

    [PSCustomObject]@{
        MdmPolicyCheckRequired         = (-not $InformationalOnly)
        MdmPolicyStatus                = $mdmPolicyStatus
        MdmPolicyStatusMessage         = $mdmPolicyStatusMessage
        MdmPolicyPresent              = ($null -ne $autoEnrollMdm -or $null -ne $useAadCredentialType)
        MdmAutoEnrollMDM              = if ($null -eq $autoEnrollMdm) { "" } else { [string]$autoEnrollMdm }
        MdmUseAADCredentialType       = if ($null -eq $useAadCredentialType) { "" } else { [string]$useAadCredentialType }
        MdmCredentialType             = $credentialType
        WorkplaceJoinPolicyPresent    = ($null -ne $autoWorkplaceJoin)
        WorkplaceJoinAutoWorkplaceJoin = if ($null -eq $autoWorkplaceJoin) { "" } else { [string]$autoWorkplaceJoin }
        AutomaticDeviceJoinTaskPresent = [bool]$joinTask.Present
        AutomaticDeviceJoinTaskState   = [string]$joinTask.State
        PolicyIssues                  = if ($InformationalOnly) { "" } else { ($issues -join " | ") }
    }
}

function Test-ProcessElevated {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-CurrentUserAdminState {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        $isLocalAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $userName = [string]$identity.Name
        $authority = ""
        $upn = ""

        if ($userName -match "^(?<authority>[^\\]+)\\") {
            $authority = $matches["authority"]
        }

        try {
            $upnOutput = & whoami.exe /upn 2>$null
            if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($upnOutput)) {
                $upn = [string]($upnOutput | Select-Object -First 1)
            }
        }
        catch {
            $upn = ""
        }

        $nonDomainAuthorities = @(
            "",
            $env:COMPUTERNAME,
            "NT AUTHORITY",
            "NT SERVICE",
            "LOCAL SERVICE",
            "NETWORK SERVICE",
            "AzureAD",
            "MicrosoftAccount",
            "WORKGROUP"
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        $hasDomainUpn = ($upn -match "^[^@\s]+@[^@\s]+\.[^@\s]+$")
        $isDomainAuthority = (
            -not [string]::IsNullOrWhiteSpace($authority) -and
            -not ($nonDomainAuthorities -contains $authority)
        )
        $isDomainUser = (
            $isDomainAuthority -and
            (-not [string]::IsNullOrWhiteSpace($env:USERDNSDOMAIN) -or $hasDomainUpn)
        )

        [PSCustomObject]@{
            UserName       = $userName
            UserAuthority  = $authority
            UserDnsDomain  = [string]$env:USERDNSDOMAIN
            UserUpn        = $upn
            IsDomainUser   = [bool]$isDomainUser
            IsLocalAdmin   = [bool]$isLocalAdmin
            IsElevated     = [bool](Test-ProcessElevated)
        }
    }
    catch {
        [PSCustomObject]@{
            UserName       = ""
            UserAuthority  = ""
            UserDnsDomain  = ""
            UserUpn        = ""
            IsDomainUser   = $false
            IsLocalAdmin   = $false
            IsElevated     = $false
        }
    }
}

function Test-DomainControllerReachable {
    param([Parameter(Mandatory = $true)][string]$DomainName)

    $nltest = Join-Path $env:WINDIR "System32\nltest.exe"
    if (-not (Test-Path -LiteralPath $nltest)) {
        return [PSCustomObject]@{
            Reachable = $false
            Method    = "nltest_not_found"
            Detail    = "nltest.exe was not found."
        }
    }

    $scOutput = & $nltest /sc_query:$DomainName 2>&1
    $scExitCode = $LASTEXITCODE
    if ($scExitCode -eq 0) {
        return [PSCustomObject]@{
            Reachable = $true
            Method    = "nltest_sc_query"
            Detail    = "nltest /sc_query returned 0."
        }
    }

    $dcOutput = & $nltest /dsgetdc:$DomainName 2>&1
    $dcExitCode = $LASTEXITCODE
    if ($dcExitCode -eq 0) {
        return [PSCustomObject]@{
            Reachable = $true
            Method    = "nltest_dsgetdc"
            Detail    = "nltest /dsgetdc returned 0."
        }
    }

    return [PSCustomObject]@{
        Reachable = $false
        Method    = "nltest"
        Detail    = "nltest failed. sc_query=$scExitCode; dsgetdc=$dcExitCode; $($scOutput -join ' ') $($dcOutput -join ' ')"
    }
}

function Start-AutomaticDeviceJoinTask {
    $taskPath = "\Microsoft\Windows\Workplace Join\"
    $taskName = "Automatic-Device-Join"

    try {
        $task = Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
        if ($task.State -eq "Disabled") {
            Enable-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction SilentlyContinue | Out-Null
            Write-SmartM365DeviceRegistrationLog "Automatic-Device-Join task was disabled and enable was requested."
        }

        Start-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
        Write-SmartM365DeviceRegistrationLog "Automatic-Device-Join task start requested."

        return [PSCustomObject]@{
            Success = $true
            Detail  = "Automatic-Device-Join task start requested."
        }
    }
    catch {
        Write-SmartM365DeviceRegistrationLog ("Automatic-Device-Join task start failed: {0}" -f $_.Exception.Message)
        return [PSCustomObject]@{
            Success = $false
            Detail  = $_.Exception.Message
        }
    }
}

function Start-IntuneAutoEnrollment {
    $deviceEnrollerPath = Join-Path $env:WINDIR "System32\deviceenroller.exe"
    if (-not (Test-Path -LiteralPath $deviceEnrollerPath)) {
        return [PSCustomObject]@{
            Success = $false
            ToolFound = $false
            ExitCode = ""
            Detail = "deviceenroller.exe not found."
        }
    }

    try {
        Write-SmartM365DeviceRegistrationLog "Running deviceenroller.exe /c /AutoEnrollMDM."
        $enrollOutput = & $deviceEnrollerPath /c /AutoEnrollMDM 2>&1
        $enrollExitCode = $LASTEXITCODE
        $detail = ConvertTo-CleanNativeOutput -Output $enrollOutput

        return [PSCustomObject]@{
            Success = ($enrollExitCode -eq 0)
            ToolFound = $true
            ExitCode = $enrollExitCode
            Detail = $detail
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            ToolFound = $true
            ExitCode = ""
            Detail = $_.Exception.Message
        }
    }
}

function Test-DisabledOrDeletedDeviceState {
    param([Parameter(Mandatory = $true)][psobject]$Dsreg)

    return (
        $Dsreg.AzureAdJoined -eq "YES" -and
        $Dsreg.DeviceAuthStatus -like "*FAILED*" -and
        $Dsreg.DeviceAuthStatus -match "disabled or deleted" -and
        -not [string]::IsNullOrWhiteSpace($Dsreg.DeviceId) -and
        -not [string]::IsNullOrWhiteSpace($Dsreg.TenantId)
    )
}

function Test-DsregDeviceHealthy {
    param(
        [Parameter(Mandatory = $false)][string]$AzureAdJoined,
        [Parameter(Mandatory = $false)][string]$DeviceAuthStatus,
        [Parameter(Mandatory = $false)][string]$KeySignTest
    )

    if ($AzureAdJoined -ne "YES") { return $false }
    if ($DeviceAuthStatus -like "*SUCCESS*") { return $true }
    if ([string]::IsNullOrWhiteSpace($DeviceAuthStatus) -and $KeySignTest -eq "PASSED") { return $true }
    return $false
}

function Test-DsregLeaveApplicable {
    param(
        [Parameter(Mandatory = $true)][psobject]$Dsreg,
        [bool]$StrictDisabledDeletedOnly = $false
    )

    $baseGuardMatched = (
        $Dsreg.AzureAdJoined -eq "YES" -and
        -not [string]::IsNullOrWhiteSpace($Dsreg.DeviceId) -and
        -not [string]::IsNullOrWhiteSpace($Dsreg.TenantId)
    )

    if (-not $baseGuardMatched) {
        return [PSCustomObject]@{ Applicable = $false; Reason = ("Base dsregcmd /leave guard did not match. AzureAdJoined={0}; DeviceIdPresent={1}; TenantIdPresent={2}" -f $Dsreg.AzureAdJoined, (-not [string]::IsNullOrWhiteSpace($Dsreg.DeviceId)), (-not [string]::IsNullOrWhiteSpace($Dsreg.TenantId))) }
    }

    if ($StrictDisabledDeletedOnly -and (Test-DisabledOrDeletedDeviceState -Dsreg $Dsreg)) {
        return [PSCustomObject]@{ Applicable = $true; Reason = "DeviceAuthStatus indicates disabled or deleted device." }
    }

    if (-not $StrictDisabledDeletedOnly) {
        if ([string]$Dsreg.DeviceAuthStatus -like "FAILED*") {
            return [PSCustomObject]@{ Applicable = $true; Reason = ("DeviceAuthStatus starts with FAILED. DeviceAuthStatus={0}" -f $Dsreg.DeviceAuthStatus) }
        }

        if ($Dsreg.AzureAdJoined -eq "YES" -and $Dsreg.KeySignTest -eq "FAILED") {
            return [PSCustomObject]@{ Applicable = $true; Reason = ("KeySignTest is FAILED on an AzureAdJoined device. DeviceAuthStatus={0}" -f $Dsreg.DeviceAuthStatus) }
        }
    }

    return [PSCustomObject]@{ Applicable = $false; Reason = ("No dsregcmd /leave guard matched. AzureAdJoined={0}; DeviceAuthStatus={1}; KeySignTest={2}" -f $Dsreg.AzureAdJoined, $Dsreg.DeviceAuthStatus, $Dsreg.KeySignTest) }
}

function Get-NextActionForDeviceRegistrationStatus {
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [bool]$IntuneEnrolled = $false
    )

    switch ($Status) {
        "HEALTHY" { if ($IntuneEnrolled) { return "NO_ACTION_ALREADY_INTUNE_OR_HEALTHY" }; return "TRIGGER_OR_WAIT_INTUNE_AUTOENROLL" }
        "HEALTHY_AFTER_ACTION" { return "RECHECK_INTUNE_ENROLLMENT" }
        "INTUNE_AUTOENROLL_TRIGGERED" { return "WAIT_AND_RECHECK_INTUNE" }
        "INTUNE_ENROLLED" { return "NO_ACTION_ALREADY_INTUNE" }
        "STALE_INTUNE_ENROLLMENT_REMOVED" { return "RETRY_INTUNE_AUTOENROLL" }
        "STALE_INTUNE_ENROLLMENT_LOCAL" { return "VALIDATE_STALE_INTUNE_TRACE_OR_ENABLE_CLEANUP" }
        "NON_INTUNE_MDM_REMOVED" { return "RETRY_INTUNE_AUTOENROLL" }
        "NON_INTUNE_MDM_ENROLLED" { return "VALIDATE_MDM_PROVIDER_OR_ENABLE_EXPLICIT_CLEANUP" }
        "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" { return "CHECK_GPO_AUTOENROLL" }
        "INTUNE_ENROLLMENT_TOOL_NOT_FOUND" { return "CHECK_WINDOWS_DEVICEENROLLER" }
        "DISABLED_OR_DELETED_DEVICE_DETECTED" { return "ALLOW_DSREG_LEAVE_OR_FIX_ENTRA_DEVICE" }
        "MISSING_DEVICE_HINT_DETECTED" { return "WAIT_AAD_CONNECT_OR_TRIGGER_JOIN" }
        "KEY_SIGN_TEST_FAILED" { return "ALLOW_DSREG_LEAVE_OR_REJOIN" }
        "ACTION_COMPLETED_RECHECK_REQUIRED" { return "WAIT_SYNC_AND_RECHECK" }
        "ACTION_SKIPPED_INTUNE_ENROLLED" { return "NO_ACTION_ALREADY_INTUNE" }
        "ADMIN_ELEVATION_REQUIRED" { return "RUN_AS_ADMIN" }
        "DOMAIN_USER_REQUIRED" { return "LOGON_WITH_DOMAIN_OR_AAD_USER" }
        "NOT_DOMAIN_JOINED" { return "CHECK_DOMAIN_JOIN_OR_PROFILE" }
        "DC_NOT_REACHABLE" { return "FIX_DOMAIN_CONNECTIVITY_OR_VPN" }
        "ERROR" { return "CHECK_LOG" }
        default { if ($IntuneEnrolled) { return "NO_ACTION_ALREADY_INTUNE" }; return "CHECK_DEVICE_REGISTRATION_LOGS" }
    }
}

function Test-MissingDeviceHint {
    param([Parameter(Mandatory = $true)][psobject]$Dsreg)

    return (
        $Dsreg.ServerErrorSubCode -eq "error_missing_device" -or
        $Dsreg.ServerMessage -match "is not found" -or
        $Dsreg.ClientErrorCode -eq "0x801c03f3" -or
        $Dsreg.ServerOperation -eq "DeviceRenew"
    )
}

function Invoke-PostActionRetry {
    param(
        [Parameter(Mandatory = $true)][string]$DsregcmdPath,
        [Parameter(Mandatory = $true)][string]$OutputDir,
        [Parameter(Mandatory = $true)][int]$Count,
        [Parameter(Mandatory = $true)][int]$SleepMinutes,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $last = $null

    for ($index = 1; $index -le $Count; $index++) {
        Write-SmartM365DeviceRegistrationLog ("{0}: retry {1}/{2}. Sleeping {3} minute(s)." -f $Context, $index, $Count, $SleepMinutes)
        Start-Sleep -Seconds ($SleepMinutes * 60)

        $last = Get-DsregStatusSnapshot -DsregcmdPath $DsregcmdPath -OutputDir $OutputDir -Phase ("{0}_retry{1}" -f $Context, $index)
        if (Test-DsregDeviceHealthy -AzureAdJoined $last.AzureAdJoined -DeviceAuthStatus $last.DeviceAuthStatus -KeySignTest $last.KeySignTest) {
            return [PSCustomObject]@{
                Success = $true
                Dsreg   = $last
                Attempts = $index
            }
        }
    }

    return [PSCustomObject]@{
        Success = $false
        Dsreg   = $last
        Attempts = $Count
    }
}

function New-ResultObject {
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][int]$ExitCode,
        [string]$Message = "",
        [psobject]$Dsreg,
        [bool]$IntuneEnrolled = $false,
        [string]$EnrollmentEvidence = "",
        [bool]$DomainJoined = $false,
        [string]$DomainName = "",
        [bool]$DcReachable = $false,
        [string]$DcCheck = "",
        [bool]$LeaveAttempted = $false,
        [string]$LeaveExitCode = "",
        [string]$Action = "None",
        [psobject]$PolicyState,
        [psobject]$UserAdminState,
        [psobject]$MdmEnrollmentState,
        [psobject]$MdmCleanupResult,
        [psobject]$IntuneAutoEnrollResult,
        [string]$LogPath = "",
        [string]$CsvPath = ""
    )

    if (-not $PolicyState) {
        $PolicyState = $Script:PolicyState
    }

    if (-not $UserAdminState) {
        $UserAdminState = $Script:UserAdminState
    }

    if (-not $MdmEnrollmentState) {
        $MdmEnrollmentState = $Script:MdmEnrollmentState
    }

    if (-not $PSBoundParameters.ContainsKey("IntuneEnrolled") -and $Script:IntuneEnrollmentState) {
        $IntuneEnrolled = [bool]$Script:IntuneEnrollmentState.Enrolled
    }

    if ([string]::IsNullOrWhiteSpace($EnrollmentEvidence) -and $Script:IntuneEnrollmentState) {
        $EnrollmentEvidence = [string]$Script:IntuneEnrollmentState.Evidence
    }

    $nextAction = Get-NextActionForDeviceRegistrationStatus -Status $Status -IntuneEnrolled $IntuneEnrolled

    $enrollmentStatusMessage = if ($IntuneEnrolled) {
        "Device enrolled in Intune."
    }
    else {
        "Device not enrolled in Intune."
    }

    $overallHealth = switch ($Status) {
        "HEALTHY" { "Healthy"; break }
        "HEALTHY_AFTER_ACTION" { "Healthy"; break }
        "ACTION_SKIPPED_INTUNE_ENROLLED" { "ManagedRepairBlocked"; break }
        "NOT_DOMAIN_JOINED" { if ($Script:RequireDomainConnectivity) { "HybridPrereqMissing" } else { "ManagedButNotHybridJoined" }; break }
        "DOMAIN_NAME_EMPTY" { "HybridPrereqMissing"; break }
        "DC_NOT_REACHABLE" { "HybridPrereqMissing"; break }
        "DOMAIN_USER_REQUIRED" { "HybridPrereqMissing"; break }
        "ADMIN_ELEVATION_REQUIRED" { "RepairBlocked"; break }
        "ERROR" { "Error"; break }
        default {
            if ($IntuneEnrolled -and -not $Script:RequireDomainConnectivity) { "ManagedNeedsAttention" } else { "AttentionRequired" }
        }
    }

    [PSCustomObject]@{
        ToolName              = $Script:ToolName
        Version               = $Script:Version
        Mode                  = $Script:ExecutionMode
        RunId                 = $Script:RunId
        Timestamp             = Get-Date
        ComputerName          = $Script:ComputerName
        CurrentUser           = if ($UserAdminState) { $UserAdminState.UserName } else { "" }
        CurrentUserAuthority  = if ($UserAdminState) { $UserAdminState.UserAuthority } else { "" }
        CurrentUserDnsDomain  = if ($UserAdminState) { $UserAdminState.UserDnsDomain } else { "" }
        CurrentUserUpn        = if ($UserAdminState) { $UserAdminState.UserUpn } else { "" }
        CurrentUserIsDomainUser = if ($UserAdminState) { $UserAdminState.IsDomainUser } else { $false }
        CurrentUserIsLocalAdmin = if ($UserAdminState) { $UserAdminState.IsLocalAdmin } else { $false }
        CurrentProcessElevated = if ($UserAdminState) { $UserAdminState.IsElevated } else { $false }
        Status                = $Status
        NextAction            = $nextAction
        OverallHealth         = $overallHealth
        ExitCode              = $ExitCode
        Message               = $Message
        Action                = $Action
        ConfigPath            = if ($script:ToolConfig -and $script:ToolConfig.PSObject.Properties.Name -contains "ConfigPath") { $script:ToolConfig.ConfigPath } else { "" }
        LogRetentionCount     = [int]$script:LogRetentionCount
        DeviceProfile         = [string]$Script:DeviceProfile
        RequireDomainConnectivity = [bool]$Script:RequireDomainConnectivity
        DomainJoined          = $DomainJoined
        DomainName            = $DomainName
        DcReachable           = $DcReachable
        DcCheck               = $DcCheck
        IntuneEnrolled        = $IntuneEnrolled
        EnrollmentStatusMessage = $enrollmentStatusMessage
        EnrollmentConfidence  = if ($Script:IntuneEnrollmentState) { $Script:IntuneEnrollmentState.Confidence } else { "" }
        EnrollmentEvidence    = $EnrollmentEvidence
        EnrollmentStrongEvidence = if ($Script:IntuneEnrollmentState) { $Script:IntuneEnrollmentState.StrongEvidence } else { "" }
        EnrollmentWeakEvidence = if ($Script:IntuneEnrollmentState) { $Script:IntuneEnrollmentState.WeakEvidence } else { "" }
        MdmAnyEnrollmentDetected = if ($MdmEnrollmentState) { $MdmEnrollmentState.AnyMdmEnrollmentDetected } else { $false }
        MdmIntuneEnrollmentDetected = if ($MdmEnrollmentState) { $MdmEnrollmentState.IntuneEnrollmentDetected } else { $false }
        MdmNonIntuneEnrollmentDetected = if ($MdmEnrollmentState) { $MdmEnrollmentState.NonIntuneMdmEnrollmentDetected } else { $false }
        MdmEnrollmentCount    = if ($MdmEnrollmentState) { $MdmEnrollmentState.EnrollmentCount } else { 0 }
        MdmIntuneEnrollmentIds = if ($MdmEnrollmentState) { $MdmEnrollmentState.IntuneEnrollmentIds } else { "" }
        MdmNonIntuneEnrollmentIds = if ($MdmEnrollmentState) { $MdmEnrollmentState.NonIntuneEnrollmentIds } else { "" }
        MdmUnconfirmedIntuneEnrollmentIds = if ($MdmEnrollmentState) { $MdmEnrollmentState.UnconfirmedIntuneEnrollmentIds } else { "" }
        MdmProviderIds        = if ($MdmEnrollmentState) { $MdmEnrollmentState.ProviderIds } else { "" }
        MdmEnrollmentDetails  = if ($MdmEnrollmentState) { $MdmEnrollmentState.EnrollmentDetails } else { "" }
        MdmIgnoredEnrollmentDetails = if ($MdmEnrollmentState) { $MdmEnrollmentState.IgnoredEnrollmentDetails } else { "" }
        MdmPolicyCheckRequired = if ($PolicyState -and $PolicyState.PSObject.Properties.Name -contains "MdmPolicyCheckRequired") { $PolicyState.MdmPolicyCheckRequired } else { $true }
        MdmPolicyStatus       = if ($PolicyState -and $PolicyState.PSObject.Properties.Name -contains "MdmPolicyStatus") { $PolicyState.MdmPolicyStatus } else { "" }
        MdmPolicyStatusMessage = if ($PolicyState -and $PolicyState.PSObject.Properties.Name -contains "MdmPolicyStatusMessage") { $PolicyState.MdmPolicyStatusMessage } else { "" }
        LeaveAttempted        = $LeaveAttempted
        LeaveExitCode         = $LeaveExitCode
        MdmPolicyPresent      = if ($PolicyState) { $PolicyState.MdmPolicyPresent } else { $false }
        MdmAutoEnrollMDM      = if ($PolicyState) { $PolicyState.MdmAutoEnrollMDM } else { "" }
        MdmUseAADCredentialType = if ($PolicyState) { $PolicyState.MdmUseAADCredentialType } else { "" }
        MdmCredentialType     = if ($PolicyState) { $PolicyState.MdmCredentialType } else { "" }
        WorkplaceJoinPolicyPresent = if ($PolicyState) { $PolicyState.WorkplaceJoinPolicyPresent } else { $false }
        WorkplaceJoinAutoWorkplaceJoin = if ($PolicyState) { $PolicyState.WorkplaceJoinAutoWorkplaceJoin } else { "" }
        AutomaticDeviceJoinTaskPresent = if ($PolicyState) { $PolicyState.AutomaticDeviceJoinTaskPresent } else { $false }
        AutomaticDeviceJoinTaskState = if ($PolicyState) { $PolicyState.AutomaticDeviceJoinTaskState } else { "" }
        PolicyIssues          = if ($PolicyState) { $PolicyState.PolicyIssues } else { "" }
        AzureAdJoined         = if ($Dsreg) { $Dsreg.AzureAdJoined } else { "" }
        EnterpriseJoined      = if ($Dsreg) { $Dsreg.EnterpriseJoined } else { "" }
        DsregDomainJoined     = if ($Dsreg) { $Dsreg.DomainJoined } else { "" }
        DeviceId              = if ($Dsreg) { $Dsreg.DeviceId } else { "" }
        TenantName            = if ($Dsreg) { $Dsreg.TenantName } else { "" }
        TenantId              = if ($Dsreg) { $Dsreg.TenantId } else { "" }
        DeviceAuthStatus      = if ($Dsreg) { $Dsreg.DeviceAuthStatus } else { "" }
        AzureAdPrt            = if ($Dsreg) { $Dsreg.AzureAdPrt } else { "" }
        ClientErrorCode       = if ($Dsreg) { $Dsreg.ClientErrorCode } else { "" }
        ServerErrorCode       = if ($Dsreg) { $Dsreg.ServerErrorCode } else { "" }
        ServerErrorSubCode    = if ($Dsreg) { $Dsreg.ServerErrorSubCode } else { "" }
        ServerOperation       = if ($Dsreg) { $Dsreg.ServerOperation } else { "" }
        ServerMessage         = if ($Dsreg) { $Dsreg.ServerMessage } else { "" }
        HttpsStatus           = if ($Dsreg) { $Dsreg.HttpsStatus } else { "" }
        RequestId             = if ($Dsreg) { $Dsreg.RequestId } else { "" }
        ErrorPhase            = if ($Dsreg) { $Dsreg.ErrorPhase } else { "" }
        KeySignTest           = if ($Dsreg) { $Dsreg.KeySignTest } else { "" }
        MdmUrl                = if ($Dsreg) { $Dsreg.MdmUrl } else { "" }
        MdmTouUrl             = if ($Dsreg) { $Dsreg.MdmTouUrl } else { "" }
        MdmComplianceUrl      = if ($Dsreg) { $Dsreg.MdmComplianceUrl } else { "" }
        DsregExitCode         = if ($Dsreg -and $Dsreg.PSObject.Properties.Name -contains "ExitCode") { $Dsreg.ExitCode } else { "" }
        DsregSnapshotPath     = if ($Dsreg -and $Dsreg.PSObject.Properties.Name -contains "SnapshotPath") { $Dsreg.SnapshotPath } else { "" }
        MdmCleanupAttempted   = [bool]$MdmCleanupResult
        MdmCleanupSuccess     = if ($MdmCleanupResult) { $MdmCleanupResult.Success } else { $false }
        MdmCleanupEnrollmentIds = if ($MdmCleanupResult) { $MdmCleanupResult.EnrollmentIds } else { "" }
        MdmCleanupBackupDir   = if ($MdmCleanupResult) { $MdmCleanupResult.BackupDir } else { "" }
        MdmCleanupRemovedItems = if ($MdmCleanupResult) { $MdmCleanupResult.RemovedItems } else { "" }
        MdmCleanupDetail      = if ($MdmCleanupResult) { $MdmCleanupResult.Detail } else { "" }
        IntuneAutoEnrollAttempted = [bool]$IntuneAutoEnrollResult
        IntuneAutoEnrollSuccess = if ($IntuneAutoEnrollResult) { $IntuneAutoEnrollResult.Success } else { $false }
        IntuneAutoEnrollExitCode = if ($IntuneAutoEnrollResult) { $IntuneAutoEnrollResult.ExitCode } else { "" }
        IntuneAutoEnrollDetail = if ($IntuneAutoEnrollResult) { $IntuneAutoEnrollResult.Detail } else { "" }
        SupportBundlePath     = ""
        SupportEmail          = [string]$Script:SupportEmail
        SupportEmailSendMode  = [string]$Script:SupportEmailSendMode
        SupportEmailStatus    = ""
        SupportEmailMessage   = ""
        SupportEmailDraftPath = ""
        SupportSummaryEmailStatus = ""
        SupportSummaryEmailMessage = ""
        SupportSummaryEmailDraftPath = ""
        JsonOutputPath        = ""
        LogPath               = $LogPath
        CsvPath               = $CsvPath
    }
}

function Write-ResultSummary {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    Write-Host ""
    Write-Host ("{0} v{1}" -f $Result.ToolName, $Result.Version) -ForegroundColor Cyan
    Write-Host ("Mode         : {0}" -f $Result.Mode)
    Write-Host ("Status       : {0}" -f $Result.Status)
    Write-Host ("Next action  : {0}" -f $Result.NextAction)
    Write-Host ("Message      : {0}" -f $Result.Message)
    Write-Host ("Computer     : {0}" -f $Result.ComputerName)
    Write-Host ("User         : {0}; DomainUser={1}; LocalAdmin={2}; Elevated={3}" -f $Result.CurrentUser, $Result.CurrentUserIsDomainUser, $Result.CurrentUserIsLocalAdmin, $Result.CurrentProcessElevated)
    Write-Host ("Profile      : {0}; DomainRequired={1}" -f $Result.DeviceProfile, $Result.RequireDomainConnectivity)
    Write-Host ("Domain       : Joined={0}; Name={1}; DCReachable={2}" -f $Result.DomainJoined, $Result.DomainName, $Result.DcReachable)
    Write-Host ("Enrollment  : {0}; Confidence={1}" -f $Result.EnrollmentStatusMessage, $Result.EnrollmentConfidence)
    if (-not [string]::IsNullOrWhiteSpace($Result.EnrollmentEvidence)) {
        Write-Host ("Evidence    : {0}" -f $Result.EnrollmentEvidence)
    }
    Write-Host ("MDM policy   : {0}" -f $Result.MdmPolicyStatusMessage)
    Write-Host ("MDM state    : Intune={0}; NonIntune={1}; UnconfirmedIntuneIds={2}" -f $Result.MdmIntuneEnrollmentDetected, $Result.MdmNonIntuneEnrollmentDetected, $Result.MdmUnconfirmedIntuneEnrollmentIds)
    Write-Host ("Policy       : MDM={0}; Credential={1}; WorkplaceJoin={2}; AutoJoinTask={3}/{4}" -f $Result.MdmAutoEnrollMDM, $Result.MdmCredentialType, $Result.WorkplaceJoinAutoWorkplaceJoin, $Result.AutomaticDeviceJoinTaskPresent, $Result.AutomaticDeviceJoinTaskState)
    if (-not [string]::IsNullOrWhiteSpace($Result.PolicyIssues)) {
        Write-Host ("Policy issues: {0}" -f $Result.PolicyIssues)
    }
    Write-Host ("Dsreg        : AzureAdJoined={0}; DeviceAuthStatus={1}; KeySignTest={2}; ExitCode={3}" -f $Result.AzureAdJoined, $Result.DeviceAuthStatus, $Result.KeySignTest, $Result.DsregExitCode)
    Write-Host ("Errors       : Client={0}; ServerSubCode={1}; Phase={2}" -f $Result.ClientErrorCode, $Result.ServerErrorSubCode, $Result.ErrorPhase)
    Write-Host ("Action       : {0}; LeaveAttempted={1}; LeaveExitCode={2}" -f $Result.Action, $Result.LeaveAttempted, $Result.LeaveExitCode)
    Write-Host ("Run log      : {0}" -f $Result.LogPath)
    Write-Host ("CSV summary  : {0}" -f $Result.CsvPath)
}

function Invoke-SmartM365DeviceRegistrationTool {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [ValidateSet("User", "Admin")]
        [string]$Mode = "User",
        [switch]$RepairDisabledDeletedDevice,
        [switch]$TriggerJoin,
        [switch]$AllowIntuneEnrolledAction,
        [switch]$AllowDsregLeave,
        [switch]$AllowRemoveStaleIntuneEnrollment,
        [switch]$AllowRemoveNonIntuneMdmEnrollment,
        [switch]$TriggerIntuneAutoEnrollment,
        [switch]$AuditOnly,
        [ValidateRange(0, 12)]
        [int]$RetryCount = 0,
        [ValidateRange(1, 120)]
        [int]$RetrySleepMinutes = 5,
        [ValidateRange(0, 365)]
        [int]$LogRetentionCount = 10,
        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,
        [switch]$SupportBundle,
        [switch]$NoTranscript
    )

    $Script:ExecutionMode = $Mode
    if ($Mode -eq "User") {
        $RepairDisabledDeletedDevice = $false
        $TriggerJoin = $false
        $AllowIntuneEnrolledAction = $false
        $AllowDsregLeave = $false
        $AllowRemoveStaleIntuneEnrollment = $false
        $AllowRemoveNonIntuneMdmEnrollment = $false
        $TriggerIntuneAutoEnrollment = $false
        $AuditOnly = $false
        $RetryCount = 0
    }

    $paths = Initialize-SmartM365DeviceRegistrationOutput -Root $OutputRoot
    $Script:UserAdminState = Get-CurrentUserAdminState
    $transcriptStarted = $false
    $dsreg = $null
    $result = $null
    $intuneState = $null
    $mdmEnrollmentState = $null
    $mdmCleanupResult = $null
    $intuneAutoEnrollResult = $null
    $intuneDetected = $false
    $eventLogPaths = @()

    try {
        Write-SmartM365DeviceRegistrationLog ("Start. RunId={0}; Mode={1}; RepairDisabledDeletedDevice={2}; AllowDsregLeave={3}; TriggerJoin={4}; TriggerIntuneAutoEnrollment={5}; AllowRemoveStaleIntuneEnrollment={6}; AllowRemoveNonIntuneMdmEnrollment={7}; AllowIntuneEnrolledAction={8}; AuditOnly={9}; RetryCount={10}; RetrySleepMinutes={11}" -f $Script:RunId, $Mode, [bool]$RepairDisabledDeletedDevice, [bool]$AllowDsregLeave, [bool]$TriggerJoin, [bool]$TriggerIntuneAutoEnrollment, [bool]$AllowRemoveStaleIntuneEnrollment, [bool]$AllowRemoveNonIntuneMdmEnrollment, [bool]$AllowIntuneEnrolledAction, [bool]$AuditOnly, $RetryCount, $RetrySleepMinutes)
        Write-SmartM365DeviceRegistrationLog ("Retention. LogRetentionCount={0}" -f $LogRetentionCount)

        if (-not $NoTranscript) {
            try {
                Start-Transcript -LiteralPath $paths.Transcript -Force | Out-Null
                $transcriptStarted = $true
                Write-SmartM365DeviceRegistrationLog ("Transcript started: {0}" -f $paths.Transcript)
            }
            catch {
                Write-SmartM365DeviceRegistrationLog ("Transcript start failed: {0}" -f $_.Exception.Message)
            }
        }

        $dsregcmdPath = Join-Path $env:WINDIR "System32\dsregcmd.exe"
        if (-not (Test-Path -LiteralPath $dsregcmdPath)) {
            $dsregcmdPath = "dsregcmd.exe"
        }

        $intuneState = Get-IntuneEnrollmentState
        $Script:IntuneEnrollmentState = $intuneState
        $intuneDetected = [bool]$intuneState.Enrolled
        Write-SmartM365DeviceRegistrationLog ("Intune enrollment detected: {0}; Evidence={1}" -f $intuneDetected, $intuneState.Evidence)

        $mdmEnrollmentState = Get-MdmEnrollmentState
        $Script:MdmEnrollmentState = $mdmEnrollmentState
        if ($mdmEnrollmentState.IntuneEnrollmentDetected) {
            $intuneDetected = $true
        }
        Write-SmartM365DeviceRegistrationLog ("MDM enrollment state: Intune={0}; NonIntune={1}; UnconfirmedIntuneIds={2}; ProviderIds={3}" -f $mdmEnrollmentState.IntuneEnrollmentDetected, $mdmEnrollmentState.NonIntuneMdmEnrollmentDetected, $mdmEnrollmentState.UnconfirmedIntuneEnrollmentIds, $mdmEnrollmentState.ProviderIds)

        $Script:PolicyState = Get-DeviceRegistrationPolicyState -InformationalOnly:$intuneDetected -InformationalReason "Device is already enrolled in Intune, so missing auto-enrollment policy is informational only."
        if ($intuneDetected) {
            Write-SmartM365DeviceRegistrationLog "MDM auto-enrollment policy checked in informational mode because Intune enrollment is already detected."
        }

        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $isDomainJoined = [bool]$computerSystem.PartOfDomain
        $domainName = [string]$computerSystem.Domain
        $dcReachable = $false
        $dcCheck = if ($Script:RequireDomainConnectivity) { "" } else { "NotRequired" }

        if ($Script:RequireDomainConnectivity -and -not $isDomainJoined) {
            $result = New-ResultObject -Status "NOT_DOMAIN_JOINED" -ExitCode 2 -Message "Device is not joined to an Active Directory domain." -IntuneEnrolled $intuneDetected -EnrollmentEvidence $intuneState.Evidence -DomainJoined $false -DomainName $domainName -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        if ($Script:RequireDomainConnectivity -and [string]::IsNullOrWhiteSpace($domainName)) {
            $result = New-ResultObject -Status "DOMAIN_NAME_EMPTY" -ExitCode 1 -Message "Device reports PartOfDomain=True but the domain name is empty." -IntuneEnrolled $intuneDetected -EnrollmentEvidence $intuneState.Evidence -DomainJoined $true -DomainName $domainName -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        if ($Script:RequireDomainConnectivity) {
            $dcTest = Test-DomainControllerReachable -DomainName $domainName
            $dcReachable = [bool]$dcTest.Reachable
            $dcCheck = [string]$dcTest.Method
            if (-not $dcTest.Reachable) {
                $result = New-ResultObject -Status "DC_NOT_REACHABLE" -ExitCode 4 -Message $dcTest.Detail -IntuneEnrolled $intuneDetected -EnrollmentEvidence $intuneState.Evidence -DomainJoined $true -DomainName $domainName -DcReachable $false -DcCheck $dcTest.Method -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                return $result
            }

            if (-not $Script:UserAdminState.IsDomainUser) {
                $message = "Domain connectivity is required, but the current Windows session is not running as an AD domain user."
                $result = New-ResultObject -Status "DOMAIN_USER_REQUIRED" -ExitCode 3 -Message $message -IntuneEnrolled $intuneDetected -EnrollmentEvidence $intuneState.Evidence -DomainJoined $true -DomainName $domainName -DcReachable $true -DcCheck $dcCheck -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                return $result
            }
        }
        else {
            Write-SmartM365DeviceRegistrationLog "AD domain connectivity is not required by configuration. Domain controller check skipped."
        }

        $dsreg = Get-DsregStatusSnapshot -DsregcmdPath $dsregcmdPath -OutputDir $paths.Output -Phase "initial"

        $deviceJoinHealthy = Test-DsregDeviceHealthy -AzureAdJoined $dsreg.AzureAdJoined -DeviceAuthStatus $dsreg.DeviceAuthStatus -KeySignTest $dsreg.KeySignTest
        $repairActionRequested = [bool]($RepairDisabledDeletedDevice -or $AllowDsregLeave -or $TriggerJoin -or $TriggerIntuneAutoEnrollment -or $AllowRemoveStaleIntuneEnrollment -or $AllowRemoveNonIntuneMdmEnrollment)

        if ($intuneDetected -and -not $AllowIntuneEnrolledAction -and $repairActionRequested) {
            $result = New-ResultObject -Status "ACTION_SKIPPED_INTUNE_ENROLLED" -ExitCode 3 -Message "Local Intune enrollment was detected. Repair action skipped unless -AllowIntuneEnrolledAction is used." -Dsreg $dsreg -IntuneEnrolled $true -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "DiagnosticOnly" -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        if ($Mode -eq "Admin" -and $repairActionRequested -and -not (Test-ProcessElevated)) {
            $result = New-ResultObject -Status "ADMIN_ELEVATION_REQUIRED" -ExitCode 3 -Message "Admin mode repair actions require an elevated PowerShell process." -Dsreg $dsreg -IntuneEnrolled $intuneDetected -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "SkippedNotElevated" -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        if ($AuditOnly) {
            $auditStatus = if ($intuneDetected) {
                "AUDIT_SUCCESS_ALREADY_INTUNE"
            }
            elseif ($deviceJoinHealthy -and -not [string]::IsNullOrWhiteSpace($mdmEnrollmentState.UnconfirmedIntuneEnrollmentIds)) {
                "AUDIT_STALE_INTUNE_ENROLLMENT_LOCAL"
            }
            elseif ($deviceJoinHealthy) {
                "AUDIT_INTUNE_MISSING"
            }
            else {
                "AUDIT_HYBRID_JOIN_UNHEALTHY"
            }
            $result = New-ResultObject -Status $auditStatus -ExitCode 3 -Message "Audit only. No repair action was executed." -Dsreg $dsreg -IntuneEnrolled $intuneDetected -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "AuditOnly" -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        if ($deviceJoinHealthy) {
            if ($intuneDetected) {
                $result = New-ResultObject -Status "HEALTHY" -ExitCode 0 -Message "Device registration is healthy and local Intune enrollment is detected." -Dsreg $dsreg -IntuneEnrolled $true -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                return $result
            }

            if ($mdmEnrollmentState.NonIntuneMdmEnrollmentDetected) {
                if ($Mode -eq "Admin" -and $AllowRemoveNonIntuneMdmEnrollment) {
                    if ($PSCmdlet.ShouldProcess("Non-Intune MDM enrollment on $Script:ComputerName", "Remove local non-Intune MDM traces")) {
                        $mdmCleanupResult = Remove-DeviceRegistrationMdmEnrollment -EnrollmentIds $mdmEnrollmentState.NonIntuneEnrollmentIds -OutputDirPath $paths.Output -RemovalLabel "Non-Intune MDM enrollment"
                    }

                    $cleanupStatus = if ($mdmCleanupResult -and $mdmCleanupResult.Success) { "NON_INTUNE_MDM_REMOVED" } else { "NON_INTUNE_MDM_REMOVE_FAILED" }
                    $result = New-ResultObject -Status $cleanupStatus -ExitCode 3 -Message "Hybrid Join is healthy, but non-Intune MDM traces were handled." -Dsreg $dsreg -IntuneEnrolled $false -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "RemoveNonIntuneMdmEnrollment" -MdmCleanupResult $mdmCleanupResult -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                    return $result
                }

                $result = New-ResultObject -Status "NON_INTUNE_MDM_ENROLLED" -ExitCode 3 -Message "Hybrid Join is healthy, but a non-Intune MDM enrollment is present. Cleanup is explicit opt-in." -Dsreg $dsreg -IntuneEnrolled $false -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                return $result
            }

            if (-not [string]::IsNullOrWhiteSpace($mdmEnrollmentState.UnconfirmedIntuneEnrollmentIds)) {
                if ($Mode -eq "Admin" -and $AllowRemoveStaleIntuneEnrollment) {
                    if ($PSCmdlet.ShouldProcess("Stale local Intune enrollment trace on $Script:ComputerName", "Remove stale local Intune traces")) {
                        $mdmCleanupResult = Remove-DeviceRegistrationMdmEnrollment -EnrollmentIds $mdmEnrollmentState.UnconfirmedIntuneEnrollmentIds -OutputDirPath $paths.Output -RemovalLabel "Stale local Intune enrollment trace"
                    }

                    $cleanupStatus = if ($mdmCleanupResult -and $mdmCleanupResult.Success) { "STALE_INTUNE_ENROLLMENT_REMOVED" } else { "STALE_INTUNE_ENROLLMENT_REMOVE_FAILED" }
                    $result = New-ResultObject -Status $cleanupStatus -ExitCode 3 -Message "Hybrid Join is healthy, but stale local Intune enrollment traces were handled." -Dsreg $dsreg -IntuneEnrolled $false -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "RemoveStaleIntuneEnrollment" -MdmCleanupResult $mdmCleanupResult -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                    return $result
                }

                $result = New-ResultObject -Status "STALE_INTUNE_ENROLLMENT_LOCAL" -ExitCode 3 -Message "Hybrid Join is healthy, but a stale local Intune enrollment trace is suspected. Cleanup is explicit opt-in." -Dsreg $dsreg -IntuneEnrolled $false -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                return $result
            }

            if ($Mode -eq "Admin" -and $TriggerIntuneAutoEnrollment) {
                if ($Script:PolicyState.MdmPolicyStatus -ne "Enabled") {
                    $result = New-ResultObject -Status "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" -ExitCode 3 -Message "Hybrid Join is healthy, but MDM auto-enrollment policy is not configured." -Dsreg $dsreg -IntuneEnrolled $false -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "SkippedPolicyMissing" -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                    return $result
                }

                if ($PSCmdlet.ShouldProcess("Intune auto-enrollment on $Script:ComputerName", "Run deviceenroller.exe /c /AutoEnrollMDM")) {
                    $intuneAutoEnrollResult = Start-IntuneAutoEnrollment
                }

                $autoEnrollStatus = if ($intuneAutoEnrollResult -and -not $intuneAutoEnrollResult.ToolFound) { "INTUNE_ENROLLMENT_TOOL_NOT_FOUND" } elseif ($intuneAutoEnrollResult -and $intuneAutoEnrollResult.Success) { "INTUNE_AUTOENROLL_TRIGGERED" } else { "INTUNE_AUTOENROLL_FAILED" }
                $result = New-ResultObject -Status $autoEnrollStatus -ExitCode 3 -Message "Hybrid Join is healthy and Intune auto-enrollment trigger was requested." -Dsreg $dsreg -IntuneEnrolled $false -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "TriggerIntuneAutoEnrollment" -IntuneAutoEnrollResult $intuneAutoEnrollResult -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                return $result
            }

            $result = New-ResultObject -Status "HEALTHY" -ExitCode 0 -Message "Device registration is healthy. Intune enrollment is not detected locally." -Dsreg $dsreg -IntuneEnrolled $false -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        $actionMessages = New-Object System.Collections.Generic.List[string]
        $leaveAttempted = $false
        $leaveExitCode = ""
        $action = "DiagnosticOnly"

        $repairActionRequested = [bool]($RepairDisabledDeletedDevice -or $AllowDsregLeave -or $TriggerJoin -or $TriggerIntuneAutoEnrollment -or $AllowRemoveStaleIntuneEnrollment -or $AllowRemoveNonIntuneMdmEnrollment)

        if ($intuneDetected -and -not $AllowIntuneEnrolledAction -and $repairActionRequested) {
            $result = New-ResultObject -Status "ACTION_SKIPPED_INTUNE_ENROLLED" -ExitCode 3 -Message "Local Intune enrollment was detected. Repair action skipped unless -AllowIntuneEnrolledAction is used." -Dsreg $dsreg -IntuneEnrolled $true -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action $action -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        $disabledOrDeleted = Test-DisabledOrDeletedDeviceState -Dsreg $dsreg
        $missingDevice = Test-MissingDeviceHint -Dsreg $dsreg

        if ($Mode -eq "Admin" -and $repairActionRequested -and -not (Test-ProcessElevated)) {
            $result = New-ResultObject -Status "ADMIN_ELEVATION_REQUIRED" -ExitCode 3 -Message "Admin mode repair actions require an elevated PowerShell process." -Dsreg $dsreg -IntuneEnrolled $intuneDetected -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -Action "SkippedNotElevated" -LogPath $paths.RunLog -CsvPath $paths.CsvPath
            return $result
        }

        if ($RepairDisabledDeletedDevice -or $AllowDsregLeave) {
            $leaveApplicable = Test-DsregLeaveApplicable -Dsreg $dsreg -StrictDisabledDeletedOnly:($RepairDisabledDeletedDevice -and -not $AllowDsregLeave)
            if ($leaveApplicable.Applicable) {
                $target = "dsregcmd /leave on $Script:ComputerName"
                if ($PSCmdlet.ShouldProcess($target, "Repair Hybrid Join device registration")) {
                    $leaveAttempted = $true
                    $action = "DsregLeave"
                    Write-SmartM365DeviceRegistrationLog ("Running dsregcmd /leave because guard matched: {0}" -f $leaveApplicable.Reason)

                    $leaveOutput = & $dsregcmdPath /leave 2>&1
                    $leaveExitCode = [string]$LASTEXITCODE
                    $leaveOutputPath = Join-Path $paths.Output ("{0}_dsreg_leave_{1}.txt" -f $Script:ComputerName, $Script:RunId)
                    $leaveOutput | Out-File -LiteralPath $leaveOutputPath -Encoding UTF8 -Force
                    $actionMessages.Add(("dsregcmd /leave exit code: {0}" -f $leaveExitCode))

                    if ($LASTEXITCODE -ne 0) {
                        $result = New-ResultObject -Status "LEAVE_FAILED" -ExitCode 1 -Message ("dsregcmd /leave failed with exit code {0}." -f $leaveExitCode) -Dsreg $dsreg -IntuneEnrolled $intuneDetected -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -LeaveAttempted $leaveAttempted -LeaveExitCode $leaveExitCode -Action $action -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                        return $result
                    }

                    $joinTask = Start-AutomaticDeviceJoinTask
                    $actionMessages.Add($joinTask.Detail)
                    $TriggerJoin = $false
                }
            }
            else {
                $actionMessages.Add($leaveApplicable.Reason)
            }
        }

        if ($TriggerJoin) {
            if ($PSCmdlet.ShouldProcess("Automatic-Device-Join on $Script:ComputerName", "Start scheduled task")) {
                $action = if ($action -eq "DiagnosticOnly") { "TriggerJoin" } else { "$action+TriggerJoin" }
                $joinTask = Start-AutomaticDeviceJoinTask
                $actionMessages.Add($joinTask.Detail)
            }
        }

        if (($RepairDisabledDeletedDevice -or $AllowDsregLeave -or $TriggerJoin) -and $RetryCount -gt 0) {
            $retry = Invoke-PostActionRetry -DsregcmdPath $dsregcmdPath -OutputDir $paths.Output -Count $RetryCount -SleepMinutes $RetrySleepMinutes -Context "post_action"

            if ($retry.Dsreg) {
                $dsreg = $retry.Dsreg
            }

            if ($retry.Success) {
                $result = New-ResultObject -Status "HEALTHY_AFTER_ACTION" -ExitCode 0 -Message ("Device registration became healthy after {0} retry attempt(s). {1}" -f $retry.Attempts, ($actionMessages -join " ")) -Dsreg $dsreg -IntuneEnrolled $intuneDetected -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -LeaveAttempted $leaveAttempted -LeaveExitCode $leaveExitCode -Action $action -LogPath $paths.RunLog -CsvPath $paths.CsvPath
                return $result
            }
        }

        if ($disabledOrDeleted) {
            $status = if (($RepairDisabledDeletedDevice -or $AllowDsregLeave) -and $leaveAttempted) { "ACTION_COMPLETED_RECHECK_REQUIRED" } else { "DISABLED_OR_DELETED_DEVICE_DETECTED" }
            $message = if (($RepairDisabledDeletedDevice -or $AllowDsregLeave) -and $leaveAttempted) { "Repair action completed. Re-run diagnostics after sync/join completes." } else { "Device appears disabled or deleted in Entra ID. Use -RepairDisabledDeletedDevice or -AllowDsregLeave to allow guarded dsregcmd /leave." }
        }
        elseif ($dsreg.AzureAdJoined -eq "YES" -and $dsreg.KeySignTest -eq "FAILED") {
            $status = "KEY_SIGN_TEST_FAILED"
            $message = "KeySignTest is FAILED. Use -AllowDsregLeave in Admin mode when this rejoin repair is intended."
        }
        elseif ($missingDevice) {
            $status = "MISSING_DEVICE_HINT_DETECTED"
            $message = "dsregcmd indicates missing device or AAD Connect timing. Trigger join or retry after synchronization."
        }
        else {
            $status = "ATTENTION_REQUIRED"
            $message = "Device registration is not healthy and no guarded repair condition matched."
        }

        if ($actionMessages.Count -gt 0) {
            $message = "$message $($actionMessages -join ' ')"
        }

        $result = New-ResultObject -Status $status -ExitCode 3 -Message $message -Dsreg $dsreg -IntuneEnrolled $intuneDetected -DomainJoined $isDomainJoined -DomainName $domainName -DcReachable $dcReachable -DcCheck $dcCheck -LeaveAttempted $leaveAttempted -LeaveExitCode $leaveExitCode -Action $action -LogPath $paths.RunLog -CsvPath $paths.CsvPath
        return $result
    }
    catch {
        Write-SmartM365DeviceRegistrationLog ("Fatal error: {0}" -f $_.Exception.Message)
        $result = New-ResultObject -Status "ERROR" -ExitCode 1 -Message $_.Exception.Message -Dsreg $dsreg -LogPath $paths.RunLog -CsvPath $paths.CsvPath
        return $result
    }
    finally {
        if ($transcriptStarted) {
            try {
                Stop-Transcript | Out-Null
                Update-SmartM365DeviceRegistrationTranscript -Path $paths.Transcript
                Write-SmartM365DeviceRegistrationLog "Transcript stopped."
            }
            catch {
                Write-SmartM365DeviceRegistrationLog ("Stop-Transcript failed: {0}" -f $_.Exception.Message)
            }
        }

        if ($SupportBundle -and $result) {
            try {
                $eventLogPaths = @(Export-DeviceRegistrationEventLogs -Paths $paths)
                $bundlePath = New-DeviceRegistrationSupportBundle -Paths $paths -Result $result -EventLogPaths $eventLogPaths
                if (-not [string]::IsNullOrWhiteSpace($bundlePath)) {
                    $result.SupportBundlePath = $bundlePath
                    $emailResult = Send-DeviceRegistrationSupportBundleEmail -BundlePath $bundlePath -Result $result -To $Script:SupportEmail -SendMode $Script:SupportEmailSendMode
                    $result.SupportEmailStatus = $emailResult.Status
                    $result.SupportEmailMessage = $emailResult.Message
                    $result.SupportEmailDraftPath = $emailResult.EmlPath
                }
            }
            catch {
                Write-SmartM365DeviceRegistrationLog ("Support bundle creation failed: {0}" -f $_.Exception.Message)
            }
        }

        if ($result) {
            try {
                Write-AtomicCsvAppend -Path $paths.CsvPath -RowObject $result -RunIdValue $Script:RunId
                Write-SmartM365DeviceRegistrationLog ("CSV summary appended: {0}" -f $paths.CsvPath)
            }
            catch {
                Write-SmartM365DeviceRegistrationLog ("CSV append failed: {0}" -f $_.Exception.Message)
            }
        }

        try {
            Invoke-DeviceRegistrationFileRetention -Paths $paths -KeepLatest $LogRetentionCount
        }
        catch {
            Write-SmartM365DeviceRegistrationLog ("Retention unexpected failure: {0}" -f $_.Exception.Message)
        }
    }
}

function Format-DeviceRegistrationResultText {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    @(
        ("Status              : {0}" -f $Result.Status),
        ("NextAction          : {0}" -f $Result.NextAction),
        ("OverallHealth       : {0}" -f $Result.OverallHealth),
        ("Mode                : {0}" -f $Result.Mode),
        ("ExitCode            : {0}" -f $Result.ExitCode),
        ("Message             : {0}" -f $Result.Message),
        ("ComputerName        : {0}" -f $Result.ComputerName),
        ("CurrentUser         : {0}" -f $Result.CurrentUser),
        ("CurrentUserAuthority : {0}" -f $Result.CurrentUserAuthority),
        ("CurrentUserDnsDomain : {0}" -f $Result.CurrentUserDnsDomain),
        ("CurrentUserUpn      : {0}" -f $Result.CurrentUserUpn),
        ("UserIsDomainUser    : {0}" -f $Result.CurrentUserIsDomainUser),
        ("UserIsLocalAdmin    : {0}" -f $Result.CurrentUserIsLocalAdmin),
        ("ProcessElevated     : {0}" -f $Result.CurrentProcessElevated),
        ("ConfigPath          : {0}" -f $Result.ConfigPath),
        ("LogRetentionCount   : {0}" -f $Result.LogRetentionCount),
        ("DeviceProfile       : {0}" -f $Result.DeviceProfile),
        ("RequireDomainConnectivity : {0}" -f $Result.RequireDomainConnectivity),
        ("DomainJoined        : {0}" -f $Result.DomainJoined),
        ("DomainName          : {0}" -f $Result.DomainName),
        ("DcReachable         : {0}" -f $Result.DcReachable),
        ("IntuneEnrolled      : {0}" -f $Result.IntuneEnrolled),
        ("EnrollmentStatus    : {0}" -f $Result.EnrollmentStatusMessage),
        ("EnrollmentConfidence: {0}" -f $Result.EnrollmentConfidence),
        ("EnrollmentEvidence  : {0}" -f $Result.EnrollmentEvidence),
        ("EnrollmentStrongEvidence : {0}" -f $Result.EnrollmentStrongEvidence),
        ("EnrollmentWeakEvidence   : {0}" -f $Result.EnrollmentWeakEvidence),
        ("MdmAnyEnrollmentDetected : {0}" -f $Result.MdmAnyEnrollmentDetected),
        ("MdmIntuneEnrollmentDetected : {0}" -f $Result.MdmIntuneEnrollmentDetected),
        ("MdmNonIntuneEnrollmentDetected : {0}" -f $Result.MdmNonIntuneEnrollmentDetected),
        ("MdmEnrollmentCount  : {0}" -f $Result.MdmEnrollmentCount),
        ("MdmIntuneEnrollmentIds : {0}" -f $Result.MdmIntuneEnrollmentIds),
        ("MdmNonIntuneEnrollmentIds : {0}" -f $Result.MdmNonIntuneEnrollmentIds),
        ("MdmUnconfirmedIntuneEnrollmentIds : {0}" -f $Result.MdmUnconfirmedIntuneEnrollmentIds),
        ("MdmProviderIds      : {0}" -f $Result.MdmProviderIds),
        ("MdmEnrollmentDetails : {0}" -f $Result.MdmEnrollmentDetails),
        ("MdmIgnoredEnrollmentDetails : {0}" -f $Result.MdmIgnoredEnrollmentDetails),
        ("MdmPolicyCheckRequired : {0}" -f $Result.MdmPolicyCheckRequired),
        ("MdmPolicyStatus     : {0}" -f $Result.MdmPolicyStatus),
        ("MdmPolicyMessage    : {0}" -f $Result.MdmPolicyStatusMessage),
        ("MdmPolicyPresent    : {0}" -f $Result.MdmPolicyPresent),
        ("MdmAutoEnrollMDM    : {0}" -f $Result.MdmAutoEnrollMDM),
        ("MdmCredentialType   : {0}" -f $Result.MdmCredentialType),
        ("WorkplaceJoinPolicy : {0}" -f $Result.WorkplaceJoinPolicyPresent),
        ("AutoWorkplaceJoin   : {0}" -f $Result.WorkplaceJoinAutoWorkplaceJoin),
        ("AutoJoinTask        : Present={0}; State={1}" -f $Result.AutomaticDeviceJoinTaskPresent, $Result.AutomaticDeviceJoinTaskState),
        ("PolicyIssues        : {0}" -f $Result.PolicyIssues),
        ("AzureAdJoined       : {0}" -f $Result.AzureAdJoined),
        ("AzureAdPrt          : {0}" -f $Result.AzureAdPrt),
        ("DeviceAuthStatus    : {0}" -f $Result.DeviceAuthStatus),
        ("KeySignTest         : {0}" -f $Result.KeySignTest),
        ("MdmUrl              : {0}" -f $Result.MdmUrl),
        ("MdmTouUrl           : {0}" -f $Result.MdmTouUrl),
        ("MdmComplianceUrl    : {0}" -f $Result.MdmComplianceUrl),
        ("DeviceId            : {0}" -f $Result.DeviceId),
        ("TenantName          : {0}" -f $Result.TenantName),
        ("TenantId            : {0}" -f $Result.TenantId),
        ("ClientErrorCode     : {0}" -f $Result.ClientErrorCode),
        ("ServerErrorSubCode  : {0}" -f $Result.ServerErrorSubCode),
        ("ServerOperation     : {0}" -f $Result.ServerOperation),
        ("ErrorPhase          : {0}" -f $Result.ErrorPhase),
        ("DsregExitCode       : {0}" -f $Result.DsregExitCode),
        ("Action              : {0}" -f $Result.Action),
        ("LeaveAttempted      : {0}" -f $Result.LeaveAttempted),
        ("LeaveExitCode       : {0}" -f $Result.LeaveExitCode),
        ("MdmCleanupAttempted : {0}" -f $Result.MdmCleanupAttempted),
        ("MdmCleanupSuccess   : {0}" -f $Result.MdmCleanupSuccess),
        ("MdmCleanupEnrollmentIds : {0}" -f $Result.MdmCleanupEnrollmentIds),
        ("MdmCleanupBackupDir : {0}" -f $Result.MdmCleanupBackupDir),
        ("MdmCleanupDetail    : {0}" -f $Result.MdmCleanupDetail),
        ("IntuneAutoEnrollAttempted : {0}" -f $Result.IntuneAutoEnrollAttempted),
        ("IntuneAutoEnrollSuccess : {0}" -f $Result.IntuneAutoEnrollSuccess),
        ("IntuneAutoEnrollExitCode : {0}" -f $Result.IntuneAutoEnrollExitCode),
        ("IntuneAutoEnrollDetail : {0}" -f $Result.IntuneAutoEnrollDetail),
        ("DsregSnapshotPath   : {0}" -f $Result.DsregSnapshotPath),
        ("SupportBundlePath   : {0}" -f $Result.SupportBundlePath),
        ("SupportEmail        : {0}" -f $Result.SupportEmail),
        ("SupportEmailStatus  : {0}" -f $Result.SupportEmailStatus),
        ("SupportEmailMessage : {0}" -f $Result.SupportEmailMessage),
        ("SupportEmailDraftPath : {0}" -f $Result.SupportEmailDraftPath),
        ("SupportSummaryEmailStatus : {0}" -f $Result.SupportSummaryEmailStatus),
        ("SupportSummaryEmailMessage : {0}" -f $Result.SupportSummaryEmailMessage),
        ("SupportSummaryEmailDraftPath : {0}" -f $Result.SupportSummaryEmailDraftPath),
        ("JsonOutputPath      : {0}" -f $Result.JsonOutputPath),
        ("LogPath             : {0}" -f $Result.LogPath),
        ("CsvPath             : {0}" -f $Result.CsvPath)
    ) -join [Environment]::NewLine
}

function Get-DeviceRegistrationFindingLines {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    $findings = New-Object System.Collections.Generic.List[string]

    switch ($Result.Status) {
        "ERROR" {
            $findings.Add(("[ERROR] {0}" -f $Result.Message))
        }
        "LEAVE_FAILED" {
            $findings.Add(("[ERROR] {0}" -f $Result.Message))
        }
        "DOMAIN_NAME_EMPTY" {
            $findings.Add(("[ERROR] {0}" -f $Result.Message))
        }
        "DC_NOT_REACHABLE" {
            $findings.Add(("[ERROR] {0}" -f $Result.Message))
        }
        "DOMAIN_USER_REQUIRED" {
            $findings.Add(("[WARNING] {0} CurrentUser={1}" -f $Result.Message, $Result.CurrentUser))
        }
        "NOT_DOMAIN_JOINED" {
            if ($Result.RequireDomainConnectivity) {
                $findings.Add(("[WARNING] Device is not joined to an Active Directory domain. Domain={0}" -f $Result.DomainName))
            }
            else {
                $findings.Add(("[INFO] Device is not AD domain joined. Hybrid Join checks are not applicable. Domain={0}" -f $Result.DomainName))
            }
        }
        "ACTION_SKIPPED_INTUNE_ENROLLED" {
            $findings.Add(("[WARNING] Repair action skipped because Intune enrollment is detected."))
        }
        "ADMIN_ELEVATION_REQUIRED" {
            $findings.Add(("[WARNING] Admin repair action requires an elevated PowerShell process."))
        }
        "DISABLED_OR_DELETED_DEVICE_DETECTED" {
            $findings.Add(("[WARNING] Device appears disabled or deleted in Entra ID."))
        }
        "MISSING_DEVICE_HINT_DETECTED" {
            $findings.Add(("[WARNING] dsregcmd indicates a missing device or synchronization timing issue."))
        }
        "ATTENTION_REQUIRED" {
            $findings.Add(("[WARNING] {0}" -f $Result.Message))
        }
        "KEY_SIGN_TEST_FAILED" {
            $findings.Add(("[WARNING] KeySignTest is FAILED. {0}" -f $Result.Message))
        }
        "STALE_INTUNE_ENROLLMENT_LOCAL" {
            $findings.Add(("[WARNING] Stale local Intune enrollment trace suspected. {0}" -f $Result.Message))
        }
        "STALE_INTUNE_ENROLLMENT_REMOVED" {
            $findings.Add(("[INFO] Stale local Intune enrollment trace was removed. {0}" -f $Result.MdmCleanupDetail))
        }
        "STALE_INTUNE_ENROLLMENT_REMOVE_FAILED" {
            $findings.Add(("[ERROR] Stale local Intune enrollment cleanup failed. {0}" -f $Result.MdmCleanupDetail))
        }
        "NON_INTUNE_MDM_ENROLLED" {
            $findings.Add(("[WARNING] Non-Intune MDM enrollment detected. {0}" -f $Result.Message))
        }
        "NON_INTUNE_MDM_REMOVED" {
            $findings.Add(("[INFO] Non-Intune MDM enrollment was removed. {0}" -f $Result.MdmCleanupDetail))
        }
        "NON_INTUNE_MDM_REMOVE_FAILED" {
            $findings.Add(("[ERROR] Non-Intune MDM cleanup failed. {0}" -f $Result.MdmCleanupDetail))
        }
        "INTUNE_AUTOENROLL_TRIGGERED" {
            $findings.Add(("[INFO] Intune auto-enrollment was triggered. {0}" -f $Result.IntuneAutoEnrollDetail))
        }
        "INTUNE_AUTOENROLL_FAILED" {
            $findings.Add(("[WARNING] Intune auto-enrollment trigger failed. {0}" -f $Result.IntuneAutoEnrollDetail))
        }
        "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" {
            $findings.Add(("[WARNING] MDM auto-enrollment policy is not configured. {0}" -f $Result.Message))
        }
        "INTUNE_ENROLLMENT_TOOL_NOT_FOUND" {
            $findings.Add(("[ERROR] deviceenroller.exe was not found."))
        }
        default {
            if ($Result.ExitCode -ne 0 -and -not [string]::IsNullOrWhiteSpace($Result.Message)) {
                $findings.Add(("[WARNING] {0}" -f $Result.Message))
            }
        }
    }

    if ($Result.Status -ne "NOT_DOMAIN_JOINED" -and -not $Result.DomainJoined) {
        if ($Result.RequireDomainConnectivity) {
            $findings.Add(("[WARNING] Device is not joined to an Active Directory domain. Domain={0}" -f $Result.DomainName))
        }
        else {
            $findings.Add(("[INFO] Device is not AD domain joined. Hybrid Join checks are not applicable. Domain={0}" -f $Result.DomainName))
        }
    }

    if ($Result.RequireDomainConnectivity -and -not $Result.CurrentUserIsDomainUser) {
        $findings.Add(("[WARNING] Current Windows session is not an AD domain user. CurrentUser={0}" -f $Result.CurrentUser))
    }

    if ($Result.IntuneEnrolled) {
        $findings.Add("[INFO] Device is enrolled in Intune.")
    }
    elseif ($Result.EnrollmentConfidence -eq "Weak") {
        $findings.Add("[WARNING] Weak Intune enrollment signals detected without a strong enrollment record. Enrollment may be stale or partially removed.")
    }
    else {
        $findings.Add("[WARNING] Device is not enrolled in Intune.")
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.MdmPolicyStatusMessage)) {
        if ($Result.MdmPolicyStatus -eq "MissingOrDisabled" -and $Result.MdmPolicyCheckRequired) {
            $findings.Add(("[WARNING] {0}" -f $Result.MdmPolicyStatusMessage))
        }
        else {
            $findings.Add(("[INFO] {0}" -f $Result.MdmPolicyStatusMessage))
        }
    }

    if (-not $Result.IntuneEnrolled -and -not [string]::IsNullOrWhiteSpace($Result.PolicyIssues)) {
        foreach ($issue in ($Result.PolicyIssues -split "\s*\|\s*")) {
            if (-not [string]::IsNullOrWhiteSpace($issue)) {
                $findings.Add(("[WARNING] {0}" -f $issue))
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.AzureAdJoined) -and $Result.AzureAdJoined -ne "YES") {
        $findings.Add(("[WARNING] AzureAdJoined is {0}." -f $Result.AzureAdJoined))
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.AzureAdPrt)) {
        $findings.Add(("[INFO] AzureAdPrt is {0}." -f $Result.AzureAdPrt))
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.DeviceAuthStatus) -and $Result.DeviceAuthStatus -notlike "*SUCCESS*") {
        $findings.Add(("[WARNING] DeviceAuthStatus is {0}." -f $Result.DeviceAuthStatus))
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.KeySignTest)) {
        if ($Result.KeySignTest -eq "FAILED") {
            $findings.Add("[WARNING] KeySignTest is FAILED.")
        }
        else {
            $findings.Add(("[INFO] KeySignTest is {0}." -f $Result.KeySignTest))
        }
    }

    if ($Result.MdmNonIntuneEnrollmentDetected) {
        $findings.Add(("[WARNING] Non-Intune MDM enrollment detected. EnrollmentIds={0}" -f $Result.MdmNonIntuneEnrollmentIds))
    }

    if (-not $Result.IntuneEnrolled -and -not [string]::IsNullOrWhiteSpace($Result.MdmUnconfirmedIntuneEnrollmentIds)) {
        $findings.Add(("[WARNING] Unconfirmed Intune enrollment trace detected. EnrollmentIds={0}" -f $Result.MdmUnconfirmedIntuneEnrollmentIds))
    }

    if ($Result.IntuneAutoEnrollAttempted) {
        $prefix = if ($Result.IntuneAutoEnrollSuccess) { "[INFO]" } else { "[WARNING]" }
        $findings.Add(("{0} Intune auto-enrollment trigger: Success={1}; ExitCode={2}; Detail={3}" -f $prefix, $Result.IntuneAutoEnrollSuccess, $Result.IntuneAutoEnrollExitCode, $Result.IntuneAutoEnrollDetail))
    }

    if ($Result.MdmCleanupAttempted) {
        $prefix = if ($Result.MdmCleanupSuccess) { "[INFO]" } else { "[ERROR]" }
        $findings.Add(("{0} MDM cleanup: Success={1}; EnrollmentIds={2}; Backup={3}; Detail={4}" -f $prefix, $Result.MdmCleanupSuccess, $Result.MdmCleanupEnrollmentIds, $Result.MdmCleanupBackupDir, $Result.MdmCleanupDetail))
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.ClientErrorCode)) {
        $findings.Add(("[ERROR] ClientErrorCode: {0}" -f $Result.ClientErrorCode))
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.ServerErrorSubCode)) {
        $findings.Add(("[ERROR] ServerErrorSubCode: {0}" -f $Result.ServerErrorSubCode))
    }

    if ($Result.Mode -eq "Admin" -and -not $Result.CurrentProcessElevated) {
        $findings.Add("[WARNING] Admin mode is selected but the current process is not elevated.")
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.SupportBundlePath)) {
        $findings.Add(("[INFO] Support bundle: {0}" -f $Result.SupportBundlePath))
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.SupportEmailStatus)) {
        $prefix = if ($Result.SupportEmailStatus -eq "Failed") { "[WARNING]" } else { "[INFO]" }
        $findings.Add(("{0} Support email: {1}" -f $prefix, $Result.SupportEmailMessage))
    }

    if (-not [string]::IsNullOrWhiteSpace($Result.SupportSummaryEmailStatus)) {
        $prefix = if ($Result.SupportSummaryEmailStatus -eq "Failed") { "[WARNING]" } else { "[INFO]" }
        $findings.Add(("{0} Support summary email: {1}" -f $prefix, $Result.SupportSummaryEmailMessage))
    }

    if ($findings.Count -eq 0) {
        $findings.Add("[OK] No issue or warning detected.")
    }

    return @($findings | Select-Object -Unique)
}

function Format-DeviceRegistrationFindingsText {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    $lines = @(Get-DeviceRegistrationFindingLines -Result $Result | Where-Object {
        $_.StartsWith("[ERROR]") -or $_.StartsWith("[WARNING]") -or $_.StartsWith("[OK]")
    })

    if ($lines.Count -eq 0) {
        $lines = @("[OK] No issue or warning detected.")
    }

    return $lines -join [Environment]::NewLine
}

function Format-DeviceRegistrationInfoText {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    $lines = @(Get-DeviceRegistrationFindingLines -Result $Result | Where-Object { $_.StartsWith("[INFO]") })

    if ($lines.Count -eq 0) {
        $lines = @("[INFO] No additional information.")
    }

    return $lines -join [Environment]::NewLine
}

function Import-SmartM365GuiSplash {
    $current = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    }
    else {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }

    while ($current) {
        $splashPath = Join-Path -Path $current -ChildPath 'SmartM365.GuiSplash.ps1'
        if (Test-Path -LiteralPath $splashPath) {
            return $splashPath
        }

        if ((Split-Path -Path $current -Leaf) -eq 'SmartM365') {
            return $null
        }

        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }

    return $null
}

function Show-SmartM365DeviceRegistrationGui {
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
    $splash = $null
    $splashHelperPath = Import-SmartM365GuiSplash
    if ($splashHelperPath) {
        . $splashHelperPath
        $splash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'Device Registration Tool'
    }
    $script:Strings = Get-DeviceRegistrationStrings

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart DeviceRegistration Tool"
        Height="900"
        Width="1240"
        MinHeight="760"
        MinWidth="1040"
        WindowStartupLocation="CenterScreen"
        Background="#F5F8FB"
        FontFamily="Segoe UI">
    <Window.Resources>
        <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="AccentSoftBrush" Color="#E6F4FF"/>
        <SolidColorBrush x:Key="InkBrush" Color="#1F2937"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#5F6B7A"/>
        <SolidColorBrush x:Key="BorderBrushSoft" Color="#DDE7F0"/>
        <SolidColorBrush x:Key="WarningBrush" Color="#B54708"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#027A48"/>
        <SolidColorBrush x:Key="DangerBrush" Color="#B42318"/>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Height" Value="34"/>
            <Setter Property="MinWidth" Value="118"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="BorderBrush" Value="#B9C8D7"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#F1F7FC"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                                <Setter Property="Cursor" Value="Arrow"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Height" Value="34"/>
            <Setter Property="MinWidth" Value="118"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                                <Setter Property="BorderBrush" Value="#106EBE"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                                <Setter Property="Cursor" Value="Arrow"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Margin" Value="0,0,6,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="#FFFFFF" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#F8FBFE"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource AccentSoftBrush}"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                                <Setter Property="Foreground" Value="#005A9E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="BorderBrush" Value="#B9C8D7"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
        </Style>
        <Style x:Key="FieldLabel" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,4"/>
        </Style>
        <Style x:Key="PanelTitle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
        </Style>
    </Window.Resources>

    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="18">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <Border HorizontalAlignment="Left" CornerRadius="12" Padding="10,4" Background="{StaticResource AccentSoftBrush}">
                        <TextBlock x:Name="EyebrowText" Text="LOCAL DEVICE REGISTRATION" Foreground="#005A9E" FontWeight="SemiBold" FontSize="12"/>
                    </Border>
                    <TextBlock x:Name="TitleText" Text="Smart DeviceRegistration Tool" Margin="0,12,0,0" FontSize="28" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}" TextWrapping="Wrap"/>
                    <TextBlock x:Name="SubtitleText" Text="Hybrid Join, Entra device registration, policy checks, and support-ready diagnostics." Margin="0,8,0,0" FontSize="14" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                </StackPanel>
                <Border x:Name="HeaderLogoLink" Grid.Column="1" Width="124" Height="88" CornerRadius="8" Background="#FFFFFF" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" Padding="4" Margin="18,0,0,0" Cursor="Hand" ToolTip="Open WorkplaceCloudHub.com">
                    <Grid>
                        <Image x:Name="HeaderLogoImage" Stretch="Uniform" MaxWidth="116" MaxHeight="80" RenderOptions.BitmapScalingMode="HighQuality" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
                        <StackPanel x:Name="HeaderLogoFallback" VerticalAlignment="Center">
                            <TextBlock Text="DRT" HorizontalAlignment="Center" FontSize="34" FontWeight="Bold" Foreground="{StaticResource AccentBrush}"/>
                            <TextBlock Text="SmartM365" HorizontalAlignment="Center" Foreground="{StaticResource MutedBrush}" FontSize="12"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <Grid Grid.Row="1" Margin="0,12,0,12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,0,12">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Border Width="10" Height="10" CornerRadius="5" Background="{StaticResource AccentBrush}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <StackPanel Grid.Column="1" VerticalAlignment="Center">
                        <TextBlock x:Name="DeviceNameText" Text="PC: COMPUTERNAME" FontSize="17" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}" TextWrapping="Wrap"/>
                        <TextBlock x:Name="ActionTitleText" Text="Device registration diagnostic" Foreground="{StaticResource AccentBrush}" FontWeight="SemiBold" FontSize="12" TextWrapping="Wrap" Margin="0,3,0,0"/>
                        <TextBlock x:Name="StatusText" Text="User mode. Diagnostic only." Foreground="{StaticResource MutedBrush}" FontSize="12" TextWrapping="Wrap" Margin="0,2,0,0"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="12,0,0,0">
                        <Button x:Name="RunButton" Content="Run diagnostic" Style="{StaticResource PrimaryButton}" MinWidth="150" Margin="0,0,8,0"/>
                        <Button x:Name="RefreshPrtButton" Content="Refresh Azure AD PRT" MinWidth="170" Margin="0,0,8,0"/>
                        <Button x:Name="RepairButton" Content="Run repair" MinWidth="150" Margin="0"/>
                    </StackPanel>
                </Grid>
            </Border>

            <TabControl x:Name="MainTabs" Grid.Row="1">
                <TabItem x:Name="SummaryTab" Header="Summary">
                    <Grid Margin="0,12,0,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="160"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Border x:Name="EnrollmentBanner" Grid.Row="0" Background="#F8FAFC" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                            <TextBlock x:Name="EnrollmentBannerText" Text="Running diagnostic..." FontSize="15" FontWeight="SemiBold" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                        </Border>
                        <Border x:Name="MdmPolicyBanner" Grid.Row="1" Background="#F8FAFC" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                            <TextBlock x:Name="MdmPolicyBannerText" Text="MDM auto-enrollment policy result will appear after diagnostic." FontSize="15" FontWeight="SemiBold" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                        </Border>
                        <Border Grid.Row="2" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,0,12">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock x:Name="IssuesTitleText" Text="Issues / warnings" Style="{StaticResource PanelTitle}"/>
                                <TextBox x:Name="IssuesBox" Grid.Row="1" FontFamily="Consolas" FontSize="13" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" TextWrapping="Wrap" BorderBrush="#E4EAF1" Background="#FAFCFE"/>
                            </Grid>
                        </Border>
                        <Border Grid.Row="3" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,0,12">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock x:Name="InformationTitleText" Text="Information" Style="{StaticResource PanelTitle}"/>
                                <TextBox x:Name="InfoBox" Grid.Row="1" FontFamily="Consolas" FontSize="13" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" TextWrapping="Wrap" BorderBrush="#E4EAF1" Background="#FAFCFE"/>
                            </Grid>
                        </Border>
                        <Border Grid.Row="4" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button x:Name="CopySummaryButton" Content="Copy/email support summary" MinWidth="190" Margin="0,0,8,0" IsEnabled="False"/>
                                <Button x:Name="CreateBundleButton" Content="Create/email support bundle" MinWidth="190" Margin="0,0,8,0" IsEnabled="False"/>
                                <Button x:Name="OpenLogButton" Content="Open latest log folder" MinWidth="160" Margin="0" IsEnabled="False"/>
                            </StackPanel>
                        </Border>
                    </Grid>
                </TabItem>
                <TabItem x:Name="DiagnosticOutputTab" Header="Diagnostic output">
                    <Border Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,12,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock x:Name="DiagnosticOutputTitleText" Text="Diagnostic output" Style="{StaticResource PanelTitle}"/>
                            <TextBox x:Name="OutputBox" Grid.Row="1" FontFamily="Consolas" FontSize="13" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" TextWrapping="NoWrap" BorderBrush="#E4EAF1" Background="#FAFCFE"/>
                        </Grid>
                    </Border>
                </TabItem>
                <TabItem x:Name="DsregTab" Header="Device status">
                    <Border Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,12,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock x:Name="DsregTitleText" Text="Device registration status" Style="{StaticResource PanelTitle}"/>
                            <TextBox x:Name="DsregOutputBox" Grid.Row="1" FontFamily="Consolas" FontSize="13" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" TextWrapping="NoWrap" BorderBrush="#E4EAF1" Background="#FAFCFE"/>
                        </Grid>
                    </Border>
                </TabItem>
                <TabItem x:Name="OptionsTab" Header="Options">
                    <Border Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,12,0,0">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel MaxWidth="560" HorizontalAlignment="Left">
                                <TextBlock x:Name="ModeLabelText" Text="Mode" Style="{StaticResource FieldLabel}"/>
                                <ComboBox x:Name="ModeCombo" SelectedIndex="0" Margin="0,0,0,12">
                                    <ComboBoxItem Tag="User" Content="User"/>
                                    <ComboBoxItem Tag="Admin" Content="Admin"/>
                                </ComboBox>

                                <Border Background="#FAFCFE" BorderBrush="#E4EAF1" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,0,0,12">
                                    <StackPanel>
                                        <CheckBox x:Name="RepairCheck" Content="Allow guarded dsregcmd /leave"/>
                                        <CheckBox x:Name="TriggerJoinCheck" Content="Trigger Automatic-Device-Join"/>
                                        <CheckBox x:Name="TriggerIntuneAutoEnrollCheck" Content="Trigger Intune auto-enrollment"/>
                                        <CheckBox x:Name="RemoveStaleIntuneCheck" Content="Remove stale local Intune enrollment trace"/>
                                        <CheckBox x:Name="RemoveNonIntuneMdmCheck" Content="Remove non-Intune MDM enrollment"/>
                                        <CheckBox x:Name="AllowIntuneCheck" Content="Allow actions if Intune enrolled"/>
                                        <CheckBox x:Name="AuditOnlyCheck" Content="Audit only" Margin="0"/>
                                    </StackPanel>
                                </Border>

                                <TextBlock x:Name="RuntimeSettingsText" Text="Runtime settings are loaded from JSON configuration." Foreground="{StaticResource MutedBrush}" FontSize="12" TextWrapping="Wrap" Margin="0,0,0,0"/>
                                <Border Background="#FAFCFE" BorderBrush="#E4EAF1" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,12,0,0">
                                    <StackPanel>
                                        <TextBlock x:Name="ConfigurationTitleText" Text="Configuration" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                        <TextBlock x:Name="ConfigSummaryText" Foreground="{StaticResource MutedBrush}" FontSize="12" TextWrapping="Wrap"/>
                                    </StackPanel>
                                </Border>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>
                </TabItem>
            </TabControl>
        </Grid>

        <Border Grid.Row="2" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12">
            <TextBlock x:Name="FooterText" Foreground="{StaticResource MutedBrush}"/>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $modeCombo = $window.FindName("ModeCombo")
    $repairCheck = $window.FindName("RepairCheck")
    $triggerJoinCheck = $window.FindName("TriggerJoinCheck")
    $triggerIntuneAutoEnrollCheck = $window.FindName("TriggerIntuneAutoEnrollCheck")
    $removeStaleIntuneCheck = $window.FindName("RemoveStaleIntuneCheck")
    $removeNonIntuneMdmCheck = $window.FindName("RemoveNonIntuneMdmCheck")
    $allowIntuneCheck = $window.FindName("AllowIntuneCheck")
    $auditOnlyCheck = $window.FindName("AuditOnlyCheck")
    $runButton = $window.FindName("RunButton")
    $refreshPrtButton = $window.FindName("RefreshPrtButton")
    $repairButton = $window.FindName("RepairButton")
    $openLogButton = $window.FindName("OpenLogButton")
    $copySummaryButton = $window.FindName("CopySummaryButton")
    $createBundleButton = $window.FindName("CreateBundleButton")
    $mainTabs = $window.FindName("MainTabs")
    $optionsTab = $window.FindName("OptionsTab")
    $summaryTab = $window.FindName("SummaryTab")
    $diagnosticOutputTab = $window.FindName("DiagnosticOutputTab")
    $dsregTab = $window.FindName("DsregTab")
    $configSummaryText = $window.FindName("ConfigSummaryText")
    $eyebrowText = $window.FindName("EyebrowText")
    $titleText = $window.FindName("TitleText")
    $subtitleText = $window.FindName("SubtitleText")
    $deviceNameText = $window.FindName("DeviceNameText")
    $actionTitleText = $window.FindName("ActionTitleText")
    $issuesTitleText = $window.FindName("IssuesTitleText")
    $informationTitleText = $window.FindName("InformationTitleText")
    $diagnosticOutputTitleText = $window.FindName("DiagnosticOutputTitleText")
    $dsregTitleText = $window.FindName("DsregTitleText")
    $modeLabelText = $window.FindName("ModeLabelText")
    $runtimeSettingsText = $window.FindName("RuntimeSettingsText")
    $configurationTitleText = $window.FindName("ConfigurationTitleText")
    $issuesBox = $window.FindName("IssuesBox")
    $infoBox = $window.FindName("InfoBox")
    $outputBox = $window.FindName("OutputBox")
    $dsregOutputBox = $window.FindName("DsregOutputBox")
    $statusText = $window.FindName("StatusText")
    $footerText = $window.FindName("FooterText")
    $enrollmentBanner = $window.FindName("EnrollmentBanner")
    $enrollmentBannerText = $window.FindName("EnrollmentBannerText")
    $mdmPolicyBanner = $window.FindName("MdmPolicyBanner")
    $mdmPolicyBannerText = $window.FindName("MdmPolicyBannerText")
    $headerLogoImage = $window.FindName("HeaderLogoImage")
    $headerLogoFallback = $window.FindName("HeaderLogoFallback")
    $headerLogoLink = $window.FindName("HeaderLogoLink")
    $brushConverter = New-Object System.Windows.Media.BrushConverter

    function Open-ExternalUrl {
        param([Parameter(Mandatory)][string]$Url)

        try {
            $psi = [System.Diagnostics.ProcessStartInfo]::new($Url)
            $psi.UseShellExecute = $true
            [System.Diagnostics.Process]::Start($psi) | Out-Null
        }
        catch {
            [System.Windows.MessageBox]::Show($window, "Unable to open:`r`n$Url`r`n`r`n$($_.Exception.Message)", 'SmartM365', 'OK', 'Warning') | Out-Null
        }
    }

    if ($headerLogoLink) {
        $headerLogoLink.Add_MouseLeftButtonUp({ Open-ExternalUrl -Url 'https://workplacecloudhub.com' })
    }

    $effectiveLogoPath = Get-EffectiveLogoPath
    if (-not [string]::IsNullOrWhiteSpace($effectiveLogoPath)) {
        try {
            $logoImage = Get-DeviceRegistrationLogoImage -Path $effectiveLogoPath
            $window.Icon = $logoImage
            $headerLogoImage.Source = $logoImage
            $headerLogoImage.Visibility = [System.Windows.Visibility]::Visible
            $headerLogoFallback.Visibility = [System.Windows.Visibility]::Collapsed
        }
        catch {
            Write-SmartM365DeviceRegistrationLog ("Logo load failed: {0}" -f $_.Exception.Message)
        }
    }

    $modeCombo.SelectedIndex = if ($Mode -eq "Admin") { 1 } else { 0 }
    $repairCheck.IsChecked = [bool]($RepairDisabledDeletedDevice -or $AllowDsregLeave)
    $triggerJoinCheck.IsChecked = [bool]$TriggerJoin
    $triggerIntuneAutoEnrollCheck.IsChecked = [bool]$TriggerIntuneAutoEnrollment
    $removeStaleIntuneCheck.IsChecked = [bool]$AllowRemoveStaleIntuneEnrollment
    $removeNonIntuneMdmCheck.IsChecked = [bool]$AllowRemoveNonIntuneMdmEnrollment
    $allowIntuneCheck.IsChecked = [bool]$AllowIntuneEnrolledAction
    $auditOnlyCheck.IsChecked = [bool]$AuditOnly
    $footerText.Text = "RunId: $Script:RunId"
    $deviceNameText.Text = "PC: $Script:ComputerName"
    $configPathText = if ($script:ToolConfig -and $script:ToolConfig.PSObject.Properties.Name -contains "ConfigPath") { $script:ToolConfig.ConfigPath } else { $script:Strings.DefaultParameters }
    $supportEmailText = if ([string]::IsNullOrWhiteSpace($Script:SupportEmail)) { $script:Strings.NotConfigured } else { "$($Script:SupportEmail) ($($Script:SupportEmailSendMode))" }
    $logoText = if (-not [string]::IsNullOrWhiteSpace($Script:LogoPath)) { $Script:LogoPath } elseif (-not [string]::IsNullOrWhiteSpace($effectiveLogoPath)) { Split-Path -Path $effectiveLogoPath -Leaf } else { $script:Strings.DefaultText }
    $languageText = if (-not [string]::IsNullOrWhiteSpace($Script:ForceLanguage)) { "Force=$($Script:ForceLanguage)" } else { "Default=$($Script:DefaultLanguage)" }
    $configSummaryText.Text = "Profile=$DeviceProfile`nRequireDomainConnectivity=$Script:RequireDomainConnectivity`nLogRetentionCount=$LogRetentionCount`nSupportEmail=$supportEmailText`nLogo=$logoText`nLanguage=$languageText`nConfig=$configPathText"

    $script:lastLogFolder = $null
    $script:lastResult = $null

    $getDsregSnapshotText = {
        param([psobject]$Result)

        if ($null -eq $Result -or [string]::IsNullOrWhiteSpace($Result.DsregSnapshotPath)) {
            return $script:Strings.DsregOutputUnavailable
        }

        if (-not (Test-Path -LiteralPath $Result.DsregSnapshotPath)) {
            return ([string]::Format($script:Strings.DsregSnapshotNotFoundFormat, $Result.DsregSnapshotPath))
        }

        try {
            return [IO.File]::ReadAllText($Result.DsregSnapshotPath)
        }
        catch {
            return ([string]::Format($script:Strings.DsregSnapshotReadFailedFormat, $_.Exception.Message))
        }
    }

    $getSelectedMode = {
        if ($modeCombo.SelectedItem) {
            $tag = [string]$modeCombo.SelectedItem.Tag
            if ($tag -eq "Admin") { return "Admin" }
        }

        return "User"
    }

    $applyLocalizedUi = {
        $window.Title = $script:Strings.WindowTitle
        if ($script:Strings.FlowDirection -eq "RightToLeft") {
            $window.FlowDirection = [System.Windows.FlowDirection]::RightToLeft
        }
        else {
            $window.FlowDirection = [System.Windows.FlowDirection]::LeftToRight
        }

        $eyebrowText.Text = $script:Strings.Eyebrow
        $titleText.Text = $script:Strings.Title
        $subtitleText.Text = $script:Strings.Subtitle
        $actionTitleText.Text = $script:Strings.ActionTitle
        $runButton.Content = $script:Strings.RunDiagnostic
        $refreshPrtButton.Content = $script:Strings.RefreshPrt
        $repairButton.Content = $script:Strings.RunRepair
        $summaryTab.Header = $script:Strings.SummaryTab
        $diagnosticOutputTab.Header = $script:Strings.DiagnosticOutputTab
        $dsregTab.Header = $script:Strings.DsregTab
        $optionsTab.Header = $script:Strings.OptionsTab
        $issuesTitleText.Text = $script:Strings.IssuesWarnings
        $informationTitleText.Text = $script:Strings.Information
        $diagnosticOutputTitleText.Text = $script:Strings.DiagnosticOutput
        $dsregTitleText.Text = $script:Strings.DsregStatus
        $copySummaryButton.Content = $script:Strings.CopyEmailSupportSummary
        $createBundleButton.Content = $script:Strings.CreateEmailSupportBundle
        $openLogButton.Content = $script:Strings.OpenLatestLogFolder
        $modeLabelText.Text = $script:Strings.Mode
        $repairCheck.Content = $script:Strings.AllowDsregLeave
        $triggerJoinCheck.Content = $script:Strings.TriggerAutomaticDeviceJoin
        $triggerIntuneAutoEnrollCheck.Content = $script:Strings.TriggerIntuneAutoEnrollment
        $removeStaleIntuneCheck.Content = $script:Strings.RemoveStaleIntuneEnrollment
        $removeNonIntuneMdmCheck.Content = $script:Strings.RemoveNonIntuneMdmEnrollment
        $allowIntuneCheck.Content = $script:Strings.AllowIntuneActions
        $auditOnlyCheck.Content = $script:Strings.AuditOnly
        $runtimeSettingsText.Text = $script:Strings.RuntimeSettingsFromJson
        $configurationTitleText.Text = $script:Strings.Configuration

        foreach ($item in @($modeCombo.Items)) {
            if ([string]$item.Tag -eq "Admin") {
                $item.Content = $script:Strings.AdminMode
            }
            elseif ([string]$item.Tag -eq "User") {
                $item.Content = $script:Strings.UserMode
            }
        }

        $configPathText = if ($script:ToolConfig -and $script:ToolConfig.PSObject.Properties.Name -contains "ConfigPath") { $script:ToolConfig.ConfigPath } else { $script:Strings.DefaultParameters }
        $supportEmailText = if ([string]::IsNullOrWhiteSpace($Script:SupportEmail)) { $script:Strings.NotConfigured } else { "$($Script:SupportEmail) ($($Script:SupportEmailSendMode))" }
        $logoText = if (-not [string]::IsNullOrWhiteSpace($Script:LogoPath)) { $Script:LogoPath } elseif (-not [string]::IsNullOrWhiteSpace($effectiveLogoPath)) { Split-Path -Path $effectiveLogoPath -Leaf } else { $script:Strings.DefaultText }
        $languageText = if (-not [string]::IsNullOrWhiteSpace($Script:ForceLanguage)) { "Force=$($Script:ForceLanguage)" } else { "Default=$($Script:DefaultLanguage)" }
        $configSummaryText.Text = "Profile=$DeviceProfile`nRequireDomainConnectivity=$Script:RequireDomainConnectivity`nLogRetentionCount=$LogRetentionCount`nSupportEmail=$supportEmailText`nLogo=$logoText`nLanguage=$languageText`nConfig=$configPathText"
    }

    $updateRepairButtonState = {
        $isAdminMode = $false
        if ((& $getSelectedMode) -eq "Admin") {
            $isAdminMode = $true
        }

        $hasRepairAction = ([bool]$repairCheck.IsChecked -or [bool]$triggerJoinCheck.IsChecked -or [bool]$triggerIntuneAutoEnrollCheck.IsChecked -or [bool]$removeStaleIntuneCheck.IsChecked -or [bool]$removeNonIntuneMdmCheck.IsChecked -or [bool]$auditOnlyCheck.IsChecked)
        $repairButton.Visibility = if ($isAdminMode) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
        $repairButton.IsEnabled = ($isAdminMode -and $hasRepairAction)
        $refreshPrtButton.Visibility = [System.Windows.Visibility]::Visible
        $hasPrt = $false
        if ($script:lastResult -and $script:lastResult.PSObject.Properties.Name -contains "AzureAdPrt") {
            $hasPrt = ([string]$script:lastResult.AzureAdPrt).Trim().ToUpperInvariant() -eq "YES"
        }

        $refreshPrtButton.IsEnabled = (-not $hasPrt)
    }

    $setModeState = {
        $isAdminMode = $false
        if ((& $getSelectedMode) -eq "Admin") {
            $isAdminMode = $true
        }

        foreach ($control in @($repairCheck, $triggerJoinCheck, $triggerIntuneAutoEnrollCheck, $removeStaleIntuneCheck, $removeNonIntuneMdmCheck, $allowIntuneCheck, $auditOnlyCheck)) {
            $control.IsEnabled = $isAdminMode
        }

        if (-not $isAdminMode) {
            $repairCheck.IsChecked = $false
            $triggerJoinCheck.IsChecked = $false
            $triggerIntuneAutoEnrollCheck.IsChecked = $false
            $removeStaleIntuneCheck.IsChecked = $false
            $removeNonIntuneMdmCheck.IsChecked = $false
            $allowIntuneCheck.IsChecked = $false
            $auditOnlyCheck.IsChecked = $false
        }

        if ($optionsTab) {
            $optionsTab.Visibility = if ($isAdminMode) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
            if (-not $isAdminMode -and $mainTabs -and $mainTabs.SelectedItem -eq $optionsTab) {
                $mainTabs.SelectedIndex = 0
            }
        }

        if ($openLogButton) {
            $openLogButton.Visibility = if ($isAdminMode) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
            $openLogButton.IsEnabled = ($isAdminMode -and [bool]$script:lastLogFolder)
        }

        $statusText.Text = if ($isAdminMode) { $script:Strings.AdminModeAvailable } else { $script:Strings.UserModeDiagnosticOnly }
        & $updateRepairButtonState
    }

    & $applyLocalizedUi
    $modeCombo.Add_SelectionChanged({ & $setModeState })
    $repairCheck.Add_Click({ & $updateRepairButtonState })
    $triggerJoinCheck.Add_Click({ & $updateRepairButtonState })
    $triggerIntuneAutoEnrollCheck.Add_Click({ & $updateRepairButtonState })
    $removeStaleIntuneCheck.Add_Click({ & $updateRepairButtonState })
    $removeNonIntuneMdmCheck.Add_Click({ & $updateRepairButtonState })
    $allowIntuneCheck.Add_Click({ & $updateRepairButtonState })
    $auditOnlyCheck.Add_Click({ & $updateRepairButtonState })
    & $setModeState

    $runDiagnostic = {
        param([bool]$DiagnosticOnly)

        $runButton.IsEnabled = $false
        $refreshPrtButton.IsEnabled = $false
        $repairButton.IsEnabled = $false
        $statusText.Text = $script:Strings.Running
        $issuesBox.Text = ""
        $infoBox.Text = ""
        $outputBox.Text = ""
        $dsregOutputBox.Text = ""
        $script:lastResult = $null
        $copySummaryButton.IsEnabled = $false
        $createBundleButton.IsEnabled = $false
        $enrollmentBanner.Background = $brushConverter.ConvertFromString("#F8FAFC")
        $enrollmentBanner.BorderBrush = $brushConverter.ConvertFromString("#DDE7F0")
        $enrollmentBannerText.Foreground = $brushConverter.ConvertFromString("#5F6B7A")
        $enrollmentBannerText.Text = $script:Strings.RunningDiagnostic
        $mdmPolicyBanner.Background = $brushConverter.ConvertFromString("#F8FAFC")
        $mdmPolicyBanner.BorderBrush = $brushConverter.ConvertFromString("#DDE7F0")
        $mdmPolicyBannerText.Foreground = $brushConverter.ConvertFromString("#5F6B7A")
        $mdmPolicyBannerText.Text = $script:Strings.CheckingMdmPolicy

        try {
            $selectedMode = & $getSelectedMode
            $isAdminMode = $selectedMode -eq "Admin"
            $effectiveRetryCount = if ($DiagnosticOnly) { 0 } else { $RetryCount }

            $result = Invoke-SmartM365DeviceRegistrationTool `
                -Mode $selectedMode `
                -RepairDisabledDeletedDevice:$false `
                -AllowDsregLeave:((-not $DiagnosticOnly) -and $isAdminMode -and [bool]$repairCheck.IsChecked) `
                -TriggerJoin:((-not $DiagnosticOnly) -and $isAdminMode -and [bool]$triggerJoinCheck.IsChecked) `
                -TriggerIntuneAutoEnrollment:((-not $DiagnosticOnly) -and $isAdminMode -and [bool]$triggerIntuneAutoEnrollCheck.IsChecked) `
                -AllowRemoveStaleIntuneEnrollment:((-not $DiagnosticOnly) -and $isAdminMode -and [bool]$removeStaleIntuneCheck.IsChecked) `
                -AllowRemoveNonIntuneMdmEnrollment:((-not $DiagnosticOnly) -and $isAdminMode -and [bool]$removeNonIntuneMdmCheck.IsChecked) `
                -AllowIntuneEnrolledAction:((-not $DiagnosticOnly) -and $isAdminMode -and [bool]$allowIntuneCheck.IsChecked) `
                -AuditOnly:((-not $DiagnosticOnly) -and $isAdminMode -and [bool]$auditOnlyCheck.IsChecked) `
                -RetryCount $effectiveRetryCount `
                -RetrySleepMinutes $RetrySleepMinutes `
                -LogRetentionCount $LogRetentionCount `
                -OutputRoot $OutputRoot `
                -SupportBundle:$false `
                -NoTranscript:$NoTranscript

            $outputBox.Text = Format-DeviceRegistrationResultText -Result $result
            $issuesBox.Text = Format-DeviceRegistrationFindingsText -Result $result
            $infoBox.Text = Format-DeviceRegistrationInfoText -Result $result
            $dsregOutputBox.Text = & $getDsregSnapshotText $result
            $statusText.Text = ([string]::Format($script:Strings.CompletedFormat, $result.Status))
            $script:lastResult = $result

            if ($result.IntuneEnrolled) {
                $enrollmentBanner.Background = $brushConverter.ConvertFromString("#ECFDF3")
                $enrollmentBanner.BorderBrush = $brushConverter.ConvertFromString("#ABEFC6")
                $enrollmentBannerText.Foreground = $brushConverter.ConvertFromString("#027A48")
                $enrollmentBannerText.Text = $script:Strings.DeviceEnrolled
            }
            else {
                $enrollmentBanner.Background = $brushConverter.ConvertFromString("#FEF3F2")
                $enrollmentBanner.BorderBrush = $brushConverter.ConvertFromString("#FECDCA")
                $enrollmentBannerText.Foreground = $brushConverter.ConvertFromString("#B42318")
                $enrollmentBannerText.Text = $script:Strings.DeviceNotEnrolled
            }

            if ($result.IntuneEnrolled -or $result.MdmPolicyStatus -eq "Enabled") {
                $mdmPolicyBanner.Background = $brushConverter.ConvertFromString("#ECFDF3")
                $mdmPolicyBanner.BorderBrush = $brushConverter.ConvertFromString("#ABEFC6")
                $mdmPolicyBannerText.Foreground = $brushConverter.ConvertFromString("#027A48")
            }
            elseif ($result.MdmPolicyStatus -eq "MissingOrDisabled") {
                $mdmPolicyBanner.Background = $brushConverter.ConvertFromString("#FEF3F2")
                $mdmPolicyBanner.BorderBrush = $brushConverter.ConvertFromString("#FECDCA")
                $mdmPolicyBannerText.Foreground = $brushConverter.ConvertFromString("#B42318")
            }
            else {
                $mdmPolicyBanner.Background = $brushConverter.ConvertFromString("#F8FAFC")
                $mdmPolicyBanner.BorderBrush = $brushConverter.ConvertFromString("#DDE7F0")
                $mdmPolicyBannerText.Foreground = $brushConverter.ConvertFromString("#5F6B7A")
            }

            $mdmPolicyBannerText.Text = if ($result.MdmPolicyStatus -eq "Enabled") {
                $script:Strings.MdmPolicyEnabled
            }
            elseif ($result.MdmPolicyStatus -eq "MissingOrDisabled" -and -not $result.MdmPolicyCheckRequired -and $result.IntuneEnrolled) {
                $script:Strings.MdmPolicyMissingInformational
            }
            elseif ($result.MdmPolicyStatus -eq "MissingOrDisabled") {
                $script:Strings.MdmPolicyMissingOrDisabled
            }
            elseif ([string]::IsNullOrWhiteSpace($result.MdmPolicyStatusMessage)) {
                $script:Strings.MdmPolicyUnavailable
            }
            else {
                $result.MdmPolicyStatusMessage
            }

            $script:lastLogFolder = Split-Path -Path $result.LogPath -Parent
            $openLogButton.IsEnabled = ($isAdminMode -and [bool]$script:lastLogFolder)
            $copySummaryButton.IsEnabled = $true
            $createBundleButton.IsEnabled = $true
        }
        catch {
            $issuesBox.Text = ("[ERROR] {0}" -f $_.Exception.Message)
            $infoBox.Text = ""
            $outputBox.Text = $_.Exception.Message
            $dsregOutputBox.Text = ""
            $statusText.Text = $script:Strings.Failed
        }
        finally {
            $runButton.IsEnabled = $true
            $runButton.Content = $script:Strings.RunDiagnostic
            & $updateRepairButtonState
        }
    }

    $runButton.Add_Click({
        & $runDiagnostic $true
    })

    $repairButton.Add_Click({
        & $runDiagnostic $false
    })

    $refreshPrtButton.Add_Click({
        $runButton.IsEnabled = $false
        $refreshPrtButton.IsEnabled = $false
        $repairButton.IsEnabled = $false
        $statusText.Text = $script:Strings.RefreshingPrt

        try {
            $refreshResult = Invoke-DeviceRegistrationPrtRefresh -OutputRoot $OutputRoot
            & $runDiagnostic $true

            $refreshLine = "[INFO] " + ([string]::Format($script:Strings.RefreshPrtCompletedFormat, $refreshResult.ExitCode, $refreshResult.OutputPath))
            if ([string]::IsNullOrWhiteSpace($infoBox.Text) -or $infoBox.Text -eq "[INFO] No additional information.") {
                $infoBox.Text = $refreshLine
            }
            else {
                $infoBox.Text = $refreshLine + [Environment]::NewLine + $infoBox.Text
            }
        }
        catch {
            $issuesBox.Text = ("[ERROR] {0}" -f $_.Exception.Message)
            $statusText.Text = $script:Strings.Failed
        }
        finally {
            $runButton.IsEnabled = $true
            & $updateRepairButtonState
        }
    })

    $openLogButton.Add_Click({
        if ($script:lastLogFolder -and (Test-Path -LiteralPath $script:lastLogFolder)) {
            Start-Process -FilePath explorer.exe -ArgumentList (ConvertTo-SafeArgument -Value $script:lastLogFolder) | Out-Null
        }
    })

    $copySummaryButton.Add_Click({
        if ($script:lastResult) {
            [System.Windows.Clipboard]::SetText((Format-DeviceRegistrationSupportSummary -Result $script:lastResult))
            if (-not [string]::IsNullOrWhiteSpace($Script:SupportEmail)) {
                $emailResult = Send-DeviceRegistrationSupportSummaryEmail -Result $script:lastResult -To $Script:SupportEmail -SendMode $Script:SupportEmailSendMode
                $script:lastResult.SupportSummaryEmailStatus = $emailResult.Status
                $script:lastResult.SupportSummaryEmailMessage = $emailResult.Message
                $script:lastResult.SupportSummaryEmailDraftPath = $emailResult.EmlPath
                $outputBox.Text = Format-DeviceRegistrationResultText -Result $script:lastResult
                $infoBox.Text = Format-DeviceRegistrationInfoText -Result $script:lastResult
                $issuesBox.Text = Format-DeviceRegistrationFindingsText -Result $script:lastResult
                $statusText.Text = if ($emailResult.Status -eq "Sent") {
                    $script:Strings.SupportSummaryCopiedSent
                }
                elseif ($emailResult.Status -eq "DraftOpened") {
                    $script:Strings.SupportSummaryCopiedDraftOpened
                }
                elseif ($emailResult.Status -eq "DraftFileCreated") {
                    $script:Strings.SupportSummaryCopiedDraftFileCreated
                }
                else {
                    $script:Strings.SupportSummaryCopied
                }
            }
            else {
                $statusText.Text = $script:Strings.SupportSummaryCopied
            }
        }
    })

    $createBundleButton.Add_Click({
        if (-not $script:lastResult) {
            return
        }

        try {
            $createBundleButton.IsEnabled = $false
            $statusText.Text = $script:Strings.CreatingSupportBundle
            $root = Split-Path -Path $script:lastResult.CsvPath -Parent
            $paths = [PSCustomObject]@{
                Root        = $root
                Logs        = Join-Path $root "Logs"
                Output      = Join-Path $root "Output"
                Transcripts = Join-Path $root "Transcripts"
            }
            $eventPaths = @(Export-DeviceRegistrationEventLogs -Paths $paths)
            $bundlePath = New-DeviceRegistrationSupportBundle -Paths $paths -Result $script:lastResult -EventLogPaths $eventPaths
            if (-not [string]::IsNullOrWhiteSpace($bundlePath)) {
                $script:lastResult.SupportBundlePath = $bundlePath
                $emailResult = Send-DeviceRegistrationSupportBundleEmail -BundlePath $bundlePath -Result $script:lastResult -To $Script:SupportEmail -SendMode $Script:SupportEmailSendMode
                $script:lastResult.SupportEmailStatus = $emailResult.Status
                $script:lastResult.SupportEmailMessage = $emailResult.Message
                $script:lastResult.SupportEmailDraftPath = $emailResult.EmlPath
                $outputBox.Text = Format-DeviceRegistrationResultText -Result $script:lastResult
                $infoBox.Text = Format-DeviceRegistrationInfoText -Result $script:lastResult
                $statusText.Text = if ($emailResult.Status -eq "Sent") {
                    $script:Strings.SupportBundleSent
                }
                elseif ($emailResult.Status -eq "DraftOpened") {
                    $script:Strings.SupportEmailDraftOpened
                }
                elseif ($emailResult.Status -eq "DraftFileCreated") {
                    $script:Strings.SupportEmailDraftFileCreated
                }
                else {
                    $script:Strings.SupportBundleCreated
                }
                $script:lastLogFolder = Split-Path -Path $bundlePath -Parent
                $openLogButton.IsEnabled = ((& $getSelectedMode) -eq "Admin" -and [bool]$script:lastLogFolder)
            }
        }
        catch {
            $issuesBox.Text = ("[ERROR] {0}" -f $_.Exception.Message)
            $statusText.Text = $script:Strings.SupportBundleFailed
        }
        finally {
            $createBundleButton.IsEnabled = $true
        }
    })

    $window.Add_ContentRendered({
        if ($splash) {
            Hide-SmartM365GuiSplash -Splash $splash
        }

        $window.Dispatcher.BeginInvoke([Action]{
            & $runDiagnostic $true
        }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle) | Out-Null
    })

    [void]$window.ShowDialog()
    if ($splash) {
        Close-SmartM365GuiSplash -Splash $splash
    }
}

$script:ToolConfig = Get-DeviceRegistrationToolConfig -Path $ConfigPath
Apply-DeviceRegistrationToolConfig -Config $script:ToolConfig

Invoke-Relaunch64BitIfNeeded
Invoke-RelaunchElevatedIfNeeded
Invoke-RelaunchStaIfNeeded

if (-not $Cli) {
    Show-SmartM365DeviceRegistrationGui
    exit 0
}

$cliResult = Invoke-SmartM365DeviceRegistrationTool `
    -Mode $Mode `
    -RepairDisabledDeletedDevice:$RepairDisabledDeletedDevice `
    -AllowDsregLeave:$AllowDsregLeave `
    -TriggerJoin:$TriggerJoin `
    -TriggerIntuneAutoEnrollment:$TriggerIntuneAutoEnrollment `
    -AllowRemoveStaleIntuneEnrollment:$AllowRemoveStaleIntuneEnrollment `
    -AllowRemoveNonIntuneMdmEnrollment:$AllowRemoveNonIntuneMdmEnrollment `
    -AllowIntuneEnrolledAction:$AllowIntuneEnrolledAction `
    -AuditOnly:$AuditOnly `
    -RetryCount $RetryCount `
    -RetrySleepMinutes $RetrySleepMinutes `
    -LogRetentionCount $LogRetentionCount `
    -OutputRoot $OutputRoot `
    -SupportBundle:$SupportBundle `
    -NoTranscript:$NoTranscript

if ($JsonOutput) {
    try {
        $jsonRoot = Split-Path -Path $cliResult.CsvPath -Parent
        $jsonDir = Join-Path $jsonRoot "Output"
        if (-not (Test-Path -LiteralPath $jsonDir)) {
            New-Item -ItemType Directory -Path $jsonDir -Force | Out-Null
        }
        $jsonPath = Join-Path $jsonDir ("{0}_result_{1}.json" -f $Script:ComputerName, $Script:RunId)
        $cliResult | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $jsonPath -Encoding UTF8 -Force
        $cliResult.JsonOutputPath = $jsonPath
    }
    catch {
        Write-SmartM365DeviceRegistrationLog ("JSON output failed: {0}" -f $_.Exception.Message)
    }

    $cliResult | ConvertTo-Json -Depth 6
}
else {
    Write-ResultSummary -Result $cliResult
}
exit $cliResult.ExitCode
