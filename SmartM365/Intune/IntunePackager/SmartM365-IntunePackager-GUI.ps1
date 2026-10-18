#Requires -Version 7.0
<#
.SYNOPSIS
    Smart IntunePackager: builds Intune Win32 packages (.intunewin) and publishes them to
    Microsoft Intune through Microsoft Graph.

.DESCRIPTION
    Standalone application: this folder works when copied out of the repository. It depends
    only on Microsoft.Graph.Authentication and on the local modules next to this script.

    Without -Action the WPF interface opens (STA, splash on interactive launch). Packaging,
    Graph sign-in and publication run in a background runspace; the window stays responsive
    and receives progress and log lines.

    CLI mode (no window, no splash) is selected with -Action:
    - Build           : builds the .intunewin defined by the profile.
    - Publish         : publishes an existing .intunewin (-IntuneWinPath or the profile output).
    - BuildAndPublish : builds, then publishes.
    - Import          : reads an existing Win32 app (-ExistingAppId, GET only) and writes its definition
                        to -ProfilePath, with its detection script and logo next to it.

    -DryRun builds the package (for Build/BuildAndPublish) and writes the planned Graph
    requests and payloads without calling Graph. -WhatIf has the same effect on publication.
    -ValidateOnly checks the configuration and the profile without building or connecting
    (CLI), or builds the window without showing it (GUI, with -Demo for the demo window).
    -Demo opens the offline demonstration with fictional data and no Graph call.

    The app is never assigned: create assignments in the Intune admin center.

.PARAMETER Tenant
    Tenant profile name from the local configuration (Tenants section). Default: test.

.PARAMETER PublishMode
    Overrides Publish.Mode of the profile (Create or UpdateExisting).

.PARAMETER ExistingAppId
    Overrides Publish.ExistingAppId of the profile (required for UpdateExisting).

.VERSION
    1.1.3

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com

.EXAMPLE
    pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-IntunePackager-GUI.ps1 -Tenant test -ProfilePath .\Examples\Contoso-SampleApp.profile.json.template -Action BuildAndPublish -DryRun
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$Tenant = 'test',
    [string]$ProfilePath = '',
    [ValidateSet('', 'Build', 'Publish', 'BuildAndPublish', 'Import')][string]$Action = '',
    [string]$IntuneWinPath = '',
    [switch]$DryRun,
    [ValidateSet('', 'Create', 'UpdateExisting')][string]$PublishMode = '',
    [string]$ExistingAppId = '',
    [switch]$UseDeviceCode,
    # Validation option: forces one Azure Storage SAS renewal during the upload.
    [switch]$TestSasRenewal,
    # UpdateExisting: publish even if the app changed in Intune since its definition was imported.
    [switch]$Force,
    # Import: connect with DeviceManagementApps.ReadWrite.All instead of DeviceManagementApps.Read.All.
    [switch]$UsePublishScope,
    [switch]$ValidateOnly,
    [switch]$Demo,
    [ValidateSet('Overview', 'Issues', 'Completed')][string]$DemoScenario = 'Overview',
    [string]$DemoView = '',
    [ValidateRange(1000, 3840)][int]$DemoWindowWidth = 1500,
    [ValidateRange(650, 2160)][int]$DemoWindowHeight = 950
)

# The demo must start before any module import or configuration access.
if ($Demo) {
    . (Join-Path $PSScriptRoot 'SmartM365.GuiDemo.ps1')
    Start-SmartM365GuiDemo -GuiScriptPath $PSCommandPath -ProfilePath (Join-Path $PSScriptRoot 'SmartM365-IntunePackager-Demo.ps1') -Scenario $DemoScenario -View $DemoView -Width $DemoWindowWidth -Height $DemoWindowHeight -ValidateOnly:$ValidateOnly
    return
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:AppName = 'Smart IntunePackager'
$script:AppVersion = '1.1.3'
$script:AppRoot = $PSScriptRoot
$script:ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-IntunePackager-GUI.config.json'
$script:ConfigTemplatePath = $script:ConfigPath + '.template'
$script:ModuleNames = @('SmartM365.IntunePackager.Core', 'SmartM365.IntunePackager.Packaging', 'SmartM365.IntunePackager.Graph')
$script:TenantExplicit = $PSBoundParameters.ContainsKey('Tenant')

foreach ($moduleName in $script:ModuleNames) {
    Import-Module (Join-Path -Path $PSScriptRoot -ChildPath ($moduleName + '.psm1')) -Force -ErrorAction Stop
}

#region Shared helpers

function Initialize-IpApplication {
    <#
    .SYNOPSIS
        Loads the local configuration (merged with the template), starts the log and applies Graph settings.
    #>
    [CmdletBinding()]
    param([string]$LogBaseName = 'SmartM365-IntunePackager')

    $config = Import-IpConfiguration -ConfigPath $script:ConfigPath -TemplatePath $script:ConfigTemplatePath
    $logDirectory = Resolve-IpPath -Path $config.Paths.LogDirectory -BaseDirectory $script:AppRoot
    if (-not $logDirectory) { $logDirectory = Join-Path $script:AppRoot 'Logs' }
    $maxFiles = [int]$config.Logging.MaxLogFiles
    if ($maxFiles -lt 1) { $maxFiles = 10 }
    $logPath = Initialize-IpLog -Directory $logDirectory -BaseName $LogBaseName -MaxFiles $maxFiles
    Set-IpGraphSettings -BaseUri ([string]$config.Graph.BaseUri) -MaxRetries ([int]$config.Graph.MaxRetries) -PollSeconds ([int]$config.Graph.PollSeconds) -PollTimeoutMinutes ([int]$config.Graph.PollTimeoutMinutes)
    Write-IpLog ("{0} {1}. Log: {2}" -f $script:AppName, $script:AppVersion, $logPath)
    Write-IpLog ("Configuration: {0}" -f $script:ConfigPath)
    return $config
}

function Resolve-IpTenantSelection {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][psobject]$Config, [Parameter(Mandatory = $true)][string]$Name)
    try { return (Get-IpTenantProfile -Config $Config -Tenant $Name) }
    catch {
        throw ("{0}. Add or complete the '{1}' entry under Tenants in {2}." -f $_.Exception.Message.TrimEnd('.'), $Name, $script:ConfigPath)
    }
}

function Get-IpProfileOutputFolder {
    param([psobject]$Definition, [psobject]$Config, [string]$BaseDirectory)
    $folder = Resolve-IpPath -Path ([string]$Definition.Package.OutputFolder) -BaseDirectory $BaseDirectory
    if (-not $folder) { $folder = Resolve-IpPath -Path ([string]$Config.Paths.DefaultOutputFolder) -BaseDirectory $script:AppRoot }
    if (-not $folder) { throw 'No output folder: set Package.OutputFolder in the profile or Paths.DefaultOutputFolder in the configuration.' }
    return $folder
}

function Get-IpExpectedPackagePath {
    param([psobject]$Definition, [string]$OutputFolder)
    return (Join-Path $OutputFolder ([IO.Path]::GetFileNameWithoutExtension([string]$Definition.Package.SetupFile) + '.intunewin'))
}

function Test-IpPackageMatchesDefinition {
    <#
    .SYNOPSIS
        Warns when the package MSI information differs from the MSI detection rule of the profile.
    #>
    param([psobject]$Definition, [psobject]$Metadata)
    if ($Definition.Detection.Type -ne 'MsiProductCode' -or $null -eq $Metadata.MsiInfo) { return }
    $packageCode = [string]$Metadata.MsiInfo.ProductCode
    $profileCode = [string]$Definition.Detection.Msi.ProductCode
    if ($packageCode -and $packageCode -ne $profileCode) {
        Write-IpLog ("The package MSI ProductCode {0} differs from the detection rule ProductCode {1}." -f $packageCode, $profileCode) -Level WARN
    }
}

function Write-IpErrorDetail {
    <#
    .SYNOPSIS
        Logs a user-facing error message, and its location and stack trace in the log file only.
    #>
    param([Management.Automation.ErrorRecord]$ErrorRecord)
    Write-IpLog $ErrorRecord.Exception.Message -Level ERROR
    $location = if ($ErrorRecord.InvocationInfo) { '{0}:{1}' -f $ErrorRecord.InvocationInfo.ScriptName, $ErrorRecord.InvocationInfo.ScriptLineNumber } else { 'unknown' }
    Write-IpLog ("Error location: {0}`n{1}" -f $location, $ErrorRecord.ScriptStackTrace) -Level ERROR -FileOnly
}

function Write-IpProcessExitNotice {
    <#
    .SYNOPSIS
        Records the final status before the process ends. When the native WAM broker is loaded,
        a crash can still happen during process shutdown: the logged result stays authoritative.
    #>
    param([string]$Status)
    Write-IpLog $Status
    if (Test-IpGraphBrokerLoaded) {
        Write-IpLog 'The WAM broker runtime (msalruntime.dll) is loaded in this process. If PowerShell then reports "Internal CLR error (0x80131506)", the crash happens during process shutdown, after the result above; the process exit code is then 0x80131506 instead of the logged value.' -Level WARN
    }
}

#endregion

#region CLI

function Invoke-IpCliImport {
    <#
    .SYNOPSIS
        CLI import: reads one Win32 app (GET only) and writes its definition profile, detection script and logo.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][psobject]$Config)

    if (-not (Test-IpGuid $ExistingAppId)) { throw 'Command line error: -Action Import requires -ExistingAppId <app id>.' }
    if ([string]::IsNullOrWhiteSpace($ProfilePath)) { throw 'Command line error: -Action Import requires -ProfilePath <new profile .json>.' }
    $target = [IO.Path]::GetFullPath($ProfilePath)
    if ((Test-Path -LiteralPath $target) -and -not $Force) { throw "The profile already exists: $target. Choose another -ProfilePath or add -Force to replace it." }
    $tenantProfile = Resolve-IpTenantSelection -Config $Config -Name $Tenant
    $scopes = if ($UsePublishScope) { Get-IpGraphScopes } else { @('DeviceManagementApps.Read.All') }
    Connect-IpGraph -TenantProfile $tenantProfile -Scopes $scopes -UseDeviceCode:$UseDeviceCode | Out-Null
    try {
        $app = Get-IpWin32App -AppId $ExistingAppId -Step 'Import'
        $import = ConvertFrom-IpWin32App -App $app
        $folder = Split-Path -Path $target -Parent
        $baseName = [IO.Path]::GetFileNameWithoutExtension($target) -replace '\.profile$', ''
        $definition = Save-IpImportAssets -Import $import -Folder $folder -BaseName $baseName -Relative
        Export-IpAppProfile -Definition $definition -Path $target -Confirm:$false
        Write-IpLog ("Definition imported from Intune: {0} ({1}), content version {2}. Profile: {3}" -f $app.displayName, $app.id, $app.committedContentVersion, $target)
        foreach ($item in $import.Preserved) { Write-IpLog ('Preserved in Intune: ' + $item) }
        foreach ($warning in $import.Warnings) { Write-IpLog $warning -Level WARN }
        Write-IpLog 'To publish: set Package.SourceFolder and Package.OutputFolder in the profile and use -Action BuildAndPublish, or use -Action Publish -IntuneWinPath with an existing package.'
    }
    finally {
        Disconnect-IpGraph
        Write-IpLog 'Disconnected from Microsoft Graph.'
    }
}

function Sync-IpImportedFromAfterPublish {
    <#
    .SYNOPSIS
        After a successful publication, records the app state read by [Verify] in ImportedFrom (in
        memory and, when a profile file is given, in that file), so the next publication with the same
        profile is not blocked by [Concurrency] because of its own previous update.
    #>
    param([Parameter(Mandatory = $true)][psobject]$Definition, [Parameter(Mandatory = $true)]$Result, [string]$ProfilePath = '')
    if (-not (Update-IpImportedFrom -Definition $Definition -VerifiedState (Get-IpObjectProperty $Result 'VerifiedState'))) { return $false }
    $imported = $Definition.ImportedFrom
    $summary = 'ImportedFrom updated: content version {0}, lastModifiedDateTime {1}.' -f (Get-IpObjectProperty $imported 'ContentVersion'), (Get-IpObjectProperty $imported 'LastModifiedDateTime')
    if (-not $ProfilePath -or -not (Test-Path -LiteralPath $ProfilePath -PathType Leaf) -or $ProfilePath -match '\.template$') { Write-IpLog $summary; return $true }
    try {
        if (Save-IpProfileImportedFrom -Path $ProfilePath -ImportedFrom $imported -Confirm:$false) { Write-IpLog ('{0} Profile saved: {1}' -f $summary, $ProfilePath) }
        else { Write-IpLog ('{0} The profile file was not imported from this app and was left unchanged: {1}' -f $summary, $ProfilePath) }
    }
    catch { Write-IpLog ('The publication succeeded but ImportedFrom could not be saved in the profile ({0}): {1}. Re-import the definition before the next publication.' -f $ProfilePath, $_.Exception.Message) -Level WARN }
    return $true
}

function Resolve-IpCliOutputFolder {
    <#
    .SYNOPSIS
        Returns the output folder of a CLI run. It is only needed to build, or to find the package
        when -IntuneWinPath is not given; publishing a given package needs no output or source folder.
    #>
    param([psobject]$Definition, [psobject]$Config, [string]$BaseDirectory, [bool]$NeedsBuild, [string]$IntuneWinPath)
    if (-not $NeedsBuild -and $IntuneWinPath) { return '' }
    return (Get-IpProfileOutputFolder -Definition $Definition -Config $Config -BaseDirectory $BaseDirectory)
}

function Invoke-IpCli {
    [CmdletBinding()]
    param()

    $config = Initialize-IpApplication
    Write-IpLog ("CLI mode. Tenant={0}; Action={1}; DryRun={2}; ValidateOnly={3}" -f $Tenant, $Action, [bool]$DryRun, [bool]$ValidateOnly)
    if ($Action -eq 'Import') { Invoke-IpCliImport -Config $config; return }

    if ([string]::IsNullOrWhiteSpace($ProfilePath)) { throw 'Command line error: -ProfilePath is required in CLI mode.' }
    $profileFullPath = (Resolve-Path -LiteralPath $ProfilePath -ErrorAction Stop).ProviderPath
    $baseDirectory = Split-Path -Path $profileFullPath -Parent
    $definition = Import-IpAppProfile -Path $profileFullPath
    if ($PublishMode) { $definition.Publish.Mode = $PublishMode }
    if ($ExistingAppId) { $definition.Publish.ExistingAppId = $ExistingAppId }
    Write-IpLog ("Profile: {0}; App={1}; PublishMode={2}" -f $profileFullPath, $definition.App.DisplayName, $definition.Publish.Mode)

    $needsBuild = $Action -in @('Build', 'BuildAndPublish') -or ($ValidateOnly -and -not $Action)
    $errors = @(Test-IpAppDefinition -Definition $definition -BaseDirectory $baseDirectory -ForPublish:(-not $needsBuild))
    if ($errors.Count -gt 0) {
        foreach ($validationError in $errors) { Write-IpLog $validationError -Level ERROR }
        throw ("The profile is invalid ({0} error(s))." -f $errors.Count)
    }
    $publishes = $Action -in @('Publish', 'BuildAndPublish')
    if ($publishes -or $ValidateOnly) { $tenantProfile = Resolve-IpTenantSelection -Config $config -Name $Tenant }

    if ($ValidateOnly) {
        Write-IpLog 'Validation completed: configuration and profile are valid. Nothing was built or published.'
        return
    }
    if (-not $Action) { throw 'Command line error: -Action is required in CLI mode.' }

    $outputFolder = Resolve-IpCliOutputFolder -Definition $definition -Config $config -BaseDirectory $baseDirectory -NeedsBuild $needsBuild -IntuneWinPath $IntuneWinPath
    $packagePath = ''
    if ($needsBuild) {
        $build = Invoke-IpPackageBuild -EngineName ([string]$config.Packaging.Engine) `
            -SourceFolder (Resolve-IpPath -Path $definition.Package.SourceFolder -BaseDirectory $baseDirectory) `
            -SetupFile ([string]$definition.Package.SetupFile) -OutputFolder $outputFolder -Settings $config -AppRoot $script:AppRoot
        $packagePath = $build.IntuneWinPath
        Test-IpPackageMatchesDefinition -Definition $definition -Metadata $build.Metadata
    }
    if ($Action -eq 'Build') { return }

    if (-not $packagePath) {
        $packagePath = if ($IntuneWinPath) { (Resolve-Path -LiteralPath $IntuneWinPath -ErrorAction Stop).ProviderPath }
        else { Get-IpExpectedPackagePath -Definition $definition -OutputFolder $outputFolder }
        if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf)) { throw "Package not found: $packagePath. Build it first or pass -IntuneWinPath." }
        Write-IpLog "Publishing existing package: $packagePath"
        Test-IpPackageMatchesDefinition -Definition $definition -Metadata (Read-IpIntuneWinMetadata -Path $packagePath)
    }

    if ($DryRun -or $WhatIfPreference) {
        $dryRunFolder = Join-Path (Split-Path $packagePath -Parent) ('DryRun\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Write-IpLog ("Dry run for tenant profile {0} ({1}). Microsoft Graph is not called." -f $tenantProfile.Name, $tenantProfile.AuthMode)
        [void](Publish-IpWin32App -Definition $definition -IntuneWinPath $packagePath -BaseDirectory $baseDirectory -UploadSettings $config.Upload -DryRun -DryRunOutputFolder $dryRunFolder)
        return
    }

    $uploadSettings = $config.Upload
    if ($TestSasRenewal) {
        $uploadSettings = Copy-IpJsonValue $config.Upload
        $uploadSettings | Add-Member -NotePropertyName ForceSasRenewalAtBlock -NotePropertyValue 3 -Force
        Write-IpLog 'Validation option -TestSasRenewal: a SAS renewal is requested after upload block 2, before block 3 (blocks counted from 1; the package needs at least 3 blocks).' -Level WARN
    }

    Connect-IpGraph -TenantProfile $tenantProfile -UseDeviceCode:$UseDeviceCode | Out-Null
    try {
        $result = Publish-IpWin32App -Definition $definition -IntuneWinPath $packagePath -BaseDirectory $baseDirectory -UploadSettings $uploadSettings -Force:$Force `
            -ProgressCallback {
            param($sent, $total)
            if ($total -gt 0) {
                $percent = [int](100 * $sent / $total)
                if ($percent -ge $script:IpNextProgress) {
                    Write-IpLog ("[Upload] {0}% ({1} / {2})" -f $percent, (Format-IpByteSize $sent), (Format-IpByteSize $total))
                    $script:IpNextProgress = [Math]::Min(100, $percent + 10)
                }
            }
        }
        if ($null -eq $result -or $result.Mode -ne 'Publish') {
            Write-IpLog 'Publication was not confirmed. Nothing was published.' -Level WARN
            return
        }
        Write-IpLog ("Result: AppId={0}; ContentVersion={1}; Tenant={2}; Assigned=False" -f $result.AppId, $result.ContentVersionId, $result.TenantId)
        foreach ($change in @($result.Changes)) { Write-IpLog ("Changed: {0}: {1} -> {2}" -f $change.Property, $change.Before, $change.After) }
        if (@($result.VerificationIssues).Count) { Write-IpLog ("Post-publication check found {0} issue(s); see the [Verify] lines above." -f @($result.VerificationIssues).Count) -Level WARN }
        [void](Sync-IpImportedFromAfterPublish -Definition $definition -Result $result -ProfilePath $profileFullPath)
        Write-IpLog ("Republish this profile as a new content version with: -PublishMode UpdateExisting -ExistingAppId {0}" -f $result.AppId)
    }
    catch {
        if (-not $_.Exception.Data['IpConcurrency']) { throw }
        # Concurrency block: one error line, then the two possible actions as exact command lines.
        Write-IpLog $_.Exception.Message -Level ERROR
        $commands = Get-IpConcurrencyCommands -ProfilePath $profileFullPath -AppId ([string]$definition.Publish.ExistingAppId) -PackagePath $(if ($needsBuild) { '' } else { $packagePath })
        Write-IpLog ('Nothing was published. Re-import the definition (replaces the profile; set the package folders again if you use BuildAndPublish): ' + $commands.Reimport)
        Write-IpLog ('Or publish anyway, replacing the portal changes on the properties managed by this tool: ' + $commands.Force)
        $_.Exception.Data['IpLogged'] = $true
        throw
    }
    finally {
        Disconnect-IpGraph
        Write-IpLog 'Disconnected from Microsoft Graph.'
    }
}

function Get-IpConcurrencyCommands {
    <#
    .SYNOPSIS
        Returns the exact one-line commands to re-import the definition or to republish with -Force.
    #>
    param([Parameter(Mandatory = $true)][string]$ProfilePath, [Parameter(Mandatory = $true)][string]$AppId, [string]$PackagePath = '', [string]$ScriptPath = $PSCommandPath, [string]$TenantName = $Tenant, [string]$PublishAction = $Action)
    $base = 'pwsh -NoProfile -ExecutionPolicy Bypass -File "{0}" -Tenant {1}' -f $ScriptPath, $TenantName
    $reimport = '{0} -Action Import -ExistingAppId {1} -ProfilePath "{2}" -Force' -f $base, $AppId, $ProfilePath
    $publish = '{0} -ProfilePath "{1}" -Action {2} -PublishMode UpdateExisting -ExistingAppId {3}' -f $base, $ProfilePath, $PublishAction, $AppId
    if ($PackagePath) { $publish += (' -IntuneWinPath "{0}"' -f $PackagePath) }
    return [pscustomobject]@{ Reimport = $reimport; Force = $publish + ' -Force' }
}

#endregion

#region GUI markup

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart IntunePackager" Width="1360" Height="900" MinWidth="1100" MinHeight="700"
        WindowStartupLocation="CenterScreen" Background="#F4F7FB" FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="12,5"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="MinWidth" Value="84"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#B9C7D8"/>
            <Setter Property="Foreground" Value="#18324A"/>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#0F6CBD"/>
            <Setter Property="BorderBrush" Value="#0F6CBD"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,4,16,8"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style x:Key="FieldLabel" TargetType="TextBlock">
            <Setter Property="Margin" Value="0,5,12,8"/>
            <Setter Property="Foreground" Value="#3D5A73"/>
            <Setter Property="VerticalAlignment" Value="Top"/>
        </Style>
        <Style x:Key="SectionTitle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#0F2A44"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
        </Style>
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#DCE4EE"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="0,0,0,12"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#E3EAF2"/>
            <Setter Property="BorderBrush" Value="#DCE4EE"/>
            <Setter Property="RowHeight" Value="26"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
        </Style>
    </Window.Resources>
    <DockPanel>
        <Border DockPanel.Dock="Top" Background="#0F2A44" Padding="16,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <Image x:Name="HeaderLogo" Height="34" Margin="0,0,14,0" ToolTip="Open WorkplaceCloudHub.com" Cursor="Hand"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock x:Name="AppTitleText" Text="Smart IntunePackager" Foreground="#FFFFFF" FontSize="18" FontWeight="SemiBold"/>
                        <TextBlock Text="Build Win32 packages and publish them to Microsoft Intune" Foreground="#B9D3EC" FontSize="12"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Tenant" Foreground="#B9D3EC" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <ComboBox x:Name="TenantCombo" Width="170" Margin="0,0,10,0" VerticalAlignment="Center"/>
                    <Button x:Name="ConnectButton" Content="Connect" Style="{StaticResource PrimaryButton}"/>
                    <Button x:Name="DisconnectButton" Content="Disconnect"/>
                    <Border Background="#1D3F61" CornerRadius="10" Padding="10,3" Margin="4,0,0,0" VerticalAlignment="Center">
                        <TextBlock x:Name="ConnectionStatusText" Text="Not connected" Foreground="#FFFFFF"/>
                    </Border>
                </StackPanel>
            </Grid>
        </Border>
        <Border x:Name="SignInBanner" DockPanel.Dock="Top" Background="#FFF4CE" BorderBrush="#E0B000" BorderThickness="0,0,0,1" Padding="16,8" Visibility="Collapsed">
            <TextBlock x:Name="SignInBannerText" TextWrapping="Wrap" Foreground="#5C4300" FontWeight="SemiBold" Text="Sign-in window opened."/>
        </Border>
        <Border DockPanel.Dock="Top" Background="#FFFFFF" BorderBrush="#DCE4EE" BorderThickness="0,0,0,1" Padding="16,8">
            <DockPanel>
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                    <Button x:Name="NewProfileButton" Content="New profile"/>
                    <Button x:Name="OpenProfileButton" Content="Open profile..."/>
                    <Button x:Name="SaveProfileButton" Content="Save profile"/>
                    <Button x:Name="SaveProfileAsButton" Content="Save as..."/>
                    <Button x:Name="ValidateButton" Content="Validate"/>
                </StackPanel>
                <TextBlock x:Name="ProfilePathText" Text="Unsaved profile" Foreground="#5A7389" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="12,0,0,0"/>
            </DockPanel>
        </Border>
        <Border DockPanel.Dock="Bottom" Background="#FFFFFF" BorderBrush="#DCE4EE" BorderThickness="0,1,0,0" Padding="16,8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="170"/>
                </Grid.RowDefinitions>
                <DockPanel Margin="0,0,0,6">
                    <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
                        <Button x:Name="CancelOperationButton" Content="Cancel operation" IsEnabled="False"/>
                        <Button x:Name="ExportLogButton" Content="Export log..."/>
                        <Button x:Name="OpenLogFolderButton" Content="Open log folder"/>
                        <Button x:Name="ClearLogButton" Content="Clear" Margin="0"/>
                    </StackPanel>
                    <ProgressBar x:Name="OperationProgress" Width="220" Height="12" Minimum="0" Maximum="100" Margin="0,0,12,0" VerticalAlignment="Center"/>
                    <TextBlock x:Name="StatusText" Text="Ready" VerticalAlignment="Center" Foreground="#18324A" TextTrimming="CharacterEllipsis"/>
                </DockPanel>
                <TextBox x:Name="LogTextBox" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" FontFamily="Consolas" FontSize="12" Background="#0F1E2D" Foreground="#D7E3EE" Margin="0" VerticalContentAlignment="Top"/>
            </Grid>
        </Border>
        <TabControl x:Name="MainTabs" Margin="16,12,16,8" Background="Transparent" BorderThickness="0">
            <TabItem Header="Package">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="0,12,0,0">
                        <Border Style="{StaticResource Card}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.ColumnSpan="3" Style="{StaticResource SectionTitle}" Text="Source and output"/>
                                <TextBlock Grid.Row="1" Style="{StaticResource FieldLabel}" Text="Source folder"/>
                                <TextBox x:Name="SourceFolderBox" Grid.Row="1" Grid.Column="1"/>
                                <Button x:Name="BrowseSourceButton" Grid.Row="1" Grid.Column="2" Content="Browse..." Margin="8,0,0,8"/>
                                <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Text="Setup file"/>
                                <TextBox x:Name="SetupFileBox" Grid.Row="2" Grid.Column="1" ToolTip="Path relative to the source folder"/>
                                <Button x:Name="BrowseSetupButton" Grid.Row="2" Grid.Column="2" Content="Browse..." Margin="8,0,0,8"/>
                                <TextBlock Grid.Row="3" Style="{StaticResource FieldLabel}" Text="Output folder"/>
                                <TextBox x:Name="OutputFolderBox" Grid.Row="3" Grid.Column="1"/>
                                <Button x:Name="BrowseOutputButton" Grid.Row="3" Grid.Column="2" Content="Browse..." Margin="8,0,0,8"/>
                                <StackPanel Grid.Row="4" Grid.Column="1" Orientation="Horizontal" Margin="0,4,0,0">
                                    <Button x:Name="DetectMetadataButton" Content="Detect MSI / PSADT"/>
                                    <Button x:Name="BuildButton" Content="Build package" Style="{StaticResource PrimaryButton}"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                        <Border Style="{StaticResource Card}">
                            <StackPanel>
                                <TextBlock Style="{StaticResource SectionTitle}" Text="Package"/>
                                <TextBlock x:Name="PackageInfoText" Text="No package built in this session." TextWrapping="Wrap" Foreground="#3D5A73"/>
                                <TextBlock Text="Packaging uses the official Microsoft Win32 Content Prep Tool (pinned version, SHA256 and Microsoft signature verified, downloaded on first use)." Margin="0,10,0,0" Foreground="#6C8295" TextWrapping="Wrap" FontSize="12"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="App information">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <Border Style="{StaticResource Card}" Margin="0,12,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="170"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Style="{StaticResource FieldLabel}" Text="Display name"/>
                            <TextBox x:Name="DisplayNameBox" Grid.Column="1"/>
                            <TextBlock Grid.Row="1" Style="{StaticResource FieldLabel}" Text="Description"/>
                            <TextBox x:Name="DescriptionBox" Grid.Row="1" Grid.Column="1" Height="70" TextWrapping="Wrap" AcceptsReturn="True" VerticalContentAlignment="Top" VerticalScrollBarVisibility="Auto"/>
                            <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Text="Publisher"/>
                            <TextBox x:Name="PublisherBox" Grid.Row="2" Grid.Column="1"/>
                            <TextBlock Grid.Row="3" Style="{StaticResource FieldLabel}" Text="Version"/>
                            <TextBox x:Name="VersionBox" Grid.Row="3" Grid.Column="1"/>
                            <TextBlock Grid.Row="4" Style="{StaticResource FieldLabel}" Text="Developer"/>
                            <TextBox x:Name="DeveloperBox" Grid.Row="4" Grid.Column="1"/>
                            <TextBlock Grid.Row="5" Style="{StaticResource FieldLabel}" Text="Owner"/>
                            <TextBox x:Name="OwnerBox" Grid.Row="5" Grid.Column="1"/>
                            <TextBlock Grid.Row="6" Style="{StaticResource FieldLabel}" Text="Information URL"/>
                            <TextBox x:Name="InformationUrlBox" Grid.Row="6" Grid.Column="1"/>
                            <TextBlock Grid.Row="7" Style="{StaticResource FieldLabel}" Text="Privacy URL"/>
                            <TextBox x:Name="PrivacyUrlBox" Grid.Row="7" Grid.Column="1"/>
                            <TextBlock Grid.Row="8" Style="{StaticResource FieldLabel}" Text="Notes"/>
                            <TextBox x:Name="NotesBox" Grid.Row="8" Grid.Column="1" Height="54" TextWrapping="Wrap" AcceptsReturn="True" VerticalContentAlignment="Top"/>
                            <TextBlock Grid.Row="9" Style="{StaticResource FieldLabel}" Text="Logo (png, jpg)"/>
                            <TextBox x:Name="LogoPathBox" Grid.Row="9" Grid.Column="1"/>
                            <StackPanel Grid.Row="9" Grid.Column="2" Orientation="Horizontal" Margin="8,0,0,8">
                                <Button x:Name="BrowseLogoButton" Content="Browse..."/>
                                <Border BorderBrush="#DCE4EE" BorderThickness="1" Width="48" Height="48" CornerRadius="4">
                                    <Image x:Name="AppIconPreview" Stretch="Uniform"/>
                                </Border>
                            </StackPanel>
                        </Grid>
                    </Border>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Program">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="0,12,0,0">
                        <Border Style="{StaticResource Card}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Style="{StaticResource FieldLabel}" Text="Install command"/>
                                <TextBox x:Name="InstallCommandBox" Grid.Column="1" FontFamily="Consolas"/>
                                <TextBlock Grid.Row="1" Style="{StaticResource FieldLabel}" Text="Uninstall command"/>
                                <TextBox x:Name="UninstallCommandBox" Grid.Row="1" Grid.Column="1" FontFamily="Consolas"/>
                                <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Text="Install behavior"/>
                                <ComboBox x:Name="RunAsCombo" Grid.Row="2" Grid.Column="1" Width="220" HorizontalAlignment="Left"/>
                                <TextBlock Grid.Row="3" Style="{StaticResource FieldLabel}" Text="Device restart behavior"/>
                                <ComboBox x:Name="RestartCombo" Grid.Row="3" Grid.Column="1" Width="220" HorizontalAlignment="Left"/>
                                <TextBlock Grid.Row="4" Style="{StaticResource FieldLabel}" Text="Max run time (minutes)"/>
                                <TextBox x:Name="MaxRunTimeBox" Grid.Row="4" Grid.Column="1" Width="120" HorizontalAlignment="Left"/>
                            </Grid>
                        </Border>
                        <Border Style="{StaticResource Card}">
                            <StackPanel>
                                <TextBlock Style="{StaticResource SectionTitle}" Text="Return codes"/>
                                <DataGrid x:Name="ReturnCodesGrid" Height="190">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Return code" Binding="{Binding ReturnCode}" Width="160"/>
                                        <DataGridComboBoxColumn Header="Code type" SelectedItemBinding="{Binding Type}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Orientation="Horizontal">
                                    <Button x:Name="AddReturnCodeButton" Content="Add"/>
                                    <Button x:Name="RemoveReturnCodeButton" Content="Remove"/>
                                    <Button x:Name="DefaultReturnCodesButton" Content="Intune defaults"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Requirements">
                <Border Style="{StaticResource Card}" Margin="0,12,0,12" VerticalAlignment="Top">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="170"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Style="{StaticResource FieldLabel}" Text="Architectures"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <CheckBox x:Name="ArchX64Check" Content="x64"/>
                            <CheckBox x:Name="ArchArm64Check" Content="arm64"/>
                            <CheckBox x:Name="ArchX86Check" Content="x86"/>
                            <CheckBox x:Name="ArchNotRestrictedCheck" Content="Not restricted (any architecture)"/>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Style="{StaticResource FieldLabel}" Text="Minimum Windows release"/>
                        <ComboBox x:Name="MinWindowsCombo" Grid.Row="1" Grid.Column="1" Width="220" HorizontalAlignment="Left"/>
                        <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Text="Minimum free disk space (MB)"/>
                        <TextBox x:Name="MinDiskSpaceBox" Grid.Row="2" Grid.Column="1" Width="120" HorizontalAlignment="Left" ToolTip="Empty: no disk space requirement"/>
                        <TextBlock Grid.Row="3" Grid.Column="1" Foreground="#6C8295" TextWrapping="Wrap" FontSize="12" Text="Sent as allowedArchitectures, minimumSupportedWindowsRelease and minimumFreeDiskSpaceInMB. Memory, processor and script requirement rules are not edited here; on an imported app they are kept as they are in Intune."/>
                    </Grid>
                </Border>
            </TabItem>
            <TabItem Header="Detection">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="0,12,0,0">
                        <Border Style="{StaticResource Card}">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Detection rule type" Style="{StaticResource FieldLabel}" Width="158"/>
                                <ComboBox x:Name="DetectionTypeCombo" Width="260"/>
                            </StackPanel>
                        </Border>
                        <Border x:Name="MsiPanel" Style="{StaticResource Card}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.ColumnSpan="2" Style="{StaticResource SectionTitle}" Text="MSI product code"/>
                                <TextBlock Grid.Row="1" Style="{StaticResource FieldLabel}" Text="Product code"/>
                                <TextBox x:Name="MsiProductCodeBox" Grid.Row="1" Grid.Column="1" FontFamily="Consolas"/>
                                <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Text="Version check"/>
                                <ComboBox x:Name="MsiVersionOperatorCombo" Grid.Row="2" Grid.Column="1" Width="220" HorizontalAlignment="Left"/>
                                <TextBlock Grid.Row="3" Style="{StaticResource FieldLabel}" Text="Product version"/>
                                <TextBox x:Name="MsiVersionBox" Grid.Row="3" Grid.Column="1" Width="220" HorizontalAlignment="Left"/>
                            </Grid>
                        </Border>
                        <Border x:Name="FilePanel" Style="{StaticResource Card}">
                            <StackPanel>
                                <TextBlock Style="{StaticResource SectionTitle}" Text="File or folder rules"/>
                                <DataGrid x:Name="FileRulesGrid" Height="170">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Path" Binding="{Binding Path}" Width="2*"/>
                                        <DataGridTextColumn Header="File or folder" Binding="{Binding FileOrFolderName}" Width="*"/>
                                        <DataGridComboBoxColumn Header="Detection method" SelectedItemBinding="{Binding DetectionType}" Width="130"/>
                                        <DataGridComboBoxColumn Header="Operator" SelectedItemBinding="{Binding Operator}" Width="140"/>
                                        <DataGridTextColumn Header="Value" Binding="{Binding DetectionValue}" Width="*"/>
                                        <DataGridCheckBoxColumn Header="32-bit on 64-bit" Binding="{Binding Check32BitOn64System}" Width="110"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Orientation="Horizontal">
                                    <Button x:Name="AddFileRuleButton" Content="Add rule"/>
                                    <Button x:Name="RemoveFileRuleButton" Content="Remove rule"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                        <Border x:Name="RegistryPanel" Style="{StaticResource Card}">
                            <StackPanel>
                                <TextBlock Style="{StaticResource SectionTitle}" Text="Registry rules"/>
                                <DataGrid x:Name="RegistryRulesGrid" Height="170">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Key path" Binding="{Binding KeyPath}" Width="2*"/>
                                        <DataGridTextColumn Header="Value name" Binding="{Binding ValueName}" Width="*"/>
                                        <DataGridComboBoxColumn Header="Detection method" SelectedItemBinding="{Binding DetectionType}" Width="130"/>
                                        <DataGridComboBoxColumn Header="Operator" SelectedItemBinding="{Binding Operator}" Width="140"/>
                                        <DataGridTextColumn Header="Value" Binding="{Binding DetectionValue}" Width="*"/>
                                        <DataGridCheckBoxColumn Header="32-bit on 64-bit" Binding="{Binding Check32BitOn64System}" Width="110"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Orientation="Horizontal">
                                    <Button x:Name="AddRegistryRuleButton" Content="Add rule"/>
                                    <Button x:Name="RemoveRegistryRuleButton" Content="Remove rule"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                        <Border x:Name="PreservedDetectionPanel" Style="{StaticResource Card}">
                            <StackPanel>
                                <TextBlock Style="{StaticResource SectionTitle}" Text="Detection kept as in Intune"/>
                                <TextBlock TextWrapping="Wrap" Foreground="#3D5A73" Text="The imported app combines several rule types or rules this tool cannot edit. Its detection rules are not sent: Intune keeps them unchanged. Choose another type to replace them."/>
                            </StackPanel>
                        </Border>
                        <Border x:Name="ScriptPanel" Style="{StaticResource Card}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.ColumnSpan="3" Style="{StaticResource SectionTitle}" Text="Custom PowerShell detection script"/>
                                <TextBlock Grid.Row="1" Style="{StaticResource FieldLabel}" Text="Script file (.ps1)"/>
                                <TextBox x:Name="ScriptPathBox" Grid.Row="1" Grid.Column="1"/>
                                <Button x:Name="BrowseScriptButton" Grid.Row="1" Grid.Column="2" Content="Browse..." Margin="8,0,0,8"/>
                                <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal">
                                    <CheckBox x:Name="EnforceSignatureCheck" Content="Enforce script signature check"/>
                                    <CheckBox x:Name="RunAs32BitCheck" Content="Run script as 32-bit process"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Publish">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="0,12,0,0">
                        <Border Style="{StaticResource Card}">
                            <StackPanel>
                                <TextBlock Style="{StaticResource SectionTitle}" Text="1. Find an existing Win32 app (update only)"/>
                                <DockPanel Margin="0,0,0,8">
                                    <Button x:Name="FindAppButton" DockPanel.Dock="Right" Content="Search" Margin="8,0,0,0"/>
                                    <TextBox x:Name="FindAppBox" Margin="0" ToolTip="Part of the display name"/>
                                </DockPanel>
                                <DataGrid x:Name="AppsGrid" Height="160" IsReadOnly="True" SelectionMode="Single">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Display name" Binding="{Binding displayName}" Width="2*"/>
                                        <DataGridTextColumn Header="Publisher" Binding="{Binding publisher}" Width="*"/>
                                        <DataGridTextColumn Header="Last modified" Binding="{Binding LastModified}" Width="130"/>
                                        <DataGridTextColumn Header="App id" Binding="{Binding id}" Width="2*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <TextBlock x:Name="SelectedAppDetailsText" Text="Select an app to read its version and committed content version." Foreground="#3D5A73" TextWrapping="Wrap" Margin="0,0,0,8"/>
                                <StackPanel Orientation="Horizontal">
                                    <Button x:Name="UseSelectedAppButton" Content="Use this app for update"/>
                                    <Button x:Name="ImportFromIntuneButton" Content="Import definition from Intune" ToolTip="Fill the form from the selected app (read-only); the package source stays empty"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource Card}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.ColumnSpan="2" Style="{StaticResource SectionTitle}" Text="2. Publication"/>
                                <StackPanel Grid.Row="1" Grid.ColumnSpan="2">
                                    <Border x:Name="UpdateModeBanner" Background="#FFF4CE" BorderBrush="#E0B000" BorderThickness="1" CornerRadius="6" Padding="12,8" Margin="0,0,0,10" Visibility="Collapsed">
                                        <DockPanel>
                                            <Button x:Name="SwitchToCreateButton" DockPanel.Dock="Right" Content="Switch to create mode" Margin="12,0,0,0" VerticalAlignment="Center"/>
                                            <TextBlock x:Name="UpdateModeText" TextWrapping="Wrap" Foreground="#5C4300" FontWeight="SemiBold" VerticalAlignment="Center" Text="Update mode"/>
                                        </DockPanel>
                                    </Border>
                                    <Border x:Name="PreservedPanel" Background="#EEF4FA" BorderBrush="#B9C7D8" BorderThickness="1" CornerRadius="6" Padding="12,8" Margin="0,0,0,10" Visibility="Collapsed">
                                        <StackPanel>
                                            <TextBlock Text="Preserved in Intune (not edited by this tool, never sent)" FontWeight="SemiBold" Foreground="#18324A"/>
                                            <TextBlock x:Name="PreservedText" TextWrapping="Wrap" Foreground="#3D5A73" Margin="0,4,0,0"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                                <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Text="Mode"/>
                                <ComboBox x:Name="PublishModeCombo" Grid.Row="2" Grid.Column="1" Width="320" HorizontalAlignment="Left"/>
                                <TextBlock Grid.Row="3" Style="{StaticResource FieldLabel}" Text="Existing app id"/>
                                <TextBox x:Name="ExistingAppIdBox" Grid.Row="3" Grid.Column="1" FontFamily="Consolas"/>
                                <StackPanel Grid.Row="4" Grid.Column="1" Orientation="Horizontal">
                                    <CheckBox x:Name="DryRunCheck" Content="Dry run (build and show Graph payloads, no Graph call)"/>
                                </StackPanel>
                                <StackPanel Grid.Row="5" Grid.Column="1" Orientation="Horizontal" Margin="0,6,0,0">
                                    <Button x:Name="PublishButton" Content="Publish package" Style="{StaticResource PrimaryButton}"/>
                                    <Button x:Name="BuildAndPublishButton" Content="Build and publish"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                        <Border x:Name="ResultPanel" Style="{StaticResource Card}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.ColumnSpan="3" Style="{StaticResource SectionTitle}" Text="3. Last result"/>
                                <TextBlock Grid.Row="1" Style="{StaticResource FieldLabel}" Text="App id"/>
                                <TextBox x:Name="ResultAppIdBox" Grid.Row="1" Grid.Column="1" IsReadOnly="True" FontFamily="Consolas"/>
                                <StackPanel Grid.Row="1" Grid.Column="2" Orientation="Horizontal" Margin="8,0,0,8">
                                    <Button x:Name="CopyAppIdButton" Content="Copy app id"/>
                                    <Button x:Name="OpenPortalButton" Content="Open in Intune"/>
                                    <Button x:Name="UseResultForUpdateButton" Content="Use for next update" Margin="0"/>
                                </StackPanel>
                                <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Text="Intune admin center"/>
                                <TextBlock x:Name="PortalLinkText" Grid.Row="2" Grid.Column="1" Margin="0,5,0,8" TextWrapping="Wrap"><Hyperlink x:Name="PortalLink"><Run x:Name="PortalLinkRun" Text="No app published in this session."/></Hyperlink></TextBlock>
                                <TextBlock x:Name="NotAssignedText" Grid.Row="3" Grid.Column="1" Foreground="#8A5A00" TextWrapping="Wrap" Text="Apps published by Smart IntunePackager are not assigned. Create assignments in the Intune admin center."/>
                            </Grid>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
    </DockPanel>
</Window>
'@

#endregion

#region GUI state and helpers

$script:Ui = @{}
$script:State = @{
    Config          = $null
    ProfilePath     = ''
    IntuneWinPath   = ''
    Connected       = $false
    ConnectedTenant = ''
    LastAppId       = ''
    LastResult      = $null
    UpdateTarget    = $null
    ImportedFrom    = $null
    Preserved       = @()
    AppDetails      = @{}
    Job             = $null
    Busy            = $false
    Worker          = $null
    Splash          = $null
}
$script:LogQueue = [Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:IpSync = [hashtable]::Synchronized(@{ Queue = $script:LogQueue; Cancel = $false; Sent = [int64]0; Total = [int64]0 })
$script:ActionButtons = @(
    'ConnectButton', 'DisconnectButton', 'NewProfileButton', 'OpenProfileButton', 'SaveProfileButton', 'SaveProfileAsButton', 'ValidateButton',
    'DetectMetadataButton', 'BuildButton', 'PublishButton', 'BuildAndPublishButton', 'FindAppButton', 'UseSelectedAppButton',
    'SwitchToCreateButton', 'UseResultForUpdateButton', 'ImportFromIntuneButton'
)
$script:ReturnCodeTypes = @('success', 'softReboot', 'hardReboot', 'retry', 'failed')
$script:Operators = @('notConfigured', 'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual')

function Show-IpMessage {
    param([string]$Text, [string]$Title = 'Smart IntunePackager', [string]$Icon = 'Information', [string]$Buttons = 'OK')
    $owner = $script:Ui['Window']
    if ($owner) { return [Windows.MessageBox]::Show($owner, $Text, $Title, $Buttons, $Icon) }
    return [Windows.MessageBox]::Show($Text, $Title, $Buttons, $Icon)
}

function Get-IpBaseDirectory {
    if ($script:State.ProfilePath) { return (Split-Path -Path $script:State.ProfilePath -Parent) }
    return $script:AppRoot
}

function Set-IpStatus {
    param([string]$Text)
    $script:Ui.StatusText.Text = $Text
}

function Update-IpLogPanel {
    $line = $null
    $builder = [Text.StringBuilder]::new()
    while ($script:LogQueue.TryDequeue([ref]$line)) { [void]$builder.AppendLine($line) }
    if ($builder.Length -gt 0) {
        $script:Ui.LogTextBox.AppendText($builder.ToString())
        $script:Ui.LogTextBox.ScrollToEnd()
    }
}

function Set-IpBusy {
    param([bool]$Busy, [string]$Title = '')
    $script:State.Busy = $Busy
    foreach ($name in $script:ActionButtons) { $script:Ui[$name].IsEnabled = -not $Busy }
    if (-not $Busy) { Update-IpConnectionStatus }
    $script:Ui.CancelOperationButton.IsEnabled = $Busy
    $script:Ui.OperationProgress.IsIndeterminate = $Busy
    if (-not $Busy) { $script:Ui.OperationProgress.Value = 0 }
    Set-IpStatus $(if ($Busy) { $Title + '...' } else { 'Ready' })
}

function Update-IpConnectionStatus {
    $connected = [bool]$script:State.Connected
    $statusText = 'Not connected'
    if ($connected) { $statusText = 'Connected: ' + $script:State.ConnectedTenant }
    elseif ($script:State.Config -and $script:Ui.TenantCombo.SelectedItem) {
        $profileState = Get-IpTenantProfileState -Config $script:State.Config -Name ([string]$script:Ui.TenantCombo.SelectedItem)
        if ($profileState.State -ne 'Configured') { $statusText = 'Not connected - profile not configured' }
    }
    $script:Ui.ConnectionStatusText.Text = $statusText
    if (-not $script:State.Busy) {
        $script:Ui.ConnectButton.IsEnabled = -not $connected
        $script:Ui.DisconnectButton.IsEnabled = $connected
        $script:Ui.TenantCombo.IsEnabled = -not $connected
    }
}

function New-IpDataTable {
    param([string[]]$Columns, [string[]]$BooleanColumns = @(), [string[]]$IntegerColumns = @())
    $table = [Data.DataTable]::new()
    foreach ($column in $Columns) {
        $type = if ($column -in $BooleanColumns) { [bool] } elseif ($column -in $IntegerColumns) { [int] } else { [string] }
        $dataColumn = $table.Columns.Add($column, $type)
        if ($type -eq [bool]) { $dataColumn.DefaultValue = $false }
    }
    return , $table
}

function ConvertFrom-IpDataTable {
    param([Data.DataTable]$Table)
    return @(foreach ($row in $Table.Rows) {
            if ($row.RowState -eq [Data.DataRowState]::Deleted) { continue }
            $item = [ordered]@{}
            foreach ($column in $Table.Columns) {
                $value = $row[$column.ColumnName]
                $item[$column.ColumnName] = if ($value -is [DBNull]) { $(if ($column.DataType -eq [bool]) { $false } else { '' }) } else { $value }
            }
            [pscustomobject]$item
        })
}

function Set-IpDataTableRows {
    param([Data.DataTable]$Table, [object[]]$Rows)
    $Table.Rows.Clear()
    foreach ($source in @($Rows)) {
        if ($null -eq $source) { continue }
        $row = $Table.NewRow()
        foreach ($column in $Table.Columns) {
            $property = $source.PSObject.Properties[$column.ColumnName]
            if ($null -ne $property -and $null -ne $property.Value -and [string]$property.Value -ne '') { $row[$column.ColumnName] = $property.Value }
        }
        $Table.Rows.Add($row)
    }
}

function Set-IpComboValue {
    param($Combo, [string]$Value)
    if ($Value -and $Combo.Items.Contains($Value)) { $Combo.SelectedItem = $Value } elseif ($Combo.Items.Count) { $Combo.SelectedIndex = 0 }
}

function Update-IpDetectionPanels {
    $type = [string]$script:Ui.DetectionTypeCombo.SelectedItem
    foreach ($pair in @(@('MsiProductCode', 'MsiPanel'), @('File', 'FilePanel'), @('Registry', 'RegistryPanel'), @('Script', 'ScriptPanel'), @('Preserved', 'PreservedDetectionPanel'))) {
        $script:Ui[$pair[1]].Visibility = if ($type -eq $pair[0]) { 'Visible' } else { 'Collapsed' }
    }
}

function Update-IpArchitectureChecks {
    $u = $script:Ui
    $free = -not [bool]$u.ArchNotRestrictedCheck.IsChecked
    foreach ($check in @($u.ArchX64Check, $u.ArchArm64Check, $u.ArchX86Check)) {
        $check.IsEnabled = $free
        if (-not $free) { $check.IsChecked = $false }
    }
}

function Update-IpPreservedPanel {
    $u = $script:Ui
    $items = @($script:State.Preserved | Where-Object { $_ })
    $show = $items.Count -gt 0 -and [string]$u.PublishModeCombo.SelectedItem -eq 'UpdateExisting'
    $u.PreservedPanel.Visibility = if ($show) { 'Visible' } else { 'Collapsed' }
    $u.PreservedText.Text = (@($items | ForEach-Object { '- ' + $_ }) -join "`n")
}

function Update-IpLogoPreview {
    $script:Ui.AppIconPreview.Source = $null
    $path = Resolve-IpPath -Path $script:Ui.LogoPathBox.Text -BaseDirectory (Get-IpBaseDirectory)
    if ($path -and (Test-Path -LiteralPath $path -PathType Leaf) -and [IO.Path]::GetExtension($path).ToLowerInvariant() -in @('.png', '.jpg', '.jpeg')) {
        try {
            $image = [Windows.Media.Imaging.BitmapImage]::new()
            $image.BeginInit()
            $image.CacheOption = [Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $image.UriSource = [Uri]$path
            $image.EndInit()
            $script:Ui.AppIconPreview.Source = $image
        }
        catch { Write-IpLog ("Logo preview unavailable: {0}" -f $_.Exception.Message) -Level WARN }
    }
}

function Set-IpFormDefinition {
    param([psobject]$Definition)
    $u = $script:Ui
    $u.SourceFolderBox.Text = [string]$Definition.Package.SourceFolder
    $u.SetupFileBox.Text = [string]$Definition.Package.SetupFile
    $u.OutputFolderBox.Text = [string]$Definition.Package.OutputFolder
    $u.DisplayNameBox.Text = [string]$Definition.App.DisplayName
    $u.DescriptionBox.Text = [string]$Definition.App.Description
    $u.PublisherBox.Text = [string]$Definition.App.Publisher
    $u.VersionBox.Text = [string]$Definition.App.Version
    $u.DeveloperBox.Text = [string]$Definition.App.Developer
    $u.OwnerBox.Text = [string]$Definition.App.Owner
    $u.InformationUrlBox.Text = [string]$Definition.App.InformationUrl
    $u.PrivacyUrlBox.Text = [string]$Definition.App.PrivacyInformationUrl
    $u.NotesBox.Text = [string]$Definition.App.Notes
    $u.LogoPathBox.Text = [string]$Definition.App.LogoPath
    $u.InstallCommandBox.Text = [string]$Definition.Program.InstallCommandLine
    $u.UninstallCommandBox.Text = [string]$Definition.Program.UninstallCommandLine
    Set-IpComboValue $u.RunAsCombo ([string]$Definition.Program.RunAsAccount)
    Set-IpComboValue $u.RestartCombo ([string]$Definition.Program.DeviceRestartBehavior)
    $u.MaxRunTimeBox.Text = [string]$Definition.Program.MaxRunTimeInMinutes
    Set-IpDataTableRows $script:ReturnCodesTable @($Definition.ReturnCodes)
    $architectures = @($Definition.Requirements.Architectures)
    $u.ArchX64Check.IsChecked = 'x64' -in $architectures
    $u.ArchArm64Check.IsChecked = 'arm64' -in $architectures
    $u.ArchX86Check.IsChecked = 'x86' -in $architectures
    $u.ArchNotRestrictedCheck.IsChecked = 'NotRestricted' -in $architectures
    Update-IpArchitectureChecks
    $u.MinDiskSpaceBox.Text = [string](Get-IpObjectProperty $Definition.Requirements 'MinimumFreeDiskSpaceInMB')
    $imported = Get-IpObjectProperty $Definition 'ImportedFrom'
    $script:State.ImportedFrom = if ($imported -and [string](Get-IpObjectProperty $imported 'AppId')) { $imported } else { $null }
    $script:State.Preserved = @(Get-IpObjectProperty $Definition 'Preserved' | Where-Object { $_ })
    Set-IpComboValue $u.MinWindowsCombo ([string]$Definition.Requirements.MinimumWindowsRelease)
    Set-IpComboValue $u.DetectionTypeCombo ([string]$Definition.Detection.Type)
    $u.MsiProductCodeBox.Text = [string]$Definition.Detection.Msi.ProductCode
    Set-IpComboValue $u.MsiVersionOperatorCombo ([string]$Definition.Detection.Msi.ProductVersionOperator)
    $u.MsiVersionBox.Text = [string]$Definition.Detection.Msi.ProductVersion
    Set-IpDataTableRows $script:FileRulesTable @($Definition.Detection.File | ForEach-Object { ConvertTo-IpDetectionItem -Item $_ -Kind File })
    Set-IpDataTableRows $script:RegistryRulesTable @($Definition.Detection.Registry | ForEach-Object { ConvertTo-IpDetectionItem -Item $_ -Kind Registry })
    $u.ScriptPathBox.Text = [string]$Definition.Detection.Script.ScriptPath
    $u.EnforceSignatureCheck.IsChecked = [bool]$Definition.Detection.Script.EnforceSignatureCheck
    $u.RunAs32BitCheck.IsChecked = [bool]$Definition.Detection.Script.RunAs32Bit
    Set-IpComboValue $u.PublishModeCombo ([string]$Definition.Publish.Mode)
    $u.ExistingAppIdBox.Text = [string]$Definition.Publish.ExistingAppId
    Update-IpDetectionPanels
    Update-IpLogoPreview
    Update-IpPreservedPanel
}

function Get-IpFormDefinition {
    $u = $script:Ui
    foreach ($grid in @($u.ReturnCodesGrid, $u.FileRulesGrid, $u.RegistryRulesGrid)) { [void]$grid.CommitEdit('Row', $true) }
    $definition = New-IpAppDefinition
    $definition.Package.SourceFolder = $u.SourceFolderBox.Text.Trim()
    $definition.Package.SetupFile = $u.SetupFileBox.Text.Trim()
    $definition.Package.OutputFolder = $u.OutputFolderBox.Text.Trim()
    $definition.App.DisplayName = $u.DisplayNameBox.Text.Trim()
    $definition.App.Description = $u.DescriptionBox.Text.Trim()
    $definition.App.Publisher = $u.PublisherBox.Text.Trim()
    $definition.App.Version = $u.VersionBox.Text.Trim()
    $definition.App.Developer = $u.DeveloperBox.Text.Trim()
    $definition.App.Owner = $u.OwnerBox.Text.Trim()
    $definition.App.InformationUrl = $u.InformationUrlBox.Text.Trim()
    $definition.App.PrivacyInformationUrl = $u.PrivacyUrlBox.Text.Trim()
    $definition.App.Notes = $u.NotesBox.Text.Trim()
    $definition.App.LogoPath = $u.LogoPathBox.Text.Trim()
    $definition.Program.InstallCommandLine = $u.InstallCommandBox.Text.Trim()
    $definition.Program.UninstallCommandLine = $u.UninstallCommandBox.Text.Trim()
    $definition.Program.RunAsAccount = [string]$u.RunAsCombo.SelectedItem
    $definition.Program.DeviceRestartBehavior = [string]$u.RestartCombo.SelectedItem
    $maxRun = 0
    $definition.Program.MaxRunTimeInMinutes = if ([int]::TryParse($u.MaxRunTimeBox.Text.Trim(), [ref]$maxRun)) { $maxRun } else { $u.MaxRunTimeBox.Text.Trim() }
    $definition.ReturnCodes = @(ConvertFrom-IpDataTable $script:ReturnCodesTable)
    $definition.Requirements.Architectures = if ($u.ArchNotRestrictedCheck.IsChecked) { @('NotRestricted') } else {
        @(
            if ($u.ArchX64Check.IsChecked) { 'x64' }
            if ($u.ArchArm64Check.IsChecked) { 'arm64' }
            if ($u.ArchX86Check.IsChecked) { 'x86' }
        )
    }
    $diskText = $u.MinDiskSpaceBox.Text.Trim()
    $disk = 0
    $definition.Requirements.MinimumFreeDiskSpaceInMB = if (-not $diskText) { $null } elseif ([int]::TryParse($diskText, [ref]$disk)) { $disk } else { $diskText }
    $definition.Requirements.MinimumWindowsRelease = [string]$u.MinWindowsCombo.SelectedItem
    $definition.Detection.Type = [string]$u.DetectionTypeCombo.SelectedItem
    $definition.Detection.Msi.ProductCode = $u.MsiProductCodeBox.Text.Trim()
    $definition.Detection.Msi.ProductVersionOperator = [string]$u.MsiVersionOperatorCombo.SelectedItem
    $definition.Detection.Msi.ProductVersion = $u.MsiVersionBox.Text.Trim()
    $definition.Detection.File = @(ConvertFrom-IpDataTable $script:FileRulesTable)
    $definition.Detection.Registry = @(ConvertFrom-IpDataTable $script:RegistryRulesTable)
    $definition.Detection.Script.ScriptPath = $u.ScriptPathBox.Text.Trim()
    $definition.Detection.Script.EnforceSignatureCheck = [bool]$u.EnforceSignatureCheck.IsChecked
    $definition.Detection.Script.RunAs32Bit = [bool]$u.RunAs32BitCheck.IsChecked
    $definition.Publish.Mode = [string]$u.PublishModeCombo.SelectedItem
    $definition.Publish.ExistingAppId = $u.ExistingAppIdBox.Text.Trim()
    if ($script:State.ImportedFrom) { $definition.ImportedFrom = $script:State.ImportedFrom }
    $definition.Preserved = @($script:State.Preserved)
    return $definition
}

function Test-IpFormDefinition {
    param([psobject]$Definition, [switch]$ForPublish, [switch]$Quiet)
    $errors = @(Test-IpAppDefinition -Definition $Definition -BaseDirectory (Get-IpBaseDirectory) -ForPublish:$ForPublish)
    if ($errors.Count -eq 0) {
        if (-not $Quiet) { Write-IpLog 'The app definition is valid.' }
        return $true
    }
    foreach ($validationError in $errors) { Write-IpLog $validationError -Level WARN }
    [void](Show-IpMessage -Text ("The app definition has {0} issue(s):`n`n{1}" -f $errors.Count, ($errors -join "`n")) -Icon Warning)
    return $false
}

#endregion

#region Dialogs

function Select-IpFolder {
    param([string]$Title, [string]$Initial)
    $dialog = [Microsoft.Win32.OpenFolderDialog]::new()
    $dialog.Title = $Title
    if ($Initial -and (Test-Path -LiteralPath $Initial -PathType Container)) { $dialog.InitialDirectory = $Initial }
    if ($dialog.ShowDialog($script:Ui.Window)) { return $dialog.FolderName }
    return ''
}

function Select-IpFile {
    param([string]$Title, [string]$Filter, [string]$Initial)
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    if ($Initial -and (Test-Path -LiteralPath $Initial -PathType Container)) { $dialog.InitialDirectory = $Initial }
    if ($dialog.ShowDialog($script:Ui.Window)) { return $dialog.FileName }
    return ''
}

function Select-IpSaveFile {
    param([string]$Title, [string]$Filter, [string]$FileName, [string]$Initial)
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.FileName = $FileName
    if ($Initial -and (Test-Path -LiteralPath $Initial -PathType Container)) { $dialog.InitialDirectory = $Initial }
    if ($dialog.ShowDialog($script:Ui.Window)) { return $dialog.FileName }
    return ''
}

#endregion

#region Background worker

function Initialize-IpWorker {
    <#
    .SYNOPSIS
        Opens the background runspace that owns packaging, the Graph session and publication.
    #>
    $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = [Threading.ApartmentState]::STA
    $runspace.ThreadOptions = [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('IpSync', $script:IpSync)
    $initialize = [PowerShell]::Create()
    $initialize.Runspace = $runspace
    [void]$initialize.AddScript({
            param($AppRoot, $ModuleNames, $LogPath, $GraphSettings)
            Set-StrictMode -Version Latest
            $ErrorActionPreference = 'Stop'
            foreach ($name in $ModuleNames) { Import-Module (Join-Path $AppRoot ($name + '.psm1')) -Force }
            Set-IpLogPath -Path $LogPath
            Set-IpLogSink -Sink { param($line, $level) $IpSync.Queue.Enqueue($line) } -Console $false
            Set-IpGraphSettings -BaseUri $GraphSettings.BaseUri -MaxRetries ([int]$GraphSettings.MaxRetries) -PollSeconds ([int]$GraphSettings.PollSeconds) -PollTimeoutMinutes ([int]$GraphSettings.PollTimeoutMinutes)
        }.ToString()).AddArgument($script:AppRoot).AddArgument($script:ModuleNames).AddArgument((Get-IpLogPath)).AddArgument($script:State.Config.Graph)
    try { [void]$initialize.Invoke() }
    finally { $initialize.Dispose() }
    $script:State.Worker = $runspace
}

function Start-IpJob {
    <#
    .SYNOPSIS
        Runs a script in the background runspace; OnSuccess and OnFailure run on the UI thread.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][scriptblock]$Script,
        [object[]]$Arguments = @(),
        # Passed to OnSuccess as its second argument (handlers are not closures).
        [hashtable]$Context = @{},
        [scriptblock]$OnSuccess,
        [scriptblock]$OnFailure
    )

    if ($script:State.Job) { Write-IpLog 'Another operation is running.' -Level WARN; return }
    $powerShell = [PowerShell]::Create()
    $powerShell.Runspace = $script:State.Worker
    [void]$powerShell.AddScript($Script.ToString())
    foreach ($argument in $Arguments) { [void]$powerShell.AddArgument($argument) }
    $script:IpSync.Cancel = $false
    $script:IpSync.Sent = [int64]0
    $script:IpSync.Total = [int64]0
    $script:State.Job = @{ PowerShell = $powerShell; Handle = $powerShell.BeginInvoke(); Title = $Title; Context = $Context; OnSuccess = $OnSuccess; OnFailure = $OnFailure }
    Set-IpBusy $true $Title
}

function Complete-IpJob {
    $job = $script:State.Job
    if (-not $job -or -not $job.Handle.IsCompleted) { return }
    $script:State.Job = $null
    $output = $null
    $failure = $null
    try {
        $output = $job.PowerShell.EndInvoke($job.Handle)
        if ($job.PowerShell.HadErrors -and $job.PowerShell.Streams.Error.Count -gt 0) { $failure = $job.PowerShell.Streams.Error[0] }
    }
    catch { $failure = $_ }
    finally { $job.PowerShell.Dispose() }
    Update-IpLogPanel
    Set-IpBusy $false

    if ($failure) {
        # EndInvoke wraps the worker error in MethodInvocationException: report the original one.
        $inner = Get-IpInnermostException $failure
        $message = $inner.Message
        Write-IpLog ("{0} failed: {1}" -f $job.Title, $message) -Level ERROR
        $record = if ($inner -is [Management.Automation.IContainsErrorRecord]) { $inner.ErrorRecord } else { $failure }
        Write-IpLog ("Error location: {0}" -f $record.ScriptStackTrace) -Level ERROR -FileOnly
        Set-IpStatus ($job.Title + ' failed')
        if ($job.OnFailure) { Invoke-IpGuarded -Context ($job.Title + ' result') -ScriptBlock $job.OnFailure -Arguments @($message, $job.Context) }
        else { [void](Show-IpMessage -Text ("{0} failed:`n`n{1}" -f $job.Title, $message) -Icon Error) }
        return
    }
    $result = Get-IpJobResult -Output $output
    Set-IpStatus ($job.Title + ' completed')
    if ($job.OnSuccess) { Invoke-IpGuarded -Context ($job.Title + ' result') -ScriptBlock $job.OnSuccess -Arguments @($result, $job.Context) }
}

function Get-IpInnermostException {
    <#
    .SYNOPSIS
        Unwraps MethodInvocationException (EndInvoke) and ActionPreferenceStopException (-ErrorAction Stop)
        to the exception that describes the actual failure.
    #>
    param([AllowNull()]$ErrorObject)
    $exception = if ($ErrorObject -is [Management.Automation.ErrorRecord]) { $ErrorObject.Exception } else { $ErrorObject }
    for ($depth = 0; $depth -lt 10 -and $null -ne $exception; $depth++) {
        if ($exception -is [Management.Automation.MethodInvocationException] -and $exception.InnerException) { $exception = $exception.InnerException; continue }
        if ($exception -is [Management.Automation.ActionPreferenceStopException] -and $exception.ErrorRecord -and $exception.ErrorRecord.Exception -and -not [object]::ReferenceEquals($exception.ErrorRecord.Exception, $exception)) { $exception = $exception.ErrorRecord.Exception; continue }
        break
    }
    return $exception
}

function Get-IpJobResult {
    <#
    .SYNOPSIS
        Returns the last object emitted by a worker job, or $null when it emitted nothing.

    .DESCRIPTION
        The output collection is tested on its Count, never on its truthiness: a collection
        holding one empty array evaluates to $false in PowerShell, which turned an empty
        search result into $null and then into a $null DataGrid row (ArgumentNullException).
    #>
    param([AllowNull()]$Output)
    if ($null -eq $Output -or $Output.Count -eq 0) { return $null }
    return , $Output[$Output.Count - 1]
}

function ConvertTo-IpGridItems {
    <#
    .SYNOPSIS
        Returns an array without $null entries for an ItemsSource (WPF fails on $null rows).
    #>
    param([AllowNull()]$Items)
    return , [object[]]@(@($Items) | Where-Object { $null -ne $_ })
}

function Invoke-IpGuarded {
    <#
    .SYNOPSIS
        Runs a UI callback; any exception is written to the log panel and file instead of closing the window.
    #>
    param([Parameter(Mandatory = $true)][string]$Context, [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock, [object[]]$Arguments = @())
    try { & $ScriptBlock @Arguments }
    catch {
        Write-IpLog ("{0} failed: {1}" -f $Context, $_.Exception.Message) -Level ERROR
        Write-IpLog ("Error location: {0}" -f $_.ScriptStackTrace) -Level ERROR -FileOnly
        try { Set-IpStatus ($Context + ' failed') } catch { Write-Verbose 'Status bar unavailable.' }
    }
}

function Register-IpDispatcherExceptionHandler {
    <#
    .SYNOPSIS
        Last-resort handler for exceptions raised on the UI thread (event handlers, timers, layout):
        logs them and keeps the window open. Identical messages are logged once per 10 seconds.
    #>
    param([Parameter(Mandatory = $true)][Windows.Threading.Dispatcher]$Dispatcher)
    $script:IpLastUiError = @{ Message = ''; Time = [datetime]::MinValue; Suppressed = 0 }
    $Dispatcher.Add_UnhandledException({
            param($source, $eventArgs)
            $eventArgs.Handled = $true
            try {
                $exception = $eventArgs.Exception
                while (($exception -is [Management.Automation.MethodInvocationException] -or $exception -is [Reflection.TargetInvocationException]) -and $exception.InnerException) { $exception = $exception.InnerException }
                $message = '{0}: {1}' -f $exception.GetType().Name, $exception.Message
                $last = $script:IpLastUiError
                $now = Get-Date
                if ($message -eq $last.Message -and ($now - $last.Time).TotalSeconds -lt 10) { $last.Suppressed++; return }
                if ($last.Suppressed -gt 0) { Write-IpLog ("Previous UI error repeated {0} more time(s)." -f $last.Suppressed) -Level WARN }
                $last.Message = $message; $last.Time = $now; $last.Suppressed = 0
                Write-IpLog ("Unexpected UI error, the window stays open: {0}" -f $message) -Level ERROR
                Write-IpLog ("UI error detail: {0}" -f $eventArgs.Exception.ToString()) -Level ERROR -FileOnly
            }
            catch { Write-Verbose 'The UI error could not be logged.' }
        })
}

function Invoke-IpTimerTick {
    Update-IpLogPanel
    if ($script:State.Job) {
        $total = [int64]$script:IpSync.Total
        if ($total -gt 0) {
            $script:Ui.OperationProgress.IsIndeterminate = $false
            $script:Ui.OperationProgress.Value = [Math]::Min(100, 100.0 * [int64]$script:IpSync.Sent / $total)
            Set-IpStatus ("{0}: {1} / {2}" -f $script:State.Job.Title, (Format-IpByteSize ([int64]$script:IpSync.Sent)), (Format-IpByteSize $total))
        }
        Complete-IpJob
    }
}

#endregion

#region Operations

function Get-IpTenantProfileState {
    <#
    .SYNOPSIS
        Returns Configured, NotConfigured (empty TenantId) or Invalid for a tenant profile, with a user message.
    #>
    param([psobject]$Config, [string]$Name)
    if (-not $Name) { return [pscustomobject]@{ State = 'Invalid'; Message = 'No tenant profile is selected.' } }
    try { $tenantProfile = Get-IpTenantProfile -Config $Config -Tenant $Name }
    catch { return [pscustomobject]@{ State = 'Invalid'; Message = $_.Exception.Message } }
    if (-not $tenantProfile.TenantId) {
        return [pscustomobject]@{ State = 'NotConfigured'; Message = ("Tenant profile '{0}' is not configured: its TenantId is empty." -f $Name) }
    }
    return [pscustomobject]@{ State = 'Configured'; Message = ''; Profile = $tenantProfile }
}

function Resolve-IpInitialTenant {
    <#
    .SYNOPSIS
        Selects the tenant shown at startup: -Tenant when passed explicitly, else the last used profile, else -Tenant's default.
    #>
    param([string[]]$Names = @(), [string]$Requested, [bool]$RequestedExplicitly, [string]$LastUsed)
    if ($RequestedExplicitly -and $Requested -in $Names) { return $Requested }
    if ($LastUsed -and $LastUsed -in $Names) { return $LastUsed }
    if ($Requested -in $Names) { return $Requested }
    if ($Names.Count) { return $Names[0] }
    return ''
}

function Save-IpLastTenant {
    param([string]$Name)
    try {
        $config = $script:State.Config
        if (-not $config.PSObject.Properties['Gui']) { $config | Add-Member -NotePropertyName Gui -NotePropertyValue ([pscustomobject]@{ LastTenant = '' }) }
        if ([string]$config.Gui.LastTenant -eq $Name) { return }
        $config.Gui.LastTenant = $Name
        Save-IpJsonFile -Path $script:ConfigPath -InputObject $config -Confirm:$false -WhatIf:$false
    }
    catch { Write-IpLog ("The last used tenant could not be saved: {0}" -f $_.Exception.Message) -Level WARN }
}

function Get-IpSignInBannerText {
    param([string]$SignInMode)
    switch ($SignInMode) {
        'Wam' { return "Sign-in window opened (Windows account broker). If you don't see it, check the taskbar or press Alt+Tab." }
        'Browser' { return 'Sign-in continues in your web browser. Complete it there, then return to this window.' }
        default { return 'Connecting to Microsoft Graph...' }
    }
}

function Invoke-IpFlashWindow {
    <#
    .SYNOPSIS
        Flashes the window in the taskbar until it is activated (FlashWindowEx), or stops flashing.
    #>
    param([Parameter(Mandatory = $true)][Windows.Window]$Window, [switch]$Stop)
    try {
        if (-not ('SmartM365.IntunePackager.NativeWindow' -as [type])) {
            $definition = @(
                '[StructLayout(LayoutKind.Sequential)]'
                'public struct FLASHWINFO { public uint cbSize; public System.IntPtr hwnd; public uint dwFlags; public uint uCount; public uint dwTimeout; }'
                '[DllImport("user32.dll")]'
                'public static extern bool FlashWindowEx(ref FLASHWINFO pwfi);'
            ) -join [Environment]::NewLine
            Add-Type -Namespace SmartM365.IntunePackager -Name NativeWindow -MemberDefinition $definition
        }
        $info = [SmartM365.IntunePackager.NativeWindow+FLASHWINFO]::new()
        $info.cbSize = [uint32][Runtime.InteropServices.Marshal]::SizeOf($info)
        $info.hwnd = [Windows.Interop.WindowInteropHelper]::new($Window).Handle
        # FLASHW_ALL (3) + FLASHW_TIMERNOFG (12): flash until the window comes to the foreground. FLASHW_STOP = 0.
        $info.dwFlags = if ($Stop) { [uint32]0 } else { [uint32]15 }
        [void][SmartM365.IntunePackager.NativeWindow]::FlashWindowEx([ref]$info)
    }
    catch { Write-Verbose ('Taskbar flashing unavailable: ' + $_.Exception.Message) }
}

function Set-IpSignInBanner {
    param([string]$Text)
    $u = $script:Ui
    if ($Text) {
        $u.SignInBannerText.Text = $Text
        $u.SignInBanner.Visibility = 'Visible'
        Invoke-IpFlashWindow -Window $u.Window
    }
    else {
        $u.SignInBanner.Visibility = 'Collapsed'
        Invoke-IpFlashWindow -Window $u.Window -Stop
        [void]$u.Window.Activate()
    }
}

function Invoke-IpConnect {
    $name = [string]$script:Ui.TenantCombo.SelectedItem
    $profileState = Get-IpTenantProfileState -Config $script:State.Config -Name $name
    if ($profileState.State -ne 'Configured') {
        $hint = if ($profileState.State -eq 'NotConfigured') { "`n`nSet Tenants.{0}.TenantId (tenant GUID or tenant.onmicrosoft.com) in:`n{1}`n`nThen restart Smart IntunePackager, or select another tenant profile." -f $name, $script:ConfigPath } else { "`n`nCheck the Tenants section of:`n{0}" -f $script:ConfigPath }
        Write-IpLog ($profileState.Message + ' Connection not attempted.') -Level WARN
        [void](Show-IpMessage -Text ($profileState.Message + $hint) -Icon Warning)
        return
    }
    $signIn = Resolve-IpSignInMode -TenantProfile $profileState.Profile
    if ($signIn -eq 'DeviceCode') {
        $message = "Tenant profile '{0}' uses DelegatedSignIn=DeviceCode. The device code is written to the PowerShell console, which is hidden in the GUI: use Browser (recommended) or Wam for the GUI, and DeviceCode with the CLI." -f $name
        Write-IpLog $message -Level WARN
        [void](Show-IpMessage -Text $message -Icon Warning)
        return
    }
    # Microsoft.Graph.Authentication gives WAM the console window as parent (no parameter accepts the WPF
    # window handle) and the console is hidden in the GUI, so the sign-in window can open behind this one.
    Set-IpSignInBanner -Text (Get-IpSignInBannerText -SignInMode $signIn)
    Start-IpJob -Title "Connect to $name" -Arguments @(, $profileState.Profile) -Context @{ Name = $name } -Script {
        param($TenantProfile)
        Connect-IpGraph -TenantProfile $TenantProfile
    } -OnSuccess {
        param($result, $context)
        Set-IpSignInBanner -Text ''
        $script:State.Connected = $true
        $script:State.ConnectedTenant = '{0} ({1})' -f $context.Name, $result.TenantId
        Update-IpConnectionStatus
        Save-IpLastTenant -Name $context.Name
    } -OnFailure {
        param($message, $context)
        Set-IpSignInBanner -Text ''
        [void](Show-IpMessage -Text ("Connect to {0} failed:`n`n{1}" -f $context.Name, $message) -Icon Error)
    }
}


function Invoke-IpDisconnect {
    Start-IpJob -Title 'Disconnect' -Script {
        Disconnect-IpGraph
        Write-IpLog 'Disconnected from Microsoft Graph.'
    } -OnSuccess {
        param($result)
        $script:State.Connected = $false
        $script:State.ConnectedTenant = ''
        Update-IpConnectionStatus
    }
}

function Invoke-IpDetectMetadata {
    param([switch]$Quiet)
    $u = $script:Ui
    $source = Resolve-IpPath -Path $u.SourceFolderBox.Text.Trim() -BaseDirectory (Get-IpBaseDirectory)
    if (-not $source -or -not (Test-Path -LiteralPath $source -PathType Container)) {
        if (-not $Quiet) { [void](Show-IpMessage -Text 'Select an existing source folder first.' -Icon Warning) }
        return
    }

    $psadt = Get-IpPsadtCommands -SourceFolder $source
    if ($psadt) {
        if (-not $u.SetupFileBox.Text.Trim()) { $u.SetupFileBox.Text = $psadt.SetupFile }
        $u.InstallCommandBox.Text = $psadt.InstallCommandLine
        $u.UninstallCommandBox.Text = $psadt.UninstallCommandLine
        Write-IpLog ("PSAppDeployToolkit v{0} detected: install and uninstall commands prefilled." -f $psadt.ToolkitVersion)
    }

    $setup = $u.SetupFileBox.Text.Trim()
    if ($setup -and [IO.Path]::GetExtension($setup) -eq '.msi') {
        $msiPath = Join-Path $source $setup
        if (-not (Test-Path -LiteralPath $msiPath -PathType Leaf)) { Write-IpLog "MSI not found: $msiPath" -Level WARN; return }
        try { $msi = Get-IpMsiMetadata -Path $msiPath }
        catch { Write-IpLog $_.Exception.Message -Level ERROR; return }
        if (-not $u.DisplayNameBox.Text.Trim()) { $u.DisplayNameBox.Text = $msi.Name }
        if ($msi.Publisher) { $u.PublisherBox.Text = $msi.Publisher }
        if ($msi.ProductVersion) { $u.VersionBox.Text = $msi.ProductVersion }
        if (-not $psadt) {
            $u.InstallCommandBox.Text = $msi.InstallCommandLine
            $u.UninstallCommandBox.Text = $msi.UninstallCommandLine
            Set-IpComboValue $u.RunAsCombo $msi.RunAsAccount
        }
        Set-IpComboValue $u.DetectionTypeCombo 'MsiProductCode'
        $u.MsiProductCodeBox.Text = $msi.ProductCode
        if ($msi.ProductVersion) { Set-IpComboValue $u.MsiVersionOperatorCombo 'greaterThanOrEqual'; $u.MsiVersionBox.Text = $msi.ProductVersion }
        Update-IpDetectionPanels
        Write-IpLog ("MSI metadata detected: Name={0}; Version={1}; Publisher={2}; ProductCode={3}" -f $msi.Name, $msi.ProductVersion, $msi.Publisher, $msi.ProductCode)
    }
    elseif (-not $psadt -and -not $Quiet) {
        Write-IpLog 'No MSI setup file or PSAppDeployToolkit package detected; fill the commands manually.'
    }
}

function Test-IpDefinitionLoaded {
    <#
    .SYNOPSIS
        Returns $false when the form holds no app definition (no profile opened and nothing entered).
    #>
    param([Parameter(Mandatory = $true)][psobject]$Definition)
    $values = @($Definition.App.DisplayName, $Definition.Program.InstallCommandLine, $Definition.Program.UninstallCommandLine, $Definition.Package.SourceFolder, $Definition.Package.SetupFile)
    return [bool](@($values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count)
}

function Get-IpNoDefinitionMessage {
    param([string]$PublishMode)
    if ($PublishMode -eq 'UpdateExisting') { return 'No app definition loaded. Open the profile used for this app, or import the definition from Intune.' }
    return 'No app definition loaded. Open a profile, or fill in the Package, App information and Program tabs.'
}

function Assert-IpDefinitionLoaded {
    param([Parameter(Mandatory = $true)][psobject]$Definition)
    if (Test-IpDefinitionLoaded -Definition $Definition) { return $true }
    $message = Get-IpNoDefinitionMessage -PublishMode ([string]$Definition.Publish.Mode)
    Write-IpLog $message -Level WARN
    [void](Show-IpMessage -Text $message -Icon Information)
    return $false
}

function Get-IpBuildRequest {
    param([psobject]$Definition)
    $base = Get-IpBaseDirectory
    return @{
        SourceFolder = Resolve-IpPath -Path $Definition.Package.SourceFolder -BaseDirectory $base
        SetupFile    = [string]$Definition.Package.SetupFile
        OutputFolder = Get-IpProfileOutputFolder -Definition $Definition -Config $script:State.Config -BaseDirectory $base
    }
}

function Invoke-IpBuild {
    param([scriptblock]$Then)
    $definition = Get-IpFormDefinition
    if (-not (Assert-IpDefinitionLoaded -Definition $definition)) { return }
    if (-not (Test-IpFormDefinition -Definition $definition -Quiet)) { return }
    try { $request = Get-IpBuildRequest -Definition $definition }
    catch { [void](Show-IpMessage -Text $_.Exception.Message -Icon Warning); return }
    $issues = @(Test-IpPackageBuildRequest -SourceFolder $request.SourceFolder -SetupFile $request.SetupFile -OutputFolder $request.OutputFolder)
    if ($issues.Count) { [void](Show-IpMessage -Text ($issues -join "`n") -Icon Warning); return }

    Start-IpJob -Title 'Build package' -Arguments @($request, $script:State.Config, $script:AppRoot) -Context @{ Definition = $definition; Then = $Then } -Script {
        param($Request, $Config, $AppRoot)
        Invoke-IpPackageBuild -EngineName ([string]$Config.Packaging.Engine) -SourceFolder $Request.SourceFolder -SetupFile $Request.SetupFile -OutputFolder $Request.OutputFolder -Settings $Config -AppRoot $AppRoot
    } -OnSuccess {
        param($result, $context)
        $script:State.IntuneWinPath = [string]$result.IntuneWinPath
        $script:Ui.PackageInfoText.Text = "File: {0}`nSHA256: {1}`nEncrypted size: {2}`nSetup file: {3}`nEngine: {4} {5}" -f $result.IntuneWinPath, $result.Sha256, (Format-IpByteSize $result.Metadata.SizeEncrypted), $result.Metadata.SetupFile, $result.Engine, $result.EngineVersion
        Test-IpPackageMatchesDefinition -Definition $context.Definition -Metadata $result.Metadata
        if ($context.Then) { & $context.Then }
    }
}

function Get-IpNextContentVersion {
    <#
    .SYNOPSIS
        Returns the expected next content version (Intune numbers content versions sequentially).
    #>
    param([AllowEmptyString()][string]$Current)
    $number = 0
    if ([string]::IsNullOrWhiteSpace($Current)) { return '1' }
    if ([int]::TryParse($Current.Trim(), [ref]$number)) { return [string]($number + 1) }
    return 'next'
}

function Get-IpUpdateBannerText {
    param([AllowNull()][psobject]$Target, [string]$AppId)
    if ($Target -and $Target.AppId -eq $AppId) {
        $current = if ($Target.ContentVersion) { $Target.ContentVersion } else { 'none' }
        return ("Update mode: {0} ({1}), current content version {2}. Click Publish package to upload content version {3}." -f $Target.DisplayName, $Target.AppId, $current, (Get-IpNextContentVersion $Target.ContentVersion))
    }
    if ($AppId) { return ("Update mode: app {0}. Use 'Use this app for update' in the search above to show its current content version." -f $AppId) }
    return 'Update mode: search for the app above and click Use this app for update, or enter its app id.'
}

function Get-IpUpdateConfirmationText {
    param([psobject]$Details, [string]$ProfileDisplayName, [string]$Tenant, [string]$PackagePath, [object[]]$Changes = @(), [object[]]$Warnings = @(), [switch]$Forced)
    $current = if ($Details.committedContentVersion) { $Details.committedContentVersion } else { 'none' }
    $lines = @(
        'Publish a new content version of an existing Win32 app?'
        ''
        ('App: {0}' -f $Details.displayName)
        ('App id: {0}' -f $Details.id)
        ('Committed content version: {0} -> {1}' -f $current, (Get-IpNextContentVersion $Details.committedContentVersion))
        ('Tenant: {0}' -f $Tenant)
        ('Package: {0}' -f [IO.Path]::GetFileName($PackagePath))
    )
    if ($ProfileDisplayName -and $ProfileDisplayName -ne $Details.displayName) {
        $lines += @('', ("The app metadata is replaced by the profile: its display name becomes '{0}'." -f $ProfileDisplayName))
    }
    $changeList = @($Changes | Where-Object { $_ })
    if ($changeList.Count) {
        $lines += @('', ('Changes sent to Intune ({0}); everything else is kept as it is:' -f $changeList.Count))
        $lines += @($changeList | Select-Object -First 12 | ForEach-Object { '- {0}: {1} -> {2}' -f $_.Property, $_.Before, $_.After })
        if ($changeList.Count -gt 12) { $lines += ('- ... and {0} more (see the log).' -f ($changeList.Count - 12)) }
    }
    elseif ($PSBoundParameters.ContainsKey('Changes')) { $lines += @('', 'No metadata change: only the new content version is published; everything else is kept as it is.') }
    $warningList = @($Warnings | Where-Object { $_ })
    if ($warningList.Count) { $lines += @('', 'Warnings:') + @($warningList | ForEach-Object { '- ' + $_ }) }
    if ($Forced) { $lines += @('', 'Forced publication: the app was modified in Intune after the import.') }
    $lines += @('', 'The app will not be assigned.')
    return ($lines -join "`n")
}

function Set-IpPublishButtonHighlight {
    param([bool]$On)
    $button = $script:Ui.PublishButton
    if ($On) {
        $button.BorderBrush = [Windows.Media.Brushes]::Goldenrod
        $button.BorderThickness = [Windows.Thickness]::new(3)
    }
    else {
        $button.ClearValue([Windows.Controls.Control]::BorderBrushProperty)
        $button.ClearValue([Windows.Controls.Control]::BorderThicknessProperty)
    }
}

function Update-IpUpdateBanner {
    $u = $script:Ui
    $update = [string]$u.PublishModeCombo.SelectedItem -eq 'UpdateExisting'
    $u.UpdateModeBanner.Visibility = if ($update) { 'Visible' } else { 'Collapsed' }
    if ($update) { $u.UpdateModeText.Text = Get-IpUpdateBannerText -Target $script:State.UpdateTarget -AppId $u.ExistingAppIdBox.Text.Trim() }
    else { Set-IpPublishButtonHighlight $false }
    Update-IpPreservedPanel
}

function Set-IpUpdateTarget {
    <#
    .SYNOPSIS
        Switches the publication to UpdateExisting for an app, shows the banner and highlights Publish package.
    #>
    param([Parameter(Mandatory = $true)][string]$AppId, [string]$DisplayName, [AllowEmptyString()][string]$ContentVersion)
    $script:State.UpdateTarget = [pscustomobject]@{ AppId = $AppId; DisplayName = $DisplayName; ContentVersion = $ContentVersion }
    $u = $script:Ui
    $u.ExistingAppIdBox.Text = $AppId
    Set-IpComboValue $u.PublishModeCombo 'UpdateExisting'
    Update-IpUpdateBanner
    Set-IpPublishButtonHighlight $true
    $u.MainTabs.SelectedIndex = 5
    $u.PublishButton.BringIntoView()
    [void]$u.PublishButton.Focus()
    Write-IpLog ("Update mode: {0} ({1}), current content version {2}; next publication uploads content version {3}." -f $DisplayName, $AppId, $(if ($ContentVersion) { $ContentVersion } else { 'none' }), (Get-IpNextContentVersion $ContentVersion))
}

function Clear-IpUpdateTarget {
    $script:State.UpdateTarget = $null
    $script:Ui.ExistingAppIdBox.Text = ''
    Set-IpComboValue $script:Ui.PublishModeCombo 'Create'
    Update-IpUpdateBanner
    Write-IpLog 'Publication mode switched to Create: the next publication creates a new Win32 app.'
}

function Set-IpSelectedAppDetailsText {
    param([psobject]$Details)
    $content = if ($Details.committedContentVersion) { $Details.committedContentVersion } else { 'none (incomplete app)' }
    $version = if ($Details.displayVersion) { $Details.displayVersion } else { 'not set' }
    $script:Ui.SelectedAppDetailsText.Text = "{0}`nVersion: {1}; committed content version: {2}; setup file: {3}; publishing state: {4}; assigned: {5}; last modified: {6}" -f $Details.displayName, $version, $content, $Details.setupFilePath, $Details.publishingState, $Details.isAssigned, (Format-IpLocalDateTime $Details.lastModifiedDateTime)
}

function Invoke-IpAppDetailsResult {
    <#
    .SYNOPSIS
        Result callback of Start-IpAppDetailsJob: caches the details, refreshes the selection and the
        update banner, then runs the follow-up action of the request.
    #>
    param([AllowNull()]$Result, [hashtable]$Context = @{})
    # The worker returns @{ Details; App }; a details object alone is accepted too.
    $details = $Result
    $app = $null
    if ($null -ne (Get-IpObjectValue $Result 'Details')) { $details = Get-IpObjectValue $Result 'Details'; $app = Get-IpObjectValue $Result 'App' }
    $appId = [string](Get-IpObjectValue $details 'id')
    if (-not $appId) { Write-IpLog 'App details were empty; nothing to show.' -Level WARN; return }
    $script:State.AppDetails[$appId] = $details
    $selected = $script:Ui.AppsGrid.SelectedItem
    if ($selected -and [string](Get-IpObjectValue $selected 'id') -eq $appId) { Set-IpSelectedAppDetailsText $details }
    $target = $script:State.UpdateTarget
    if ($target -and $target.AppId -eq $appId) {
        $target.ContentVersion = [string]$details.committedContentVersion
        $target.DisplayName = [string]$details.displayName
        Update-IpUpdateBanner
    }
    switch ($Context.Purpose) {
        'UseForUpdate' { Set-IpUpdateTarget -AppId $appId -DisplayName $details.displayName -ContentVersion $details.committedContentVersion }
        'ConfirmPublish' {
            if ($null -eq $app) { [void](Show-IpMessage -Text 'The app to update could not be read; nothing was published.' -Icon Error); return }
            Confirm-IpUpdatePublication -Request $Context.Request -Current $app
        }
        'Import' {
            if ($null -eq $app) { [void](Show-IpMessage -Text 'The app could not be read; nothing was imported.' -Icon Error); return }
            Invoke-IpImportResult -App $app
        }
        'Refresh' { Write-IpLog ("App {0} now has committed content version {1}." -f $appId, $details.committedContentVersion) }
    }
}

function Start-IpAppDetailsJob {
    <#
    .SYNOPSIS
        Reads one app (GET by id, read-only) and runs the follow-up action on the UI thread:
        Show (selection), UseForUpdate, ConfirmPublish, Import or Refresh (after a publication).
    #>
    param([Parameter(Mandatory = $true)][string]$AppId, [ValidateSet('Show', 'UseForUpdate', 'ConfirmPublish', 'Import', 'Refresh')][string]$Purpose = 'Show', [hashtable]$Context = @{})
    $Context.AppId = $AppId
    $Context.Purpose = $Purpose
    Start-IpJob -Title $(if ($Purpose -eq 'Import') { 'Import definition from Intune' } else { 'Read app details' }) -Arguments @($AppId) -Context $Context -Script {
        param($AppId)
        $app = Get-IpWin32App -AppId $AppId -Step 'AppDetails'
        [pscustomobject]@{ Details = (ConvertTo-IpWin32AppSummary -App $app); App = $app }
    } -OnSuccess {
        param($result, $context)
        Invoke-IpAppDetailsResult -Result $result -Context $context
    } -OnFailure {
        param($message, $context)
        if ($context.Purpose -in @('Show', 'UseForUpdate')) { $script:Ui.SelectedAppDetailsText.Text = 'App details unavailable: ' + $message }
        if ($context.Purpose -eq 'ConfirmPublish') { [void](Show-IpMessage -Text ("The app to update could not be read; nothing was published.`n`n{0}" -f $message) -Icon Error) }
        if ($context.Purpose -eq 'Import') { [void](Show-IpMessage -Text ("The app could not be read; nothing was imported.`n`n{0}" -f $message) -Icon Error) }
    }
}

function Get-IpUpdatePreview {
    <#
    .SYNOPSIS
        Computes, without any network call, the differential update of the current app: changes,
        warnings (MSI product code, MSI to non-MSI), concurrency status and current timestamp.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Request, [Parameter(Mandatory = $true)]$Current, [string]$BaseDirectory = '')
    $definition = $Request.Definition
    $metadata = Read-IpIntuneWinMetadata -Path $Request.PackagePath
    $body = New-IpWin32LobAppBody -Definition $definition -IntuneWinMetadata $metadata -BaseDirectory $BaseDirectory
    $update = New-IpUpdatePatchBody -AppBody $body -Current $Current -ContentVersionId 'next' -Definition $definition
    $warnings = [Collections.Generic.List[string]]::new()
    foreach ($warning in $update.Warnings) { $warnings.Add($warning) }
    $packageCode = ([string](Get-IpObjectProperty (Get-IpObjectProperty $metadata 'MsiInfo') 'ProductCode')).ToUpperInvariant()
    if ($packageCode) {
        $ruleCodes = @((Get-IpAppRuleSet -App $Current).Detection | Where-Object { (Get-IpRuleKind ([string](Get-IpObjectProperty $_ '@odata.type'))) -eq 'ProductCode' } | ForEach-Object { ([string](Get-IpObjectProperty $_ 'productCode')).ToUpperInvariant() } | Select-Object -Unique)
        foreach ($code in @($ruleCodes | Where-Object { $_ -and $_ -ne $packageCode })) {
            $warnings.Add(("The new MSI ProductCode {0} differs from the ProductCode detection rule in Intune ({1}): devices with the new version may not be detected unless the detection is updated." -f $packageCode, $code))
        }
        if ($definition.Detection.Type -eq 'MsiProductCode' -and ([string]$definition.Detection.Msi.ProductCode).ToUpperInvariant() -ne $packageCode) {
            $warnings.Add(("The new MSI ProductCode {0} differs from the ProductCode detection rule of the form ({1})." -f $packageCode, $definition.Detection.Msi.ProductCode))
        }
    }
    return [pscustomobject]@{
        Update              = $update
        Warnings            = $warnings.ToArray()
        Concurrency         = Test-IpImportConcurrency -Definition $definition -Current $Current
        CurrentLastModified = ConvertTo-IpIsoUtc (Get-IpObjectProperty $Current 'lastModifiedDateTime')
    }
}

function Get-IpConcurrencyQuestion {
    param([string]$Message)
    return (@(
            'The app was modified in Intune after its definition was imported.'
            ''
            $Message
            ''
            'Yes: re-import the definition from Intune (the form is replaced, your edits in the form are lost).'
            'No: publish anyway (force). Properties that differ between the form and Intune are replaced, including portal changes.'
            'Cancel: do nothing.'
        ) -join "`n")
}

function Confirm-IpUpdatePublication {
    <#
    .SYNOPSIS
        UpdateExisting confirmation: concurrency choice, list of changes and warnings, then publication.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Request, [Parameter(Mandatory = $true)]$Current)
    try { $preview = Get-IpUpdatePreview -Request $Request -Current $Current -BaseDirectory (Get-IpBaseDirectory) }
    catch {
        Write-IpLog ("The update could not be prepared: {0}" -f $_.Exception.Message) -Level ERROR
        [void](Show-IpMessage -Text ("The update could not be prepared; nothing was published.`n`n{0}" -f $_.Exception.Message) -Icon Error)
        return
    }
    $force = $false
    if ($preview.Concurrency) {
        Write-IpLog $preview.Concurrency -Level WARN
        $choice = [string](Show-IpMessage -Text (Get-IpConcurrencyQuestion $preview.Concurrency) -Icon Warning -Buttons YesNoCancel)
        if ($choice -eq 'Yes') { Invoke-IpImportResult -App $Current; return }
        if ($choice -ne 'No') { Write-IpLog 'Publication canceled by the operator.'; return }
        $force = $true
        Write-IpLog 'Publication forced by the operator despite the concurrent modification.' -Level WARN
    }
    $text = Get-IpUpdateConfirmationText -Details (ConvertTo-IpWin32AppSummary -App $Current) -ProfileDisplayName $Request.Definition.App.DisplayName -Tenant $script:State.ConnectedTenant -PackagePath $Request.PackagePath -Changes @($preview.Update.Changes) -Warnings @($preview.Warnings) -Forced:$force
    $answer = Show-IpMessage -Text $text -Icon $(if (@($preview.Warnings).Count) { 'Warning' } else { 'Question' }) -Buttons YesNo
    if ([string]$answer -ne 'Yes') { Write-IpLog 'Publication canceled by the operator.'; return }
    $Request.Force = $force
    $Request.ExpectedLastModified = $preview.CurrentLastModified
    Start-IpPublishJob -Request $Request
}

function Invoke-IpImportFromIntune {
    $app = $script:Ui.AppsGrid.SelectedItem
    if (-not $app -or -not [string](Get-IpObjectValue $app 'id')) { [void](Show-IpMessage -Text 'Select an app in the list first.' -Icon Warning); return }
    if (-not $script:State.Connected) { [void](Show-IpMessage -Text 'Connect to Microsoft Graph first.' -Icon Warning); return }
    if (Test-IpDefinitionLoaded -Definition (Get-IpFormDefinition)) {
        $answer = Show-IpMessage -Text ("Replace the current form with the definition of '{0}' read from Intune?`n`nUnsaved changes in the form are lost." -f $app.displayName) -Icon Question -Buttons YesNo
        if ([string]$answer -ne 'Yes') { return }
    }
    if ($script:State.Job) { [void](Show-IpMessage -Text 'Wait for the running operation to finish, then click again.' -Icon Information); return }
    Start-IpAppDetailsJob -AppId ([string]$app.id) -Purpose Import
}

function Invoke-IpImportResult {
    <#
    .SYNOPSIS
        Fills the form from a Win32 app read from Intune (no network call). The detection script and
        the logo are written under %LOCALAPPDATA%\SmartM365\IntunePackager\Imported\<AppId>.
    #>
    param([Parameter(Mandatory = $true)]$App)
    $import = ConvertFrom-IpWin32App -App $App
    $folder = Join-Path (Resolve-IpPath -Path '%LOCALAPPDATA%\SmartM365\IntunePackager\Imported' -BaseDirectory $script:AppRoot) ([string]$App.id)
    $definition = Save-IpImportAssets -Import $import -Folder $folder -BaseName ([string]$App.displayName)
    $script:State.ProfilePath = ''
    $script:State.IntuneWinPath = ''
    Set-IpFormDefinition $definition
    $script:Ui.ProfilePathText.Text = "Imported from Intune: {0} (unsaved profile: set the package source and output folders, then save)" -f $App.displayName
    $script:Ui.PackageInfoText.Text = 'No package built in this session.'
    Set-IpUpdateTarget -AppId ([string]$App.id) -DisplayName ([string]$App.displayName) -ContentVersion ([string]$App.committedContentVersion)
    Write-IpLog ("Definition imported from Intune: {0} ({1}), content version {2}. Package source left empty." -f $App.displayName, $App.id, $App.committedContentVersion)
    foreach ($item in $import.Preserved) { Write-IpLog ('Preserved in Intune: ' + $item) }
    foreach ($warning in $import.Warnings) { Write-IpLog $warning -Level WARN }
    $lines = @(("Definition imported from '{0}' (content version {1})." -f $App.displayName, $App.committedContentVersion), '', 'Next: set the source and output folders on the Package tab, build, then publish.')
    if (@($import.Preserved).Count) { $lines += @('', 'Preserved in Intune (not edited, never sent):') + @($import.Preserved | ForEach-Object { '- ' + $_ }) }
    if (@($import.Warnings).Count) { $lines += @('', 'Warnings:') + @($import.Warnings | ForEach-Object { '- ' + $_ }) }
    [void](Show-IpMessage -Text ($lines -join "`n") -Icon $(if (@($import.Warnings).Count) { 'Warning' } else { 'Information' }))
}

function Invoke-IpPublish {
    $definition = Get-IpFormDefinition
    if (-not (Assert-IpDefinitionLoaded -Definition $definition)) { return }
    if (-not (Test-IpFormDefinition -Definition $definition -ForPublish -Quiet)) { return }
    $dryRun = [bool]$script:Ui.DryRunCheck.IsChecked
    try { $buildRequest = Get-IpBuildRequest -Definition $definition }
    catch { [void](Show-IpMessage -Text $_.Exception.Message -Icon Warning); return }

    $packagePath = $script:State.IntuneWinPath
    if (-not $packagePath -or -not (Test-Path -LiteralPath $packagePath -PathType Leaf)) { $packagePath = Get-IpExpectedPackagePath -Definition $definition -OutputFolder $buildRequest.OutputFolder }
    if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf)) {
        [void](Show-IpMessage -Text "Package not found: $packagePath`n`nBuild the package first." -Icon Warning)
        return
    }

    $request = @{
        Definition           = $definition
        PackagePath          = $packagePath
        DryRun               = $dryRun
        DryRunFolder         = $(if ($dryRun) { Join-Path (Split-Path $packagePath -Parent) ('DryRun\' + (Get-Date -Format 'yyyyMMdd-HHmmss')) } else { '' })
        Force                = $false
        ExpectedLastModified = ''
    }
    if ($dryRun) { Start-IpPublishJob -Request $request; return }
    if (-not $script:State.Connected) { [void](Show-IpMessage -Text 'Connect to Microsoft Graph first, or select Dry run.' -Icon Warning); return }

    if ($definition.Publish.Mode -eq 'UpdateExisting') {
        # Read the app first: the confirmation lists the changes against its current state.
        Start-IpAppDetailsJob -AppId $definition.Publish.ExistingAppId -Purpose ConfirmPublish -Context @{ Request = $request }
        return
    }
    $answer = Show-IpMessage -Text ("Publish '{0}' as a new Win32 app on tenant {1}?`n`nThe app will not be assigned." -f $definition.App.DisplayName, $script:State.ConnectedTenant) -Icon Question -Buttons YesNo
    if ([string]$answer -ne 'Yes') { Write-IpLog 'Publication canceled by the operator.'; return }
    Start-IpPublishJob -Request $request
}

function Start-IpPublishJob {
    param([Parameter(Mandatory = $true)][hashtable]$Request)
    $force = [bool]$Request['Force']
    $expected = [string]$Request['ExpectedLastModified']
    Start-IpJob -Title $(if ($Request.DryRun) { 'Dry run' } else { 'Publish' }) -Arguments @($Request.Definition, $Request.PackagePath, (Get-IpBaseDirectory), $script:State.Config.Upload, $Request.DryRun, $Request.DryRunFolder, $force, $expected) -Context $Request -Script {
        param($Definition, $PackagePath, $BaseDirectory, $Upload, $DryRun, $DryRunFolder, $Force, $ExpectedLastModified)
        $publishArgs = @{
            Definition       = $Definition
            IntuneWinPath    = $PackagePath
            BaseDirectory    = $BaseDirectory
            UploadSettings   = $Upload
            ProgressCallback = { param($sent, $total) $IpSync.Sent = [int64]$sent; $IpSync.Total = [int64]$total }
            CancelCheck      = { [bool]$IpSync.Cancel }
            Confirm          = $false
        }
        if ($DryRun) { $publishArgs.DryRun = $true; $publishArgs.DryRunOutputFolder = $DryRunFolder }
        if ($Force) { $publishArgs.Force = $true }
        if ($ExpectedLastModified) { $publishArgs.ExpectedLastModified = $ExpectedLastModified }
        Publish-IpWin32App @publishArgs
    } -OnSuccess {
        param($result, $context)
        if ($null -eq $result) { return }
        if ($result.Mode -eq 'DryRun') {
            [void](Show-IpMessage -Text ("Dry run completed: {0} Graph requests planned, Microsoft Graph was not called.`n`nPayloads: {1}" -f @($result.Plan).Count, $context.DryRunFolder))
            return
        }
        Set-IpPublishResult -Result $result
        # Record our own update in ImportedFrom so the next publication is not blocked by [Concurrency].
        if (Sync-IpImportedFromAfterPublish -Definition $context.Definition -Result $result -ProfilePath ([string]$script:State.ProfilePath)) { $script:State.ImportedFrom = $context.Definition.ImportedFrom }
        [void](Show-IpMessage -Text (Get-IpPublishResultText -Result $result) -Icon $(if (@($result.VerificationIssues).Count) { 'Warning' } else { 'Information' }))
        # Refresh the selected app details and the update banner with the new committed version.
        Start-IpAppDetailsJob -AppId $result.AppId -Purpose Refresh
    }
}

function Get-IpPublishResultText {
    param([Parameter(Mandatory = $true)][psobject]$Result)
    $lines = @('Publication completed.', '', ('App: {0}' -f $Result.DisplayName), ('App id: {0}' -f $Result.AppId), ('Content version: {0}' -f $Result.ContentVersionId))
    if ($Result.PublishMode -eq 'UpdateExisting') { $lines += ('Metadata changes: {0}' -f @($Result.Changes).Count) }
    $issues = @($Result.VerificationIssues)
    if ($issues.Count) { $lines += @('', 'Post-publication check:') + @($issues | ForEach-Object { '- ' + $_ }) }
    else { $lines += @('', 'Post-publication check: the content version is committed and the preserved properties are unchanged.') }
    $lines += @('', 'The app is not assigned. Create assignments in the Intune admin center.')
    return ($lines -join "`n")
}

function Set-IpPublishResult {
    param([Parameter(Mandatory = $true)][psobject]$Result)
    $script:State.LastAppId = [string]$Result.AppId
    $script:State.LastResult = $Result
    $script:Ui.ResultAppIdBox.Text = [string]$Result.AppId
    $script:Ui.PortalLinkRun.Text = [string]$Result.PortalUrl
    $script:Ui.PortalLink.NavigateUri = [Uri][string]$Result.PortalUrl
    $script:Ui.MainTabs.SelectedIndex = 5
    $target = $script:State.UpdateTarget
    if ($target -and $target.AppId -eq [string]$Result.AppId) { $target.ContentVersion = [string]$Result.ContentVersionId; Update-IpUpdateBanner }
}

function Get-IpObjectValue {
    param($InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-IpAppGridRow {
    param([psobject]$App)
    $modified = Get-IpObjectValue $App 'lastModifiedDateTime'
    return [pscustomobject]@{
        id                   = [string](Get-IpObjectValue $App 'id')
        displayName          = [string](Get-IpObjectValue $App 'displayName')
        publisher            = [string](Get-IpObjectValue $App 'publisher')
        lastModifiedDateTime = $modified
        LastModified         = Format-IpLocalDateTime $modified
    }
}

function Set-IpSearchResults {
    <#
    .SYNOPSIS
        Result callback of the Win32 app search: binds the rows and explains an empty result.

    .DESCRIPTION
        ConvertTo-IpGridItems returns its array as one pipeline object, so it is iterated with
        foreach, never piped: piping passed the whole (empty) array to ConvertTo-IpAppGridRow.
    #>
    param([AllowNull()]$Result, [string]$Filter)
    $rows = [Collections.Generic.List[object]]::new()
    foreach ($app in (ConvertTo-IpGridItems $Result)) {
        $row = ConvertTo-IpAppGridRow $app
        if ($row.id) { $rows.Add($row) }
    }
    $script:Ui.AppsGrid.ItemsSource = [object[]]$rows.ToArray()
    $script:Ui.SelectedAppDetailsText.Text = if ($rows.Count) { 'Select an app to read its version and committed content version.' }
    elseif ($Filter) { "No Win32 app matches '{0}'. A newly created app can take a few minutes to appear in the list." -f $Filter }
    else { 'No Win32 app was found in this tenant.' }
    if (-not $rows.Count) { Write-IpLog ("No Win32 app matches '{0}'." -f $Filter) }
}

function Invoke-IpFindApps {
    if (-not $script:State.Connected) { [void](Show-IpMessage -Text 'Connect to Microsoft Graph first.' -Icon Warning); return }
    $text = $script:Ui.FindAppBox.Text.Trim()
    Start-IpJob -Title 'Search Win32 apps' -Arguments @($text) -Context @{ Filter = $text } -Script {
        param($Text)
        $apps = @(Find-IpWin32App -DisplayName $Text -Contains | Sort-Object displayName)
        Write-IpLog ("Found {0} Win32 app(s) matching '{1}'." -f $apps.Count, $Text)
        , $apps
    } -OnSuccess {
        param($result, $context)
        Set-IpSearchResults -Result $result -Filter $context.Filter
    }
}

function Invoke-IpShowSelectedAppDetails {
    <#
    .SYNOPSIS
        Reads the Win32 properties of the selected app (one GET by id); the search grid only lists base properties.
    #>
    $app = $script:Ui.AppsGrid.SelectedItem
    if (-not $app -or -not $script:State.Connected -or -not [string](Get-IpObjectValue $app 'id')) { return }
    if ($script:State.Job) { $script:Ui.SelectedAppDetailsText.Text = 'Details are loaded when no operation is running; select the app again.'; return }
    $script:Ui.SelectedAppDetailsText.Text = 'Reading app details...'
    Start-IpAppDetailsJob -AppId ([string]$app.id) -Purpose Show
}

function Invoke-IpUseSelectedAppForUpdate {
    $app = $script:Ui.AppsGrid.SelectedItem
    if (-not $app -or -not [string](Get-IpObjectValue $app 'id')) { [void](Show-IpMessage -Text 'Select an app in the list first.' -Icon Warning); return }
    $details = $script:State.AppDetails[[string]$app.id]
    if ($details) { Set-IpUpdateTarget -AppId $details.id -DisplayName $details.displayName -ContentVersion $details.committedContentVersion; return }
    if ($script:State.Job) { [void](Show-IpMessage -Text 'Wait for the running operation to finish, then click again.' -Icon Information); return }
    Start-IpAppDetailsJob -AppId ([string]$app.id) -Purpose UseForUpdate
}

function Invoke-IpUseResultForUpdate {
    $result = $script:State.LastResult
    if (-not $result) { return }
    Set-IpUpdateTarget -AppId ([string]$result.AppId) -DisplayName ([string]$result.DisplayName) -ContentVersion ([string]$result.ContentVersionId)
}

function Save-IpProfile {
    param([switch]$As)
    $path = $script:State.ProfilePath
    if ($As -or -not $path) {
        $name = if ($script:Ui.DisplayNameBox.Text.Trim()) { ($script:Ui.DisplayNameBox.Text.Trim() -replace '[^A-Za-z0-9.-]+', '-') + '.profile.json' } else { 'App.profile.json' }
        $path = Select-IpSaveFile -Title 'Save app profile' -Filter 'App profile (*.json)|*.json' -FileName $name -Initial (Resolve-IpPath -Path ([string]$script:State.Config.Paths.ProfilesFolder) -BaseDirectory $script:AppRoot)
        if (-not $path) { return }
    }
    $definition = Get-IpFormDefinition
    Export-IpAppProfile -Definition $definition -Path $path -Confirm:$false
    $script:State.ProfilePath = $path
    $script:Ui.ProfilePathText.Text = $path
    Write-IpLog "App profile saved: $path"
}

function Open-IpProfile {
    param([string]$Path)
    if (-not $Path) {
        $Path = Select-IpFile -Title 'Open app profile' -Filter 'App profile (*.json;*.json.template)|*.json;*.json.template|All files (*.*)|*.*' -Initial (Resolve-IpPath -Path ([string]$script:State.Config.Paths.ProfilesFolder) -BaseDirectory $script:AppRoot)
        if (-not $Path) { return }
    }
    try {
        $full = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
        $definition = Import-IpAppProfile -Path $full
        $script:State.ProfilePath = $full
        $script:State.IntuneWinPath = ''
        $script:Ui.ProfilePathText.Text = $full
        Set-IpFormDefinition $definition
        Write-IpLog "App profile loaded: $full"
    }
    catch {
        Write-IpErrorDetail -ErrorRecord $_
        [void](Show-IpMessage -Text $_.Exception.Message -Icon Error)
    }
}

function Stop-IpWorker {
    <#
    .SYNOPSIS
        Cancels any running operation, disconnects Graph in the worker and closes the runspace.
    #>
    $runspace = $script:State.Worker
    if (-not $runspace) { return }
    $job = $script:State.Job
    if ($job) {
        $script:IpSync.Cancel = $true
        try { $job.PowerShell.Stop() } catch { Write-IpLog ("Unable to stop the running operation: {0}" -f $_.Exception.Message) -Level WARN }
        $job.PowerShell.Dispose()
        $script:State.Job = $null
    }
    $disconnect = [PowerShell]::Create()
    $disconnect.Runspace = $runspace
    try {
        [void]$disconnect.AddScript({ Disconnect-IpGraph }.ToString())
        $handle = $disconnect.BeginInvoke()
        if (-not $handle.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds(15))) { Write-IpLog 'Microsoft Graph disconnection did not complete within 15 seconds.' -Level WARN }
        else { [void]$disconnect.EndInvoke($handle) }
    }
    catch { Write-IpLog ("Graph disconnection on close failed: {0}" -f $_.Exception.Message) -Level WARN }
    finally { $disconnect.Dispose() }
    $runspace.Close()
    $runspace.Dispose()
    $script:State.Worker = $null
}

#endregion

#region GUI startup

function Initialize-IpWindow {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $window = [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new([xml]$xaml))
    $script:Ui.Window = $window
    foreach ($match in [regex]::Matches($xaml, 'x:Name="(?<Name>[A-Za-z][A-Za-z0-9_]*)"')) {
        $name = $match.Groups['Name'].Value
        $script:Ui[$name] = $window.FindName($name)
    }
    $window.Title = '{0} {1}' -f $script:AppName, $script:AppVersion
    $icon = Join-Path $script:AppRoot 'WorkplaceCloudHub.ico'
    if (Test-Path -LiteralPath $icon) { $window.Icon = [Windows.Media.Imaging.BitmapFrame]::Create([Uri]$icon) }
    $logo = Join-Path $script:AppRoot 'WorkplaceCloudHub-lockup-WPF.png'
    if (Test-Path -LiteralPath $logo) { $script:Ui.HeaderLogo.Source = [Windows.Media.Imaging.BitmapFrame]::Create([Uri]$logo) }

    $u = $script:Ui
    $u.RunAsCombo.ItemsSource = @('system', 'user')
    $u.RestartCombo.ItemsSource = @('suppress', 'allow', 'basedOnReturnCode', 'force')
    $u.MinWindowsCombo.ItemsSource = @(Get-IpWindowsReleases)
    $u.DetectionTypeCombo.ItemsSource = @('MsiProductCode', 'File', 'Registry', 'Script', 'Preserved')
    $u.MsiVersionOperatorCombo.ItemsSource = $script:Operators
    $u.PublishModeCombo.ItemsSource = @('Create', 'UpdateExisting')
    $u.TenantCombo.ItemsSource = @($script:State.Config.Tenants.PSObject.Properties.Name)
    $lastTenant = if ($script:State.Config.PSObject.Properties['Gui']) { [string]$script:State.Config.Gui.LastTenant } else { '' }
    $initialTenant = Resolve-IpInitialTenant -Names @($u.TenantCombo.Items) -Requested $Tenant -RequestedExplicitly $script:TenantExplicit -LastUsed $lastTenant
    if ($initialTenant) { $u.TenantCombo.SelectedItem = $initialTenant }
    Write-IpLog ("Tenant profile preselected: {0}. No connection is made until you click Connect." -f $(if ($initialTenant) { $initialTenant } else { 'none' }))

    $script:ReturnCodesTable = New-IpDataTable -Columns @('ReturnCode', 'Type') -IntegerColumns @('ReturnCode')
    $script:FileRulesTable = New-IpDataTable -Columns @('Path', 'FileOrFolderName', 'DetectionType', 'Operator', 'DetectionValue', 'Check32BitOn64System') -BooleanColumns @('Check32BitOn64System')
    $script:RegistryRulesTable = New-IpDataTable -Columns @('KeyPath', 'ValueName', 'DetectionType', 'Operator', 'DetectionValue', 'Check32BitOn64System') -BooleanColumns @('Check32BitOn64System')
    $u.ReturnCodesGrid.ItemsSource = $script:ReturnCodesTable.DefaultView
    $u.FileRulesGrid.ItemsSource = $script:FileRulesTable.DefaultView
    $u.RegistryRulesGrid.ItemsSource = $script:RegistryRulesTable.DefaultView
    $u.ReturnCodesGrid.Columns[1].ItemsSource = $script:ReturnCodeTypes
    $u.FileRulesGrid.Columns[2].ItemsSource = @('exists', 'doesNotExist', 'modifiedDate', 'createdDate', 'version', 'sizeInMB')
    $u.FileRulesGrid.Columns[3].ItemsSource = $script:Operators
    $u.RegistryRulesGrid.Columns[2].ItemsSource = @('exists', 'doesNotExist', 'string', 'integer', 'version')
    $u.RegistryRulesGrid.Columns[3].ItemsSource = $script:Operators

    Register-IpEventHandlers
    Set-IpFormDefinition (New-IpAppDefinition)
    Update-IpConnectionStatus
    return $window
}

function Register-IpEventHandlers {
    $u = $script:Ui
    $u.HeaderLogo.Add_MouseLeftButtonUp({ Start-Process 'https://workplacecloudhub.com' })
    $u.ConnectButton.Add_Click({ Invoke-IpConnect })
    $u.DisconnectButton.Add_Click({ Invoke-IpDisconnect })
    $u.NewProfileButton.Add_Click({
            $script:State.ProfilePath = ''
            $script:State.IntuneWinPath = ''
            $script:Ui.ProfilePathText.Text = 'Unsaved profile'
            $script:Ui.PackageInfoText.Text = 'No package built in this session.'
            Set-IpFormDefinition (New-IpAppDefinition)
            Write-IpLog 'New app profile.'
        })
    $u.OpenProfileButton.Add_Click({ Open-IpProfile })
    $u.SaveProfileButton.Add_Click({ Save-IpProfile })
    $u.SaveProfileAsButton.Add_Click({ Save-IpProfile -As })
    $u.ValidateButton.Add_Click({ [void](Test-IpFormDefinition -Definition (Get-IpFormDefinition)) })
    $u.BrowseSourceButton.Add_Click({
            $path = Select-IpFolder -Title 'Select the source folder' -Initial (Resolve-IpPath -Path $script:Ui.SourceFolderBox.Text -BaseDirectory (Get-IpBaseDirectory))
            if ($path) { $script:Ui.SourceFolderBox.Text = $path; Invoke-IpDetectMetadata -Quiet }
        })
    $u.BrowseSetupButton.Add_Click({
            $source = Resolve-IpPath -Path $script:Ui.SourceFolderBox.Text -BaseDirectory (Get-IpBaseDirectory)
            $path = Select-IpFile -Title 'Select the setup file' -Filter 'Setup files (*.msi;*.exe;*.ps1;*.cmd;*.bat)|*.msi;*.exe;*.ps1;*.cmd;*.bat|All files (*.*)|*.*' -Initial $source
            if (-not $path) { return }
            if (-not $source -or -not $path.StartsWith($source.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
                [void](Show-IpMessage -Text 'The setup file must be inside the source folder.' -Icon Warning)
                return
            }
            $script:Ui.SetupFileBox.Text = [IO.Path]::GetRelativePath($source, $path)
            Invoke-IpDetectMetadata -Quiet
        })
    $u.BrowseOutputButton.Add_Click({
            $path = Select-IpFolder -Title 'Select the output folder' -Initial (Resolve-IpPath -Path $script:Ui.OutputFolderBox.Text -BaseDirectory (Get-IpBaseDirectory))
            if ($path) { $script:Ui.OutputFolderBox.Text = $path }
        })
    $u.BrowseLogoButton.Add_Click({
            $path = Select-IpFile -Title 'Select the app logo' -Filter 'Images (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg'
            if ($path) { $script:Ui.LogoPathBox.Text = $path; Update-IpLogoPreview }
        })
    $u.LogoPathBox.Add_LostFocus({ Update-IpLogoPreview })
    $u.BrowseScriptButton.Add_Click({
            $path = Select-IpFile -Title 'Select the detection script' -Filter 'PowerShell script (*.ps1)|*.ps1'
            if ($path) { $script:Ui.ScriptPathBox.Text = $path }
        })
    $u.DetectMetadataButton.Add_Click({ Invoke-IpDetectMetadata })
    $u.BuildButton.Add_Click({ Invoke-IpBuild })
    $u.PublishButton.Add_Click({ Invoke-IpPublish })
    $u.BuildAndPublishButton.Add_Click({ Invoke-IpBuild -Then { Invoke-IpPublish } })
    $u.FindAppButton.Add_Click({ Invoke-IpFindApps })
    $u.AppsGrid.Add_SelectionChanged({ Invoke-IpShowSelectedAppDetails })
    $u.UseSelectedAppButton.Add_Click({ Invoke-IpUseSelectedAppForUpdate })
    $u.ImportFromIntuneButton.Add_Click({ Invoke-IpImportFromIntune })
    $u.ArchNotRestrictedCheck.Add_Click({ Update-IpArchitectureChecks })
    $u.SwitchToCreateButton.Add_Click({ Clear-IpUpdateTarget })
    $u.PublishModeCombo.Add_SelectionChanged({ Update-IpUpdateBanner })
    $u.ExistingAppIdBox.Add_TextChanged({ Update-IpUpdateBanner })
    $u.TenantCombo.Add_SelectionChanged({ Update-IpConnectionStatus })
    $u.CopyAppIdButton.Add_Click({ if ($script:State.LastAppId) { [Windows.Clipboard]::SetText($script:State.LastAppId); Write-IpLog 'App id copied to the clipboard.' } })
    $u.OpenPortalButton.Add_Click({ if ($script:State.LastAppId) { Start-Process (Get-IpIntunePortalAppUrl -AppId $script:State.LastAppId) } })
    $u.PortalLink.Add_RequestNavigate({ param($source, $eventArgs) if ($script:State.LastAppId) { Start-Process $eventArgs.Uri.AbsoluteUri }; $eventArgs.Handled = $true })
    $u.UseResultForUpdateButton.Add_Click({ Invoke-IpUseResultForUpdate })
    $u.DetectionTypeCombo.Add_SelectionChanged({ Update-IpDetectionPanels })
    $u.AddReturnCodeButton.Add_Click({ $row = $script:ReturnCodesTable.NewRow(); $row['ReturnCode'] = 0; $row['Type'] = 'failed'; $script:ReturnCodesTable.Rows.Add($row) })
    $u.RemoveReturnCodeButton.Add_Click({ $item = $script:Ui.ReturnCodesGrid.SelectedItem; if ($item -is [Data.DataRowView]) { $item.Row.Delete() } })
    $u.DefaultReturnCodesButton.Add_Click({ Set-IpDataTableRows $script:ReturnCodesTable @(Get-IpDefaultReturnCodes) })
    $u.AddFileRuleButton.Add_Click({ $row = $script:FileRulesTable.NewRow(); $row['DetectionType'] = 'exists'; $row['Operator'] = 'notConfigured'; $script:FileRulesTable.Rows.Add($row) })
    $u.RemoveFileRuleButton.Add_Click({ $item = $script:Ui.FileRulesGrid.SelectedItem; if ($item -is [Data.DataRowView]) { $item.Row.Delete() } })
    $u.AddRegistryRuleButton.Add_Click({ $row = $script:RegistryRulesTable.NewRow(); $row['KeyPath'] = 'HKEY_LOCAL_MACHINE\SOFTWARE\'; $row['DetectionType'] = 'exists'; $row['Operator'] = 'notConfigured'; $script:RegistryRulesTable.Rows.Add($row) })
    $u.RemoveRegistryRuleButton.Add_Click({ $item = $script:Ui.RegistryRulesGrid.SelectedItem; if ($item -is [Data.DataRowView]) { $item.Row.Delete() } })
    $u.CancelOperationButton.Add_Click({
            if ($script:State.Job) {
                $script:IpSync.Cancel = $true
                Write-IpLog 'Cancellation requested; the operation stops at the next block or state check.' -Level WARN
            }
        })
    $u.ExportLogButton.Add_Click({
            $path = Select-IpSaveFile -Title 'Export log' -Filter 'Text file (*.txt)|*.txt' -FileName ('SmartM365-IntunePackager-Log-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
            if ($path) { [IO.File]::WriteAllText($path, $script:Ui.LogTextBox.Text, [Text.UTF8Encoding]::new($false)); Write-IpLog "Log exported: $path" }
        })
    $u.OpenLogFolderButton.Add_Click({ $folder = Split-Path (Get-IpLogPath) -Parent; if ($folder) { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $folder) } })
    $u.ClearLogButton.Add_Click({ $script:Ui.LogTextBox.Clear() })

    $u.Window.Add_Closing({
            param($source, $eventArgs)
            if ($script:State.Job) {
                $answer = Show-IpMessage -Text ("'{0}' is still running. Close and cancel it?" -f $script:State.Job.Title) -Icon Warning -Buttons YesNo
                if ([string]$answer -ne 'Yes') { $eventArgs.Cancel = $true }
            }
        })
}

function Start-IpGui {
    $script:State.Config = Initialize-IpApplication -LogBaseName 'SmartM365-IntunePackager-GUI'
    Set-IpLogSink -Sink { param($line, $level) $script:LogQueue.Enqueue($line) } -Console $false

    if (-not $ValidateOnly) {
        $splashPath = Join-Path $script:AppRoot 'SmartM365.GuiSplash.ps1'
        if (Test-Path -LiteralPath $splashPath) {
            . $splashPath
            $script:State.Splash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'IntunePackager' -WindowIconPath (Join-Path $script:AppRoot 'WorkplaceCloudHub.ico')
        }
    }

    $window = Initialize-IpWindow
    Register-IpDispatcherExceptionHandler -Dispatcher $window.Dispatcher
    if ($ProfilePath) { Open-IpProfile -Path $ProfilePath }
    Initialize-IpWorker
    Write-IpLog ("GUI ready. Graph scope: {0}. Published apps are never assigned." -f ((Get-IpGraphScopes) -join ', '))
    Update-IpLogPanel

    if ($ValidateOnly) {
        Stop-IpWorker
        Write-IpLog 'GUI validation completed: the window was built without being displayed.'
        return
    }

    $timer = [Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(200)
    $timer.Add_Tick({ Invoke-IpTimerTick })
    $timer.Start()
    $window.Add_ContentRendered({ if ($script:State.Splash) { Hide-SmartM365GuiSplash -Splash $script:State.Splash } })
    try {
        [void]$window.ShowDialog()
    }
    finally {
        $timer.Stop()
        Stop-IpWorker
        if ($script:State.Splash) { Close-SmartM365GuiSplash -Splash $script:State.Splash }
        Set-IpLogSink -Sink $null -Console $false
        Write-IpProcessExitNotice -Status 'GUI closed.'
    }
}

#endregion

#region Entry point

$script:IpNextProgress = 10
if ($Action -or ($ValidateOnly -and $ProfilePath)) {
    $exitCode = 0
    try { Invoke-IpCli }
    catch {
        $exitCode = 1
        # Errors already logged by the CLI (concurrency block) are not logged a second time.
        if (-not $_.Exception.Data['IpLogged']) { Write-IpErrorDetail -ErrorRecord $_ }
    }
    Write-IpProcessExitNotice -Status ("CLI completed with exit code {0}." -f $exitCode)
    exit $exitCode
}

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    # WPF requires STA: relaunch this script in an STA PowerShell process with the same GUI arguments.
    $quote = { param($value) '"' + ([string]$value).Replace('"', '\"') + '"' }
    $arguments = '-NoProfile -STA -ExecutionPolicy Bypass -File {0} -Tenant {1}' -f (& $quote $PSCommandPath), (& $quote $Tenant)
    if ($ProfilePath) { $arguments += ' -ProfilePath ' + (& $quote ((Resolve-Path -LiteralPath $ProfilePath).ProviderPath)) }
    if ($ValidateOnly) { $arguments += ' -ValidateOnly' }
    Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList $arguments | Out-Null
    return
}

try { Start-IpGui }
catch {
    Write-IpErrorDetail -ErrorRecord $_
    if ($script:State.Splash) { Close-SmartM365GuiSplash -Splash $script:State.Splash }
    try { [void][Windows.MessageBox]::Show($_.Exception.Message, $script:AppName, 'OK', 'Error') } catch { Write-Verbose 'Message box unavailable.' }
    exit 1
}

#endregion

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBPCtYFnSyy+wtM
# 0cNq2T/b7EbRWJwAIV3g7paF4ckxr6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBHr/7o6La1b0bwsmI2pu05X+88vcjx+m9KBA8m9lmyaMA0GCSqG
# SIb3DQEBAQUABIIBgB3RfG+r2zQd5nxA6jPy1pToZm+5WaJOqlidB8DDd75Kk3L9
# p704b18EPzYUEDDEH40rkS2QRObUQho0EPCDAlLXMlersGYeoi2ORw2DcgjfQgnM
# 7GOdTwgDhheBdH9QKQWtHZZWRiLrOHhjKXVlCxyHWDh9E/R2Mj3kPKdb3z2VEwF5
# GWW/1yYdp6snO4mjifhCiLAYIMOFOUUtjJUtAtw9fqtmJh+A4YP4WxUsiXR/7Kxv
# tfW3162igShlxhux+iNWCzw+hEDsjbNAKwdDWt7WGIcx47sC4FwJkn7szNKASMzi
# MImKX1jN5ByF/sdAWUuGpNwnXF+qaD6HS5ufZEURNnougyHXCqm6DcdwaVRAOhz9
# j874s5ihXXF8fn9f6qLGZ4uOrsvgYIjPGB43xZ33N4qZnq9uOPEx4rKRasLYFbiB
# wfE3DM43NQc5ovhgV5zzMwgWHBfOySQYJxY/mekWrpxma6sicBCC6KCtSKtvgs9o
# FnEp26hTqsv9Rc3bTqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzNaMC8GCSqGSIb3DQEJBDEiBCDIPcuBAtlqaBMGWHGSMmVT8TeDc2DnQhHCARfq
# go1syzANBgkqhkiG9w0BAQEFAASCAgAT8vAhv7td2vq3eJAXAKptdQYJOFShaX7p
# BIgIgjlVIALqVtwaZVF10pUgqYDp4eT1GBd+hTvK/3QMxls15XTjth3S46EkyhTP
# fx1AHagxSYRfY4+XEOx1UdVQ2LQ+32ObpW2P9LWIhot1vMMz5f57e70lRAxKKTI7
# DVbkjaxRFKrpLrEBZtrIFAWu8r+zME2ZooK6/OndfLKYG4PMc+vLFLuG+lKUXxbf
# JPmsffTe/PInph0evNu8ynHPrXigU2DGh+KJHlEwuZqFRzISDRKozoH9quLJd6tU
# LHT9kdpC0B01aHaTov22VMBUv8N22pDDZ7zO6fo53Zb7xU6bPh1+7hIvPunM6BAw
# sX2M9TDAEyoXh//+l0tVTJvLqdX3P/VupWjrVYDhdcUPM6eJp5VNuD7q4QAzxuzT
# fOefK7dXTPo+dXSF/k14YpFdVqV/pW4pTpbgDmYP/JEtbW7M6oXeKPxVH5B3+8sc
# ruOHQkO9UCixQnIiwtZ2v1GHJnp/1yXK8r7BtbtimufytJSbetU10vlLcU71IgcJ
# rrh5s9lma+o+YRcIh0f3IwCD3wKLum4yd50xJk9OLC3jh7CMP52H73ML0SP2cDM/
# isdhRPusfp1he45bK85Ajw+T60D4j0mL8cURcqGEe/ITVNFnyCLcKytLrtkKeEoq
# V85CNEzubw==
# SIG # End signature block
