#Requires -Version 7.0

<#
.SYNOPSIS
    Opens Smart Intune Remediation GUI.

.DESCRIPTION
    Lists local SmartM365 remediation packages and Intune deviceHealthScripts,
    edits local detection/remediation scripts, publishes local packages, resets
    Intune execution history by duplicating a cloud remediation and deleting the
    previous object, and exports the execution report to CSV.

    Authentication is delegated and interactive only. The GUI does not use
    application IDs, certificates, client secrets, stored credentials, or
    unattended authentication.

.REQUIREMENTS
    PowerShell module:
    - Microsoft.Graph.Authentication

    Delegated Microsoft Graph permissions:
    - DeviceManagementScripts.ReadWrite.All
    - DeviceManagementConfiguration.Read.All

.VERSION
    1.1

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version: 1.1
#>

[CmdletBinding()]
param(
    [ValidateSet('beta')]
    [string]$GraphApiVersion = 'beta',

    [string]$RemediationRoot = '',

    [ValidateSet('InteractiveBrowser', 'DeviceCode')]
    [string]$GraphAuthMode = 'InteractiveBrowser',

    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    throw "WPF requires an STA runspace. Start the GUI with: pwsh -STA -NoProfile -File `"$PSCommandPath`""
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

$script:GuiSplash = $null
if (-not $ValidateOnly) {
    $splashHelperPath = Import-SmartM365GuiSplash
}
if (-not $ValidateOnly -and $splashHelperPath) {
    . $splashHelperPath
    $script:GuiSplash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'Intune Remediation'
}

$script:BaseScopes = @(
    'DeviceManagementScripts.ReadWrite.All',
    'DeviceManagementConfiguration.Read.All',
    'Group.Read.All'
)
$script:ReportExportScopes = @()
$script:GraphBaseUri = "https://graph.microsoft.com/$GraphApiVersion"
$script:GraphAuthMode = $GraphAuthMode
$script:GraphContext = $null
$script:IsGraphConnected = $false
$script:CloudRemediations = @()
$script:LocalPackages = @()
$script:SelectedLocalPackage = $null
$script:SelectedCloudRemediation = $null
$script:EditorDetectionPath = ''
$script:EditorRemediationPath = ''
$script:IsImportExcelAvailable = $false
$script:ImportExcelSupportChecked = $false
$script:WorkspaceRoot = Split-Path -Path $PSScriptRoot -Parent
$script:DefaultRemediationRoot = Join-Path -Path $script:WorkspaceRoot -ChildPath 'Packages'
$script:LogDirectory = Join-Path -Path $PSScriptRoot -ChildPath 'Logs'
$script:LogPath = Join-Path -Path $script:LogDirectory -ChildPath 'SmartM365-IntuneRemediation-GUI.log'
$script:MaxLogFileCount = 10
$script:ConfigPath = Join-Path -Path $script:WorkspaceRoot -ChildPath 'SmartM365-IntuneRemediation-GUI.config.json'
$script:ConfiguredRemediationRoot = $script:DefaultRemediationRoot
$script:PublishSourceNamePrefix = 'SmartM365-'
$script:PublishTargetNamePrefix = 'SmartM365-'
$script:GroupDisplayNameCache = @{}
$script:PSScriptAnalyzerExcludeRules = @(
    'PSUseBOMForUnicodeEncodedFile'
)
$script:Ui = @{}

function Initialize-GuiLogRotation {
    [CmdletBinding()]
    param()

    try {
        $logDirectory = Split-Path -Path $script:LogPath -Parent
        if ([string]::IsNullOrWhiteSpace($logDirectory)) {
            return
        }

        [IO.Directory]::CreateDirectory($logDirectory) | Out-Null

        if (Test-Path -LiteralPath $script:LogPath -PathType Leaf) {
            $currentLog = Get-Item -LiteralPath $script:LogPath -ErrorAction Stop
            if ($currentLog.Length -gt 0) {
                $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
                $archivePath = Join-Path -Path $logDirectory -ChildPath ('SmartM365-IntuneRemediation-GUI-{0}.log' -f $timestamp)
                if (Test-Path -LiteralPath $archivePath -PathType Leaf) {
                    $suffix = ([guid]::NewGuid().ToString('N')).Substring(0, 8)
                    $archivePath = Join-Path -Path $logDirectory -ChildPath ('SmartM365-IntuneRemediation-GUI-{0}-{1}.log' -f $timestamp, $suffix)
                }

                Move-Item -LiteralPath $script:LogPath -Destination $archivePath -Force -ErrorAction Stop
            }
        }

        $archiveLimit = [Math]::Max(0, $script:MaxLogFileCount - 1)
        $oldArchives = @(
            Get-ChildItem -LiteralPath $logDirectory -Filter 'SmartM365-IntuneRemediation-GUI-*.log' -File -ErrorAction SilentlyContinue |
                Sort-Object -Property LastWriteTime -Descending |
                Select-Object -Skip $archiveLimit
        )

        foreach ($oldArchive in $oldArchives) {
            Remove-Item -LiteralPath $oldArchive.FullName -Force -ErrorAction Stop
        }
    }
    catch {
        Write-Verbose "Unable to rotate GUI log file '$script:LogPath': $($_.Exception.Message)"
    }
}

Initialize-GuiLogRotation

function Write-GuiLog {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Message)

    $line = '[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $fileLines = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { '{0} {1}' -f $timestamp, $_ })
    if ($script:Ui.ContainsKey('LogTextBox') -and $null -ne $script:Ui.LogTextBox) {
        $script:Ui.LogTextBox.AppendText("$line`r`n")
        $script:Ui.LogTextBox.ScrollToEnd()
    }
    else {
        Write-Host $line
    }

    try {
        $logDirectory = Split-Path -Path $script:LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDirectory)) {
            [IO.Directory]::CreateDirectory($logDirectory) | Out-Null
        }

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [IO.File]::AppendAllText($script:LogPath, (($fileLines -join "`r`n") + "`r`n"), $utf8NoBom)
    }
    catch {
        Write-Verbose "Unable to write GUI log file '$script:LogPath': $($_.Exception.Message)"
    }
}

function Show-GuiError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message
    )

    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Error') | Out-Null
    Write-GuiLog "$Title - $Message"
}

function Get-GuiGraphConnectionFailureMessage {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $message = $ErrorRecord.Exception.Message
    if ([string]::IsNullOrWhiteSpace($message)) {
        $message = 'Microsoft Graph sign-in did not complete.'
    }

    if ($message -match 'User canceled authentication|authentication.*canceled|timed out|timeout|browser') {
        return @"
$message

The Microsoft Graph browser sign-in was canceled or did not complete. Click Connect Graph again, keep the browser tab open, finish sign-in and consent, then return to the manager.

If the browser window does not open or the localhost redirect is blocked, launch the GUI from a visible PowerShell window with:
pwsh -STA -NoProfile -File .\Intune\Remediation\GUI\SmartM365-IntuneRemediation-GUI.ps1 -GraphAuthMode DeviceCode
"@
    }

    return $message
}

function Set-GuiButtonEnabled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][bool]$Enabled
    )

    if ($script:Ui.ContainsKey($Name) -and $null -ne $script:Ui[$Name]) {
        $script:Ui[$Name].IsEnabled = $Enabled
    }
}

function Set-GuiButtonToolTip {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()][string]$Text
    )

    if ($script:Ui.ContainsKey($Name) -and $null -ne $script:Ui[$Name]) {
        $script:Ui[$Name].ToolTip = $Text
    }
}

function Update-ActionButtonsState {
    [CmdletBinding()]
    param()

    $hasLocalPackage = $null -ne $script:SelectedLocalPackage
    $hasCloudRemediation = $null -ne $script:SelectedCloudRemediation
    $hasLocalRoot = -not [string]::IsNullOrWhiteSpace($script:ConfiguredRemediationRoot)

    Set-GuiButtonEnabled -Name 'AnalyzeLocalButton' -Enabled $hasLocalPackage
    Set-GuiButtonEnabled -Name 'ChooseLocalRootButton' -Enabled $true
    Set-GuiButtonEnabled -Name 'SaveLocalButton' -Enabled $hasLocalPackage
    Set-GuiButtonEnabled -Name 'SaveAllButton' -Enabled $hasLocalRoot
    Set-GuiButtonEnabled -Name 'PublishButton' -Enabled ($hasLocalPackage -and $script:IsGraphConnected)
    Set-GuiButtonEnabled -Name 'PublishDetectionOnlyButton' -Enabled ($hasLocalPackage -and $script:IsGraphConnected)
    Set-GuiButtonEnabled -Name 'PublishAllButton' -Enabled ($hasLocalRoot -and $script:IsGraphConnected)

    Set-GuiButtonEnabled -Name 'SaveCloudScriptsButton' -Enabled ($hasCloudRemediation -and $script:IsGraphConnected)
    Set-GuiButtonEnabled -Name 'AnalyzeCloudButton' -Enabled ($hasCloudRemediation -and $script:IsGraphConnected)
    Set-GuiButtonEnabled -Name 'SaveAllCloudButton' -Enabled $script:IsGraphConnected
    Set-GuiButtonEnabled -Name 'CompareLocalCloudButton' -Enabled ($hasLocalPackage -and $hasCloudRemediation -and $script:IsGraphConnected)
    Set-GuiButtonEnabled -Name 'ResetHistoryButton' -Enabled ($hasCloudRemediation -and $script:IsGraphConnected)
    Set-GuiButtonEnabled -Name 'DeleteCloudButton' -Enabled ($hasCloudRemediation -and $script:IsGraphConnected)
    Set-GuiButtonEnabled -Name 'ExportReportButton' -Enabled ($hasCloudRemediation -and $script:IsGraphConnected)

    if (-not $script:IsGraphConnected) {
        $connectMessage = 'Click Connect Graph first to use Intune cloud actions.'
        foreach ($buttonName in @(
                'PublishButton',
                'PublishDetectionOnlyButton',
                'PublishAllButton',
                'SaveCloudScriptsButton',
                'AnalyzeCloudButton',
                'SaveAllCloudButton',
                'CompareLocalCloudButton',
                'ResetHistoryButton',
                'DeleteCloudButton',
                'ExportReportButton'
            )) {
            Set-GuiButtonToolTip -Name $buttonName -Text $connectMessage
        }
        return
    }

    foreach ($buttonName in @(
            'PublishButton',
            'PublishDetectionOnlyButton',
            'PublishAllButton',
            'SaveCloudScriptsButton',
            'AnalyzeCloudButton',
            'SaveAllCloudButton',
            'CompareLocalCloudButton',
            'ResetHistoryButton',
            'DeleteCloudButton',
            'ExportReportButton'
        )) {
        Set-GuiButtonToolTip -Name $buttonName -Text $null
    }
}

function Update-ConnectionStatus {
    [CmdletBinding()]
    param()

    if (-not $script:Ui.ContainsKey('StatusText') -or $null -eq $script:Ui.StatusText) {
        return
    }

    if ($script:IsGraphConnected -and $null -ne $script:GraphContext) {
        $account = if ([string]::IsNullOrWhiteSpace($script:GraphContext.Account)) { 'interactive account' } else { [string]$script:GraphContext.Account }
        $tenant = if ([string]::IsNullOrWhiteSpace($script:GraphContext.TenantId)) { 'tenant connected' } else { [string]$script:GraphContext.TenantId }
        $script:Ui.StatusText.Text = "Interactive Graph | $GraphApiVersion`r`nConnected: $account`r`nTenant: $tenant"
        return
    }

    $script:Ui.StatusText.Text = "Interactive Graph | $GraphApiVersion`r`nNot connected`r`nClick Connect Graph first for Intune actions"
}

function Set-GuiBusy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][bool]$IsBusy,
        [string]$Message = 'Working...'
    )

    if (-not $script:Ui.ContainsKey('BusyOverlay') -or $null -eq $script:Ui.BusyOverlay) {
        return
    }

    if ($IsBusy) {
        $script:Ui.BusyMessage.Text = $Message
        $script:Ui.BusyOverlay.Visibility = [System.Windows.Visibility]::Visible
        $script:Ui.BusyOverlay.IsHitTestVisible = $true
        if ($null -ne $script:Ui.BusyProgress) {
            $script:Ui.BusyProgress.IsIndeterminate = $true
        }
        if ($null -ne $window) {
            $window.Cursor = [System.Windows.Input.Cursors]::Wait
        }
    }
    else {
        if ($null -ne $script:Ui.BusyProgress) {
            $script:Ui.BusyProgress.IsIndeterminate = $false
        }
        $script:Ui.BusyOverlay.Visibility = [System.Windows.Visibility]::Collapsed
        $script:Ui.BusyOverlay.IsHitTestVisible = $false
        if ($null -ne $window) {
            $window.Cursor = $null
        }
    }

    try {
        $script:Ui.BusyOverlay.UpdateLayout()
        $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
    }
    catch {
        Write-Verbose "Unable to refresh busy overlay: $($_.Exception.Message)"
    }
}

function Get-DefaultRemediationRoot {
    [CmdletBinding()]
    param()

    if (-not [string]::IsNullOrWhiteSpace($RemediationRoot)) {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RemediationRoot)
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ConfiguredRemediationRoot)) {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($script:ConfiguredRemediationRoot)
    }

    throw 'Local remediation script folder is not configured.'
}

function Import-RequiredGraphModule {
    [CmdletBinding()]
    param()

    $moduleName = 'Microsoft.Graph.Authentication'
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        throw @"
Required module '$moduleName' is not installed.

Install it with:
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

Or install the full SDK with:
Install-Module Microsoft.Graph -Scope CurrentUser
"@
    }

    Import-Module $moduleName -ErrorAction Stop
}

function Connect-GuiGraph {
    [CmdletBinding()]
    param(
        [string[]]$Scopes = $script:BaseScopes
    )

    Import-RequiredGraphModule

    $script:IsGraphConnected = $false
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    $connectParams = @{
        Scopes       = @($Scopes | Select-Object -Unique)
        ContextScope = 'Process'
        NoWelcome    = $true
        ErrorAction  = 'Stop'
    }

    if ($script:GraphAuthMode -eq 'DeviceCode') {
        $connectParams.UseDeviceCode = $true
    }

    Write-GuiLog ("Connecting to Microsoft Graph using {0} with scopes: {1}" -f $script:GraphAuthMode, (($connectParams.Scopes) -join ', '))
    Connect-MgGraph @connectParams | Out-Null
    $script:GraphContext = Get-MgContext
    if ($null -eq $script:GraphContext) {
        throw 'Microsoft Graph connection failed. Get-MgContext returned no context.'
    }

    $account = if ([string]::IsNullOrWhiteSpace($script:GraphContext.Account)) { 'interactive account' } else { $script:GraphContext.Account }
    Write-GuiLog ("Connected to tenant {0} as {1}" -f $script:GraphContext.TenantId, $account)
    $script:IsGraphConnected = $true
    Update-ActionButtonsState
    Update-ConnectionStatus
}

function Ensure-GuiGraphConnection {
    [CmdletBinding()]
    param(
        [string[]]$Scopes = $script:BaseScopes
    )

    if ($script:IsGraphConnected -and $null -ne $script:GraphContext) {
        return
    }

    $context = $null
    try { $context = Get-MgContext -ErrorAction SilentlyContinue } catch {}
    if ($null -eq $context) {
        Connect-GuiGraph -Scopes $Scopes
        return
    }

    $existingScopes = @($context.Scopes)
    $missingScopes = @($Scopes | Where-Object { $_ -notin $existingScopes })
    if ($missingScopes.Count -gt 0) {
        $mergedScopes = @($existingScopes + $Scopes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        Write-GuiLog ("Additional consent required for: {0}" -f ($missingScopes -join ', '))
        Connect-GuiGraph -Scopes $mergedScopes
        return
    }

    $script:GraphContext = $context
    $script:IsGraphConnected = $true
    Update-ActionButtonsState
    Update-ConnectionStatus
}

function Get-ObjectValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject[$Name] }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Get-IntValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return 0
    }

    try {
        return [int]$Value
    }
    catch {
        return 0
    }
}

function Invoke-GraphGetAllPages {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Uri)

    $items = @()
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop
        $value = Get-ObjectValue -InputObject $page -Name 'value'
        if ($null -ne $value) {
            foreach ($item in @($value)) { $items += $item }
        }
        $nextUri = Get-ObjectValue -InputObject $page -Name '@odata.nextLink'
    }
    return $items
}

function Invoke-GraphJsonRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('POST', 'PATCH')]
        [string]$Method,

        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)]$Body
    )

    $json = $Body | ConvertTo-Json -Depth 30
    try {
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $json -ContentType 'application/json' -ErrorAction Stop
    }
    catch {
        $details = ''
        if ($null -ne $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
            $details = $_.ErrorDetails.Message
        }
        elseif (-not [string]::IsNullOrWhiteSpace($_.Exception.Message)) {
            $details = $_.Exception.Message
        }

        Write-GuiLog ("Graph {0} failed: {1}" -f $Method, $Uri)
        if (-not [string]::IsNullOrWhiteSpace($details)) {
            Write-GuiLog ("Graph error detail: {0}" -f ($details -replace '\s+', ' ').Trim())
        }
        throw
    }
}

function ConvertFrom-GraphScriptContent {
    [CmdletBinding()]
    param([AllowNull()]$Content)

    if ($null -eq $Content) { return '' }
    if ($Content -is [byte[]]) { return [Text.Encoding]::UTF8.GetString($Content) }

    $contentString = [string]$Content
    if ([string]::IsNullOrWhiteSpace($contentString)) { return '' }

    try {
        $bytes = [Convert]::FromBase64String($contentString)
        return [Text.Encoding]::UTF8.GetString($bytes)
    }
    catch {
        return $contentString
    }
}

function ConvertTo-GraphBinaryString {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $bytes = [System.IO.File]::ReadAllBytes($LiteralPath)
    return [Convert]::ToBase64String($bytes)
}

function ConvertTo-SafeFileName {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $safe = $Value
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$char, '_')
    }
    $safe = ($safe -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'UnnamedRemediation' }
    if ($safe.Length -gt 120) { return $safe.Substring(0, 120).Trim() }
    return $safe
}

function Get-UserDocumentsPath {
    [CmdletBinding()]
    param()

    $documentsPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    if (-not [string]::IsNullOrWhiteSpace($documentsPath) -and
        (Test-Path -LiteralPath $documentsPath -PathType Container)) {
        return $documentsPath
    }

    $profileDocuments = Join-Path -Path $HOME -ChildPath 'Documents'
    if (Test-Path -LiteralPath $profileDocuments -PathType Container) {
        return $profileDocuments
    }

    return $HOME
}

function Get-ExecutionReportSavePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$Timestamp,
        [Parameter(Mandatory = $true)][ValidateSet('CSV', 'Excel')][string]$Format
    )

    $safeName = ConvertTo-SafeFileName -Value $DisplayName
    $extension = if ($Format -eq 'Excel') { '.xlsx' } else { '.csv' }
    $filter = if ($Format -eq 'Excel') { 'Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*' } else { 'CSV files (*.csv)|*.csv|All files (*.*)|*.*' }
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Title = 'Save Intune remediation execution report'
    $dialog.Filter = $filter
    $dialog.FileName = "$safeName-ExecutionReport-$Timestamp$extension"
    $dialog.InitialDirectory = Get-UserDocumentsPath
    $dialog.DefaultExt = $extension
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $true

    $result = $dialog.ShowDialog()
    if ($result -ne $true) { return '' }

    return $dialog.FileName
}

function Show-ExecutionReportFormatDialog {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][bool]$ExcelAvailable)

    $formatWindow = New-Object System.Windows.Window
    $formatWindow.Title = 'Export execution report'
    $formatWindow.Width = 360
    $formatWindow.Height = 190
    $formatWindow.WindowStartupLocation = 'CenterOwner'
    $formatWindow.ResizeMode = 'NoResize'
    $formatWindow.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(246, 250, 253))
    if ($script:Ui.ContainsKey('Window') -and $null -ne $script:Ui.Window) {
        $formatWindow.Owner = $script:Ui.Window
    }

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = [System.Windows.Thickness]::new(18)
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))

    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = if ($ExcelAvailable) {
        'Choose the export format for the execution report.'
    }
    else {
        'Excel export is unavailable because ImportExcel could not be installed. CSV export is available.'
    }
    $label.TextWrapping = 'Wrap'
    $label.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(15, 36, 55))
    $label.FontSize = 14
    $label.FontWeight = 'SemiBold'
    $label.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($label, 0)
    $grid.Children.Add($label) | Out-Null

    $buttons = New-Object System.Windows.Controls.StackPanel
    $buttons.Orientation = 'Horizontal'
    $buttons.HorizontalAlignment = 'Right'
    $buttons.VerticalAlignment = 'Bottom'
    [System.Windows.Controls.Grid]::SetRow($buttons, 2)

    $selectedFormat = ''
    $buttonDefinitions = @()
    if ($ExcelAvailable) {
        $buttonDefinitions += @{ Label = 'Excel'; Value = 'Excel'; Width = 88 }
    }
    $buttonDefinitions += @{ Label = 'CSV'; Value = 'CSV'; Width = 88 }
    $buttonDefinitions += @{ Label = 'Cancel'; Value = ''; Width = 88 }

    foreach ($buttonDefinition in $buttonDefinitions) {
        $button = New-Object System.Windows.Controls.Button
        $button.Content = $buttonDefinition.Label
        $button.Width = $buttonDefinition.Width
        $button.Height = 34
        $button.Margin = [System.Windows.Thickness]::new(8, 0, 0, 0)
        $button.Tag = $buttonDefinition.Value
        $button.FontSize = 12
        $button.FontWeight = 'SemiBold'
        $button.Cursor = [System.Windows.Input.Cursors]::Hand
        $button.Padding = [System.Windows.Thickness]::new(12, 4, 12, 4)
        if ([string]$buttonDefinition.Value -eq 'Excel') {
            $button.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 132, 184))
            $button.Foreground = [System.Windows.Media.Brushes]::White
            $button.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 132, 184))
        }
        else {
            $button.Background = [System.Windows.Media.Brushes]::White
            $button.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(15, 36, 55))
            $button.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(190, 207, 220))
        }
        $button.Add_Click({
                $script:ExecutionReportFormatDialogResult = [string]$this.Tag
                $formatWindow.DialogResult = $true
                $formatWindow.Close()
            })
        $buttons.Children.Add($button) | Out-Null
    }

    $grid.Children.Add($buttons) | Out-Null
    $formatWindow.Content = $grid
    $script:ExecutionReportFormatDialogResult = ''
    [void]$formatWindow.ShowDialog()
    $selectedFormat = [string]$script:ExecutionReportFormatDialogResult
    $script:ExecutionReportFormatDialogResult = ''
    return $selectedFormat
}

function Get-SaveAllArchivePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][string]$Timestamp,
        [Parameter(Mandatory = $true)][ValidateSet('Local', 'Cloud')][string]$Scope
    )

    $safeTenantName = ConvertTo-SafeFileName -Value $TenantName
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Title = "Save all SmartM365 $Scope remediations"
    $dialog.Filter = 'ZIP files (*.zip)|*.zip|All files (*.*)|*.*'
    $dialog.FileName = "SmartM365-IntuneRemediation-$Scope-$safeTenantName-$Timestamp.zip"
    $dialog.InitialDirectory = Get-UserDocumentsPath
    $dialog.DefaultExt = '.zip'
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $true

    $result = $dialog.ShowDialog()
    if ($result -ne $true) { return '' }

    return $dialog.FileName
}

function Get-CloudScriptCopyFolderPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$DisplayName)

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select where to save script copies for $DisplayName"
    $dialog.ShowNewFolderButton = $true
    $dialog.SelectedPath = Get-UserDocumentsPath

    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return ''
    }

    return $dialog.SelectedPath
}

function Save-Utf8NoBom {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowNull()][string]$Content
    )

    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $parentPath = Split-Path -Path $fullPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parentPath)) {
        [IO.Directory]::CreateDirectory($parentPath) | Out-Null
    }

    if ($null -eq $Content) { $Content = '' }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($fullPath, $Content, $utf8NoBom)
}

function Initialize-GuiConfigurationFromTemplate {
    [CmdletBinding()]
    param()

    if (Test-Path -LiteralPath $script:ConfigPath) { return $false }

    $templatePath = Join-Path -Path $script:WorkspaceRoot -ChildPath 'SmartM365-IntuneRemediation-GUI.config.template.json'
    if (-not (Test-Path -LiteralPath $templatePath)) { return $false }

    Copy-Item -LiteralPath $templatePath -Destination $script:ConfigPath -ErrorAction Stop
    Write-GuiLog "Created GUI local configuration from template: $script:ConfigPath"
    Write-Host ((@(
        'Created Intune Remediation GUI local JSON from template.',
        "Local JSON: $script:ConfigPath",
        "Template: $templatePath",
        'Review the generated local JSON values; continuing with default template values unless edited before next run.'
    )) -join [Environment]::NewLine) -ForegroundColor Yellow
    return $true
}

function Read-GuiConfiguration {
    [CmdletBinding()]
    param()

    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        Initialize-GuiConfigurationFromTemplate | Out-Null
    }

    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        return
    }

    try {
        $config = Get-Content -LiteralPath $script:ConfigPath -Raw | ConvertFrom-Json -ErrorAction Stop
        $root = [string](Get-ObjectValue -InputObject $config -Name 'LocalRemediationRoot')

        $sourceNamePrefix = Get-ObjectValue -InputObject $config -Name 'PublishSourceNamePrefix'
        if ($null -ne $sourceNamePrefix -and -not [string]::IsNullOrWhiteSpace([string]$sourceNamePrefix)) {
            $script:PublishSourceNamePrefix = [string]$sourceNamePrefix
        }

        $targetNamePrefix = Get-ObjectValue -InputObject $config -Name 'PublishTargetNamePrefix'
        if ($null -ne $targetNamePrefix) {
            $script:PublishTargetNamePrefix = [string]$targetNamePrefix
        }

        Write-GuiLog ("Publish name prefix mapping loaded: '{0}' -> '{1}'" -f $script:PublishSourceNamePrefix, $script:PublishTargetNamePrefix)

        if ([string]::IsNullOrWhiteSpace($root)) {
            return
        }

        $configRoot = if ([System.IO.Path]::IsPathRooted($root)) {
            $root
        }
        else {
            Join-Path -Path $script:WorkspaceRoot -ChildPath $root
        }
        $resolvedRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($configRoot)
        if (Test-Path -LiteralPath $resolvedRoot -PathType Container) {
            $script:ConfiguredRemediationRoot = $resolvedRoot
            Write-GuiLog "Configured local script folder loaded: $resolvedRoot"
        }
        else {
            Write-GuiLog "Configured LocalRemediationRoot not found: $resolvedRoot"
        }
    }
    catch {
        Write-GuiLog "Failed to load GUI configuration: $($_.Exception.Message)" -Level WARN
    }
}
function Save-GuiConfiguration {
    [CmdletBinding()]
    param()

    $localRootValue = if ($script:ConfiguredRemediationRoot -eq $script:DefaultRemediationRoot) {
        'Packages'
    }
    else {
        $script:ConfiguredRemediationRoot
    }

    $config = [ordered]@{
        LocalRemediationRoot     = $localRootValue
        PublishSourceNamePrefix  = $script:PublishSourceNamePrefix
        PublishTargetNamePrefix  = $script:PublishTargetNamePrefix
    }

    Save-Utf8NoBom -Path $script:ConfigPath -Content ($config | ConvertTo-Json -Depth 5)
    Write-GuiLog "Local script folder saved in GUI configuration: $script:ConfiguredRemediationRoot"
}

function Select-LocalRemediationRoot {
    [CmdletBinding()]
    param([switch]$Required)

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select the local folder that contains SmartM365 Intune remediation scripts'
    $dialog.ShowNewFolderButton = $false
    if (-not [string]::IsNullOrWhiteSpace($script:ConfiguredRemediationRoot) -and
        (Test-Path -LiteralPath $script:ConfiguredRemediationRoot -PathType Container)) {
        $dialog.SelectedPath = $script:ConfiguredRemediationRoot
    }

    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        if ($Required) {
            Write-GuiLog 'Local script folder selection cancelled. Local packages are not loaded.'
        }
        return ''
    }

    $selectedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($dialog.SelectedPath)
    if (-not (Test-Path -LiteralPath $selectedPath -PathType Container)) {
        throw "Selected local script folder does not exist: $selectedPath"
    }

    $script:ConfiguredRemediationRoot = $selectedPath
    Save-GuiConfiguration
    Update-ActionButtonsState
    return $selectedPath
}

function Ensure-LocalRemediationRootConfigured {
    [CmdletBinding()]
    param([switch]$PromptIfMissing)

    if (-not [string]::IsNullOrWhiteSpace($script:ConfiguredRemediationRoot) -and
        (Test-Path -LiteralPath $script:ConfiguredRemediationRoot -PathType Container)) {
        return $script:ConfiguredRemediationRoot
    }

    $script:ConfiguredRemediationRoot = ''
    Update-ActionButtonsState
    if (-not $PromptIfMissing) {
        return ''
    }

    return Select-LocalRemediationRoot -Required
}

function Initialize-LocalRemediationRootConfiguration {
    [CmdletBinding()]
    param()

    Read-GuiConfiguration

    if ([string]::IsNullOrWhiteSpace($RemediationRoot)) {
        return
    }

    $resolvedRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RemediationRoot)
    if (-not (Test-Path -LiteralPath $resolvedRoot -PathType Container)) {
        Write-GuiLog "Command line local script folder does not exist: $resolvedRoot"
        return
    }

    $script:ConfiguredRemediationRoot = $resolvedRoot
    Save-GuiConfiguration
}

function Ensure-PSScriptAnalyzer {
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
        Write-GuiLog 'PSScriptAnalyzer not found. Installing for current user...'
        Install-Module -Name PSScriptAnalyzer -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
        Write-GuiLog 'PSScriptAnalyzer installed.'
    }

    Import-Module PSScriptAnalyzer -ErrorAction Stop
}

function Invoke-PSScriptAnalyzerForContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ScriptName,
        [AllowEmptyString()][string]$Content
    )

    if ([string]::IsNullOrWhiteSpace($Content)) {
        return @()
    }

    Ensure-PSScriptAnalyzer

    $safeName = ConvertTo-SafeFileName -Value $ScriptName
    $tempPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365-PSScriptAnalyzer-{0}-{1}.ps1" -f ([guid]::NewGuid().ToString('N')), $safeName)
    try {
        Save-Utf8NoBom -Path $tempPath -Content $Content
        $findings = @(Invoke-ScriptAnalyzer -Path $tempPath -Severity Error, Warning -ExcludeRule $script:PSScriptAnalyzerExcludeRules -ErrorAction Stop)
        foreach ($finding in $findings) {
            [pscustomobject]@{
                ScriptName = $ScriptName
                Severity   = [string]$finding.Severity
                RuleName   = [string]$finding.RuleName
                Line       = [int]$finding.Line
                Message    = [string]$finding.Message
            }
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Format-PSScriptAnalyzerFindings {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Findings)

    $lines = @()
    foreach ($finding in @($Findings | Select-Object -First 12)) {
        $lines += ('{0} [{1}] line {2}: {3} - {4}' -f $finding.ScriptName, $finding.Severity, $finding.Line, $finding.RuleName, $finding.Message)
    }

    if (@($Findings).Count -gt 12) {
        $lines += ('...and {0} more finding(s).' -f (@($Findings).Count - 12))
    }

    return ($lines -join "`r`n")
}

function Set-ObjectNoteProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        $property.Value = $Value
        return
    }

    Add-Member -InputObject $InputObject -MemberType NoteProperty -Name $Name -Value $Value -Force
}

function Get-NormalizedScriptContent {
    [CmdletBinding()]
    param([AllowNull()][string]$Content)

    if ($null -eq $Content) {
        return ''
    }

    return (@(ConvertTo-CompareLines -Content ([string]$Content)) -join "`n")
}

function Get-ContentSha256 {
    [CmdletBinding()]
    param([AllowNull()][string]$Content)

    $normalizedContent = Get-NormalizedScriptContent -Content $Content
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalizedContent)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return [System.BitConverter]::ToString($sha256.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-ScriptContentSignature {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$DetectionContent,
        [AllowNull()][string]$RemediationContent
    )

    return '{0}|{1}' -f (Get-ContentSha256 -Content $DetectionContent), (Get-ContentSha256 -Content $RemediationContent)
}

function Test-ScriptContentEqual {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$LeftContent,
        [AllowNull()][string]$RightContent
    )

    $leftLines = @(ConvertTo-CompareLines -Content $LeftContent)
    $rightLines = @(ConvertTo-CompareLines -Content $RightContent)
    if ($leftLines.Count -ne $rightLines.Count) {
        return $false
    }

    for ($i = 0; $i -lt $leftLines.Count; $i++) {
        if ($leftLines[$i] -ne $rightLines[$i]) {
            return $false
        }
    }

    return $true
}

function Get-LocalPackageContentSignature {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Package)

    $detectionContent = if (-not [string]::IsNullOrWhiteSpace([string]$Package.DetectionPath)) {
        [IO.File]::ReadAllText([string]$Package.DetectionPath)
    }
    else {
        ''
    }

    $remediationContent = if (-not [string]::IsNullOrWhiteSpace([string]$Package.RemediationPath)) {
        [IO.File]::ReadAllText([string]$Package.RemediationPath)
    }
    else {
        ''
    }

    return Get-ScriptContentSignature -DetectionContent $detectionContent -RemediationContent $remediationContent
}

function Get-LocalPackageScriptContent {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Package)

    $detectionContent = if (-not [string]::IsNullOrWhiteSpace([string]$Package.DetectionPath)) {
        [IO.File]::ReadAllText([string]$Package.DetectionPath)
    }
    else {
        ''
    }

    $remediationContent = if (-not [string]::IsNullOrWhiteSpace([string]$Package.RemediationPath)) {
        [IO.File]::ReadAllText([string]$Package.RemediationPath)
    }
    else {
        ''
    }

    return [pscustomobject]@{
        DetectionContent   = $detectionContent
        RemediationContent = $remediationContent
    }
}

function Set-CloudRemediationContentSignature {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$CloudRemediation,
        [Parameter(Mandatory = $true)]$Detail
    )

    $detectionContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $Detail -Name 'detectionScriptContent')
    $remediationContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $Detail -Name 'remediationScriptContent')

    Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'DetectionContentHash' -Value (Get-ContentSha256 -Content $detectionContent)
    Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'RemediationContentHash' -Value (Get-ContentSha256 -Content $remediationContent)
    Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'ContentSignature' -Value (Get-ScriptContentSignature -DetectionContent $detectionContent -RemediationContent $remediationContent)
    Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'DetectionNormalizedContent' -Value (Get-NormalizedScriptContent -Content $detectionContent)
    Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'RemediationNormalizedContent' -Value (Get-NormalizedScriptContent -Content $remediationContent)
    Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'ContentSignatureAvailable' -Value $true
    Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'ContentSignatureDetails' -Value 'Cloud content signature loaded.'
}

function Ensure-CloudRemediationContentSignature {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$CloudRemediation)

    if ($null -eq $CloudRemediation) { return $false }

    $available = $false
    try { $available = [bool]$CloudRemediation.ContentSignatureAvailable } catch {}
    if ($available -and -not [string]::IsNullOrWhiteSpace([string]$CloudRemediation.ContentSignature)) {
        return $true
    }

    try {
        $detail = Get-CloudRemediationDetail -Id ([string]$CloudRemediation.Id)
        Set-CloudRemediationContentSignature -CloudRemediation $CloudRemediation -Detail $detail
        return $true
    }
    catch {
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'ContentSignatureAvailable' -Value $false
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'ContentSignatureDetails' -Value $_.Exception.Message
        return $false
    }
}

function Get-PSScriptAnalyzerStatusFromFindings {
    [CmdletBinding()]
    param([AllowNull()]$Findings)

    $findingsList = @($Findings)
    if (@($findingsList | Where-Object { $_.Severity -eq 'Error' }).Count -gt 0) {
        return 'Error'
    }

    if (@($findingsList | Where-Object { $_.Severity -eq 'Warning' }).Count -gt 0) {
        return 'Warning'
    }

    return 'OK'
}

function Get-PSScriptAnalyzerStatusSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Items)

    $itemsList = @($Items)
    $ok = @($itemsList | Where-Object { $_.PSScriptAnalyzerStatus -eq 'OK' }).Count
    $warning = @($itemsList | Where-Object { $_.PSScriptAnalyzerStatus -eq 'Warning' }).Count
    $errorCount = @($itemsList | Where-Object { $_.PSScriptAnalyzerStatus -eq 'Error' }).Count
    $notTested = @($itemsList | Where-Object { $_.PSScriptAnalyzerStatus -eq 'Not tested' }).Count
    $cloudDetailUnavailable = @($itemsList | Where-Object { $_.PSScriptAnalyzerStatus -eq 'Cloud detail unavailable' }).Count
    $summary = "OK=$ok, Warning=$warning, Error=$errorCount, Not tested=$notTested"
    if ($cloudDetailUnavailable -gt 0) {
        $summary = "$summary, Cloud detail unavailable=$cloudDetailUnavailable"
    }

    return $summary
}

function Test-GraphNotFoundException {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Exception)

    $message = [string]$Exception.Message
    if ($message -match '\b404\b' -or $message -match 'NotFound' -or $message -match 'Not Found') {
        return $true
    }

    try {
        $responseStatusCode = $Exception.ResponseStatusCode
        if ($null -ne $responseStatusCode -and [string]$responseStatusCode -match 'NotFound|404') {
            return $true
        }
    }
    catch {}

    try {
        if ($null -ne $Exception.Response -and [int]$Exception.Response.StatusCode -eq 404) {
            return $true
        }
    }
    catch {}

    return $false
}

function Set-PSScriptAnalyzerUnavailableStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Items,
        [Parameter(Mandatory = $true)][string]$Message
    )

    foreach ($item in @($Items)) {
        Set-ObjectNoteProperty -InputObject $item -Name 'PSScriptAnalyzerStatus' -Value 'Error'
        Set-ObjectNoteProperty -InputObject $item -Name 'PSScriptAnalyzerDetails' -Value $Message
    }
}

function Update-LocalPackagePSScriptAnalyzerStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Package)

    $packageName = [string]$Package.DisplayName
    $scriptNames = @()
    try {
        $findings = @()
        $detectionPath = [string]$Package.DetectionPath
        if (-not [string]::IsNullOrWhiteSpace($detectionPath)) {
            $detectionName = [IO.Path]::GetFileName($detectionPath)
            $scriptNames += $detectionName
            $findings += Invoke-PSScriptAnalyzerForContent `
                -ScriptName $detectionName `
                -Content ([IO.File]::ReadAllText($detectionPath))
        }

        $remediationPath = [string]$Package.RemediationPath
        if (-not [string]::IsNullOrWhiteSpace($remediationPath)) {
            $remediationName = [IO.Path]::GetFileName($remediationPath)
            $scriptNames += $remediationName
            $findings += Invoke-PSScriptAnalyzerForContent `
                -ScriptName $remediationName `
                -Content ([IO.File]::ReadAllText($remediationPath))
        }

        $status = Get-PSScriptAnalyzerStatusFromFindings -Findings $findings
        $details = if (@($findings).Count -gt 0) {
            Format-PSScriptAnalyzerFindings -Findings $findings
        }
        else {
            "OK - $($scriptNames -join ', ')"
        }

        Set-ObjectNoteProperty -InputObject $Package -Name 'PSScriptAnalyzerStatus' -Value $status
        Set-ObjectNoteProperty -InputObject $Package -Name 'PSScriptAnalyzerDetails' -Value $details

        if ($status -ne 'OK') {
            Write-GuiLog ("Local PSScriptAnalyzer {0} for package '{1}' ({2})" -f $status, $packageName, ($scriptNames -join ', '))
        }
    }
    catch {
        Set-ObjectNoteProperty -InputObject $Package -Name 'PSScriptAnalyzerStatus' -Value 'Error'
        Set-ObjectNoteProperty -InputObject $Package -Name 'PSScriptAnalyzerDetails' -Value $_.Exception.Message
        Write-GuiLog ("Local PSScriptAnalyzer Error for package '{0}' - {1}" -f $packageName, $_.Exception.Message)
    }
}

function Update-CloudRemediationPSScriptAnalyzerStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$CloudRemediation)

    $displayName = [string]$CloudRemediation.DisplayName
    try {
        $detail = Get-CloudRemediationDetail -Id ([string]$CloudRemediation.Id)
        $detectionScriptName = "$displayName-Cloud-Detection.ps1"
        $remediationScriptName = "$displayName-Cloud-Remediation.ps1"
        $detectionContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'detectionScriptContent')
        $remediationContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'remediationScriptContent')
        Set-CloudRemediationContentSignature -CloudRemediation $CloudRemediation -Detail $detail

        $findings = @()
        $findings += Invoke-PSScriptAnalyzerForContent -ScriptName $detectionScriptName -Content $detectionContent
        $findings += Invoke-PSScriptAnalyzerForContent -ScriptName $remediationScriptName -Content $remediationContent

        $status = Get-PSScriptAnalyzerStatusFromFindings -Findings $findings
        $details = if (@($findings).Count -gt 0) {
            Format-PSScriptAnalyzerFindings -Findings $findings
        }
        else {
            "OK - $detectionScriptName, $remediationScriptName"
        }

        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PSScriptAnalyzerStatus' -Value $status
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PSScriptAnalyzerDetails' -Value $details

        if ($status -ne 'OK') {
            Write-GuiLog ("Cloud PSScriptAnalyzer {0} for remediation '{1}' ({2}, {3})" -f $status, $displayName, $detectionScriptName, $remediationScriptName)
        }
    }
    catch {
        $message = $_.Exception.Message
        if (Test-GraphNotFoundException -Exception $_.Exception) {
            Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PSScriptAnalyzerStatus' -Value 'Cloud detail unavailable'
            Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PSScriptAnalyzerDetails' -Value "Graph returned NotFound while loading the cloud script content. The remediation may be stale or broken in Intune. Id: $($CloudRemediation.Id). Detail: $message"
            Write-GuiLog ("Cloud detail unavailable for remediation '{0}' ({1}) - {2}" -f $displayName, $CloudRemediation.Id, $message)
        }
        else {
            Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PSScriptAnalyzerStatus' -Value 'Error'
            Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PSScriptAnalyzerDetails' -Value $message
            Write-GuiLog ("Cloud PSScriptAnalyzer Error for remediation '{0}' - {1}" -f $displayName, $message)
        }
    }
}

function Update-LocalPackagesPSScriptAnalyzerStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Packages)

    $packageList = @($Packages)
    if ($packageList.Count -eq 0) { return }

    try {
        Ensure-PSScriptAnalyzer
    }
    catch {
        $message = "PSScriptAnalyzer unavailable - $($_.Exception.Message)"
        Set-PSScriptAnalyzerUnavailableStatus -Items $packageList -Message $message
        Write-GuiLog $message
        return
    }

    for ($i = 0; $i -lt $packageList.Count; $i++) {
        $package = $packageList[$i]
        Set-GuiBusy -IsBusy $true -Message ("Analyzing local scripts {0}/{1}: {2}" -f ($i + 1), $packageList.Count, $package.DisplayName)
        Update-LocalPackagePSScriptAnalyzerStatus -Package $package
    }

    Write-GuiLog ("Local PSScriptAnalyzer status: {0}" -f (Get-PSScriptAnalyzerStatusSummary -Items $packageList))
}

function Update-CloudRemediationsPSScriptAnalyzerStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$CloudRemediations)

    $cloudList = @($CloudRemediations)
    if ($cloudList.Count -eq 0) { return }

    try {
        Ensure-PSScriptAnalyzer
    }
    catch {
        $message = "PSScriptAnalyzer unavailable - $($_.Exception.Message)"
        Set-PSScriptAnalyzerUnavailableStatus -Items $cloudList -Message $message
        Write-GuiLog $message
        return
    }

    for ($i = 0; $i -lt $cloudList.Count; $i++) {
        $cloudRemediation = $cloudList[$i]
        Set-GuiBusy -IsBusy $true -Message ("Analyzing cloud scripts {0}/{1}: {2}" -f ($i + 1), $cloudList.Count, $cloudRemediation.DisplayName)
        Update-CloudRemediationPSScriptAnalyzerStatus -CloudRemediation $cloudRemediation
    }

    Write-GuiLog ("Cloud PSScriptAnalyzer status: {0}" -f (Get-PSScriptAnalyzerStatusSummary -Items $cloudList))
}

function Test-LocalEditorScriptsWithPSScriptAnalyzer {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($script:EditorDetectionPath)) {
        throw 'Select a local package before running PSScriptAnalyzer.'
    }

    $detectionScriptName = [IO.Path]::GetFileName($script:EditorDetectionPath)
    $remediationScriptName = if ([string]::IsNullOrWhiteSpace($script:EditorRemediationPath)) { '' } else { [IO.Path]::GetFileName($script:EditorRemediationPath) }

    $findings = @()
    $findings += Invoke-PSScriptAnalyzerForContent `
        -ScriptName $detectionScriptName `
        -Content $script:Ui.LocalDetectionEditor.Text

    if (-not [string]::IsNullOrWhiteSpace($script:EditorRemediationPath)) {
        $findings += Invoke-PSScriptAnalyzerForContent `
            -ScriptName $remediationScriptName `
            -Content $script:Ui.LocalRemediationEditor.Text
    }

    $errors = @($findings | Where-Object { $_.Severity -eq 'Error' })
    if ($errors.Count -gt 0) {
        $message = Format-PSScriptAnalyzerFindings -Findings $errors
        Show-GuiError -Title 'PSScriptAnalyzer errors' -Message $message
        throw 'PSScriptAnalyzer found blocking errors.'
    }

    $warnings = @($findings | Where-Object { $_.Severity -eq 'Warning' })
    if ($warnings.Count -gt 0) {
        $message = "PSScriptAnalyzer found warning(s):`r`n`r`n$(Format-PSScriptAnalyzerFindings -Findings $warnings)`r`n`r`nContinue anyway?"
        $answer = [System.Windows.MessageBox]::Show($message, 'PSScriptAnalyzer warnings', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') {
            throw 'PSScriptAnalyzer warnings were not approved.'
        }
    }

    $scriptList = @($detectionScriptName, $remediationScriptName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    Write-GuiLog ("PSScriptAnalyzer check completed for local script(s): {0}" -f ($scriptList -join ', '))
}

function Test-LocalEditorDetectionWithPSScriptAnalyzer {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($script:EditorDetectionPath)) {
        throw 'Select a local package before running PSScriptAnalyzer.'
    }

    $detectionScriptName = [IO.Path]::GetFileName($script:EditorDetectionPath)
    $findings = @(Invoke-PSScriptAnalyzerForContent `
            -ScriptName $detectionScriptName `
            -Content $script:Ui.LocalDetectionEditor.Text)

    $errors = @($findings | Where-Object { $_.Severity -eq 'Error' })
    if ($errors.Count -gt 0) {
        $message = Format-PSScriptAnalyzerFindings -Findings $errors
        Show-GuiError -Title 'PSScriptAnalyzer errors' -Message $message
        throw 'PSScriptAnalyzer found blocking errors in the detection script.'
    }

    $warnings = @($findings | Where-Object { $_.Severity -eq 'Warning' })
    if ($warnings.Count -gt 0) {
        $message = "PSScriptAnalyzer found detection warning(s):`r`n`r`n$(Format-PSScriptAnalyzerFindings -Findings $warnings)`r`n`r`nContinue anyway?"
        $answer = [System.Windows.MessageBox]::Show($message, 'PSScriptAnalyzer warnings', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') {
            throw 'PSScriptAnalyzer warnings were not approved.'
        }
    }

    Write-GuiLog ("PSScriptAnalyzer check completed for local detection script: {0}" -f $detectionScriptName)
}

function Test-LocalPackagesWithPSScriptAnalyzer {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Packages)

    $packageList = @($Packages)
    if ($packageList.Count -eq 0) {
        throw 'No local package found for PSScriptAnalyzer.'
    }

    $findings = @()
    foreach ($package in $packageList) {
        $detectionPath = [string]$package.DetectionPath
        if (-not [string]::IsNullOrWhiteSpace($detectionPath)) {
            $findings += Invoke-PSScriptAnalyzerForContent `
                -ScriptName ([IO.Path]::GetFileName($detectionPath)) `
                -Content ([IO.File]::ReadAllText($detectionPath))
        }

        $remediationPath = [string]$package.RemediationPath
        if (-not [string]::IsNullOrWhiteSpace($remediationPath)) {
            $findings += Invoke-PSScriptAnalyzerForContent `
                -ScriptName ([IO.Path]::GetFileName($remediationPath)) `
                -Content ([IO.File]::ReadAllText($remediationPath))
        }
    }

    $errors = @($findings | Where-Object { $_.Severity -eq 'Error' })
    if ($errors.Count -gt 0) {
        $message = Format-PSScriptAnalyzerFindings -Findings $errors
        Show-GuiError -Title 'PSScriptAnalyzer errors' -Message $message
        throw 'PSScriptAnalyzer found blocking errors before publish all.'
    }

    $warnings = @($findings | Where-Object { $_.Severity -eq 'Warning' })
    if ($warnings.Count -gt 0) {
        $message = "PSScriptAnalyzer found warning(s) before publishing all local packages:`r`n`r`n$(Format-PSScriptAnalyzerFindings -Findings $warnings)`r`n`r`nContinue publishing anyway?"
        $answer = [System.Windows.MessageBox]::Show($message, 'PSScriptAnalyzer warnings', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') {
            throw 'PSScriptAnalyzer warnings were not approved.'
        }
    }

    Write-GuiLog ("PSScriptAnalyzer check completed for {0} local package(s)." -f $packageList.Count)
}

function Test-CloudViewerScriptsWithPSScriptAnalyzer {
    [CmdletBinding()]
    param()

    if ($null -eq $script:SelectedCloudRemediation) {
        throw 'Select a cloud remediation before running PSScriptAnalyzer.'
    }

    $displayName = [string]$script:SelectedCloudRemediation.DisplayName
    $detectionScriptName = "$displayName-Cloud-Detection.ps1"
    $remediationScriptName = "$displayName-Cloud-Remediation.ps1"
    $findings = @()
    $findings += Invoke-PSScriptAnalyzerForContent `
        -ScriptName $detectionScriptName `
        -Content $script:Ui.CloudDetectionEditor.Text
    $findings += Invoke-PSScriptAnalyzerForContent `
        -ScriptName $remediationScriptName `
        -Content $script:Ui.CloudRemediationEditor.Text

    $errors = @($findings | Where-Object { $_.Severity -eq 'Error' })
    if ($errors.Count -gt 0) {
        $message = Format-PSScriptAnalyzerFindings -Findings $errors
        Show-GuiError -Title 'Cloud PSScriptAnalyzer errors' -Message $message
        throw 'PSScriptAnalyzer found cloud script errors.'
    }

    $warnings = @($findings | Where-Object { $_.Severity -eq 'Warning' })
    if ($warnings.Count -gt 0) {
        $message = Format-PSScriptAnalyzerFindings -Findings $warnings
        [System.Windows.MessageBox]::Show($message, 'Cloud PSScriptAnalyzer warnings', 'OK', 'Warning') | Out-Null
        Write-GuiLog ("Cloud PSScriptAnalyzer completed for remediation '{0}' with {1} warning(s). Scripts: {2}" -f $displayName, $warnings.Count, (@($detectionScriptName, $remediationScriptName) -join ', '))
        return
    }

    Write-GuiLog ("Cloud PSScriptAnalyzer check completed for remediation '{0}'. Scripts: {1}" -f $displayName, (@($detectionScriptName, $remediationScriptName) -join ', '))
}

function Get-ScriptHelpDescription {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $headerLineCount = 0
    foreach ($line in [System.IO.File]::ReadLines($LiteralPath)) {
        $headerLineCount++
        if ($headerLineCount -gt 80) { break }
        if ($line -match '^\s*#\s*Description\s*:\s*(?<Text>.+?)\s*$') {
            return ($Matches['Text'] -replace '\s+', ' ').Trim()
        }
    }
    return ''
}

function ConvertTo-PackageDisplayName {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$DetectionScriptPath)

    $name = [System.IO.Path]::GetFileNameWithoutExtension($DetectionScriptPath)
    $name = $name -replace '-DetectionOnly$', ''
    $name = $name -replace '-Detection$', ''
    return $name
}

function ConvertTo-PublishedDisplayName {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$DisplayName)

    if ([string]::IsNullOrEmpty($script:PublishSourceNamePrefix)) {
        return $DisplayName
    }

    if ($DisplayName.StartsWith($script:PublishSourceNamePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ('{0}{1}' -f $script:PublishTargetNamePrefix, $DisplayName.Substring($script:PublishSourceNamePrefix.Length))
    }

    return $DisplayName
}

function Get-PackageDescription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$DetectionScriptPath,
        [AllowNull()][string]$RemediationScriptPath
    )

    $detectionDescription = Get-ScriptHelpDescription -LiteralPath $DetectionScriptPath
    $remediationDescription = ''
    if (-not [string]::IsNullOrWhiteSpace($RemediationScriptPath)) {
        $remediationDescription = Get-ScriptHelpDescription -LiteralPath $RemediationScriptPath
    }

    if (-not [string]::IsNullOrWhiteSpace($detectionDescription) -and
        -not [string]::IsNullOrWhiteSpace($remediationDescription) -and
        $detectionDescription -ne $remediationDescription) {
        return ("Detection: {0}`nRemediation: {1}" -f $detectionDescription, $remediationDescription)
    }

    if (-not [string]::IsNullOrWhiteSpace($detectionDescription)) { return $detectionDescription }
    if (-not [string]::IsNullOrWhiteSpace($remediationDescription)) { return $remediationDescription }
    return ''
}

function Get-MatchingLocalRemediationPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$DetectionScriptPath)

    $detectionName = [System.IO.Path]::GetFileName($DetectionScriptPath)
    $folder = Split-Path -Path $DetectionScriptPath -Parent
    $exactCandidate = Join-Path -Path $folder -ChildPath ($detectionName -replace '-Detection\.ps1$', '-Remediation.ps1')
    if (Test-Path -LiteralPath $exactCandidate -PathType Leaf) {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($exactCandidate)
    }

    $remediationCandidates = @(Get-ChildItem -LiteralPath $folder -File -Filter '*-Remediation.ps1' -ErrorAction SilentlyContinue)
    if ($remediationCandidates.Count -eq 1) {
        return $remediationCandidates[0].FullName
    }

    return ''
}

function New-LocalPackageFromDetection {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$DetectionScriptPath)

    $detectionName = [System.IO.Path]::GetFileName($DetectionScriptPath)
    if ($detectionName -notmatch '^.+-(Detection|DetectionOnly)\.ps1$') {
        return $null
    }

    $isDetectionOnly = $detectionName -match '-DetectionOnly\.ps1$'
    $remediationPath = ''
    if (-not $isDetectionOnly) {
        $remediationPath = Get-MatchingLocalRemediationPath -DetectionScriptPath $DetectionScriptPath
    }

    $description = Get-PackageDescription -DetectionScriptPath $DetectionScriptPath -RemediationScriptPath $remediationPath
    if ([string]::IsNullOrWhiteSpace($description)) {
        $description = ConvertTo-PackageDisplayName -DetectionScriptPath $DetectionScriptPath
    }

    $relativeFolder = [System.IO.Path]::GetRelativePath((Get-DefaultRemediationRoot), (Split-Path -Path $DetectionScriptPath -Parent))
    $localPackage = [pscustomobject]@{
        DisplayName     = ConvertTo-PackageDisplayName -DetectionScriptPath $DetectionScriptPath
        Description     = $description
        Folder          = $relativeFolder
        DetectionPath   = $DetectionScriptPath
        RemediationPath = $remediationPath
        DetectionOnly   = ($isDetectionOnly -or [string]::IsNullOrWhiteSpace($remediationPath))
        PublishedDisplayName = ''
        IntuneStatus    = 'Unknown'
        IntuneDetails   = 'Connect Graph to check whether this script content already exists in Intune.'
        CloudDisplayName = ''
        CloudId         = ''
        DetectionContentHash = ''
        RemediationContentHash = ''
        ContentSignature = ''
        PSScriptAnalyzerStatus  = 'Not tested'
        PSScriptAnalyzerDetails = ''
    }
    Set-ObjectNoteProperty -InputObject $localPackage -Name 'ContentSignature' -Value (Get-LocalPackageContentSignature -Package $localPackage)
    return $localPackage
}

function Get-LocalPackages {
    [CmdletBinding()]
    param()

    $root = Get-DefaultRemediationRoot
    if (-not (Test-Path -LiteralPath $root)) {
        throw "Local remediation root not found: $root"
    }

    $packages = @()
    $detectionFiles = @(
        [System.IO.Directory]::EnumerateFiles($root, '*-Detection.ps1', [System.IO.SearchOption]::AllDirectories)
        [System.IO.Directory]::EnumerateFiles($root, '*-DetectionOnly.ps1', [System.IO.SearchOption]::AllDirectories)
    ) | Sort-Object -Unique

    foreach ($file in $detectionFiles) {
        $package = New-LocalPackageFromDetection -DetectionScriptPath $file
        if ($null -ne $package) { $packages += $package }
    }

    return @($packages | Sort-Object Folder, DisplayName)
}

function Update-LocalPackagesIntuneStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Packages,
        [AllowNull()]$CloudRemediations
    )

    $packageList = @($Packages)
    if ($packageList.Count -eq 0) { return }

    foreach ($package in $packageList) {
        $publishedDisplayName = ConvertTo-PublishedDisplayName -DisplayName ([string]$package.DisplayName)
        Set-ObjectNoteProperty -InputObject $package -Name 'PublishedDisplayName' -Value $publishedDisplayName
        Set-ObjectNoteProperty -InputObject $package -Name 'CloudDisplayName' -Value ''
        Set-ObjectNoteProperty -InputObject $package -Name 'CloudId' -Value ''
    }

    if (-not $script:IsGraphConnected) {
        foreach ($package in $packageList) {
            Set-ObjectNoteProperty -InputObject $package -Name 'IntuneStatus' -Value 'Unknown'
            Set-ObjectNoteProperty -InputObject $package -Name 'IntuneDetails' -Value ("Connect Graph to compare local script content with Intune. Expected published name: {0}" -f $package.PublishedDisplayName)
        }
        return
    }

    $cloudByContentSignature = @{}
    $cloudByDisplayName = @{}
    foreach ($cloudRemediation in @($CloudRemediations)) {
        if ($null -eq $cloudRemediation) { continue }

        [void](Ensure-CloudRemediationContentSignature -CloudRemediation $cloudRemediation)

        $contentSignature = [string]$cloudRemediation.ContentSignature
        if (-not [string]::IsNullOrWhiteSpace($contentSignature)) {
            if (-not $cloudByContentSignature.ContainsKey($contentSignature)) {
                $cloudByContentSignature[$contentSignature] = @()
            }

            $cloudByContentSignature[$contentSignature] = @($cloudByContentSignature[$contentSignature]) + $cloudRemediation
        }

        $displayName = [string]$cloudRemediation.DisplayName
        if ([string]::IsNullOrWhiteSpace($displayName)) { continue }

        if (-not $cloudByDisplayName.ContainsKey($displayName)) {
            $cloudByDisplayName[$displayName] = @()
        }

        $cloudByDisplayName[$displayName] = @($cloudByDisplayName[$displayName]) + $cloudRemediation
    }

    foreach ($package in $packageList) {
        $publishedDisplayName = [string]$package.PublishedDisplayName
        Set-ObjectNoteProperty -InputObject $package -Name 'ContentSignature' -Value (Get-LocalPackageContentSignature -Package $package)

        $contentSignature = [string]$package.ContentSignature
        if (-not [string]::IsNullOrWhiteSpace($contentSignature) -and $cloudByContentSignature.ContainsKey($contentSignature)) {
            $matches = @($cloudByContentSignature[$contentSignature])
            $ids = @($matches | ForEach-Object { [string]$_.Id }) -join ', '
            $names = @($matches | ForEach-Object { [string]$_.DisplayName }) -join ', '

            if ($matches.Count -eq 1) {
                Set-ObjectNoteProperty -InputObject $package -Name 'IntuneStatus' -Value 'Content match'
            }
            else {
                Set-ObjectNoteProperty -InputObject $package -Name 'IntuneStatus' -Value 'Content duplicate'
            }

            Set-ObjectNoteProperty -InputObject $package -Name 'CloudDisplayName' -Value $names
            Set-ObjectNoteProperty -InputObject $package -Name 'CloudId' -Value $ids
            Set-ObjectNoteProperty -InputObject $package -Name 'IntuneDetails' -Value ("Content already exists in Intune. Expected published name: '{0}'. Match name(s): {1}. Id(s): {2}" -f $publishedDisplayName, $names, $ids)
        }
        elseif ($cloudByDisplayName.ContainsKey($publishedDisplayName)) {
            $nameMatches = @($cloudByDisplayName[$publishedDisplayName])
            $localContent = Get-LocalPackageScriptContent -Package $package
            $manualContentMatches = @(
                foreach ($nameMatch in $nameMatches) {
                    if (-not (Ensure-CloudRemediationContentSignature -CloudRemediation $nameMatch)) {
                        continue
                    }

                    $cloudDetectionContent = [string](Get-ObjectValue -InputObject $nameMatch -Name 'DetectionNormalizedContent')
                    $cloudRemediationContent = [string](Get-ObjectValue -InputObject $nameMatch -Name 'RemediationNormalizedContent')
                    $detectionMatches = Test-ScriptContentEqual -LeftContent $localContent.DetectionContent -RightContent $cloudDetectionContent
                    $remediationMatches = Test-ScriptContentEqual -LeftContent $localContent.RemediationContent -RightContent $cloudRemediationContent
                    if ($detectionMatches -and $remediationMatches) {
                        $nameMatch
                    }
                }
            )

            if ($manualContentMatches.Count -gt 0) {
                $ids = @($manualContentMatches | ForEach-Object { [string]$_.Id }) -join ', '
                $names = @($manualContentMatches | ForEach-Object { [string]$_.DisplayName }) -join ', '
                if ($manualContentMatches.Count -eq 1) {
                    Set-ObjectNoteProperty -InputObject $package -Name 'IntuneStatus' -Value 'Content match'
                }
                else {
                    Set-ObjectNoteProperty -InputObject $package -Name 'IntuneStatus' -Value 'Content duplicate'
                }

                Set-ObjectNoteProperty -InputObject $package -Name 'CloudDisplayName' -Value $names
                Set-ObjectNoteProperty -InputObject $package -Name 'CloudId' -Value $ids
                Set-ObjectNoteProperty -InputObject $package -Name 'IntuneDetails' -Value ("Content match confirmed by line comparison. Expected published name: '{0}'. Match name(s): {1}. Id(s): {2}" -f $publishedDisplayName, $names, $ids)
            }
            else {
                $ids = @($nameMatches | ForEach-Object { [string]$_.Id }) -join ', '
                Set-ObjectNoteProperty -InputObject $package -Name 'IntuneStatus' -Value 'Name only'
                Set-ObjectNoteProperty -InputObject $package -Name 'CloudDisplayName' -Value $publishedDisplayName
                Set-ObjectNoteProperty -InputObject $package -Name 'CloudId' -Value $ids
                Set-ObjectNoteProperty -InputObject $package -Name 'IntuneDetails' -Value ("A remediation exists with the expected name '{0}', but the Detection/Remediation content does not match the local package. Id(s): {1}" -f $publishedDisplayName, $ids)
            }
        }
        else {
            Set-ObjectNoteProperty -InputObject $package -Name 'IntuneStatus' -Value 'No content match'
            Set-ObjectNoteProperty -InputObject $package -Name 'IntuneDetails' -Value ("No Intune remediation has the same Detection/Remediation content. Expected published name: {0}" -f $publishedDisplayName)
        }
    }
}

function Get-LocalRemediationRootWarnings {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Root)

    $warnings = @()
    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        return @("Local script folder does not exist: $Root")
    }

    $allPs1Files = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.ps1' -ErrorAction Stop)
    $detectionFiles = @($allPs1Files | Where-Object { $_.Name -match '^.+-Detection\.ps1$' })
    $detectionOnlyFiles = @($allPs1Files | Where-Object { $_.Name -match '^.+-DetectionOnly\.ps1$' })
    $remediationFiles = @($allPs1Files | Where-Object { $_.Name -match '^.+-Remediation\.ps1$' })

    if (($detectionFiles.Count + $detectionOnlyFiles.Count) -eq 0) {
        $warnings += 'No local package found. Expected files ending with -Detection.ps1 or -DetectionOnly.ps1.'
    }

    $usedRemediationPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($detectionFile in $detectionFiles) {
        $expectedName = $detectionFile.Name -replace '-Detection\.ps1$', '-Remediation.ps1'
        $matchedRemediationPath = Get-MatchingLocalRemediationPath -DetectionScriptPath $detectionFile.FullName
        if (-not [string]::IsNullOrWhiteSpace($matchedRemediationPath)) {
            [void]$usedRemediationPaths.Add($matchedRemediationPath)
            continue
        }

        $folderRemediations = @(Get-ChildItem -LiteralPath $detectionFile.DirectoryName -File -Filter '*-Remediation.ps1' -ErrorAction SilentlyContinue)
        if ($folderRemediations.Count -gt 1) {
            $warnings += ("Ambiguous remediation scripts for {0}. Expected {1}, but {2} remediation files exist in the same folder." -f $detectionFile.Name, $expectedName, $folderRemediations.Count)
        }
        else {
            $warnings += ("Missing remediation script for {0}. Expected {1}, or exactly one -Remediation.ps1 file in the same folder. This package is treated as detection-only." -f $detectionFile.Name, $expectedName)
        }
    }

    foreach ($remediationFile in $remediationFiles) {
        if (-not $usedRemediationPaths.Contains($remediationFile.FullName)) {
            $expectedDetectionName = $remediationFile.Name -replace '-Remediation\.ps1$', '-Detection.ps1'
            $expectedDetectionOnlyName = $remediationFile.Name -replace '-Remediation\.ps1$', '-DetectionOnly.ps1'
            $warnings += ("Remediation script without matching detection: {0}. Expected {1} or {2} in the same folder." -f $remediationFile.Name, $expectedDetectionName, $expectedDetectionOnlyName)
        }
    }

    $similarFiles = @(
        $allPs1Files |
            Where-Object {
                $_.Name -match '(Detection|Remediation)' -and
                $_.Name -notmatch '^.+-(Detection|DetectionOnly|Remediation)\.ps1$'
            } |
            Select-Object -First 10
    )
    foreach ($file in $similarFiles) {
        $warnings += ("Ignored script with non-standard remediation name: {0}" -f $file.Name)
    }

    if ($warnings.Count -gt 25) {
        $remainingCount = $warnings.Count - 25
        $warnings = @($warnings | Select-Object -First 25)
        $warnings += "...and $remainingCount more local folder warning(s)."
    }

    return $warnings
}

function Get-CloudRemediations {
    [CmdletBinding()]
    param()

    Ensure-GuiGraphConnection
    $listUri = "$script:GraphBaseUri/deviceManagement/deviceHealthScripts"
    $items = Invoke-GraphGetAllPages -Uri $listUri

    $rows = @()
    foreach ($item in @($items)) {
        $runSummary = Get-ObjectValue -InputObject $item -Name 'runSummary'
        $roleScopeTagIds = Get-ObjectValue -InputObject $item -Name 'roleScopeTagIds'
        $rows += [pscustomobject]@{
            Id                    = Get-ObjectValue -InputObject $item -Name 'id'
            DisplayName           = Get-ObjectValue -InputObject $item -Name 'displayName'
            Description           = Get-ObjectValue -InputObject $item -Name 'description'
            Publisher             = Get-ObjectValue -InputObject $item -Name 'publisher'
            Version               = Get-ObjectValue -InputObject $item -Name 'version'
            CreatedDateTime       = Get-ObjectValue -InputObject $item -Name 'createdDateTime'
            LastModifiedDateTime  = Get-ObjectValue -InputObject $item -Name 'lastModifiedDateTime'
            RunAsAccount          = Get-ObjectValue -InputObject $item -Name 'runAsAccount'
            RunAs32Bit            = Get-ObjectValue -InputObject $item -Name 'runAs32Bit'
            EnforceSignatureCheck = Get-ObjectValue -InputObject $item -Name 'enforceSignatureCheck'
            RoleScopeTagIds       = if ($roleScopeTagIds) { @($roleScopeTagIds) -join ';' } else { '' }
            DetectionState        = Get-ObjectValue -InputObject $runSummary -Name 'detectionState'
            RemediationState      = Get-ObjectValue -InputObject $runSummary -Name 'remediationState'
            PortalStatus          = 'Loading'
            NoIssueDetectedDeviceCount = 0
            IssueDetectedDeviceCount = 0
            IssueRemediatedDeviceCount = 0
            IssueReoccurredDeviceCount = 0
            IssueRemediatedCumulativeDeviceCount = 0
            DetectionContentHash = ''
            RemediationContentHash = ''
            ContentSignature = ''
            ContentSignatureAvailable = $false
            ContentSignatureDetails = ''
            PSScriptAnalyzerStatus  = 'Not tested'
            PSScriptAnalyzerDetails = ''
            Raw                   = $item
        }
    }

    return @($rows | Sort-Object DisplayName)
}

function Get-CloudRemediationDetail {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Id)

    Ensure-GuiGraphConnection
    return Invoke-MgGraphRequest -Method GET -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$Id" -ErrorAction Stop
}

function Get-CloudRemediationRunSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Id)

    Ensure-GuiGraphConnection
    $summary = Invoke-MgGraphRequest -Method GET -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$Id/runSummary" -ErrorAction Stop
    $value = Get-ObjectValue -InputObject $summary -Name 'value'
    if ($null -ne $value) {
        return $value
    }

    return $summary
}

function Update-CloudRemediationPortalColumns {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$CloudRemediation)

    $displayName = [string]$CloudRemediation.DisplayName
    try {
        $assignments = @(Get-CloudAssignments -Id ([string]$CloudRemediation.Id))
        $status = if ($assignments.Count -gt 0) { 'Active' } else { 'Not deployed' }
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PortalStatus' -Value $status

        $summary = Get-CloudRemediationRunSummary -Id ([string]$CloudRemediation.Id)
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'NoIssueDetectedDeviceCount' -Value (Get-IntValue (Get-ObjectValue -InputObject $summary -Name 'noIssueDetectedDeviceCount'))
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'IssueDetectedDeviceCount' -Value (Get-IntValue (Get-ObjectValue -InputObject $summary -Name 'issueDetectedDeviceCount'))
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'IssueRemediatedDeviceCount' -Value (Get-IntValue (Get-ObjectValue -InputObject $summary -Name 'issueRemediatedDeviceCount'))
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'IssueReoccurredDeviceCount' -Value (Get-IntValue (Get-ObjectValue -InputObject $summary -Name 'issueReoccurredDeviceCount'))
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'IssueRemediatedCumulativeDeviceCount' -Value (Get-IntValue (Get-ObjectValue -InputObject $summary -Name 'issueRemediatedCumulativeDeviceCount'))
    }
    catch {
        Set-ObjectNoteProperty -InputObject $CloudRemediation -Name 'PortalStatus' -Value 'Unknown'
        Write-GuiLog ("Cloud portal columns failed for remediation '{0}' - {1}" -f $displayName, $_.Exception.Message)
    }
}

function Update-CloudRemediationsPortalColumns {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$CloudRemediations)

    $cloudList = @($CloudRemediations)
    if ($cloudList.Count -eq 0) { return }

    for ($i = 0; $i -lt $cloudList.Count; $i++) {
        $cloudRemediation = $cloudList[$i]
        Set-GuiBusy -IsBusy $true -Message ("Loading cloud status {0}/{1}: {2}" -f ($i + 1), $cloudList.Count, $cloudRemediation.DisplayName)
        Update-CloudRemediationPortalColumns -CloudRemediation $cloudRemediation
    }
}

function Get-ConnectedTenantName {
    [CmdletBinding()]
    param()

    Ensure-GuiGraphConnection

    try {
        $organizations = @(Invoke-GraphGetAllPages -Uri "$script:GraphBaseUri/organization?`$select=id,displayName,verifiedDomains")
        $organization = $organizations | Select-Object -First 1
        if ($null -ne $organization) {
            $displayName = [string](Get-ObjectValue -InputObject $organization -Name 'displayName')
            if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                return $displayName
            }

            $verifiedDomains = @(Get-ObjectValue -InputObject $organization -Name 'verifiedDomains')
            $defaultDomain = $verifiedDomains |
                Where-Object { [bool](Get-ObjectValue -InputObject $_ -Name 'isDefault') } |
                Select-Object -First 1
            if ($null -ne $defaultDomain) {
                $domainName = [string](Get-ObjectValue -InputObject $defaultDomain -Name 'name')
                if (-not [string]::IsNullOrWhiteSpace($domainName)) {
                    return $domainName
                }
            }
        }
    }
    catch {
        Write-GuiLog "Tenant display name lookup failed - $($_.Exception.Message)"
    }

    if ($null -ne $script:GraphContext -and -not [string]::IsNullOrWhiteSpace($script:GraphContext.Account)) {
        $account = [string]$script:GraphContext.Account
        if ($account -like '*@*') {
            return ($account -split '@')[-1]
        }
    }

    if ($null -ne $script:GraphContext -and -not [string]::IsNullOrWhiteSpace($script:GraphContext.TenantId)) {
        return [string]$script:GraphContext.TenantId
    }

    return 'UnknownTenant'
}

function Get-CloudAssignments {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Id)

    Ensure-GuiGraphConnection
    return @(Invoke-GraphGetAllPages -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$Id/assignments")
}

function Get-GroupDisplayName {
    [CmdletBinding()]
    param([AllowNull()][string]$GroupId)

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        return ''
    }

    if ($script:GroupDisplayNameCache.ContainsKey($GroupId)) {
        return [string]$script:GroupDisplayNameCache[$GroupId]
    }

    try {
        $group = Invoke-MgGraphRequest -Method GET -Uri "$script:GraphBaseUri/groups/$GroupId?`$select=id,displayName" -ErrorAction Stop
        $displayName = [string](Get-ObjectValue -InputObject $group -Name 'displayName')
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = ''
        }

        $script:GroupDisplayNameCache[$GroupId] = $displayName
        return $displayName
    }
    catch {
        Write-GuiLog "Group name lookup failed for $GroupId - $($_.Exception.Message)"
        $script:GroupDisplayNameCache[$GroupId] = ''
        return ''
    }
}

function ConvertTo-AssignmentArchiveObject {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Assignment)

    $target = Get-ObjectValue -InputObject $Assignment -Name 'target'
    $targetType = [string](Get-ObjectValue -InputObject $target -Name '@odata.type')
    $groupId = [string](Get-ObjectValue -InputObject $target -Name 'groupId')
    $groupName = Get-GroupDisplayName -GroupId $groupId

    return [ordered]@{
        id                    = Get-ObjectValue -InputObject $Assignment -Name 'id'
        runRemediationScript  = Get-ObjectValue -InputObject $Assignment -Name 'runRemediationScript'
        runSchedule           = Get-ObjectValue -InputObject $Assignment -Name 'runSchedule'
        targetODataType       = $targetType
        targetGroupId         = $groupId
        targetGroupName       = $groupName
        target                = $target
    }
}

function New-DeviceHealthScriptBodyFromLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Package,
        [Parameter(Mandatory = $true)][string]$Publisher
    )

    $publishedDisplayName = ConvertTo-PublishedDisplayName -DisplayName ([string]$Package.DisplayName)
    $body = [ordered]@{
        '@odata.type'             = '#microsoft.graph.deviceHealthScript'
        publisher                 = $Publisher
        version                   = '1.0'
        displayName               = $publishedDisplayName
        description               = $Package.Description
        detectionScriptContent    = ConvertTo-GraphBinaryString -LiteralPath $Package.DetectionPath
        remediationScriptContent  = ''
        runAsAccount             = 'system'
        enforceSignatureCheck     = $false
        runAs32Bit                = $false
        roleScopeTagIds           = @('0')
        deviceHealthScriptType    = 'deviceHealthScript'
    }

    if (-not [string]::IsNullOrWhiteSpace($Package.RemediationPath)) {
        $body['remediationScriptContent'] = ConvertTo-GraphBinaryString -LiteralPath $Package.RemediationPath
    }

    return $body
}

function New-DeviceHealthScriptPatchBodyFromLocal {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Package)

    $body = [ordered]@{
        '@odata.type'             = '#microsoft.graph.deviceHealthScript'
        description               = $Package.Description
        detectionScriptContent    = ConvertTo-GraphBinaryString -LiteralPath $Package.DetectionPath
        remediationScriptContent  = ''
        runAsAccount             = 'system'
        enforceSignatureCheck     = $false
        runAs32Bit                = $false
    }

    if (-not [string]::IsNullOrWhiteSpace($Package.RemediationPath)) {
        $body['remediationScriptContent'] = ConvertTo-GraphBinaryString -LiteralPath $Package.RemediationPath
    }

    return $body
}

function New-DeviceHealthScriptBodyFromCloudDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Detail,
        [AllowNull()][string]$DisplayNameOverride
    )

    $roleScopeTagIds = Get-ObjectValue -InputObject $Detail -Name 'roleScopeTagIds'
    if ($null -eq $roleScopeTagIds) { $roleScopeTagIds = @('0') }
    $displayName = if ([string]::IsNullOrWhiteSpace($DisplayNameOverride)) {
        Get-ObjectValue -InputObject $Detail -Name 'displayName'
    }
    else {
        $DisplayNameOverride
    }

    return [ordered]@{
        '@odata.type'             = '#microsoft.graph.deviceHealthScript'
        publisher                 = Get-ObjectValue -InputObject $Detail -Name 'publisher'
        version                   = Get-ObjectValue -InputObject $Detail -Name 'version'
        displayName               = $displayName
        description               = Get-ObjectValue -InputObject $Detail -Name 'description'
        detectionScriptContent    = Get-ObjectValue -InputObject $Detail -Name 'detectionScriptContent'
        remediationScriptContent  = Get-ObjectValue -InputObject $Detail -Name 'remediationScriptContent'
        runAsAccount             = Get-ObjectValue -InputObject $Detail -Name 'runAsAccount'
        enforceSignatureCheck     = [bool](Get-ObjectValue -InputObject $Detail -Name 'enforceSignatureCheck')
        runAs32Bit                = [bool](Get-ObjectValue -InputObject $Detail -Name 'runAs32Bit')
        roleScopeTagIds           = @($roleScopeTagIds)
        deviceHealthScriptType    = 'deviceHealthScript'
    }
}

function Publish-LocalPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Package,
        [switch]$ConfirmExistingUpdate
    )

    Ensure-GuiGraphConnection

    $publisher = if ($null -ne $script:GraphContext -and -not [string]::IsNullOrWhiteSpace($script:GraphContext.Account)) {
        [string]$script:GraphContext.Account
    }
    else {
        'Interactive Graph account'
    }

    $body = New-DeviceHealthScriptBodyFromLocal -Package $Package -Publisher $publisher
    $publishedDisplayName = ConvertTo-PublishedDisplayName -DisplayName ([string]$Package.DisplayName)
    $escapedName = $publishedDisplayName.Replace("'", "''")
    $filter = [System.Uri]::EscapeDataString("displayName eq '$escapedName'")
    $existing = @(Invoke-GraphGetAllPages -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts?`$filter=$filter")

    if ($existing.Count -gt 1) {
        throw "More than one Intune remediation uses displayName '$publishedDisplayName'. Rename duplicates before updating."
    }

    if ($existing.Count -eq 1) {
        $id = [string](Get-ObjectValue -InputObject $existing[0] -Name 'id')
        if ($ConfirmExistingUpdate -and -not $ValidateOnly) {
            $message = "A remediation named '$publishedDisplayName' already exists in Intune.`r`n`r`nLocal package: $($Package.DisplayName)`r`n`r`nThis action will update the existing remediation instead of creating a new one.`r`n`r`nContinue?"
            $answer = [System.Windows.MessageBox]::Show($message, 'Existing remediation found', 'YesNo', 'Warning')
            if ($answer -ne 'Yes') {
                Write-GuiLog "Publish cancelled - Existing cloud remediation was not updated: $publishedDisplayName ($id)"
                return $null
            }

            Write-GuiLog "Existing cloud remediation found; updating: $publishedDisplayName ($id). Local package: $($Package.DisplayName)"
        }

        $patchBody = New-DeviceHealthScriptPatchBodyFromLocal -Package $Package
        Invoke-GraphJsonRequest -Method PATCH -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$id" -Body $patchBody | Out-Null
        Write-GuiLog "Updated cloud remediation: $publishedDisplayName ($id). Local package: $($Package.DisplayName)"
        return $id
    }

    $created = Invoke-GraphJsonRequest -Method POST -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts" -Body $body
    $createdId = [string](Get-ObjectValue -InputObject $created -Name 'id')
    Write-GuiLog "Created cloud remediation: $publishedDisplayName ($createdId). Local package: $($Package.DisplayName)"
    return $createdId
}

function New-DetectionOnlyPublishPackage {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Package)

    $description = Get-ScriptHelpDescription -LiteralPath ([string]$Package.DetectionPath)
    if ([string]::IsNullOrWhiteSpace($description)) {
        $description = [string]$Package.DisplayName
    }

    return [pscustomobject]@{
        DisplayName     = [string]$Package.DisplayName
        Description     = $description
        Folder          = [string]$Package.Folder
        DetectionPath   = [string]$Package.DetectionPath
        RemediationPath = ''
        DetectionOnly   = $true
    }
}

function Get-PublishAllLocalPackagesPlan {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Packages)

    Ensure-GuiGraphConnection

    $packagesToCheck = @($Packages)
    $items = @()

    Set-GuiBusy -IsBusy $true -Message ("Checking existing Intune remediations for {0} local package(s)..." -f $packagesToCheck.Count)
    try {
        for ($i = 0; $i -lt $packagesToCheck.Count; $i++) {
            $package = $packagesToCheck[$i]
            $publishedDisplayName = ConvertTo-PublishedDisplayName -DisplayName ([string]$package.DisplayName)
            Set-GuiBusy -IsBusy $true -Message ("Checking existing Intune remediation {0}/{1}: {2}" -f ($i + 1), $packagesToCheck.Count, $publishedDisplayName)

            $escapedName = $publishedDisplayName.Replace("'", "''")
            $filter = [System.Uri]::EscapeDataString("displayName eq '$escapedName'")
            $existing = @(Invoke-GraphGetAllPages -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts?`$filter=$filter")

            $status = 'Create'
            $ids = @()
            if ($existing.Count -eq 1) {
                $status = 'Update'
                $ids = @([string](Get-ObjectValue -InputObject $existing[0] -Name 'id'))
            }
            elseif ($existing.Count -gt 1) {
                $status = 'Duplicate'
                $ids = @($existing | ForEach-Object { [string](Get-ObjectValue -InputObject $_ -Name 'id') })
            }

            $items += [pscustomobject]@{
                Package     = $package
                DisplayName = $publishedDisplayName
                LocalName   = [string]$package.DisplayName
                Status      = $status
                Ids         = $ids
            }
        }
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }

    $createCount = @($items | Where-Object { $_.Status -eq 'Create' }).Count
    $updateCount = @($items | Where-Object { $_.Status -eq 'Update' }).Count
    $duplicateItems = @($items | Where-Object { $_.Status -eq 'Duplicate' })

    Write-GuiLog ("Publish all preflight: Total={0}, Create={1}, Update={2}, Duplicate={3}" -f $items.Count, $createCount, $updateCount, $duplicateItems.Count)
    foreach ($duplicateItem in $duplicateItems) {
        Write-GuiLog ("Publish all preflight duplicate: {0} ({1})" -f $duplicateItem.DisplayName, ($duplicateItem.Ids -join ', '))
    }

    return [pscustomobject]@{
        Items      = $items
        Create     = @($items | Where-Object { $_.Status -eq 'Create' })
        Update     = @($items | Where-Object { $_.Status -eq 'Update' })
        Duplicate  = $duplicateItems
    }
}

function Format-PublishAllLocalPackagesPlanMessage {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Plan)

    $lines = @(
        'Publish ALL to Intune summary:',
        '',
        ('Local packages: {0}' -f @($Plan.Items).Count),
        ('Existing remediations to update: {0}' -f @($Plan.Update).Count),
        ('New remediations to create: {0}' -f @($Plan.Create).Count),
        ('Name prefix mapping: "{0}" -> "{1}"' -f $script:PublishSourceNamePrefix, $script:PublishTargetNamePrefix)
    )

    if (@($Plan.Duplicate).Count -gt 0) {
        $lines += ''
        $lines += 'Duplicate Intune remediation names found. Publish ALL cannot continue until these are cleaned up:'
        foreach ($item in @($Plan.Duplicate)) {
            $lines += ('- {0} ({1})' -f $item.DisplayName, ($item.Ids -join ', '))
        }

        return ($lines -join "`r`n")
    }

    if (@($Plan.Update).Count -gt 0) {
        $lines += ''
        $lines += 'These existing Intune remediations will be updated:'
        foreach ($item in @($Plan.Update)) {
            if ($item.LocalName -ne $item.DisplayName) {
                $lines += ('- {0} (local: {1})' -f $item.DisplayName, $item.LocalName)
            }
            else {
                $lines += ('- {0}' -f $item.DisplayName)
            }
        }
    }

    if (@($Plan.Create).Count -gt 0) {
        $lines += ''
        $lines += 'These new Intune remediations will be created:'
        foreach ($item in @($Plan.Create)) {
            if ($item.LocalName -ne $item.DisplayName) {
                $lines += ('- {0} (local: {1})' -f $item.DisplayName, $item.LocalName)
            }
            else {
                $lines += ('- {0}' -f $item.DisplayName)
            }
        }
    }

    $lines += ''
    $lines += 'Continue?'
    return ($lines -join "`r`n")
}

function Publish-AllLocalPackages {
    [CmdletBinding()]
    param([AllowNull()]$Packages)

    Ensure-GuiGraphConnection

    $packages = if ($null -eq $Packages) { @(Get-LocalPackages) } else { @($Packages) }
    if ($packages.Count -eq 0) {
        throw 'No local package found to publish.'
    }

    Set-GuiBusy -IsBusy $true -Message "Running PSScriptAnalyzer on $($packages.Count) local package(s)..."
    try {
        Test-LocalPackagesWithPSScriptAnalyzer -Packages $packages

        $publishedIds = @()
        for ($i = 0; $i -lt $packages.Count; $i++) {
            $package = $packages[$i]
            $publishedDisplayName = ConvertTo-PublishedDisplayName -DisplayName ([string]$package.DisplayName)
            Set-GuiBusy -IsBusy $true -Message ("Publishing local package {0}/{1}: {2}" -f ($i + 1), $packages.Count, $publishedDisplayName)
            $publishedIds += Publish-LocalPackage -Package $package
        }

        Write-GuiLog ("Publish all completed: {0} local package(s) published." -f $publishedIds.Count)
        return $publishedIds
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
}

function Copy-CloudAssignments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourceId,
        [Parameter(Mandatory = $true)][string]$TargetId,
        [AllowNull()]$Assignments
    )

    $assignments = if ($null -eq $Assignments) { @(Get-CloudAssignments -Id $SourceId) } else { @($Assignments) }
    if ($assignments.Count -eq 0) {
        Write-GuiLog 'No assignments to copy.'
        return
    }

    $assignmentBody = @()
    foreach ($assignment in @($assignments)) {
        $target = Get-ObjectValue -InputObject $assignment -Name 'target'
        if ($null -eq $target) { continue }

        $assignmentBody += [ordered]@{
            '@odata.type'        = '#microsoft.graph.deviceHealthScriptAssignment'
            target               = $target
            runRemediationScript = [bool](Get-ObjectValue -InputObject $assignment -Name 'runRemediationScript')
            runSchedule          = Get-ObjectValue -InputObject $assignment -Name 'runSchedule'
        }
    }

    if ($assignmentBody.Count -eq 0) {
        Write-GuiLog 'No reusable assignments were found.'
        return
    }

    $body = [ordered]@{
        deviceHealthScriptAssignments = @($assignmentBody)
    }
    Invoke-GraphJsonRequest -Method POST -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$TargetId/assign" -Body $body | Out-Null
    Write-GuiLog "Copied assignments from $SourceId to $TargetId`: $($assignmentBody.Count)"
}

function Reset-CloudRemediationHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$CloudRemediation)

    Ensure-GuiGraphConnection

    $id = [string]$CloudRemediation.Id
    $displayName = [string]$CloudRemediation.DisplayName
    $detail = Get-CloudRemediationDetail -Id $id
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $duplicateDisplayName = "$displayName - Duplicate - $timestamp"
    $body = New-DeviceHealthScriptBodyFromCloudDetail -Detail $detail -DisplayNameOverride $duplicateDisplayName

    Write-GuiLog "Creating duplicate remediation for $displayName as '$duplicateDisplayName'..."
    $created = Invoke-GraphJsonRequest -Method POST -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts" -Body $body
    $newId = [string](Get-ObjectValue -InputObject $created -Name 'id')
    if ([string]::IsNullOrWhiteSpace($newId)) {
        throw 'Graph did not return an id for the duplicated remediation.'
    }

    try {
        Write-GuiLog "Duplicate remediation created: $duplicateDisplayName ($newId)"

        $assignments = @(Get-CloudAssignments -Id $id)
        if ($assignments.Count -gt 0) {
            Copy-CloudAssignments -SourceId $id -TargetId $newId -Assignments $assignments
        }
        else {
            Write-GuiLog "No assignments to copy for $displayName ($id)."
        }

        Invoke-MgGraphRequest -Method DELETE -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$id" -ErrorAction Stop | Out-Null
        Write-GuiLog "Original remediation deleted after duplicate validation: $displayName ($id)"

        Write-GuiLog "History reset completed: old $id deleted, replacement $newId kept as '$duplicateDisplayName'."
        return [pscustomobject]@{
            Id          = $newId
            DisplayName = $duplicateDisplayName
        }
    }
    catch {
        Write-GuiLog "History reset failed after duplicate creation. Cleaning up duplicate $newId..."
        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$newId" -ErrorAction Stop | Out-Null
            Write-GuiLog "Duplicate cleanup completed: $newId deleted."
        }
        catch {
            Write-GuiLog "Duplicate cleanup failed for $newId - $($_.Exception.Message)"
        }

        throw
    }
}

function Remove-CloudRemediation {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$CloudRemediation)

    Ensure-GuiGraphConnection

    $id = [string]$CloudRemediation.Id
    $displayName = [string]$CloudRemediation.DisplayName
    if ([string]::IsNullOrWhiteSpace($id)) {
        throw 'Selected cloud remediation does not have an Intune id.'
    }

    try {
        Invoke-MgGraphRequest -Method DELETE -Uri "$script:GraphBaseUri/deviceManagement/deviceHealthScripts/$id" -ErrorAction Stop | Out-Null
        Write-GuiLog "Deleted cloud remediation: $displayName ($id)"
    }
    catch {
        Write-GuiLog ("Delete cloud remediation failed: {0} ({1}) - {2}" -f $displayName, $id, $_.Exception.Message)
        throw
    }
}

function Test-ImportExcelModuleAvailable {
    [CmdletBinding()]
    param()

    return $null -ne (Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1)
}

function Ensure-ImportExcelModuleAvailable {
    [CmdletBinding()]
    param()

    if (Test-ImportExcelModuleAvailable) {
        $script:IsImportExcelAvailable = $true
        return $true
    }

    try {
        Write-GuiLog 'ImportExcel module not found. Attempting installation for current user...'
        [void](Get-Command -Name Install-Module -ErrorAction Stop)
        Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -AcceptLicense -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        $script:IsImportExcelAvailable = $true
        Write-GuiLog 'ImportExcel module installed successfully.'
        return $true
    }
    catch {
        $script:IsImportExcelAvailable = $false
        Write-GuiLog "ImportExcel module installation failed. Excel export will not be offered. $($_.Exception.Message)"
        return $false
    }
}

function Initialize-ImportExcelSupport {
    [CmdletBinding()]
    param()

    if ($script:ImportExcelSupportChecked) {
        return $script:IsImportExcelAvailable
    }

    $script:ImportExcelSupportChecked = $true
    if (Test-ImportExcelModuleAvailable) {
        $script:IsImportExcelAvailable = $true
        Write-GuiLog 'ImportExcel module detected. Excel execution report export is available.'
        return $true
    }

    $message = @"
Excel export requires the PowerShell module ImportExcel.

This module creates .xlsx workbooks without using Excel COM automation, which is safer and more reliable for this tool.

Install ImportExcel for the current user from PowerShell Gallery now?

If you choose No, the Excel option will be hidden and CSV export will remain available.
"@

    $answer = [System.Windows.MessageBox]::Show($message, 'Enable Excel export', 'YesNo', 'Question')
    if ($answer -ne 'Yes') {
        $script:IsImportExcelAvailable = $false
        Write-GuiLog 'ImportExcel installation declined by user. Excel execution report export will not be offered.'
        return $false
    }

    try {
        Set-GuiBusy -IsBusy $true -Message 'Installing ImportExcel module for Excel export...'
        return (Ensure-ImportExcelModuleAvailable)
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
}

function Export-ExecutionCsvToExcelWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][string]$ExcelPath,
        [Parameter(Mandatory = $true)]$CloudRemediation
    )

    Import-Module ImportExcel -ErrorAction Stop

    $rows = @(Import-Csv -LiteralPath $CsvPath)
    $uniqueDevices = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.DeviceId) } | Select-Object -ExpandProperty DeviceId -Unique)
    $detectionStatusCount = @($rows | Group-Object -Property DetectionStatus | Sort-Object Name | ForEach-Object {
            [pscustomobject]@{
                Type   = 'DetectionStatus'
                Status = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { '(blank)' } else { [string]$_.Name }
                Count  = $_.Count
            }
        })
    $remediationStatusCount = @($rows | Group-Object -Property RemediationStatus | Sort-Object Name | ForEach-Object {
            [pscustomobject]@{
                Type   = 'RemediationStatus'
                Status = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { '(blank)' } else { [string]$_.Name }
                Count  = $_.Count
            }
        })
    $summary = @(
        [pscustomobject]@{ Name = 'Remediation'; Value = [string]$CloudRemediation.DisplayName }
        [pscustomobject]@{ Name = 'Id'; Value = [string]$CloudRemediation.Id }
        [pscustomobject]@{ Name = 'Author'; Value = [string]$CloudRemediation.Publisher }
        [pscustomobject]@{ Name = 'Status'; Value = [string]$CloudRemediation.PortalStatus }
        [pscustomobject]@{ Name = 'ExportedAt'; Value = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
        [pscustomobject]@{ Name = 'RowCount'; Value = $rows.Count }
        [pscustomobject]@{ Name = 'UniqueDeviceCount'; Value = $uniqueDevices.Count }
    )

    $targetFolder = Split-Path -Path $ExcelPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($targetFolder)) {
        [IO.Directory]::CreateDirectory($targetFolder) | Out-Null
    }

    if (Test-Path -LiteralPath $ExcelPath -PathType Leaf) {
        Remove-Item -LiteralPath $ExcelPath -Force -ErrorAction Stop
    }

    $summary | Export-Excel -Path $ExcelPath -WorksheetName 'Summary' -TableName 'Summary' -AutoSize -BoldTopRow
    @($detectionStatusCount + $remediationStatusCount) |
        Export-Excel -Path $ExcelPath -WorksheetName 'Status counts' -TableName 'StatusCounts' -AutoSize -BoldTopRow
    if ($rows.Count -gt 0) {
        $rows | Export-Excel -Path $ExcelPath -WorksheetName 'Raw data' -TableName 'RawData' -AutoSize -FreezeTopRow -BoldTopRow
    }
    else {
        [pscustomobject]@{ Message = 'No execution rows were returned by Intune.' } |
            Export-Excel -Path $ExcelPath -WorksheetName 'Raw data' -TableName 'RawData' -AutoSize -BoldTopRow
    }
}

function Export-ExecutionReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$CloudRemediation,
        [Parameter(Mandatory = $true)][ValidateSet('CSV', 'Excel')][string]$Format
    )

    Ensure-GuiGraphConnection -Scopes @($script:BaseScopes + $script:ReportExportScopes)

    $policyId = [string]$CloudRemediation.Id
    $displayName = [string]$CloudRemediation.DisplayName
    $safeName = ConvertTo-SafeFileName -Value $displayName
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    if ($Format -eq 'Excel' -and -not (Test-ImportExcelModuleAvailable)) {
        throw 'Excel export was selected, but ImportExcel is not available.'
    }

    $targetPath = Get-ExecutionReportSavePath -DisplayName $displayName -Timestamp $timestamp -Format $Format
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        Write-GuiLog 'Execution report export cancelled.'
        return $null
    }

    $targetPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($targetPath)
    $targetFolder = Split-Path -Path $targetPath -Parent
    if ([string]::IsNullOrWhiteSpace($targetFolder)) {
        throw 'The selected report path has no parent folder.'
    }

    $exportFolder = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365-IntuneRemediationReport-{0}-{1}" -f $timestamp, ([guid]::NewGuid().ToString('N')))
    [IO.Directory]::CreateDirectory($exportFolder) | Out-Null

    Set-GuiBusy -IsBusy $true -Message "Starting execution report export for $displayName..."
    try {
        $select = @(
            'DeviceId',
            'DeviceName',
            'UserName',
            'UPN',
            'OSVersion',
            'ModifiedTime',
            'PolicyId',
            'DetectionStatus',
            'RemediationStatus',
            'PreRemediationDetectionScriptOutput',
            'PreRemediationDetectionScriptError',
            'RemediationScriptOutputDetails',
            'RemediationScriptErrorDetails',
            'PostRemediationDetectionScriptOutput',
            'PostRemediationDetectionScriptError',
            'NextExecutionDateTime'
        )
        $body = [ordered]@{
            reportName       = 'DeviceRunStatesByProactiveRemediation'
            filter           = "(PolicyId eq '$policyId')"
            select           = $select
            format           = 'csv'
            localizationType = 'ReplaceLocalizableValues'
        }

        Write-GuiLog "Starting execution report export for $displayName..."
        $job = Invoke-GraphJsonRequest -Method POST -Uri "$script:GraphBaseUri/deviceManagement/reports/exportJobs" -Body $body
        $jobId = [string](Get-ObjectValue -InputObject $job -Name 'id')
        if ([string]::IsNullOrWhiteSpace($jobId)) {
            throw 'Graph did not return an export job id.'
        }

        $status = ''
        $url = ''
        for ($i = 0; $i -lt 60; $i++) {
            Set-GuiBusy -IsBusy $true -Message 'Waiting for Intune report generation...'
            Start-Sleep -Seconds 5
            $jobState = Invoke-MgGraphRequest -Method GET -Uri "$script:GraphBaseUri/deviceManagement/reports/exportJobs('$jobId')" -ErrorAction Stop
            $status = [string](Get-ObjectValue -InputObject $jobState -Name 'status')
            Write-GuiLog "Report job status: $status"

            if ($status -eq 'completed') {
                $url = [string](Get-ObjectValue -InputObject $jobState -Name 'url')
                break
            }
            if ($status -in @('failed', 'unknownFutureValue')) {
                throw "Report export job ended with status: $status"
            }
        }

        if ([string]::IsNullOrWhiteSpace($url)) {
            throw 'Report export job did not complete before timeout.'
        }

        Set-GuiBusy -IsBusy $true -Message 'Downloading execution report export...'
        $zipPath = Join-Path -Path $exportFolder -ChildPath "$safeName-$timestamp.zip"
        Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop

        Set-GuiBusy -IsBusy $true -Message 'Extracting execution report export...'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $exportFolder, $true)
        $csv = Get-ChildItem -LiteralPath $exportFolder -Filter '*.csv' -File -Recurse | Select-Object -First 1
        if ($null -eq $csv) {
            throw "Report ZIP downloaded, but no CSV file was found in $exportFolder"
        }

        [IO.Directory]::CreateDirectory($targetFolder) | Out-Null
        if ($Format -eq 'Excel') {
            Set-GuiBusy -IsBusy $true -Message 'Creating execution Excel workbook...'
            Export-ExecutionCsvToExcelWorkbook -CsvPath $csv.FullName -ExcelPath $targetPath -CloudRemediation $CloudRemediation
            Write-GuiLog "Execution report Excel workbook: $targetPath"
        }
        else {
            Move-Item -LiteralPath $csv.FullName -Destination $targetPath -Force
            Write-GuiLog "Execution report CSV: $targetPath"
        }

        return [pscustomobject]@{
            Path   = $targetPath
            Format = $Format
        }
    }
    finally {
        if (Test-Path -LiteralPath $exportFolder) {
            Remove-Item -LiteralPath $exportFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        Set-GuiBusy -IsBusy $false
    }
}

function Export-LocalRemediationsArchive {
    [CmdletBinding()]
    param()

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $tenantName = if ($script:IsGraphConnected) { Get-ConnectedTenantName } else { 'Local' }
    $targetZipPath = Get-SaveAllArchivePath -TenantName $tenantName -Timestamp $timestamp -Scope 'Local'
    if ([string]::IsNullOrWhiteSpace($targetZipPath)) {
        Write-GuiLog 'Save all local archive cancelled.'
        return ''
    }

    $targetZipPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($targetZipPath)
    $tempRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365-IntuneRemediation-Local-{0}-{1}" -f $timestamp, ([guid]::NewGuid().ToString('N')))
    Set-GuiBusy -IsBusy $true -Message 'Creating local remediation archive...'

    try {
        $localPackages = @(Get-LocalPackages)
        Set-GuiBusy -IsBusy $true -Message "Copying $($localPackages.Count) local package(s)..."
        $localRoot = Get-DefaultRemediationRoot
        $localExportRoot = Join-Path -Path $tempRoot -ChildPath 'Local'
        [IO.Directory]::CreateDirectory($localExportRoot) | Out-Null

        Copy-Item -LiteralPath $localRoot -Destination (Join-Path -Path $localExportRoot -ChildPath 'Packages') -Recurse -Force
        $localPackages | Export-Csv -LiteralPath (Join-Path -Path $localExportRoot -ChildPath 'LocalPackages.csv') -NoTypeInformation -Encoding UTF8

        $manifest = [ordered]@{
            CreatedAtUtc         = (Get-Date).ToUniversalTime().ToString('o')
            ArchiveScope         = 'Local'
            TenantName           = $tenantName
            TenantId             = if ($null -ne $script:GraphContext) { [string]$script:GraphContext.TenantId } else { '' }
            Account              = if ($null -ne $script:GraphContext) { [string]$script:GraphContext.Account } else { '' }
            LocalRemediationRoot = $localRoot
            LocalPackageCount    = $localPackages.Count
        }
        Save-Utf8NoBom -Path (Join-Path -Path $tempRoot -ChildPath 'Manifest.json') -Content ($manifest | ConvertTo-Json -Depth 20)

        $targetFolder = Split-Path -Path $targetZipPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($targetFolder)) {
            [IO.Directory]::CreateDirectory($targetFolder) | Out-Null
        }

        $archiveItems = Get-ChildItem -LiteralPath $tempRoot -Force
        Set-GuiBusy -IsBusy $true -Message 'Compressing local remediation archive...'
        Compress-Archive -LiteralPath $archiveItems.FullName -DestinationPath $targetZipPath -Force
        Write-GuiLog ("Save all local archive created: {0}" -f $targetZipPath)
        Write-GuiLog ("Local archive content: {0} local package(s)." -f $localPackages.Count)
        return $targetZipPath
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Set-GuiBusy -IsBusy $false
    }
}

function Export-CloudRemediationsArchive {
    [CmdletBinding()]
    param()

    Ensure-GuiGraphConnection

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $tenantName = Get-ConnectedTenantName
    $targetZipPath = Get-SaveAllArchivePath -TenantName $tenantName -Timestamp $timestamp -Scope 'Cloud'
    if ([string]::IsNullOrWhiteSpace($targetZipPath)) {
        Write-GuiLog 'Save all cloud archive cancelled.'
        return ''
    }

    $targetZipPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($targetZipPath)
    $tempRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365-IntuneRemediation-Cloud-{0}-{1}" -f $timestamp, ([guid]::NewGuid().ToString('N')))
    Set-GuiBusy -IsBusy $true -Message 'Creating cloud remediation archive...'

    try {
        $cloudRemediations = @(Get-CloudRemediations)
        Set-GuiBusy -IsBusy $true -Message "Exporting $($cloudRemediations.Count) cloud remediation(s)..."

        $cloudExportRoot = Join-Path -Path $tempRoot -ChildPath 'Cloud'
        [IO.Directory]::CreateDirectory($cloudExportRoot) | Out-Null

        $cloudIndex = 0
        foreach ($cloudRemediation in $cloudRemediations) {
            $cloudIndex++
            $displayName = [string]$cloudRemediation.DisplayName
            $id = [string]$cloudRemediation.Id
            if ([string]::IsNullOrWhiteSpace($id)) { continue }
            Set-GuiBusy -IsBusy $true -Message ("Exporting cloud remediation {0}/{1}: {2}" -f $cloudIndex, $cloudRemediations.Count, $displayName)
            $idPrefix = $id.Substring(0, [Math]::Min(8, $id.Length))
            $safeCloudName = ConvertTo-SafeFileName -Value ("{0}_{1}" -f $displayName, $idPrefix)
            $cloudFolder = Join-Path -Path $cloudExportRoot -ChildPath $safeCloudName
            [IO.Directory]::CreateDirectory($cloudFolder) | Out-Null

            $detail = Get-CloudRemediationDetail -Id $id
            $detectionContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'detectionScriptContent')
            $remediationContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'remediationScriptContent')
            Save-Utf8NoBom -Path (Join-Path -Path $cloudFolder -ChildPath 'Detection.ps1') -Content $detectionContent
            Save-Utf8NoBom -Path (Join-Path -Path $cloudFolder -ChildPath 'Remediation.ps1') -Content $remediationContent

            $metadata = [ordered]@{
                Id                    = $id
                DisplayName           = Get-ObjectValue -InputObject $detail -Name 'displayName'
                Description           = Get-ObjectValue -InputObject $detail -Name 'description'
                Publisher             = Get-ObjectValue -InputObject $detail -Name 'publisher'
                Version               = Get-ObjectValue -InputObject $detail -Name 'version'
                CreatedDateTime       = Get-ObjectValue -InputObject $detail -Name 'createdDateTime'
                LastModifiedDateTime  = Get-ObjectValue -InputObject $detail -Name 'lastModifiedDateTime'
                RunAsAccount          = Get-ObjectValue -InputObject $detail -Name 'runAsAccount'
                RunAs32Bit            = Get-ObjectValue -InputObject $detail -Name 'runAs32Bit'
                EnforceSignatureCheck = Get-ObjectValue -InputObject $detail -Name 'enforceSignatureCheck'
                RoleScopeTagIds       = Get-ObjectValue -InputObject $detail -Name 'roleScopeTagIds'
            }
            Save-Utf8NoBom -Path (Join-Path -Path $cloudFolder -ChildPath 'Metadata.json') -Content ($metadata | ConvertTo-Json -Depth 20)

            try {
                $assignments = @(Get-CloudAssignments -Id $id)
                $assignmentArchive = @(
                    foreach ($assignment in $assignments) {
                        ConvertTo-AssignmentArchiveObject -Assignment $assignment
                    }
                )
                Save-Utf8NoBom -Path (Join-Path -Path $cloudFolder -ChildPath 'Assignments.json') -Content ($assignmentArchive | ConvertTo-Json -Depth 30)
            }
            catch {
                Save-Utf8NoBom -Path (Join-Path -Path $cloudFolder -ChildPath 'Assignments.json') -Content "[]"
                Write-GuiLog "Assignments export failed for $displayName - $($_.Exception.Message)"
            }
        }

        $manifest = [ordered]@{
            CreatedAtUtc          = (Get-Date).ToUniversalTime().ToString('o')
            ArchiveScope          = 'Cloud'
            TenantName            = $tenantName
            TenantId              = if ($null -ne $script:GraphContext) { [string]$script:GraphContext.TenantId } else { '' }
            Account               = if ($null -ne $script:GraphContext) { [string]$script:GraphContext.Account } else { '' }
            GraphApiVersion       = $GraphApiVersion
            CloudRemediationCount = $cloudRemediations.Count
        }
        Save-Utf8NoBom -Path (Join-Path -Path $tempRoot -ChildPath 'Manifest.json') -Content ($manifest | ConvertTo-Json -Depth 20)

        $targetFolder = Split-Path -Path $targetZipPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($targetFolder)) {
            [IO.Directory]::CreateDirectory($targetFolder) | Out-Null
        }

        $archiveItems = Get-ChildItem -LiteralPath $tempRoot -Force
        Set-GuiBusy -IsBusy $true -Message 'Compressing cloud remediation archive...'
        Compress-Archive -LiteralPath $archiveItems.FullName -DestinationPath $targetZipPath -Force
        Write-GuiLog ("Save all cloud archive created: {0}" -f $targetZipPath)
        Write-GuiLog ("Cloud archive content: {0} cloud remediation(s)." -f $cloudRemediations.Count)
        return $targetZipPath
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Set-GuiBusy -IsBusy $false
    }
}

function Refresh-LocalGrid {
    [CmdletBinding()]
    param()

    $busyWasVisible = (
        $script:Ui.ContainsKey('BusyOverlay') -and
        $null -ne $script:Ui.BusyOverlay -and
        $script:Ui.BusyOverlay.Visibility -eq [System.Windows.Visibility]::Visible
    )

    if (-not $busyWasVisible -and -not $ValidateOnly) {
        Set-GuiBusy -IsBusy $true -Message 'Refreshing local packages and running PSScriptAnalyzer...'
    }

    try {
        $localRoot = Ensure-LocalRemediationRootConfigured -PromptIfMissing:(!$ValidateOnly)
        if ([string]::IsNullOrWhiteSpace($localRoot)) {
            $script:LocalPackages = @()
            $script:SelectedLocalPackage = $null
            $script:EditorDetectionPath = ''
            $script:EditorRemediationPath = ''
            $script:Ui.LocalGrid.ItemsSource = $null
            $script:Ui.PackageDetails.Text = 'Local script folder is not configured.'
            $script:Ui.LocalDetectionEditor.Text = ''
            $script:Ui.LocalRemediationEditor.Text = ''
            Update-ActionButtonsState
            return
        }

        $script:LocalPackages = @(Get-LocalPackages)
        $structureWarnings = @(Get-LocalRemediationRootWarnings -Root $localRoot)
        foreach ($warning in $structureWarnings) {
            Write-GuiLog "Local folder warning - $warning"
        }
        if (-not $ValidateOnly) {
            Update-LocalPackagesPSScriptAnalyzerStatus -Packages $script:LocalPackages
        }
        Update-LocalPackagesIntuneStatus -Packages $script:LocalPackages -CloudRemediations $script:CloudRemediations
        $script:Ui.LocalGrid.ItemsSource = $null
        $script:Ui.LocalGrid.ItemsSource = $script:LocalPackages
        if ($script:LocalPackages.Count -eq 0) {
            $script:Ui.PackageDetails.Text = @"
Local script folder
Path: $localRoot
No local package found.

Expected naming:
- <PackageName>-Detection.ps1
- <PackageName>-Remediation.ps1
- <PackageName>-DetectionOnly.ps1
"@
        }
        Write-GuiLog "Local packages loaded: $($script:LocalPackages.Count)"
    }
    finally {
        if (-not $busyWasVisible) {
            Set-GuiBusy -IsBusy $false
        }
    }
}

function Refresh-CloudGrid {
    [CmdletBinding()]
    param()

    $busyWasVisible = (
        $script:Ui.ContainsKey('BusyOverlay') -and
        $null -ne $script:Ui.BusyOverlay -and
        $script:Ui.BusyOverlay.Visibility -eq [System.Windows.Visibility]::Visible
    )

    if (-not $busyWasVisible -and -not $ValidateOnly) {
        Set-GuiBusy -IsBusy $true -Message 'Refreshing cloud remediations and running PSScriptAnalyzer...'
    }

    try {
        $script:CloudRemediations = @(Get-CloudRemediations)
        if (-not $ValidateOnly) {
            Update-CloudRemediationsPortalColumns -CloudRemediations $script:CloudRemediations
            Update-CloudRemediationsPSScriptAnalyzerStatus -CloudRemediations $script:CloudRemediations
        }
        Update-CloudGridItemsSource
        Update-LocalPackagesIntuneStatus -Packages $script:LocalPackages -CloudRemediations $script:CloudRemediations
        $script:Ui.LocalGrid.ItemsSource = $null
        $script:Ui.LocalGrid.ItemsSource = $script:LocalPackages
        Write-GuiLog "Cloud remediations loaded: $($script:CloudRemediations.Count)"
    }
    finally {
        if (-not $busyWasVisible) {
            Set-GuiBusy -IsBusy $false
        }
    }
}

function Set-EditorFromLocalPackage {
    [CmdletBinding()]
    param([AllowNull()]$Package)

    $script:SelectedLocalPackage = $Package
    if ($null -eq $Package) {
        $script:EditorDetectionPath = ''
        $script:EditorRemediationPath = ''
        $script:Ui.LocalDetectionEditor.Text = ''
        $script:Ui.LocalRemediationEditor.Text = ''
        $script:Ui.PackageDetails.Text = ''
        Update-ActionButtonsState
        return
    }

    $script:EditorDetectionPath = [string]$Package.DetectionPath
    $script:EditorRemediationPath = [string]$Package.RemediationPath
    $script:Ui.LocalDetectionEditor.IsReadOnly = $false
    $script:Ui.LocalRemediationEditor.IsReadOnly = $false
    Update-ActionButtonsState
    $script:Ui.LocalDetectionEditor.Text = [IO.File]::ReadAllText($script:EditorDetectionPath)
    $script:Ui.LocalRemediationEditor.Text = if ([string]::IsNullOrWhiteSpace($script:EditorRemediationPath)) { '' } else { [IO.File]::ReadAllText($script:EditorRemediationPath) }
    $script:Ui.PackageDetails.Text = @"
Local package
Name: $($Package.DisplayName)
Folder: $($Package.Folder)
Detection: $($Package.DetectionPath)
Remediation: $($Package.RemediationPath)
Detection only: $($Package.DetectionOnly)
Expected Intune name: $($Package.PublishedDisplayName)
In Intune: $($Package.IntuneStatus)
Intune detail: $($Package.IntuneDetails)
PSScriptAnalyzer: $($Package.PSScriptAnalyzerStatus)
Description: $($Package.Description)
"@
}

function Set-EditorFromCloudRemediation {
    [CmdletBinding()]
    param([AllowNull()]$CloudRemediation)

    $script:SelectedCloudRemediation = $CloudRemediation
    if ($null -eq $CloudRemediation) {
        $script:Ui.CloudDetails.Text = ''
        Clear-CloudScriptViewer
        Update-ActionButtonsState
        return
    }

    Clear-CloudScriptViewer
    $script:Ui.CloudDetails.Text = @"
Cloud remediation
Name: $($CloudRemediation.DisplayName)
Id: $($CloudRemediation.Id)
Author: $($CloudRemediation.Publisher)
Status: $($CloudRemediation.PortalStatus)
Version: $($CloudRemediation.Version)
Run as: $($CloudRemediation.RunAsAccount)
Modified: $($CloudRemediation.LastModifiedDateTime)
Without issues: $($CloudRemediation.NoIssueDetectedDeviceCount)
With issues: $($CloudRemediation.IssueDetectedDeviceCount)
Issue fixed: $($CloudRemediation.IssueRemediatedDeviceCount)
Recurred: $($CloudRemediation.IssueReoccurredDeviceCount)
Total remediated: $($CloudRemediation.IssueRemediatedCumulativeDeviceCount)
PSScriptAnalyzer: $($CloudRemediation.PSScriptAnalyzerStatus)
Description: $($CloudRemediation.Description)
"@

    Update-ActionButtonsState

    try {
        Export-CloudScriptsToEditor
    }
    catch {
        $message = $_.Exception.Message
        if (Test-GraphNotFoundException -Exception $_.Exception) {
            $unavailableMessage = "Cloud detail unavailable. Graph returned NotFound while loading this remediation's script content. This Intune object may be stale or broken. Id: $($CloudRemediation.Id)"
            Set-CloudScriptViewerUnavailable -Message $unavailableMessage
            $script:Ui.CloudDetails.Text = "$($script:Ui.CloudDetails.Text)`r`nCloud detail: unavailable (Graph NotFound)`r`n"
            Write-GuiLog "Cloud detail unavailable for selected remediation '$($CloudRemediation.DisplayName)' ($($CloudRemediation.Id)) - $message"
        }
        else {
            Set-CloudScriptViewerUnavailable -Message "Cloud script content load failed. $message"
            Write-GuiLog "Cloud script content load failed - $message"
        }
    }
}

function Update-CloudGridItemsSource {
    [CmdletBinding()]
    param()

    $activeCloudRemediations = @($script:CloudRemediations | Where-Object { [string]$_.PortalStatus -eq 'Active' })
    $otherCloudRemediations = @($script:CloudRemediations | Where-Object { [string]$_.PortalStatus -ne 'Active' })

    $script:Ui.CloudActiveGrid.ItemsSource = $null
    $script:Ui.CloudOtherGrid.ItemsSource = $null
    $script:Ui.CloudActiveGrid.ItemsSource = $activeCloudRemediations
    $script:Ui.CloudOtherGrid.ItemsSource = $otherCloudRemediations
}

function Save-EditorFiles {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($script:EditorDetectionPath)) {
        throw 'Select a local package before saving.'
    }

    Test-LocalEditorScriptsWithPSScriptAnalyzer

    $modifiedScripts = @()

    $detectionContent = $script:Ui.LocalDetectionEditor.Text
    $currentDetectionContent = if (Test-Path -LiteralPath $script:EditorDetectionPath) {
        [IO.File]::ReadAllText($script:EditorDetectionPath)
    }
    else {
        ''
    }

    if ($currentDetectionContent -ne $detectionContent) {
        Save-Utf8NoBom -Path $script:EditorDetectionPath -Content $detectionContent
        $modifiedScripts += [IO.Path]::GetFileName($script:EditorDetectionPath)
    }

    if (-not [string]::IsNullOrWhiteSpace($script:EditorRemediationPath)) {
        $remediationContent = $script:Ui.LocalRemediationEditor.Text
        $currentRemediationContent = if (Test-Path -LiteralPath $script:EditorRemediationPath) {
            [IO.File]::ReadAllText($script:EditorRemediationPath)
        }
        else {
            ''
        }

        if ($currentRemediationContent -ne $remediationContent) {
            Save-Utf8NoBom -Path $script:EditorRemediationPath -Content $remediationContent
            $modifiedScripts += [IO.Path]::GetFileName($script:EditorRemediationPath)
        }
    }

    $packageName = if ($null -eq $script:SelectedLocalPackage) { 'selected package' } else { [string]$script:SelectedLocalPackage.DisplayName }
    if ($modifiedScripts.Count -gt 0) {
        Write-GuiLog ("Local package saved: {0}" -f $packageName)
        foreach ($scriptName in $modifiedScripts) {
            Write-GuiLog ("Modified local script: {0}" -f $scriptName)
        }
    }
    else {
        Write-GuiLog ("No local script changes detected for package: {0}" -f $packageName)
    }

    Refresh-LocalGrid
}

function Save-DetectionEditorFile {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($script:EditorDetectionPath)) {
        throw 'Select a local package before saving detection.'
    }

    Test-LocalEditorDetectionWithPSScriptAnalyzer

    $detectionContent = $script:Ui.LocalDetectionEditor.Text
    $currentDetectionContent = if (Test-Path -LiteralPath $script:EditorDetectionPath) {
        [IO.File]::ReadAllText($script:EditorDetectionPath)
    }
    else {
        ''
    }

    $packageName = if ($null -eq $script:SelectedLocalPackage) { 'selected package' } else { [string]$script:SelectedLocalPackage.DisplayName }
    if ($currentDetectionContent -ne $detectionContent) {
        Save-Utf8NoBom -Path $script:EditorDetectionPath -Content $detectionContent
        Write-GuiLog ("Local package detection saved: {0}" -f $packageName)
        Write-GuiLog ("Modified local script: {0}" -f ([IO.Path]::GetFileName($script:EditorDetectionPath)))
    }
    else {
        Write-GuiLog ("No local detection changes detected for package: {0}" -f $packageName)
    }
}

function Clear-CloudScriptViewer {
    [CmdletBinding()]
    param()

    $script:Ui.CloudDetectionEditor.Text = ''
    $script:Ui.CloudRemediationEditor.Text = ''
    $script:Ui.CloudDetectionEditor.IsReadOnly = $true
    $script:Ui.CloudRemediationEditor.IsReadOnly = $true
}

function Set-CloudScriptViewerUnavailable {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Message)

    Clear-CloudScriptViewer
    $script:Ui.CloudDetectionEditor.Text = $Message
}

function ConvertTo-CompareLines {
    [CmdletBinding()]
    param([AllowNull()][string]$Content)

    if ([string]::IsNullOrEmpty($Content)) {
        return @()
    }

    $normalized = $Content -replace "`r`n", "`n" -replace "`r", "`n"
    return @($normalized.Split("`n"))
}

function ConvertTo-CompactDiffLine {
    [CmdletBinding()]
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '<missing>' }

    $text = $Value -replace "`t", '    '
    if ($text.Length -gt 180) {
        return $text.Substring(0, 180) + '...'
    }

    return $text
}

function Get-ScriptComparisonText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [AllowNull()][string]$LocalContent,
        [AllowNull()][string]$CloudContent
    )

    $localLines = @(ConvertTo-CompareLines -Content $LocalContent)
    $cloudLines = @(ConvertTo-CompareLines -Content $CloudContent)
    $maxCount = [Math]::Max($localLines.Count, $cloudLines.Count)
    $differentLines = New-Object System.Collections.Generic.List[int]

    for ($i = 0; $i -lt $maxCount; $i++) {
        $localLine = if ($i -lt $localLines.Count) { $localLines[$i] } else { $null }
        $cloudLine = if ($i -lt $cloudLines.Count) { $cloudLines[$i] } else { $null }
        if ($localLine -ne $cloudLine) {
            $differentLines.Add($i + 1)
        }
    }

    $result = New-Object System.Collections.Generic.List[string]
    if ($differentLines.Count -eq 0) {
        $result.Add("[$Label] Identical. Lines=$($localLines.Count)")
        return ($result -join "`r`n")
    }

    $result.Add("[$Label] Different. LocalLines=$($localLines.Count); CloudLines=$($cloudLines.Count); DifferentLineCount=$($differentLines.Count)")
    $result.Add('')
    $result.Add("[$Label] First differences")

    foreach ($lineNumber in @($differentLines | Select-Object -First 40)) {
        $index = $lineNumber - 1
        $localLine = if ($index -lt $localLines.Count) { $localLines[$index] } else { $null }
        $cloudLine = if ($index -lt $cloudLines.Count) { $cloudLines[$index] } else { $null }
        $result.Add(("Line {0}" -f $lineNumber))
        $result.Add(("  LOCAL: {0}" -f (ConvertTo-CompactDiffLine -Value $localLine)))
        $result.Add(("  CLOUD: {0}" -f (ConvertTo-CompactDiffLine -Value $cloudLine)))
    }

    if ($differentLines.Count -gt 40) {
        $result.Add(("... {0} additional differing line(s) not shown." -f ($differentLines.Count - 40)))
    }

    return ($result -join "`r`n")
}

function Update-LocalPackageIntuneStatusFromComparison {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Package,
        [Parameter(Mandatory = $true)]$CloudRemediation,
        [Parameter(Mandatory = $true)][bool]$DetectionIdentical,
        [Parameter(Mandatory = $true)][bool]$RemediationIdentical
    )

    if (-not ($DetectionIdentical -and $RemediationIdentical)) {
        return
    }

    Set-ObjectNoteProperty -InputObject $Package -Name 'IntuneStatus' -Value 'Content match'
    Set-ObjectNoteProperty -InputObject $Package -Name 'CloudDisplayName' -Value ([string]$CloudRemediation.DisplayName)
    Set-ObjectNoteProperty -InputObject $Package -Name 'CloudId' -Value ([string]$CloudRemediation.Id)
    Set-ObjectNoteProperty -InputObject $Package -Name 'IntuneDetails' -Value ("Content match confirmed by manual comparison with '{0}' ({1})." -f $CloudRemediation.DisplayName, $CloudRemediation.Id)

    if ($script:Ui.ContainsKey('LocalGrid') -and $null -ne $script:Ui.LocalGrid -and $null -ne $script:Ui.LocalGrid.Items) {
        $script:Ui.LocalGrid.Items.Refresh()
    }
}

function Show-ComparisonWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $compareWindow = New-Object System.Windows.Window
    $compareWindow.Title = $Title
    $compareWindow.Width = 1050
    $compareWindow.Height = 760
    $compareWindow.WindowStartupLocation = 'CenterOwner'
    $compareWindow.Owner = $window

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = [System.Windows.Thickness]::new(12)
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)

    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Text = $Text
    $textBox.IsReadOnly = $true
    $textBox.AcceptsReturn = $true
    $textBox.AcceptsTab = $true
    $textBox.FontFamily = New-Object System.Windows.Media.FontFamily('Consolas')
    $textBox.FontSize = 13
    $textBox.HorizontalScrollBarVisibility = 'Auto'
    $textBox.VerticalScrollBarVisibility = 'Auto'
    [System.Windows.Controls.Grid]::SetRow($textBox, 0)
    $grid.Children.Add($textBox) | Out-Null

    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = 'Close'
    $closeButton.Width = 90
    $closeButton.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
    $closeButton.HorizontalAlignment = 'Right'
    $closeButton.Add_Click({ $compareWindow.Close() })
    [System.Windows.Controls.Grid]::SetRow($closeButton, 1)
    $grid.Children.Add($closeButton) | Out-Null

    $compareWindow.Content = $grid
    $compareWindow.ShowDialog() | Out-Null
}

function Compare-SelectedLocalAndCloudScripts {
    [CmdletBinding()]
    param()

    if ($null -eq $script:SelectedLocalPackage) {
        throw 'Select a local package first.'
    }

    if ($null -eq $script:SelectedCloudRemediation) {
        throw 'Select a cloud remediation first.'
    }

    Ensure-GuiGraphConnection

    $localPackage = $script:SelectedLocalPackage
    $cloudRemediation = $script:SelectedCloudRemediation
    $detail = Get-CloudRemediationDetail -Id ([string]$cloudRemediation.Id)

    $localDetectionContent = if ($script:EditorDetectionPath -eq [string]$localPackage.DetectionPath) {
        [string]$script:Ui.LocalDetectionEditor.Text
    }
    else {
        [IO.File]::ReadAllText([string]$localPackage.DetectionPath)
    }

    $localRemediationPath = [string]$localPackage.RemediationPath
    $localRemediationContent = if ([string]::IsNullOrWhiteSpace($localRemediationPath)) {
        ''
    }
    elseif ($script:EditorRemediationPath -eq $localRemediationPath) {
        [string]$script:Ui.LocalRemediationEditor.Text
    }
    else {
        [IO.File]::ReadAllText($localRemediationPath)
    }

    $cloudDetectionContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'detectionScriptContent')
    $cloudRemediationContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'remediationScriptContent')

    $detectionCompare = Get-ScriptComparisonText -Label 'Detection' -LocalContent $localDetectionContent -CloudContent $cloudDetectionContent
    $remediationCompare = Get-ScriptComparisonText -Label 'Remediation' -LocalContent $localRemediationContent -CloudContent $cloudRemediationContent
    $detectionIdentical = Test-ScriptContentEqual -LeftContent $localDetectionContent -RightContent $cloudDetectionContent
    $remediationIdentical = Test-ScriptContentEqual -LeftContent $localRemediationContent -RightContent $cloudRemediationContent
    $overallStatus = if ($detectionIdentical -and $remediationIdentical) { 'Content match' } else { 'Different' }

    Update-LocalPackageIntuneStatusFromComparison `
        -Package $localPackage `
        -CloudRemediation $cloudRemediation `
        -DetectionIdentical $detectionIdentical `
        -RemediationIdentical $remediationIdentical

    $report = @"
Local package: $($localPackage.DisplayName)
Local folder: $($localPackage.Folder)
Cloud remediation: $($cloudRemediation.DisplayName)
Cloud id: $($cloudRemediation.Id)
Overall: $overallStatus

$detectionCompare

$remediationCompare
"@

    Write-GuiLog ("Compared local package '{0}' with cloud remediation '{1}'. Detection={2}; Remediation={3}; Overall={4}" -f $localPackage.DisplayName, $cloudRemediation.DisplayName, $(if ($detectionIdentical) { 'Identical' } else { 'Different' }), $(if ($remediationIdentical) { 'Identical' } else { 'Different' }), $overallStatus)
    Show-ComparisonWindow -Title 'Compare local and cloud scripts' -Text $report
}

function Export-CloudScriptsToEditor {
    [CmdletBinding()]
    param()

    if ($null -eq $script:SelectedCloudRemediation) {
        throw 'Select a cloud remediation first.'
    }

    Clear-CloudScriptViewer
    $detail = Get-CloudRemediationDetail -Id ([string]$script:SelectedCloudRemediation.Id)
    $script:Ui.CloudDetectionEditor.Text = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'detectionScriptContent')
    $script:Ui.CloudRemediationEditor.Text = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'remediationScriptContent')
    $script:Ui.CloudDetectionEditor.IsReadOnly = $true
    $script:Ui.CloudRemediationEditor.IsReadOnly = $true
    Write-GuiLog 'Cloud script content loaded in the cloud viewer as read-only.'
}

function Save-CloudScriptCopies {
    [CmdletBinding()]
    param()

    if ($null -eq $script:SelectedCloudRemediation) {
        throw 'Select a cloud remediation first.'
    }

    Ensure-GuiGraphConnection

    $displayName = [string]$script:SelectedCloudRemediation.DisplayName
    $targetFolder = Get-CloudScriptCopyFolderPath -DisplayName $displayName
    if ([string]::IsNullOrWhiteSpace($targetFolder)) {
        Write-GuiLog 'Cloud script copy save cancelled.'
        return ''
    }

    Set-GuiBusy -IsBusy $true -Message "Saving cloud script copies for $displayName..."
    try {
        $detail = Get-CloudRemediationDetail -Id ([string]$script:SelectedCloudRemediation.Id)
        $detectionContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'detectionScriptContent')
        $remediationContent = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name 'remediationScriptContent')

        $safeName = ConvertTo-SafeFileName -Value $displayName
        $targetFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($targetFolder)
        [IO.Directory]::CreateDirectory($targetFolder) | Out-Null

        $detectionPath = Join-Path -Path $targetFolder -ChildPath "$safeName-Detection.ps1"
        $remediationPath = Join-Path -Path $targetFolder -ChildPath "$safeName-Remediation.ps1"
        Save-Utf8NoBom -Path $detectionPath -Content $detectionContent

        $savedFiles = @($detectionPath)
        if (-not [string]::IsNullOrWhiteSpace($remediationContent)) {
            Save-Utf8NoBom -Path $remediationPath -Content $remediationContent
            $savedFiles += $remediationPath
        }

        $script:Ui.CloudDetectionEditor.Text = $detectionContent
        $script:Ui.CloudRemediationEditor.Text = $remediationContent
        Write-GuiLog ("Cloud script copies saved for remediation '{0}': {1}" -f $displayName, (($savedFiles | ForEach-Object { [IO.Path]::GetFileName($_) }) -join ', '))
        return ($savedFiles -join "`r`n")
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SmartM365 Intune Remediation"
        Width="1320"
        Height="860"
        MinWidth="1100"
        MinHeight="720"
        WindowStartupLocation="CenterScreen"
        WindowState="Maximized"
        Background="#F5F8FB"
        FontFamily="Segoe UI" FontSize="12" UseLayoutRounding="True" SnapsToDevicePixels="True">
    <Window.Resources>
        <SolidColorBrush x:Key="InkBrush" Color="#1F2937"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#475569"/>
        <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="PanelBorderBrush" Color="#D8E4F0"/>
        <SolidColorBrush x:Key="HeaderBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="AccentDarkBrush" Color="#005A9E"/>
        <SolidColorBrush x:Key="SoftAccentBrush" Color="#E7F3FF"/>

        <Style TargetType="Button">
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Padding" Value="13,8"/>
            <Setter Property="MinWidth" Value="96"/>
            <Setter Property="ToolTipService.ShowOnDisabled" Value="True"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#E8F7FC"/>
                                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#D3EFF8"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource AccentDarkBrush}"/>
                                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="#5ED7FF"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#005F86"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.55"/>
                                <Setter Property="Foreground" Value="#E6F8FF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Foreground" Value="#9A3412"/>
            <Setter Property="BorderBrush" Value="#FDBA74"/>
            <Setter Property="Background" Value="#FFF7ED"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#FFEDD5"/>
                                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="#FB923C"/>
                                <Setter Property="Foreground" Value="#7C2D12"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#FED7AA"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="#0F2437"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="CanUserResizeRows" Value="False"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="RowHeight" Value="34"/>
            <Setter Property="ColumnHeaderHeight" Value="36"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#D8E4F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#EEF4FA"/>
            <Setter Property="AlternatingRowBackground" Value="#F8FBFE"/>
            <Setter Property="RowBackground" Value="#FFFFFF"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#F3F8FD"/>
            <Setter Property="Foreground" Value="#1F2937"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="11,8"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="BorderBrush" Value="#D8E4F0"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="#EDF3F7"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#F8FBFE"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#E6F4FF"/>
                    <Setter Property="Foreground" Value="#005A9E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#0F2437"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="11,0"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="Transparent"/>
                    <Setter Property="Foreground" Value="#005A9E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#475569"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="0,0,6,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder"
                                Background="#F5F8FB"
                                BorderBrush="#D7E3EC"
                                BorderThickness="1"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header"
                                              HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#E6F4FF"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                                <Setter Property="Foreground" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#F8FBFE"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="BorderBrush" Value="#C9D6E2"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
        </Style>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="0"/>
            <RowDefinition Height="2*"/>
            <RowDefinition Height="3*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0"
                CornerRadius="8"
                BorderBrush="{StaticResource PanelBorderBrush}"
                BorderThickness="1"
                Background="{StaticResource PanelBrush}"
                Padding="20"
                Margin="0,0,0,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="190"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <Border Background="{StaticResource SoftAccentBrush}"
                            BorderBrush="#B5DCFF"
                            BorderThickness="1"
                            CornerRadius="8"
                            Padding="10,3"
                            HorizontalAlignment="Left"
                            Margin="0,0,0,14">
                        <TextBlock Text="SMARTM365 INTUNE"
                                   Foreground="{StaticResource AccentDarkBrush}"
                                   FontWeight="SemiBold"/>
                    </Border>
                    <TextBlock Text="SmartM365 Intune Remediation"
                               Foreground="{StaticResource InkBrush}"
                               FontSize="28"
                               FontWeight="SemiBold"
                               Margin="0,0,0,14"/>
                    <StackPanel Orientation="Horizontal">
                        <Button x:Name="ConnectButton" Content="Connect Graph" Style="{StaticResource PrimaryButton}"/>
                        <Button x:Name="RefreshButton" Content="Refresh all"/>
                    </StackPanel>
                </StackPanel>
                <Border x:Name="LogoLink"
                        Grid.Column="1"
                        BorderBrush="{StaticResource PanelBorderBrush}"
                        BorderThickness="1"
                        CornerRadius="8"
                        Background="#F8FBFF"
                        Padding="12"
                        Margin="18,0,18,0"
                        Cursor="Hand"
                        ToolTip="Open WorkplaceCloudHub.com">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                        <Image x:Name="LogoImage" Width="128" Height="72" Stretch="Uniform" SnapsToDevicePixels="True" RenderOptions.BitmapScalingMode="HighQuality"/>
                        <TextBlock Text="workplacecloudhub.com" Foreground="{StaticResource MutedBrush}" FontSize="11" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                    </StackPanel>
                </Border>
                <Border Grid.Column="2"
                        VerticalAlignment="Center"
                        Background="#F8FBFF"
                        BorderBrush="{StaticResource PanelBorderBrush}"
                        BorderThickness="1"
                        CornerRadius="8"
                        Padding="14,9"
                        MinWidth="250">
                    <TextBlock x:Name="StatusText"
                               Foreground="{StaticResource MutedBrush}"
                               FontSize="12"
                               FontWeight="SemiBold"
                               TextAlignment="Right"/>
                </Border>
            </Grid>
        </Border>

        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <GroupBox Grid.Column="0" Header="Local SmartM365 packages" Padding="8">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="8"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <WrapPanel Grid.Row="0">
                        <Button x:Name="AnalyzeLocalButton" Content="PSScriptAnalyzer"/>
                        <Button x:Name="ChooseLocalRootButton" Content="Local folder"/>
                        <Button x:Name="SaveLocalButton" Content="Save script"/>
                        <Button x:Name="SaveAllButton" Content="Save all local"/>
                        <Button x:Name="PublishButton" Content="Publish to Intune"/>
                        <Button x:Name="PublishDetectionOnlyButton" Content="Publish detection only"/>
                        <Button x:Name="PublishAllButton" Content="Publish ALL to Intune"/>
                    </WrapPanel>
                    <DataGrid x:Name="LocalGrid" Grid.Row="2">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" Width="2*"/>
                            <DataGridTextColumn Header="Folder" Binding="{Binding Folder}" Width="1.5*"/>
                            <DataGridTextColumn Header="In Intune" Binding="{Binding IntuneStatus}" Width="130">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                        <Setter Property="ToolTip" Value="{Binding IntuneDetails}"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridCheckBoxColumn Header="Detection only" Binding="{Binding DetectionOnly}" Width="110"/>
                            <DataGridTextColumn Header="PSSA" Binding="{Binding PSScriptAnalyzerStatus}" Width="62">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                        <Setter Property="ToolTip" Value="{Binding PSScriptAnalyzerDetails}"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </GroupBox>

            <GroupBox Grid.Column="2" Header="Intune cloud remediations" Padding="8">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="8"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <WrapPanel Grid.Row="0">
                        <Button x:Name="AnalyzeCloudButton" Content="PSScriptAnalyzer"/>
                        <Button x:Name="SaveCloudScriptsButton" Content="Save scripts as"/>
                        <Button x:Name="SaveAllCloudButton" Content="Save all cloud"/>
                        <Button x:Name="CompareLocalCloudButton" Content="Compare local/cloud"/>
                        <Button x:Name="ResetHistoryButton" Content="Reset history" Style="{StaticResource DangerButton}"/>
                        <Button x:Name="DeleteCloudButton" Content="Delete" Style="{StaticResource DangerButton}"/>
                        <Button x:Name="ExportReportButton" Content="Export execution"/>
                    </WrapPanel>
                    <TabControl Grid.Row="2">
                        <TabItem Header="Active">
                            <DataGrid x:Name="CloudActiveGrid">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Script package name" Binding="{Binding DisplayName}" Width="300" MinWidth="260"/>
                                    <DataGridTextColumn Header="Status" Binding="{Binding PortalStatus}" Width="95" MinWidth="85"/>
                                    <DataGridTextColumn Header="Without issues" Binding="{Binding NoIssueDetectedDeviceCount}" Width="110" MinWidth="105"/>
                                    <DataGridTextColumn Header="With issues" Binding="{Binding IssueDetectedDeviceCount}" Width="95"/>
                                    <DataGridTextColumn Header="Issue fixed" Binding="{Binding IssueRemediatedDeviceCount}" Width="90"/>
                                    <DataGridTextColumn Header="Recurred" Binding="{Binding IssueReoccurredDeviceCount}" Width="85"/>
                                    <DataGridTextColumn Header="Remediated" Binding="{Binding IssueRemediatedCumulativeDeviceCount}" Width="105"/>
                                    <DataGridTextColumn Header="PSSA" Binding="{Binding PSScriptAnalyzerStatus}" Width="70" MinWidth="62">
                                        <DataGridTextColumn.ElementStyle>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="FontWeight" Value="SemiBold"/>
                                                <Setter Property="ToolTip" Value="{Binding PSScriptAnalyzerDetails}"/>
                                            </Style>
                                        </DataGridTextColumn.ElementStyle>
                                    </DataGridTextColumn>
                                </DataGrid.Columns>
                            </DataGrid>
                        </TabItem>
                        <TabItem Header="Not deployed">
                            <DataGrid x:Name="CloudOtherGrid">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Script package name" Binding="{Binding DisplayName}" Width="300" MinWidth="260"/>
                                    <DataGridTextColumn Header="Status" Binding="{Binding PortalStatus}" Width="95" MinWidth="85"/>
                                    <DataGridTextColumn Header="Without issues" Binding="{Binding NoIssueDetectedDeviceCount}" Width="110" MinWidth="105"/>
                                    <DataGridTextColumn Header="With issues" Binding="{Binding IssueDetectedDeviceCount}" Width="95"/>
                                    <DataGridTextColumn Header="Issue fixed" Binding="{Binding IssueRemediatedDeviceCount}" Width="90"/>
                                    <DataGridTextColumn Header="Recurred" Binding="{Binding IssueReoccurredDeviceCount}" Width="85"/>
                                    <DataGridTextColumn Header="Remediated" Binding="{Binding IssueRemediatedCumulativeDeviceCount}" Width="105"/>
                                    <DataGridTextColumn Header="PSSA" Binding="{Binding PSScriptAnalyzerStatus}" Width="70" MinWidth="62">
                                        <DataGridTextColumn.ElementStyle>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="FontWeight" Value="SemiBold"/>
                                                <Setter Property="ToolTip" Value="{Binding PSScriptAnalyzerDetails}"/>
                                            </Style>
                                        </DataGridTextColumn.ElementStyle>
                                    </DataGridTextColumn>
                                </DataGrid.Columns>
                            </DataGrid>
                        </TabItem>
                    </TabControl>
                </Grid>
            </GroupBox>
        </Grid>

        <TabControl Grid.Row="3" Margin="0,12,0,0">
            <TabItem Header="Scripts">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <GroupBox Grid.Column="0" Header="Local script editor" Padding="8">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="105"/>
                                <RowDefinition Height="8"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBox x:Name="PackageDetails"
                                     Grid.Row="0"
                                     IsReadOnly="True"
                                     TextWrapping="Wrap"
                                     VerticalScrollBarVisibility="Auto"/>
                            <TabControl Grid.Row="2">
                                <TabItem Header="Detection">
                                    <TextBox x:Name="LocalDetectionEditor"
                                             AcceptsReturn="True"
                                             AcceptsTab="True"
                                             FontFamily="Consolas"
                                             FontSize="13"
                                             HorizontalScrollBarVisibility="Auto"
                                             VerticalScrollBarVisibility="Auto"/>
                                </TabItem>
                                <TabItem Header="Remediation">
                                    <TextBox x:Name="LocalRemediationEditor"
                                             AcceptsReturn="True"
                                             AcceptsTab="True"
                                             FontFamily="Consolas"
                                             FontSize="13"
                                             HorizontalScrollBarVisibility="Auto"
                                             VerticalScrollBarVisibility="Auto"/>
                                </TabItem>
                            </TabControl>
                        </Grid>
                    </GroupBox>

                    <GroupBox Grid.Column="2" Header="Cloud script viewer" Padding="8">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="105"/>
                                <RowDefinition Height="8"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBox x:Name="CloudDetails"
                                     Grid.Row="0"
                                     IsReadOnly="True"
                                     TextWrapping="Wrap"
                                     VerticalScrollBarVisibility="Auto"/>
                            <TabControl Grid.Row="2">
                                <TabItem Header="Detection">
                                    <TextBox x:Name="CloudDetectionEditor"
                                             IsReadOnly="True"
                                             AcceptsReturn="True"
                                             AcceptsTab="True"
                                             FontFamily="Consolas"
                                             FontSize="13"
                                             HorizontalScrollBarVisibility="Auto"
                                             VerticalScrollBarVisibility="Auto"/>
                                </TabItem>
                                <TabItem Header="Remediation">
                                    <TextBox x:Name="CloudRemediationEditor"
                                             IsReadOnly="True"
                                             AcceptsReturn="True"
                                             AcceptsTab="True"
                                             FontFamily="Consolas"
                                             FontSize="13"
                                             HorizontalScrollBarVisibility="Auto"
                                             VerticalScrollBarVisibility="Auto"/>
                                </TabItem>
                            </TabControl>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>
            <TabItem Header="Activity">
                <TextBox x:Name="LogTextBox"
                         IsReadOnly="True"
                         FontFamily="Consolas"
                         FontSize="12"
                         TextWrapping="NoWrap"
                         HorizontalScrollBarVisibility="Auto"
                         VerticalScrollBarVisibility="Auto"/>
            </TabItem>
        </TabControl>

        <Border x:Name="BusyOverlay"
                Grid.RowSpan="4"
                Panel.ZIndex="100"
                Background="#99091F33"
                Visibility="Collapsed"
                IsHitTestVisible="False">
            <Border Width="420"
                    MinHeight="128"
                    HorizontalAlignment="Center"
                    VerticalAlignment="Center"
                    Background="#FFFFFF"
                    BorderBrush="#8EB8D8"
                    BorderThickness="1"
                    CornerRadius="8"
                    Padding="24">
                <StackPanel>
                    <TextBlock Text="Action in progress"
                               Foreground="#061F36"
                               FontSize="18"
                               FontWeight="SemiBold"
                               HorizontalAlignment="Center"/>
                    <TextBlock x:Name="BusyMessage"
                               Text="Working..."
                               Foreground="#29465E"
                               TextWrapping="Wrap"
                               TextAlignment="Center"
                               Margin="0,10,0,16"/>
                    <ProgressBar x:Name="BusyProgress"
                                 Height="12"
                                 IsIndeterminate="True"/>
                </StackPanel>
            </Border>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$script:Ui['Window'] = $window
$smartM365RootPath = Split-Path -Path (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent) -Parent
$logoIconPath = Join-Path -Path $PSScriptRoot -ChildPath 'WorkplaceCloudHub.ico'
if (Test-Path -LiteralPath $logoIconPath) {
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$logoIconPath)
}
$logoImagePath = Join-Path -Path $smartM365RootPath -ChildPath 'WorkplaceCloudHub.ico'

foreach ($name in @(
        'LogoImage',
        'LogoLink',
        'ConnectButton',
        'RefreshButton',
        'AnalyzeLocalButton',
        'ChooseLocalRootButton',
        'SaveLocalButton',
        'SaveAllButton',
        'PublishButton',
        'PublishDetectionOnlyButton',
        'PublishAllButton',
        'SaveCloudScriptsButton',
        'AnalyzeCloudButton',
        'SaveAllCloudButton',
        'CompareLocalCloudButton',
        'ResetHistoryButton',
        'DeleteCloudButton',
        'ExportReportButton',
        'StatusText',
        'LocalGrid',
        'CloudActiveGrid',
        'CloudOtherGrid',
        'PackageDetails',
        'CloudDetails',
        'LocalDetectionEditor',
        'LocalRemediationEditor',
        'CloudDetectionEditor',
        'CloudRemediationEditor',
        'LogTextBox',
        'BusyOverlay',
        'BusyMessage',
        'BusyProgress'
    )) {
    $script:Ui[$name] = $window.FindName($name)
}

if (Test-Path -LiteralPath $logoImagePath) {
    $script:Ui.LogoImage.Source = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$logoImagePath)
}

function Open-ExternalUrl {
    param([Parameter(Mandatory)][string]$Url)

    try {
        $psi = [System.Diagnostics.ProcessStartInfo]::new($Url)
        $psi.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    }
    catch {
        Show-GuiError -Title 'Open link failed' -Message "Unable to open:`r`n$Url`r`n`r`n$($_.Exception.Message)"
    }
}

if ($script:Ui.LogoLink) {
    $script:Ui.LogoLink.Add_MouseLeftButtonUp({ Open-ExternalUrl -Url 'https://workplacecloudhub.com' })
}

Initialize-LocalRemediationRootConfiguration
Update-ConnectionStatus
Update-ActionButtonsState
if (-not $ValidateOnly -and -not $script:IsGraphConnected) {
    Write-GuiLog 'Graph is not connected. Click Connect Graph before using Intune cloud actions such as publish, reset history, save cloud, or export execution.'
}

$script:Ui.ConnectButton.Add_Click({
    try {
        Set-GuiBusy -IsBusy $true -Message 'Connecting to Microsoft Graph...'
        Connect-GuiGraph
        Set-GuiBusy -IsBusy $true -Message 'Loading Intune cloud remediations...'
        Refresh-CloudGrid
    }
    catch {
        Show-GuiError -Title 'Graph connection failed' -Message (Get-GuiGraphConnectionFailureMessage -ErrorRecord $_)
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
})

$script:Ui.RefreshButton.Add_Click({
    try {
        Set-GuiBusy -IsBusy $true -Message 'Refreshing local SmartM365 packages...'
        Refresh-LocalGrid
        if (-not $script:IsGraphConnected) {
            Write-GuiLog 'Refresh all completed for local packages only. Connect Graph first to refresh Intune cloud remediations.'
            [System.Windows.MessageBox]::Show("Local packages refreshed.`r`n`r`nTo refresh Intune cloud remediations, click Connect Graph first.", 'Graph connection required', 'OK', 'Information') | Out-Null
            return
        }

        Set-GuiBusy -IsBusy $true -Message 'Refreshing Intune cloud remediations...'
        Refresh-CloudGrid
    }
    catch {
        Show-GuiError -Title 'Refresh failed' -Message $_.Exception.Message
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
})

$script:Ui.AnalyzeLocalButton.Add_Click({
    try {
        Test-LocalEditorScriptsWithPSScriptAnalyzer
        $packageName = if ($null -eq $script:SelectedLocalPackage) { 'selected package' } else { [string]$script:SelectedLocalPackage.DisplayName }
        [System.Windows.MessageBox]::Show("Local PSScriptAnalyzer completed without blocking errors for:`r`n$packageName", 'PSScriptAnalyzer', 'OK', 'Information') | Out-Null
    }
    catch {
        Write-GuiLog "PSScriptAnalyzer failed - $($_.Exception.Message)"
    }
})

$script:Ui.ChooseLocalRootButton.Add_Click({
    try {
        $selectedPath = Select-LocalRemediationRoot
        if ([string]::IsNullOrWhiteSpace($selectedPath)) { return }
        Refresh-LocalGrid
    }
    catch {
        Show-GuiError -Title 'Local folder selection failed' -Message $_.Exception.Message
    }
})

$script:Ui.SaveLocalButton.Add_Click({
    try { Save-EditorFiles }
    catch { Show-GuiError -Title 'Save failed' -Message $_.Exception.Message }
})

$script:Ui.SaveAllButton.Add_Click({
    try {
        $zipPath = Export-LocalRemediationsArchive
        if ([string]::IsNullOrWhiteSpace($zipPath)) { return }
        [System.Windows.MessageBox]::Show("Local archive created:`r`n$zipPath", 'Save all local completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'Save all local failed' -Message $_.Exception.Message
    }
})

$script:Ui.PublishButton.Add_Click({
    try {
        $packageToPublish = $script:SelectedLocalPackage
        if ($null -eq $packageToPublish) { throw 'Select a local package first.' }
        Save-EditorFiles
        $publishedId = Publish-LocalPackage -Package $packageToPublish -ConfirmExistingUpdate
        if ([string]::IsNullOrWhiteSpace($publishedId)) { return }
        Refresh-CloudGrid
        $publishedDisplayName = ConvertTo-PublishedDisplayName -DisplayName ([string]$packageToPublish.DisplayName)
        [System.Windows.MessageBox]::Show("Publish completed:`r`n$publishedDisplayName`r`nLocal package: $($packageToPublish.DisplayName)`r`nId: $publishedId", 'Publish completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'Publish failed' -Message $_.Exception.Message
    }
})

$script:Ui.PublishDetectionOnlyButton.Add_Click({
    try {
        $packageToPublish = $script:SelectedLocalPackage
        if ($null -eq $packageToPublish) { throw 'Select a local package first.' }

        $publishedDisplayName = ConvertTo-PublishedDisplayName -DisplayName ([string]$packageToPublish.DisplayName)
        $message = @"
Publish detection only for this Intune remediation?

Published name: $publishedDisplayName
Local package: $($packageToPublish.DisplayName)

Only the detection script will be sent to Intune.
The remediation script content in Intune will be empty for this package.
Local remediation files are not deleted.

Continue?
"@
        $answer = [System.Windows.MessageBox]::Show($message, 'Publish detection only', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') {
            Write-GuiLog "Publish detection only cancelled: $publishedDisplayName"
            return
        }

        Save-DetectionEditorFile
        $detectionOnlyPackage = New-DetectionOnlyPublishPackage -Package $packageToPublish
        $publishedId = Publish-LocalPackage -Package $detectionOnlyPackage -ConfirmExistingUpdate
        if ([string]::IsNullOrWhiteSpace($publishedId)) { return }
        Refresh-CloudGrid
        [System.Windows.MessageBox]::Show("Publish detection only completed:`r`n$publishedDisplayName`r`nLocal package: $($packageToPublish.DisplayName)`r`nId: $publishedId", 'Publish detection only completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'Publish detection only failed' -Message $_.Exception.Message
    }
})

$script:Ui.PublishAllButton.Add_Click({
    try {
        if ($null -ne $script:SelectedLocalPackage -and -not [string]::IsNullOrWhiteSpace($script:EditorDetectionPath)) {
            Save-EditorFiles
        }

        $packages = @(Get-LocalPackages)
        if ($packages.Count -eq 0) { throw 'No local package found to publish.' }

        $publishPlan = Get-PublishAllLocalPackagesPlan -Packages $packages
        $message = Format-PublishAllLocalPackagesPlanMessage -Plan $publishPlan
        if (@($publishPlan.Duplicate).Count -gt 0) {
            [System.Windows.MessageBox]::Show($message, 'Publish all blocked', 'OK', 'Error') | Out-Null
            return
        }

        $answer = [System.Windows.MessageBox]::Show($message, 'Publish all local packages', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return }

        $publishedIds = Publish-AllLocalPackages -Packages $packages
        Refresh-CloudGrid
        [System.Windows.MessageBox]::Show("Publish all completed:`r`n$(@($publishedIds).Count) local package(s) published.", 'Publish all completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'Publish all failed' -Message $_.Exception.Message
    }
})

$script:Ui.SaveCloudScriptsButton.Add_Click({
    try {
        $savedFiles = Save-CloudScriptCopies
        if ([string]::IsNullOrWhiteSpace($savedFiles)) { return }
        [System.Windows.MessageBox]::Show("Cloud script copies saved:`r`n$savedFiles", 'Save script copies completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'Save script copies failed' -Message $_.Exception.Message
    }
})

$script:Ui.SaveAllCloudButton.Add_Click({
    try {
        $zipPath = Export-CloudRemediationsArchive
        if ([string]::IsNullOrWhiteSpace($zipPath)) { return }
        [System.Windows.MessageBox]::Show("Cloud archive created:`r`n$zipPath", 'Save all cloud completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'Save all cloud failed' -Message $_.Exception.Message
    }
})

$script:Ui.CompareLocalCloudButton.Add_Click({
    try {
        Set-GuiBusy -IsBusy $true -Message 'Comparing selected local and cloud scripts...'
        Compare-SelectedLocalAndCloudScripts
    }
    catch {
        Show-GuiError -Title 'Compare local/cloud failed' -Message $_.Exception.Message
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
})

$script:Ui.AnalyzeCloudButton.Add_Click({
    try {
        Test-CloudViewerScriptsWithPSScriptAnalyzer
        $remediationName = if ($null -eq $script:SelectedCloudRemediation) { 'selected remediation' } else { [string]$script:SelectedCloudRemediation.DisplayName }
        [System.Windows.MessageBox]::Show("Cloud PSScriptAnalyzer completed without blocking errors for:`r`n$remediationName", 'PSScriptAnalyzer', 'OK', 'Information') | Out-Null
    }
    catch {
        Write-GuiLog "Cloud PSScriptAnalyzer failed - $($_.Exception.Message)"
    }
})

$script:Ui.ResetHistoryButton.Add_Click({
    try {
        if ($null -eq $script:SelectedCloudRemediation) { throw 'Select a cloud remediation first.' }
        $displayName = [string]$script:SelectedCloudRemediation.DisplayName
        $message = "Create a duplicate of '$displayName', copy assignments, then delete the old Intune object? This resets execution history."
        $answer = [System.Windows.MessageBox]::Show($message, 'Reset execution history', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return }
        $replacement = Reset-CloudRemediationHistory -CloudRemediation $script:SelectedCloudRemediation
        Refresh-CloudGrid
        [System.Windows.MessageBox]::Show("Execution history reset completed:`r`nOld remediation deleted: $displayName`r`nReplacement kept as: $($replacement.DisplayName)`r`nReplacement Id: $($replacement.Id)", 'History reset completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'History reset failed' -Message $_.Exception.Message
    }
})

$script:Ui.DeleteCloudButton.Add_Click({
    try {
        if ($null -eq $script:SelectedCloudRemediation) { throw 'Select a cloud remediation first.' }

        $displayName = [string]$script:SelectedCloudRemediation.DisplayName
        $id = [string]$script:SelectedCloudRemediation.Id
        $message = @"
You are about to permanently delete this Intune remediation:

Name: $displayName
Id: $id

This deletes only the selected cloud remediation object from Intune.
It does not create a replacement, does not reset history safely, and does not delete local scripts.

Continue with deletion?
"@
        $answer = [System.Windows.MessageBox]::Show($message, 'Delete Intune remediation', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') {
            Write-GuiLog "Delete cloud remediation cancelled: $displayName ($id)"
            return
        }

        Set-GuiBusy -IsBusy $true -Message "Deleting Intune remediation: $displayName"
        Remove-CloudRemediation -CloudRemediation $script:SelectedCloudRemediation
        $script:Ui.CloudActiveGrid.SelectedItem = $null
        $script:Ui.CloudOtherGrid.SelectedItem = $null
        Set-EditorFromCloudRemediation -CloudRemediation $null
        Refresh-CloudGrid
        [System.Windows.MessageBox]::Show("Cloud remediation deleted:`r`n$displayName`r`nId: $id", 'Delete completed', 'OK', 'Information') | Out-Null
    }
    catch {
        Show-GuiError -Title 'Delete failed' -Message $_.Exception.Message
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
})

$script:Ui.ExportReportButton.Add_Click({
    try {
        if ($null -eq $script:SelectedCloudRemediation) { throw 'Select a cloud remediation first.' }
        $format = Show-ExecutionReportFormatDialog -ExcelAvailable $script:IsImportExcelAvailable
        if ([string]::IsNullOrWhiteSpace($format)) {
            Write-GuiLog 'Execution report export cancelled.'
            return
        }

        $result = Export-ExecutionReport -CloudRemediation $script:SelectedCloudRemediation -Format $format
        if ($null -eq $result -or [string]::IsNullOrWhiteSpace([string]$result.Path)) { return }

        if ([string]$result.Format -eq 'Excel') {
            $openResult = [System.Windows.MessageBox]::Show(
                "Excel workbook exported:`r`n$($result.Path)`r`n`r`nOpen it now?",
                'Execution report exported',
                'YesNo',
                'Information'
            )
            if ($openResult -eq 'Yes') {
                Start-Process -FilePath ([string]$result.Path)
            }
        }
        else {
            [System.Windows.MessageBox]::Show("CSV exported:`r`n$($result.Path)", 'Execution report exported', 'OK', 'Information') | Out-Null
        }
    }
    catch {
        Show-GuiError -Title 'Report export failed' -Message $_.Exception.Message
    }
})

$script:Ui.LocalGrid.Add_SelectionChanged({
    try {
        Set-EditorFromLocalPackage -Package $script:Ui.LocalGrid.SelectedItem
    }
    catch {
        Show-GuiError -Title 'Local package load failed' -Message $_.Exception.Message
    }
})

$script:Ui.CloudActiveGrid.Add_SelectionChanged({
    try {
        if ($null -ne $script:Ui.CloudActiveGrid.SelectedItem) {
            $script:Ui.CloudOtherGrid.SelectedItem = $null
            Set-EditorFromCloudRemediation -CloudRemediation $script:Ui.CloudActiveGrid.SelectedItem
        }
    }
    catch {
        Show-GuiError -Title 'Cloud remediation selection failed' -Message $_.Exception.Message
    }
})

$script:Ui.CloudOtherGrid.Add_SelectionChanged({
    try {
        if ($null -ne $script:Ui.CloudOtherGrid.SelectedItem) {
            $script:Ui.CloudActiveGrid.SelectedItem = $null
            Set-EditorFromCloudRemediation -CloudRemediation $script:Ui.CloudOtherGrid.SelectedItem
        }
    }
    catch {
        Show-GuiError -Title 'Cloud remediation selection failed' -Message $_.Exception.Message
    }
})

if ($ValidateOnly) {
    try {
        Refresh-LocalGrid
    }
    catch {
        Write-GuiLog $_.Exception.Message
    }

    Write-Output 'Smart Intune Remediation GUI validation completed.'
    return
}

$script:InitialLocalLoadCompleted = $false
$window.Add_ContentRendered({
    if ($script:GuiSplash) {
        Hide-SmartM365GuiSplash -Splash $script:GuiSplash
    }

    if ($script:InitialLocalLoadCompleted) { return }
    $script:InitialLocalLoadCompleted = $true
    try {
        Initialize-ImportExcelSupport | Out-Null
        Set-GuiBusy -IsBusy $true -Message 'Loading local packages and running PSScriptAnalyzer...'
        Refresh-LocalGrid
    }
    catch {
        Write-GuiLog $_.Exception.Message
    }
    finally {
        Set-GuiBusy -IsBusy $false
    }
})

[void]$window.ShowDialog()
if ($script:GuiSplash) {
    Close-SmartM365GuiSplash -Splash $script:GuiSplash
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
$script:IsGraphConnected = $false
Update-ConnectionStatus

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZu/jg0HMQ/n56
# RGKiTXOMfYzpjiVkPdQjziQjQJTZLaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDXYZTsAIMRjLAJ9xJl
# 3VkFLZLpIHnqzroOIspQleO8NzANBgkqhkiG9w0BAQEFAASCAYBRpUQDlwCoEWk9
# Ml86/YZw/Zbed4Yj69Ty37g201hHE19V/UVZPT+dNo0VkdiduYK/W2o6f5WHKnDc
# tyDdgRMdR/C9COS+eAXUWAYeciBhO5JpWMqSvoNSJ4/hrwfDfRveSJIMl7gS2H4q
# PpexjIy1ge3kHvQs1eSwwvxsIqJyVAVITKMkoJlPpP4YNUXeASYcZINTgFoc9yIJ
# zGrQ9rxzW5O/uCxuffb5r4EbNHyDlflqdVE8A8xLnVXFPAXKBYhpDFIrwc4v2pH4
# bmEjpejaw8hn/1NLRo12qtCfLOPVBQNOqh8hp3S/fhyDevEoF4CtyODNwf8v9Cg3
# EJDMUXLuM10ibkmGl+BIDaNOVX6rt+uj1MVPI6OnagVqNwtawrTsRJtl+pY3rBMa
# /jxaeLsL4TZt9oLMm/W7uOa2MCBc8gUvyFjIMjzO76k1A03FpOCM+V6NyQJRSs8i
# Vl0Bc+VpG1sGQFWmbKVQOVdFu7uiEFD/96ECx/BSBGXkzf6O/k4=
# SIG # End signature block
