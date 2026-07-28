<#
.SYNOPSIS
Provides the Automatic LOT workflow for the Intune Hybrid Join WPF launcher.

.VERSION
1.0.0
#>

Set-StrictMode -Version Latest

function Get-AutomaticGuiDialogOwner {
    if ($script:AutomaticInventoryProgressState -and
        $script:AutomaticInventoryProgressState.Window -and
        $script:AutomaticInventoryProgressState.Window.IsVisible) {
        return $script:AutomaticInventoryProgressState.Window
    }
    return $window
}

function Show-GuiError {
    param([string]$Message)
    [System.Windows.MessageBox]::Show((Get-AutomaticGuiDialogOwner), $Message, 'SmartM365', 'OK', 'Error') | Out-Null
}

function Show-AutomaticGuiWarningYesNo {
    param(
        [string]$Message,
        [string]$Title
    )

    $result = [System.Windows.MessageBox]::Show((Get-AutomaticGuiDialogOwner), $Message, $Title, 'YesNo', 'Warning')
    return ($result -eq [System.Windows.MessageBoxResult]::Yes)
}

function New-AutomaticInventoryProgressState {
    param(
        [string]$Title = 'Preparing automatic Hybrid Join LOT preview',
        [string]$InitialStage = 'Checking inventory caches...',
        [string]$InitialDetail = 'Please wait.'
    )

    $progressXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="500" Height="235" WindowStyle="None" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="#F5F8FB" FontFamily="Segoe UI" FontSize="12"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
    <Border BorderBrush="#B9DDF7" BorderThickness="1" CornerRadius="10" Background="White" Padding="22">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock x:Name="ProgressTitleText" FontSize="18" FontWeight="SemiBold" Foreground="#1F2937"/>
            <TextBlock Grid.Row="1" x:Name="ProgressStageText" Margin="0,14,0,0" FontWeight="SemiBold" Foreground="#005A9E" TextWrapping="Wrap"/>
            <TextBlock Grid.Row="2" x:Name="ProgressDetailText" Margin="0,6,0,0" Foreground="#475569" TextWrapping="Wrap" MaxHeight="62"/>
            <ProgressBar Grid.Row="3" Height="8" Margin="0,18,0,0" IsIndeterminate="True" Foreground="#0078D4"/>
            <TextBlock Grid.Row="4" x:Name="ProgressElapsedText" Text="Elapsed: 0 s" Margin="0,10,0,0" Foreground="#64748B" HorizontalAlignment="Right"/>
        </Grid>
    </Border>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$progressXaml)
    $progressWindow = [Windows.Markup.XamlReader]::Load($reader)
    $progressWindow.Owner = $window
    $progressWindow.FindName('ProgressTitleText').Text = $Title
    $progressWindow.FindName('ProgressStageText').Text = $InitialStage
    $progressWindow.FindName('ProgressDetailText').Text = $InitialDetail
    [pscustomobject]@{
        Window = $progressWindow
        StageText = $progressWindow.FindName('ProgressStageText')
        DetailText = $progressWindow.FindName('ProgressDetailText')
        ElapsedText = $progressWindow.FindName('ProgressElapsedText')
        StartedUtc = [datetime]::UtcNow
        Result = $null
        ErrorRecord = $null
    }
}

function Update-AutomaticInventoryProgress {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$Stage,
        [string]$Detail = ''
    )

    if (-not $State.Window -or -not $State.Window.IsVisible) { return }
    $State.StageText.Text = $Stage
    $State.DetailText.Text = if ([string]::IsNullOrWhiteSpace($Detail)) { 'Please wait.' } else { $Detail }
    $elapsedSeconds = [math]::Max(0, [math]::Floor(([datetime]::UtcNow - $State.StartedUtc).TotalSeconds))
    $State.ElapsedText.Text = "Elapsed: $elapsedSeconds s"
    $State.Window.UpdateLayout()
    $State.Window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Get-AutomaticInventoryFileInfo {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][double]$FreshnessHours,
        [Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra')][string]$SourceName,
        [string]$ExpectedTenantId = ''
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{
            Source = $SourceName; Path = $Path; Exists = $false; Fresh = $false
            AgeHours = [double]::PositiveInfinity; TenantId = ''; AuthenticationMode = ''
            Scope = ''; ContentVerified = $false; Detail = 'Cache missing'
        }
    }

    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $ageHours = ((Get-Date) - $item.LastWriteTime).TotalHours
    $firstRow = @(Import-Csv -LiteralPath $Path | Select-Object -First 1)
    $row = if ($firstRow.Count -gt 0) { $firstRow[0] } else { $null }
    $contentVerified = $false
    $tenantId = ''
    $authenticationMode = ''
    $scope = ''
    if ($row) {
        if ($SourceName -eq 'AD') {
            $hasName = ($row.PSObject.Properties['ComputerName'] -or $row.PSObject.Properties['Name'])
            $contentVerified = [bool]($hasName -and $row.PSObject.Properties['Enabled'] -and $row.PSObject.Properties['OperatingSystem'])
        }
        else {
            $tenantId = if ($row.PSObject.Properties['InventoryTenantId']) { [string]$row.InventoryTenantId } else { '' }
            $authenticationMode = if ($row.PSObject.Properties['InventoryAuthenticationMode']) { [string]$row.InventoryAuthenticationMode } else { '' }
            $scope = if ($row.PSObject.Properties['InventoryScope']) { [string]$row.InventoryScope } else { '' }
            $expectedScope = if ($SourceName -eq 'Intune') { 'AllManagedDevices' } else { 'AllEntraDevices' }
            $tenantMatches = (-not [string]::IsNullOrWhiteSpace($tenantId) -and
                ([string]::IsNullOrWhiteSpace($ExpectedTenantId) -or $tenantId -ieq $ExpectedTenantId))
            $contentVerified = ($tenantMatches -and
                $authenticationMode -in @('DelegatedInteractive','DelegatedExistingSession') -and
                $scope -eq $expectedScope)
        }
    }

    $fresh = ($ageHours -le $FreshnessHours -and $contentVerified)
    $detail = "Age={0:N1}h; TTL={1}h; content={2}" -f $ageHours, $FreshnessHours, $contentVerified
    if ($SourceName -ne 'AD') {
        $detail += "; tenant=$tenantId; auth=$authenticationMode; scope=$scope"
    }
    [pscustomobject]@{
        Source = $SourceName
        Path = $Path
        Exists = $true
        Fresh = $fresh
        AgeHours = $ageHours
        TenantId = $tenantId
        AuthenticationMode = $authenticationMode
        Scope = $scope
        ContentVerified = $contentVerified
        Detail = $detail
    }
}

function Get-AutomaticProcessLogDetail {
    param(
        [string[]]$Paths,
        [int]$TailLines = 10
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $content = (@(Get-Content -LiteralPath $path -Tail $TailLines -ErrorAction SilentlyContinue) -join [Environment]::NewLine).Trim()
        if (-not [string]::IsNullOrWhiteSpace($content)) { $parts.Add($content) }
    }
    return ($parts.ToArray() -join [Environment]::NewLine)
}

function Invoke-AutomaticPowerShellProcess {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$Activity,
        [switch]$Interactive,
        [scriptblock]$ProgressCallback
    )

    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) { throw "Script not found: $ScriptPath" }
    $stderrPath = "$LogPath.stderr.txt"
    $argumentParts = New-Object System.Collections.Generic.List[string]
    foreach ($value in @('-NoProfile','-ExecutionPolicy','Bypass','-File',$ScriptPath) + @($Arguments)) {
        $argumentParts.Add((ConvertTo-CmdArgument -Value ([string]$value)))
    }

    if ($ProgressCallback) { & $ProgressCallback $Activity 'Starting the PowerShell inventory process...' }
    $startParameters = @{
        FilePath = 'powershell.exe'
        ArgumentList = ($argumentParts -join ' ')
        PassThru = $true
        RedirectStandardOutput = $LogPath
        RedirectStandardError = $stderrPath
    }
    if (-not $Interactive) { $startParameters.WindowStyle = 'Hidden' }
    $process = Start-Process @startParameters
    $nextUpdateUtc = [datetime]::MinValue
    while (-not $process.HasExited) {
        $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        if ($ProgressCallback -and [datetime]::UtcNow -ge $nextUpdateUtc) {
            & $ProgressCallback $Activity (Get-AutomaticProcessLogDetail -Paths @($LogPath,$stderrPath) -TailLines 3)
            $nextUpdateUtc = [datetime]::UtcNow.AddSeconds(1)
        }
        Start-Sleep -Milliseconds 200
        $process.Refresh()
    }
    $process.Refresh()
    if ($process.ExitCode -ne 0) {
        $detail = Get-AutomaticProcessLogDetail -Paths @($stderrPath,$LogPath) -TailLines 20
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'The inventory process produced no error text.' }
        throw "Inventory refresh failed with exit code $($process.ExitCode). $detail"
    }
}

function Get-AutomaticInventorySnapshot {
    param(
        [switch]$ForceRefresh,
        [scriptblock]$ProgressCallback
    )

    $automaticRoot = Join-Path (Get-RunsRoot -RootPath $toolkitRoot) 'AutomaticLotInventory'
    $sourceRunPath = Join-Path $automaticRoot ('Sources-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmssfff'))
    New-Item -ItemType Directory -Path $sourceRunPath -Force | Out-Null
    $paths = @{ AD = ''; Intune = ''; Entra = '' }
    $details = New-Object System.Collections.Generic.List[string]
    $configuredTenantId = Get-ConfiguredValue 'EHJIR_INTUNE_TENANT_ID'

    if ($ProgressCallback) { & $ProgressCallback 'Checking root inventory caches...' 'AD TTL: 12h; Intune and Entra TTL: 2h.' }
    $adRoot = Join-Path $toolkitRoot 'DevicesAD.csv'
    $adOutput = Join-Path $sourceRunPath 'DevicesAD.csv'
    $adInfo = Get-AutomaticInventoryFileInfo -Path $adRoot -FreshnessHours 12 -SourceName AD
    if (-not $ForceRefresh -and $adInfo.Fresh) {
        Copy-Item -LiteralPath $adRoot -Destination $adOutput -Force
        $paths.AD = $adOutput
        $details.Add("AD: reused recent root cache; $($adInfo.Detail); snapshot=$adOutput")
    }
    else {
        try {
            $adExporter = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1'
            Invoke-AutomaticPowerShellProcess -ScriptPath $adExporter -Arguments @(
                '-OutputPath',$adOutput,'-ForceRefresh'
            ) -LogPath (Join-Path $sourceRunPath 'DevicesAD.refresh.log') -Activity 'Reading a fresh complete AD forest inventory...' -ProgressCallback $ProgressCallback
            $generatedAdInfo = Get-AutomaticInventoryFileInfo -Path $adOutput -FreshnessHours 12 -SourceName AD
            if (-not $generatedAdInfo.Fresh) { throw "Generated AD inventory is invalid: $($generatedAdInfo.Detail)" }
            $paths.AD = $adOutput
            $details.Add("AD: generated isolated automatic snapshot $adOutput")
        }
        catch {
            $useFallback = $false
            if ($adInfo.Fresh) {
                $useFallback = Show-AutomaticGuiWarningYesNo -Title 'AD refresh failed' -Message ((@(
                    'The automatic AD snapshot could not be generated:'
                    $_.Exception.Message
                    ''
                    "A verified recent root cache is available: $adRoot"
                    $adInfo.Detail
                    ''
                    'Use this root cache as fallback?'
                )) -join [Environment]::NewLine)
            }
            if (-not $useFallback) { throw }
            Copy-Item -LiteralPath $adRoot -Destination $adOutput -Force
            $paths.AD = $adOutput
            $details.Add("AD: explicitly accepted recent root fallback; snapshot=$adOutput")
        }
    }

    $intuneRoot = Join-Path $toolkitRoot 'DevicesIntune.csv'
    $entraRoot = Join-Path $toolkitRoot 'DevicesEntra.csv'
    $intuneOutput = Join-Path $sourceRunPath 'DevicesIntune.csv'
    $entraOutput = Join-Path $sourceRunPath 'DevicesEntra.csv'
    $intuneInfo = Get-AutomaticInventoryFileInfo -Path $intuneRoot -FreshnessHours 2 -SourceName Intune -ExpectedTenantId $configuredTenantId
    $entraInfo = Get-AutomaticInventoryFileInfo -Path $entraRoot -FreshnessHours 2 -SourceName Entra -ExpectedTenantId $configuredTenantId
    $graphCachesMatch = ($intuneInfo.Fresh -and $entraInfo.Fresh -and $intuneInfo.TenantId -ieq $entraInfo.TenantId)

    if (-not $ForceRefresh -and $graphCachesMatch) {
        Copy-Item -LiteralPath $intuneRoot -Destination $intuneOutput -Force
        Copy-Item -LiteralPath $entraRoot -Destination $entraOutput -Force
        $paths.Intune = $intuneOutput
        $paths.Entra = $entraOutput
        $details.Add("Intune: reused verified root cache; $($intuneInfo.Detail); snapshot=$intuneOutput")
        $details.Add("Entra: reused verified root cache; $($entraInfo.Detail); snapshot=$entraOutput")
    }
    else {
        try {
            $graphExporter = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-Export-AutomaticGraphInventories.ps1'
            $graphArguments = @('-IntuneOutputPath',$intuneOutput,'-EntraOutputPath',$entraOutput)
            if (-not [string]::IsNullOrWhiteSpace($configuredTenantId)) {
                $graphArguments += @('-TenantId',$configuredTenantId)
            }
            Invoke-AutomaticPowerShellProcess -ScriptPath $graphExporter -Arguments $graphArguments `
                -LogPath (Join-Path $sourceRunPath 'GraphInventories.refresh.log') `
                -Activity 'Waiting for delegated Graph sign-in and inventory export...' `
                -Interactive -ProgressCallback $ProgressCallback

            $generatedIntuneInfo = Get-AutomaticInventoryFileInfo -Path $intuneOutput -FreshnessHours 2 -SourceName Intune -ExpectedTenantId $configuredTenantId
            if (-not $generatedIntuneInfo.Fresh) { throw "Generated Intune inventory is invalid: $($generatedIntuneInfo.Detail)" }
            $paths.Intune = $intuneOutput
            $details.Add("Intune: generated delegated snapshot; $($generatedIntuneInfo.Detail); snapshot=$intuneOutput")
            $generatedEntraInfo = Get-AutomaticInventoryFileInfo -Path $entraOutput -FreshnessHours 2 -SourceName Entra -ExpectedTenantId $generatedIntuneInfo.TenantId
            if ($generatedEntraInfo.Fresh) {
                $paths.Entra = $entraOutput
                $details.Add("Entra: generated delegated snapshot; $($generatedEntraInfo.Detail); snapshot=$entraOutput")
            }
            else {
                Remove-Item -LiteralPath $entraOutput -Force -ErrorAction SilentlyContinue
                $details.Add("Entra: optional enrichment unavailable; $($generatedEntraInfo.Detail)")
            }
        }
        catch {
            $useFallback = $false
            if ($intuneInfo.Fresh) {
                $useFallback = Show-AutomaticGuiWarningYesNo -Title 'Graph refresh failed' -Message ((@(
                    'The automatic Graph snapshot could not be generated:'
                    $_.Exception.Message
                    ''
                    "A verified recent Intune root cache is available: $intuneRoot"
                    $intuneInfo.Detail
                    ''
                    'Use this Intune cache as fallback? Entra enrichment will be reused only when its tenant and provenance match.'
                )) -join [Environment]::NewLine)
            }
            if (-not $useFallback) { throw }
            Copy-Item -LiteralPath $intuneRoot -Destination $intuneOutput -Force
            $paths.Intune = $intuneOutput
            $details.Add("Intune: explicitly accepted recent root fallback; snapshot=$intuneOutput")
            if ($entraInfo.Fresh -and $entraInfo.TenantId -ieq $intuneInfo.TenantId) {
                Copy-Item -LiteralPath $entraRoot -Destination $entraOutput -Force
                $paths.Entra = $entraOutput
                $details.Add("Entra: reused matching recent root cache; snapshot=$entraOutput")
            }
            else {
                $details.Add('Entra: optional enrichment unavailable after Graph refresh failure.')
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$paths.AD)) { throw 'Required AD inventory is unavailable.' }
    if ([string]::IsNullOrWhiteSpace([string]$paths.Intune)) { throw 'Required Intune inventory is unavailable.' }
    if ($ProgressCallback) { & $ProgressCallback 'Inventory sources ready.' ($details.ToArray() -join [Environment]::NewLine) }
    [pscustomobject]@{
        AdInventoryCsv = [string]$paths.AD
        IntuneInventoryCsv = [string]$paths.Intune
        EntraInventoryCsv = [string]$paths.Entra
        EntraAvailable = (-not [string]::IsNullOrWhiteSpace([string]$paths.Entra))
        SourceRunPath = $sourceRunPath
        SourceDetails = $details.ToArray()
        ForceInventoryRefresh = [bool]$ForceRefresh
    }
}

function Invoke-AutomaticLotSelection {
    param(
        [Parameter(Mandatory = $true)]$InventoryContext,
        [switch]$Create,
        [scriptblock]$ProgressCallback
    )

    $engine = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-New-AutomaticLot.ps1'
    if (-not (Test-Path -LiteralPath $engine -PathType Leaf)) { throw "Automatic LOT engine not found: $engine" }
    $parameters = @{
        AdInventoryCsv = [string]$InventoryContext.AdInventoryCsv
        IntuneInventoryCsv = [string]$InventoryContext.IntuneInventoryCsv
        EntraInventoryCsv = [string]$InventoryContext.EntraInventoryCsv
        ToolkitRoot = $toolkitRoot
        LotName = [string]$controls.AutomaticLotNameText.Text
        ComputerNamePrefix = @([string]$controls.AutomaticNamePrefixText.Text)
        ComputerNameContains = @([string]$controls.AutomaticNameContainsText.Text)
        AdLastLogonMaxAgeDays = [int](Get-IntText -TextBox $controls.AutomaticLastLogonDaysText -Default 45 -Minimum 1)
        EvidenceRoot = Join-Path (Get-RunsRoot -RootPath $toolkitRoot) 'AutomaticLotInventory'
    }
    if ([bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked) { $parameters.ExcludeStaleAd = $true }
    if ($Create) { $parameters.Create = $true }
    if ($ProgressCallback) { $parameters.ProgressCallback = $ProgressCallback }
    $results = @(& $engine @parameters)
    $summaryResults = @($results | Where-Object { $_ -and $_.PSObject.Properties['Summary'] })
    if ($summaryResults.Count -ne 1) {
        throw "Automatic LOT engine returned $($summaryResults.Count) result object(s); exactly one is required."
    }
    return $summaryResults[0]
}

function Format-AutomaticLotSummary {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)]$InventoryContext
    )

    $summary = $Result.Summary
    @(
        "Selected: $($summary.SelectedDevices) device(s)"
        "Needs Hybrid Join: $($summary.NeedsHybridJoin)"
        "Hybrid Join pending: $($summary.HybridJoinPending)"
        "Needs Intune enrollment: $($summary.NeedsIntuneEnrollment)"
        "Safety exclusions: $($summary.ExcludedDevices)"
        "Name/stale filters: $($summary.FilterExcludedDevices)"
        "Already in Intune: $($summary.IntunePresentExcluded)"
        "Disabled AD: $($summary.ADDisabledExcluded)"
        "Windows Server: $($summary.WindowsServerExcluded)"
        "Ambiguous AD/Entra: $([int]$summary.ADNameCollisions + [int]$summary.EntraAmbiguousExcluded)"
        "Entra enrichment: $(if ($InventoryContext.EntraAvailable) { 'Available' } else { 'Unavailable - local diagnosis will decide' })"
        ''
        ($InventoryContext.SourceDetails -join [Environment]::NewLine)
    ) -join [Environment]::NewLine
}

function Get-AutomaticPreviewSignature {
    @(
        ([string]$controls.AutomaticNamePrefixText.Text).Trim().ToUpperInvariant()
        ([string]$controls.AutomaticNameContainsText.Text).Trim().ToUpperInvariant()
        [bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked
        (Get-IntText -TextBox $controls.AutomaticLastLogonDaysText -Default 45 -Minimum 1)
        [bool]$controls.AutomaticForceRefreshCheck.IsChecked
    ) -join '|'
}

function Get-AutomaticGeneratedLotName {
    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @(([string]$controls.AutomaticNamePrefixText.Text) -split ';')) {
        $segment = [regex]::Replace($candidate.Trim().ToUpperInvariant(), '[^A-Z0-9_-]+', '-').Trim('-_')
        if (-not [string]::IsNullOrWhiteSpace($segment) -and -not $segments.Contains($segment)) { $segments.Add($segment) }
    }
    $prefixSegment = if ($segments.Count -gt 0) { "-$($segments -join '-')" } else { '' }
    return 'LOT-AUTO-IHJ{0}-{1}' -f $prefixSegment, $script:AutomaticLotNameTimestamp
}

function Update-AutomaticGeneratedLotName {
    param([switch]$Force)

    $generated = Get-AutomaticGeneratedLotName
    $current = [string]$controls.AutomaticLotNameText.Text
    if ($Force -or [string]::IsNullOrWhiteSpace($current) -or $current -eq $script:AutomaticLastGeneratedLotName) {
        $controls.AutomaticLotNameText.Text = $generated
        $script:AutomaticLastGeneratedLotName = $generated
    }
}

function Set-AutomaticPreviewStale {
    $script:AutomaticPreviewSignature = ''
    $controls.AutomaticCreateButton.IsEnabled = $false
    $controls.AutomaticSummaryText.Text = 'Filters changed. Refresh the preview before creating the LOT.'
}

function Update-AutomaticLotPreview {
    param(
        [switch]$ForceInventoryRefresh,
        [scriptblock]$ProgressCallback
    )

    $context = Get-AutomaticInventorySnapshot -ForceRefresh:$ForceInventoryRefresh -ProgressCallback $ProgressCallback
    if ($ProgressCallback) { & $ProgressCallback 'Loading inventories and building the selection preview...' $context.SourceRunPath }
    $result = Invoke-AutomaticLotSelection -InventoryContext $context -ProgressCallback $ProgressCallback
    $script:AutomaticPreviewContext = $context
    $script:AutomaticPreviewResult = $result
    $script:AutomaticPreviewSignature = Get-AutomaticPreviewSignature
    $controls.AutomaticSummaryText.Text = Format-AutomaticLotSummary -Result $result -InventoryContext $context
    $controls.AutomaticEvidencePathText.Text = [string]$result.Summary.EvidencePath
    $controls.AutomaticCreateButton.IsEnabled = ([int]$result.Summary.SelectedDevices -gt 0)
    Add-Status -Title 'Automatic preview' -Message ("Selected {0} Hybrid Join repair candidate(s); excluded {1}." -f $result.Summary.SelectedDevices,$result.Summary.ExcludedDevices)
    return $result
}

function Invoke-AutomaticModalOperation {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Stage,
        [string]$Detail = 'Please wait.',
        [Parameter(Mandatory = $true)][scriptblock]$Operation
    )

    $state = New-AutomaticInventoryProgressState -Title $Title -InitialStage $Stage -InitialDetail $Detail
    $script:AutomaticInventoryProgressState = $state
    $progressCallback = {
        param([string]$ProgressStage,[string]$ProgressDetail)
        Update-AutomaticInventoryProgress -State $state -Stage $ProgressStage -Detail $ProgressDetail
    }.GetNewClosure()
    $work = {
        try { $state.Result = & $Operation $progressCallback }
        catch { $state.ErrorRecord = $_ }
        finally {
            if ($state.Window -and $state.Window.IsVisible) { $state.Window.Close() }
        }
    }.GetNewClosure()
    $state.Window.Add_ContentRendered({
        [void]$state.Window.Dispatcher.BeginInvoke([action]$work, [System.Windows.Threading.DispatcherPriority]::Background)
    }.GetNewClosure())
    try { [void]$state.Window.ShowDialog() }
    finally { $script:AutomaticInventoryProgressState = $null }
    if ($state.ErrorRecord) { throw $state.ErrorRecord }
    return $state.Result
}

function Confirm-AutomaticLotCreate {
    param([Parameter(Mandatory = $true)]$Result)

    $summary = $Result.Summary
    $message = @(
        "Create $($controls.AutomaticLotNameText.Text)?"
        ''
        "Selected devices: $($summary.SelectedDevices)"
        "Needs Hybrid Join: $($summary.NeedsHybridJoin)"
        "Hybrid Join pending: $($summary.HybridJoinPending)"
        "Needs Intune enrollment: $($summary.NeedsIntuneEnrollment)"
        "Safety exclusions: $($summary.ExcludedDevices)"
        "Filtered out: $($summary.FilterExcludedDevices)"
        "Entra enrichment available: $($script:AutomaticPreviewContext.EntraAvailable)"
        ''
        'The LOT will be created but not launched.'
    ) -join [Environment]::NewLine
    return Show-AutomaticGuiWarningYesNo -Title 'Create automatic Hybrid Join LOT' -Message $message
}

function Initialize-AutomaticLotGui {
    $script:AutomaticInventoryProgressState = $null
    $script:AutomaticPreviewContext = $null
    $script:AutomaticPreviewResult = $null
    $script:AutomaticPreviewSignature = ''
    $script:AutomaticLotNameTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:AutomaticLastGeneratedLotName = ''
    $controls.AutomaticLastLogonDaysText.Text = '45'
    $controls.AutomaticLastLogonDaysText.IsEnabled = $false
    $controls.AutomaticCreateButton.IsEnabled = $false
    $controls.AutomaticSummaryText.Text = 'No automatic selection has been calculated yet.'
    Update-AutomaticGeneratedLotName -Force

    $controls.AutomaticNamePrefixText.Add_TextChanged({
        Update-AutomaticGeneratedLotName
        Set-AutomaticPreviewStale
    })
    $controls.AutomaticNameContainsText.Add_TextChanged({ Set-AutomaticPreviewStale })
    $controls.AutomaticForceRefreshCheck.Add_Click({ Set-AutomaticPreviewStale })
    $controls.AutomaticExcludeStaleAdCheck.Add_Click({
        $controls.AutomaticLastLogonDaysText.IsEnabled = [bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked
        Set-AutomaticPreviewStale
    })
    $controls.AutomaticLastLogonDaysText.Add_TextChanged({ Set-AutomaticPreviewStale })
    $controls.AutomaticPreviewButton.Add_Click({
        try {
            $force = [bool]$controls.AutomaticForceRefreshCheck.IsChecked
            [void](Invoke-AutomaticModalOperation -Title 'Preparing automatic Hybrid Join LOT preview' `
                -Stage 'Checking inventory caches...' `
                -Operation { param($callback) Update-AutomaticLotPreview -ForceInventoryRefresh:$force -ProgressCallback $callback }.GetNewClosure())
            if ($force) {
                $controls.AutomaticForceRefreshCheck.IsChecked = $false
                $script:AutomaticPreviewSignature = Get-AutomaticPreviewSignature
            }
        }
        catch {
            Add-Status -Title 'Automatic preview failed' -Message $_.Exception.Message
            Show-GuiError $_.Exception.Message
        }
    })
    $controls.AutomaticCreateButton.Add_Click({
        try {
            $currentSignature = Get-AutomaticPreviewSignature
            if ($null -eq $script:AutomaticPreviewResult -or $script:AutomaticPreviewSignature -ne $currentSignature) {
                throw 'The automatic preview is stale. Click Refresh and preview first.'
            }
            if ([int]$script:AutomaticPreviewResult.Summary.SelectedDevices -le 0) { throw 'No eligible device is selected.' }
            if (-not (Confirm-AutomaticLotCreate -Result $script:AutomaticPreviewResult)) { return }
            $lotName = [string]$controls.AutomaticLotNameText.Text
            $creation = Invoke-AutomaticModalOperation -Title 'Creating automatic Hybrid Join LOT' `
                -Stage 'Validating the confirmed selection...' -Detail $lotName `
                -Operation {
                    param($callback)
                    $created = Invoke-AutomaticLotSelection -InventoryContext $script:AutomaticPreviewContext -Create -ProgressCallback $callback
                    & $callback 'Reading the created LOT...' ([string]$created.Summary.LotPath)
                    $folder = Get-Item -LiteralPath ([string]$created.Summary.LotPath) -ErrorAction Stop
                    $lot = Get-LotSummary -Folder $folder
                    & $callback 'Refreshing the LOT list...' 'The LOT is created but will not be launched.'
                    $script:SelectedLot = $lot
                    Refresh-LotList
                    [pscustomobject]@{ Result = $created; Lot = $lot }
                }
            $script:AutomaticPreviewResult = $creation.Result
            $controls.AutomaticSummaryText.Text = Format-AutomaticLotSummary -Result $creation.Result -InventoryContext $script:AutomaticPreviewContext
            $controls.AutomaticEvidencePathText.Text = [string]$creation.Result.Summary.EvidencePath
            $controls.MainTabs.SelectedIndex = 0
            Add-Status -Title 'Automatic LOT created' -Message ("Created {0} with {1} device(s). Open Existing LOT to launch it." -f $creation.Lot.Name,$creation.Lot.DeviceCount)
        }
        catch {
            Add-Status -Title 'Automatic LOT creation failed' -Message $_.Exception.Message
            Show-GuiError $_.Exception.Message
        }
    })
    $controls.AutomaticOpenEvidenceButton.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace([string]$controls.AutomaticEvidencePathText.Text)) {
            Open-FolderPath -Path ([string]$controls.AutomaticEvidencePathText.Text)
        }
    })
}

Initialize-AutomaticLotGui
