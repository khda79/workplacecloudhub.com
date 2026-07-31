<#
.SYNOPSIS
Provides the Automatic LOT workflow for the Intune Hybrid Join WPF launcher.

.VERSION
1.0.1
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
    $controls.AutomaticSummaryText.Text = 'Filters changed. Create will refresh the preview before creating the LOT.'
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
            $force = [bool]$controls.AutomaticForceRefreshCheck.IsChecked
            if ($force -or $null -eq $script:AutomaticPreviewResult -or $script:AutomaticPreviewSignature -ne $currentSignature) {
                [void](Invoke-AutomaticModalOperation -Title 'Preparing automatic Hybrid Join LOT preview' `
                    -Stage 'Checking inventory caches...' `
                    -Operation { param($callback) Update-AutomaticLotPreview -ForceInventoryRefresh:$force -ProgressCallback $callback }.GetNewClosure())
                if ($force) {
                    $controls.AutomaticForceRefreshCheck.IsChecked = $false
                    $script:AutomaticPreviewSignature = Get-AutomaticPreviewSignature
                }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBajlST3Ciljxwx
# 7k4o2ydenPAEsikfXunp3hwYfZr6NKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJrOLIiJl60XAnra0Oamk9UDG7DN4eFK0dezNmhkg69vMA0GCSqG
# SIb3DQEBAQUABIIBgGwf0EciaXjFTTewq7irgraSq0Jm8w+phaEf/PF5NQQZqYGH
# MoBntvzDltD3aq6COFgcxBOqx3ABi813bmvHlSxqMmvq7dLz4RyUtYYJWZL8eNPR
# qDcmXuo1Cxytnw8AGWH0BuefRW6ZKsn//0WwJ1pPt82e0RTI6gy5zTI0XBWP0rjy
# UaTajGUYuUp30FEIcvoszBen995Z+Phfd9Y0VDoCjkN9lZin18o5ZCiO0UIe13R1
# X7FGcgfo9tYbMZQv4tbQhCYC03hIDZ7lrC94VBBECypBpg1VZpk34Vyz9E3J/z+H
# R9dvSgioR9JvNiZGl6tjOBEF9wusLPK6WlljDZARA/1nBB9I+AsPjX2x/CTqJJoX
# PT3qcq3N0Hxo6WZtELsFXKlrmq+8baMvBLELWLQCfM4wXvzjnQiJGlw+sJC1ywRd
# 9orBz9kffuFy4oA3wp7buzt+qNbtmvLY1qhvxj6W+pEm4FR+GgebyTam/ojzEfr3
# bxIbPStB2B6GUY4696GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MzEwNjU4
# NTBaMC8GCSqGSIb3DQEJBDEiBCDCU8vrz7rSXRTkim5XBD6jOORVEMD7Xr6qJOe7
# ZjJdMDANBgkqhkiG9w0BAQEFAASCAgBOhJoCa14b1eNicPDi+2nD3ThSL7cqdKZl
# 6KWgNs+1x8SVLyl1/RlZ5hI5wx14kPeHh3/HImwoi5JxDI4TDwb2+T9fYgDKvMJw
# RfPnLg6G+P91Iugb8e1++JiQFTNwf0E6iyNHA3/fm/abXp53pm1vW8CS6eQMmpYs
# PLJBFvkwPDsuVe6BZO+DvkOVS73OLR4rBcJGW4gllEKNVYYA039IOONeThX6EREv
# 9fU89IMhOCTD9v6B7l51sSytnQEs1ZqVHBHj1dWaBxA9Nky5uwzD0xqH626yflVy
# 7gNKgQVlMmpwzMLPzEIedUWABrGJIPmRArReQL0gKLK/Q6v20qfucIZ/BZrJhWLU
# CoyLCchcv62YOuTgiaVzxhSvMffa3py3srN9gbiq7WJS8zjmGSbRNpRiCP8XFJMI
# fR0bb2vlKbsWiwAGYYOgFLtoi7+tHQ9nsMEB0O2J/VyqmDIfW0QeKtTGRZfRAMM6
# DXMjzRKTzQle2SkiCXdG8XhIfJTc6G9WBawjnXxbnqUAD9gftfdnEkhiL4vlATVd
# ZX/xlwoVXgqv0YMf7DlP4Q5Xq4sLGP6mZyWGWXe1yFBC+ceLC8QFyT+AjhMBJn9T
# 1HYnqtqjYHctIVg6putQ9d76V/wmAOzgcrUm5U4GxZQy0ODYXkGT5XU2TL+ld++M
# FlZCOqMs4A==
# SIG # End signature block
