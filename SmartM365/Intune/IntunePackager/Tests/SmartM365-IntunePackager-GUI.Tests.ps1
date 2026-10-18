#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
.SYNOPSIS
    Pester 5 tests for the GUI worker result handling and UI exception guards.

.DESCRIPTION
    The functions are extracted from SmartM365-IntunePackager-GUI.ps1 through its AST, so the
    script entry point is not executed. WPF tests need an STA thread; the test runner relaunches
    itself with -STA when needed.

    Regression covered: an empty Win32 app search result was turned into $null by a truthiness
    test on the worker output, bound as a $null DataGrid row, and WPF raised
    ArgumentNullException (Parameter 'key') during layout, which closed the window.

.VERSION
    1.2.0
#>

BeforeDiscovery {
    $script:IsSta = [Threading.Thread]::CurrentThread.GetApartmentState() -eq [Threading.ApartmentState]::STA
}

BeforeAll {
    $script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Force
    Set-IpLogPath -Path ''
    $guiPath = Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.ps1'
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "GUI script parse failed: $guiPath" }
    foreach ($name in @('Get-IpJobResult', 'ConvertTo-IpGridItems', 'Invoke-IpGuarded', 'Register-IpDispatcherExceptionHandler', 'Get-IpInnermostException', 'Get-IpNextContentVersion', 'Get-IpUpdateBannerText', 'Get-IpUpdateConfirmationText', 'Resolve-IpInitialTenant', 'Get-IpTenantProfileState', 'ConvertTo-IpAppGridRow',
            'Get-IpObjectValue', 'Set-IpSearchResults', 'Invoke-IpShowSelectedAppDetails', 'Invoke-IpUseSelectedAppForUpdate', 'Invoke-IpAppDetailsResult', 'Set-IpSelectedAppDetailsText',
            'Test-IpDefinitionLoaded', 'Get-IpNoDefinitionMessage', 'Get-IpSignInBannerText', 'Invoke-IpFlashWindow')) {
        $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        if (-not $definition) { throw "GUI function not found: $name" }
        . ([scriptblock]::Create($definition.Extent.Text))
    }

    # Result callbacks are scriptblock arguments of Start-IpJob: extract them as written in the GUI.
    function Get-GuiJobCallback {
        param([string]$FunctionName, [string]$Parameter)
        $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName }, $true)
        $call = $function.Find({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Start-IpJob' }, $true)
        $elements = @($call.CommandElements)
        for ($i = 0; $i -lt $elements.Count - 1; $i++) {
            if ($elements[$i] -is [Management.Automation.Language.CommandParameterAst] -and $elements[$i].ParameterName -eq $Parameter) {
                return $elements[$i + 1].ScriptBlock.GetScriptBlock()
            }
        }
        throw "Callback $Parameter not found in $FunctionName"
    }
    $script:SearchCallback = Get-GuiJobCallback -FunctionName 'Invoke-IpFindApps' -Parameter 'OnSuccess'
    $script:SearchWorker = Get-GuiJobCallback -FunctionName 'Invoke-IpFindApps' -Parameter 'Script'

    # UI stubs used by the extracted functions.
    function Show-IpMessage { param($Text, $Title, $Icon, $Buttons) $script:Messages.Add([string]$Text); return 'OK' }
    function Set-IpStatus { param($Text) $script:Status = $Text }
    function Start-IpJob { param($Title, $Script, $Arguments, $Context, $OnSuccess, $OnFailure) $script:StartedJobs.Add($Title) }

    function New-FakeUi {
        $script:Ui = @{
            AppsGrid               = [pscustomobject]@{ ItemsSource = $null; SelectedItem = $null }
            SelectedAppDetailsText = [pscustomobject]@{ Text = '' }
        }
        $script:State = @{ Connected = $true; Job = $null; AppDetails = @{}; UpdateTarget = $null; ConnectedTenant = 'prod (t1)' }
        $script:Messages = [Collections.Generic.List[string]]::new()
        $script:StartedJobs = [Collections.Generic.List[string]]::new()
    }

    function Invoke-SearchResultLikeGui {
        param([object[]]$Apps, [string]$Filter)
        # Same path as the GUI: worker output -> Get-IpJobResult -> guarded OnSuccess callback.
        $script:WorkerApps = $Apps
        $output = Invoke-WorkerScript ('$apps = @($script:WorkerAppsJson | ConvertFrom-Json); , $apps' -replace '\$script:WorkerAppsJson', ("'" + ((ConvertTo-Json -InputObject @($Apps) -Depth 5 -Compress) -replace "'", "''") + "'"))
        $result = Get-IpJobResult -Output $output
        Invoke-IpGuarded -Context 'Search Win32 apps result' -ScriptBlock $script:SearchCallback -Arguments @($result, @{ Filter = $Filter })
    }

    function Invoke-WorkerScript {
        param([string]$Script)
        $powerShell = [PowerShell]::Create()
        try {
            [void]$powerShell.AddScript($Script)
            return , $powerShell.EndInvoke($powerShell.BeginInvoke())
        }
        finally { $powerShell.Dispose() }
    }

    function Start-LogCapture {
        $script:Logged = [Collections.Generic.List[string]]::new()
        Set-IpLogSink -Sink { param($line, $level) $script:Logged.Add($line) } -Console $false
    }

    function New-TestWindow {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        $window = [Windows.Window]::new()
        $window.Width = 500; $window.Height = 300; $window.Left = -3000; $window.Top = -3000
        $window.ShowInTaskbar = $false; $window.WindowStyle = 'None'; $window.ShowActivated = $false
        $panel = [Windows.Controls.StackPanel]::new()
        $grid = [Windows.Controls.DataGrid]::new()
        $grid.AutoGenerateColumns = $false
        $column = [Windows.Controls.DataGridTextColumn]::new()
        $column.Binding = [Windows.Data.Binding]::new('displayName')
        [void]$grid.Columns.Add($column)
        $button = [Windows.Controls.Button]::new()
        [void]$panel.Children.Add($button)
        [void]$panel.Children.Add($grid)
        $window.Content = $panel
        return [pscustomobject]@{ Window = $window; Grid = $grid; Button = $button }
    }

    function Invoke-DispatcherPump {
        param($Window, [int]$Rounds = 5)
        for ($i = 0; $i -lt $Rounds; $i++) { $Window.Dispatcher.Invoke([Action] {}, [Windows.Threading.DispatcherPriority]::ApplicationIdle) }
    }
}

AfterAll { Set-IpLogSink -Sink $null -Console $false }

Describe 'Worker result extraction' {
    It 'reproduces the empty search result case: the former truthiness test returned $null' {
        $output = Invoke-WorkerScript '$apps = @(); , $apps'
        $output.Count | Should -Be 1
        [bool]($output -and $output.Count) | Should -BeFalse
        $result = Get-IpJobResult -Output $output
        $null -eq $result | Should -BeFalse
        , $result | Should -BeOfType [object[]]
        $result.Count | Should -Be 0
    }

    It 'returns $null when the worker emitted nothing' {
        Get-IpJobResult -Output (Invoke-WorkerScript '$null = 1') | Should -BeNullOrEmpty
        Get-IpJobResult -Output $null | Should -BeNullOrEmpty
    }

    It 'returns the last emitted object' {
        $result = Get-IpJobResult -Output (Invoke-WorkerScript "'first'; [pscustomobject]@{ AppId = 'a1' }")
        $result.AppId | Should -Be 'a1'
    }

    It 'keeps a one-element array as an array' {
        $result = Get-IpJobResult -Output (Invoke-WorkerScript ", @([pscustomobject]@{ id = 'x1' })")
        , $result | Should -BeOfType [object[]]
        $result[0].id | Should -Be 'x1'
    }
}

Describe 'Grid items' {
    It 'never yields $null rows' -ForEach @(
        @{ Name = 'null'; Items = $null; Count = 0 }
        @{ Name = 'empty array'; Items = @(); Count = 0 }
        @{ Name = 'array with a null'; Items = @($null, [pscustomobject]@{ id = 'a' }); Count = 1 }
        @{ Name = 'single object'; Items = [pscustomobject]@{ id = 'b' }; Count = 1 }
    ) {
        $items = ConvertTo-IpGridItems $Items
        , $items | Should -BeOfType [object[]]
        $items.Count | Should -Be $Count
        @($items | Where-Object { $null -eq $_ }).Count | Should -Be 0
    }
}

Describe 'Guarded UI callbacks' {
    It 'logs an exception instead of throwing it' {
        Start-LogCapture
        { Invoke-IpGuarded -Context 'Search Win32 apps result' -ScriptBlock { param($result) throw "boom $result" } -Arguments @('x') } | Should -Not -Throw
        ($script:Logged -join "`n") | Should -Match '\[ERROR\] Search Win32 apps result failed: boom x'
    }

    It 'passes the result and the context to the callback' {
        $script:Seen = $null
        Invoke-IpGuarded -Context 'c' -ScriptBlock { param($result, $context) $script:Seen = '{0}|{1}' -f $result, $context.Key } -Arguments @('r', @{ Key = 'k' })
        $script:Seen | Should -Be 'r|k'
    }
}

Describe 'WPF regression' -Skip:(-not $script:IsSta) {
    It 'binds an empty search result to the DataGrid without an error' {
        $test = New-TestWindow
        $script:UiErrors = [Collections.Generic.List[string]]::new()
        $test.Window.Dispatcher.Add_UnhandledException({ param($source, $eventArgs) $eventArgs.Handled = $true; $script:UiErrors.Add($eventArgs.Exception.GetBaseException().Message) })
        $test.Window.Show()
        try {
            $result = Get-IpJobResult -Output (Invoke-WorkerScript '$apps = @(); , $apps')
            $test.Grid.ItemsSource = ConvertTo-IpGridItems $result
            Invoke-DispatcherPump $test.Window
            $script:UiErrors.Count | Should -Be 0
            $test.Grid.Items.Count | Should -Be 0
        }
        finally { $test.Window.Close() }
    }

    It 'logs handler and layout exceptions once and keeps the window open' {
        $test = New-TestWindow
        Start-LogCapture
        Register-IpDispatcherExceptionHandler -Dispatcher $test.Window.Dispatcher
        $test.Button.Add_Click({ throw 'search handler failure' })
        $test.Window.Show()
        try {
            [void]$test.Window.Dispatcher.BeginInvoke([Action] { $test.Button.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent)) })
            # The former bug: a $null row makes WPF throw ArgumentNullException on every layout pass.
            [void]$test.Window.Dispatcher.BeginInvoke([Action] { $test.Grid.ItemsSource = @($null) })
            { Invoke-DispatcherPump $test.Window } | Should -Not -Throw
            $test.Window.IsVisible | Should -BeTrue
            $errors = @($script:Logged | Where-Object { $_ -match 'Unexpected UI error' })
            @($errors | Where-Object { $_ -match 'search handler failure' }).Count | Should -Be 1
            @($errors | Where-Object { $_ -match "ArgumentNullException: Value cannot be null. \(Parameter 'key'\)" }).Count | Should -Be 1
        }
        finally {
            $test.Grid.ItemsSource = $null
            $test.Window.Close()
        }
    }
}

Describe 'Update workflow helpers' {
    It 'computes the next content version' -ForEach @(
        @{ Current = '1'; Next = '2' }, @{ Current = ' 9 '; Next = '10' }, @{ Current = ''; Next = '1' }, @{ Current = 'v2'; Next = 'next' }
    ) {
        Get-IpNextContentVersion $Current | Should -Be $Next
    }

    It 'describes the update target in the banner' {
        $target = [pscustomobject]@{ AppId = 'a1'; DisplayName = 'Contoso Sample App'; ContentVersion = '1' }
        Get-IpUpdateBannerText -Target $target -AppId 'a1' | Should -Be 'Update mode: Contoso Sample App (a1), current content version 1. Click Publish package to upload content version 2.'
        Get-IpUpdateBannerText -Target $target -AppId 'other' | Should -Match '^Update mode: app other\.'
        Get-IpUpdateBannerText -Target $null -AppId '' | Should -Match 'search for the app above'
    }

    It 'confirms an update with the app name, its id and the version change' {
        $details = [pscustomobject]@{ id = 'a1'; displayName = 'Contoso Sample App'; committedContentVersion = '1' }
        $text = Get-IpUpdateConfirmationText -Details $details -ProfileDisplayName 'Contoso Sample App' -Tenant 'prod (t1)' -PackagePath 'C:\Out\ContosoSampleApp.intunewin'
        $text | Should -Match 'App: Contoso Sample App'
        $text | Should -Match 'App id: a1'
        $text | Should -Match 'Committed content version: 1 -> 2'
        $text | Should -Match 'Package: ContosoSampleApp\.intunewin'
        $text | Should -Not -Match 'display name becomes'
    }

    It 'warns in the confirmation when the profile renames the app' {
        $details = [pscustomobject]@{ id = 'a1'; displayName = 'Old name'; committedContentVersion = '' }
        $text = Get-IpUpdateConfirmationText -Details $details -ProfileDisplayName 'New name' -Tenant 't' -PackagePath 'x.intunewin'
        $text | Should -Match 'Committed content version: none -> 1'
        $text | Should -Match "display name becomes 'New name'"
    }

    It 'formats the grid row with a local last-modified time' {
        $utc = [datetime]::new(2026, 10, 18, 9, 31, 0, [DateTimeKind]::Utc)
        $row = ConvertTo-IpAppGridRow ([pscustomobject]@{ id = 'a1'; displayName = 'Contoso'; publisher = 'Contoso Ltd'; lastModifiedDateTime = $utc })
        $row.LastModified | Should -Be $utc.ToLocalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    }
}

Describe 'Tenant selection' {
    It 'prefers an explicit -Tenant, then the last used profile' {
        Resolve-IpInitialTenant -Names @('test', 'prod') -Requested 'test' -RequestedExplicitly $true -LastUsed 'prod' | Should -Be 'test'
        Resolve-IpInitialTenant -Names @('test', 'prod') -Requested 'test' -RequestedExplicitly $false -LastUsed 'prod' | Should -Be 'prod'
        Resolve-IpInitialTenant -Names @('test', 'prod') -Requested 'test' -RequestedExplicitly $false -LastUsed 'gone' | Should -Be 'test'
        Resolve-IpInitialTenant -Names @('prod') -Requested 'test' -RequestedExplicitly $true -LastUsed '' | Should -Be 'prod'
    }

    It 'reports a profile without TenantId as not configured' {
        $config = '{"Tenants":{"test":{"TenantId":"","AuthMode":"Delegated"},"prod":{"TenantId":"b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e","AuthMode":"Delegated"},"bad":{"TenantId":"<id>","AuthMode":"Delegated"}}}' | ConvertFrom-Json
        $test = Get-IpTenantProfileState -Config $config -Name 'test'
        $test.State | Should -Be 'NotConfigured'
        $test.Message | Should -Be "Tenant profile 'test' is not configured: its TenantId is empty."
        (Get-IpTenantProfileState -Config $config -Name 'prod').State | Should -Be 'Configured'
        (Get-IpTenantProfileState -Config $config -Name 'bad').State | Should -Be 'Invalid'
        (Get-IpTenantProfileState -Config $config -Name 'missing').Message | Should -Match "tenant profile 'missing' was not found"
    }
}

Describe 'Error unwrapping' {
    It 'reports the error behind ActionPreferenceStopException and MethodInvocationException' {
        # Same chain as the worker: ErrorActionPreference=Stop, then EndInvoke on the UI thread.
        $powerShell = [PowerShell]::Create()
        try {
            [void]$powerShell.AddScript('$ErrorActionPreference = "Stop"; Write-Error "InteractiveBrowserCredential authentication failed: User canceled authentication."')
            $failure = $null
            try { [void]$powerShell.EndInvoke($powerShell.BeginInvoke()) } catch { $failure = $_ }
        }
        finally { $powerShell.Dispose() }
        $failure.Exception | Should -BeOfType [Management.Automation.MethodInvocationException]
        $failure.Exception.InnerException | Should -BeOfType [Management.Automation.ActionPreferenceStopException]
        (Get-IpInnermostException $failure).Message | Should -Be 'InteractiveBrowserCredential authentication failed: User canceled authentication.'
        (Get-IpInnermostException ([Exception]::new('plain'))).Message | Should -Be 'plain'
    }
}

Describe 'Search result callback (regression: 0 result)' {
    BeforeEach { New-FakeUi; Start-LogCapture }

    It 'shows an explicit message for an empty result without any error' {
        Invoke-SearchResultLikeGui -Apps @() -Filter 'constoso'
        @($script:Logged | Where-Object { $_ -match '\[ERROR\]' }).Count | Should -Be 0
        @($script:Ui.AppsGrid.ItemsSource).Count | Should -Be 0
        $script:Ui.SelectedAppDetailsText.Text | Should -Match "^No Win32 app matches 'constoso'\."
        ($script:Logged -join "`n") | Should -Match "No Win32 app matches 'constoso'\."
    }

    It 'binds one result as one row' {
        Invoke-SearchResultLikeGui -Apps @([pscustomobject]@{ id = 'a1'; displayName = 'Contoso Sample App'; publisher = 'Contoso Ltd'; lastModifiedDateTime = '2026-10-18T09:31:00Z' }) -Filter 'contoso'
        @($script:Logged | Where-Object { $_ -match '\[ERROR\]' }).Count | Should -Be 0
        $rows = @($script:Ui.AppsGrid.ItemsSource)
        $rows.Count | Should -Be 1
        $rows[0].id | Should -Be 'a1'
        $rows[0].LastModified | Should -Match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$'
    }

    It 'binds several results and skips rows without id' {
        Invoke-SearchResultLikeGui -Apps @([pscustomobject]@{ id = 'a1'; displayName = 'A' }, [pscustomobject]@{ displayName = 'No id' }, [pscustomobject]@{ id = 'a2'; displayName = 'B' }) -Filter ''
        @($script:Ui.AppsGrid.ItemsSource.id) | Should -Be @('a1', 'a2')
    }

    It 'accepts a $null result' {
        Invoke-IpGuarded -Context 'Search Win32 apps result' -ScriptBlock $script:SearchCallback -Arguments @($null, @{ Filter = 'x' })
        @($script:Logged | Where-Object { $_ -match '\[ERROR\]' }).Count | Should -Be 0
        $script:Ui.SelectedAppDetailsText.Text | Should -Match "No Win32 app matches 'x'"
    }

    It 'returns an empty array from the real worker script for 0 matches' {
        $output = Invoke-WorkerScript ("function Find-IpWin32App { param(`$DisplayName, [switch]`$Contains) }`nfunction Write-IpLog { param(`$Message) }`n& {`n" + $script:SearchWorker.ToString() + "`n} 'constoso'")
        $result = Get-IpJobResult -Output $output
        , $result | Should -BeOfType [object[]]
        $result.Count | Should -Be 0
    }
}

Describe 'Selection and details with an empty or partial result' {
    BeforeEach { New-FakeUi; Start-LogCapture }

    It 'ignores a selection change on an empty grid' {
        { Invoke-IpShowSelectedAppDetails } | Should -Not -Throw
        $script:StartedJobs.Count | Should -Be 0
    }

    It 'ignores a selected row without id' {
        $script:Ui.AppsGrid.SelectedItem = [pscustomobject]@{ displayName = 'No id' }
        { Invoke-IpShowSelectedAppDetails } | Should -Not -Throw
        $script:StartedJobs.Count | Should -Be 0
    }

    It 'asks to select an app before using it for update' {
        { Invoke-IpUseSelectedAppForUpdate } | Should -Not -Throw
        $script:Messages | Should -Contain 'Select an app in the list first.'
    }

    It 'handles empty app details without error' {
        { Invoke-IpAppDetailsResult -Result $null -Context @{ Purpose = 'Show' } } | Should -Not -Throw
        { Invoke-IpAppDetailsResult -Result ([pscustomobject]@{}) -Context @{ Purpose = 'Refresh' } } | Should -Not -Throw
        @($script:Logged | Where-Object { $_ -match 'App details were empty' }).Count | Should -Be 2
    }

    It 'shows the details of the selected app' {
        $script:Ui.AppsGrid.SelectedItem = [pscustomobject]@{ id = 'a1' }
        Invoke-IpAppDetailsResult -Result ([pscustomobject]@{ id = 'a1'; displayName = 'Contoso'; displayVersion = '2.4.0'; committedContentVersion = '2'; setupFilePath = 'x.msi'; publishingState = 'published'; isAssigned = 'False'; lastModifiedDateTime = $null }) -Context @{ Purpose = 'Show' }
        $script:Ui.SelectedAppDetailsText.Text | Should -Match 'committed content version: 2'
        $script:State.AppDetails['a1'].displayVersion | Should -Be '2.4.0'
    }
}

Describe 'Definition loaded check' {
    It 'detects an empty form and explains what to do in update mode' {
        $definition = New-IpAppDefinition
        Test-IpDefinitionLoaded -Definition $definition | Should -BeFalse
        Get-IpNoDefinitionMessage -PublishMode 'UpdateExisting' | Should -Be 'No app definition loaded. Open the profile used for this app, or import the definition from Intune.'
        Get-IpNoDefinitionMessage -PublishMode 'Create' | Should -Match '^No app definition loaded\.'
        $definition.App.DisplayName = 'Contoso'
        Test-IpDefinitionLoaded -Definition $definition | Should -BeTrue
    }
}

Describe 'Sign-in banner' {
    It 'tells the operator where to find the WAM window' {
        Get-IpSignInBannerText -SignInMode 'Wam' | Should -Be "Sign-in window opened (Windows account broker). If you don't see it, check the taskbar or press Alt+Tab."
        Get-IpSignInBannerText -SignInMode 'Browser' | Should -Match 'web browser'
    }

    It 'flashes and stops flashing a window without error' -Skip:(-not $script:IsSta) {
        $test = New-TestWindow
        $test.Window.Show()
        try {
            { Invoke-IpFlashWindow -Window $test.Window } | Should -Not -Throw
            { Invoke-IpFlashWindow -Window $test.Window -Stop } | Should -Not -Throw
            ('SmartM365.IntunePackager.NativeWindow' -as [type]) | Should -Not -BeNullOrEmpty
        }
        finally { $test.Window.Close() }
    }
}