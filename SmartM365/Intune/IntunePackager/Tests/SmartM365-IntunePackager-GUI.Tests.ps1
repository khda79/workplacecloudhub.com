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
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCaOLt8KcePyIH+
# 8AMVszJ2eJui3kcSc56I73V5P3i3CaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMFerOtOtxT42hs2s4+9px06hON41FUNL9vlM7j+pqZiMA0GCSqG
# SIb3DQEBAQUABIIBgCysX7A2ohxzqsFPEu0wszqDW7HY2zwseYOj3MlKN6wAMZlm
# GAbF7ef+Ia+JY7XoWQ6Eym487f2wYpzGuWsgkQ6bqw+l5tlOdtLWlXIUD59fGcF0
# 93X8uveFgPgm9AuCY+U+srVTXUbUd21mg6uvDyo0vUjcTNDKeiVAbTNpvh7RJ1Z5
# RHTVOoV5PFpmdOWXmXjECS21TzMqBBbmbrwIePr/QowzJYjMuZOd9U3ebm1xwDAH
# 0B3WWGmSlUFwI1hS0uGYI2hNFLlRotXucSBZlYcxgrqpVUMrnjHl7J57nBYTM9/g
# 0d8HxkNcoUh2XnsnBVASbQWHy9v8uLsp4TsoY+0TOS/O4jWwigGspSo/BsGgHvBl
# mUA+FoLwfVsNIVubsU9l4gnk70GqxywYX8U/kVRmJyQ/c9aW2JbHoeGWMst5BD01
# +7/lTu8fje7myXNf/oa8pVh1tULzzCOU1g4bKqQWyof2ecywMz7URWJWIUrwCZmI
# yNf0Yx2RjYtGBnJye6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzZaMC8GCSqGSIb3DQEJBDEiBCDGpDxj+II5904FPsRuuhFo67w/1XVSns9oXQ7i
# Hfm/ZzANBgkqhkiG9w0BAQEFAASCAgBB6zN2XwMeqYcKwOsI2t8T3hLFjBMeRJXd
# NjYRIMOSVwwnV8jsxzken+1mUkO3yRuFKFZhrjm8f1GlUu6uGrqYXocw8SR4ERLJ
# H++HcO05IbCAhNhlnEVZ2SL7lxlVMkO+AwOhtFmpSMlrV5OJVW/HF9bcFJk8on/1
# ugkkcngttl/s4YJ9xFQjWVtseqyiMZbxAB1lgXyx1gJOfTWyNPEcU08utUm0kc17
# AbH4ubYfM+GaGFlAAGPJOTk2c8vhiKBu6U2V3bL6WoIbD8qb76IRfh+vQxFmfyHO
# ZdIFROnKGw6KvYkf2bVtER+xDsvhJEI8oUZkEBRY9/kbs+DjGHNjEbgj4kCaObC8
# KsBZ5iqjmMsO0y1/D6lXk0yPuXuzVYaYXVHrI0uySAcv+epjKhNkdFjexR/76gv+
# /CCaSQW0DEQWxTGFjKYojrrp2UzOfqK793Gca8hcO5XWhy+TqOQOcqq34kM1Dy2b
# aH8kBr3yE86yfBf63KKSHoeAniMxqkBgE+gQ2ajvxDDXUVcXlg03XzaFO87hHwpR
# /JwhI9je6mHm8prC42HcizIJyU9O3pXRxmzt7h9YaPuppGYrtxTHr8j1FnI8qjzA
# N80KUQoXSxbTWxcvYUhMsQZBr5BKqFX3PgLsF0nuhCttBaLWveKUkBmpE849+FM6
# PyBbEaynTw==
# SIG # End signature block
