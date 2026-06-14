<#
.SYNOPSIS
Common GUI launcher for SmartM365 Exchange user notification campaigns.

.DESCRIPTION
Provides a Windows Forms launcher for the three ExchangeUserNotifications campaigns:
Exchange migration, archive enablement, and migration mailbox size reduction.

The GUI does not reimplement campaign logic. It validates local files, builds the
PowerShell command line, launches the selected campaign script, and displays output.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $pwsh = (Get-Process -Id $PID).Path
    Start-Process -FilePath $pwsh -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', $PSCommandPath) | Out-Null
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-ExchangeUserNotificationsRoot {
    $current = $PSScriptRoot
    while ($current) {
        if ((Split-Path -Path $current -Leaf) -eq 'ExchangeUserNotifications') { return $current }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    throw 'ExchangeUserNotifications root not found.'
}

function New-Button {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 120, [int]$H = 30)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.SetBounds($X, $Y, $W, $H)
    return $button
}

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 120, [int]$H = 20)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.SetBounds($X, $Y, $W, $H)
    return $label
}

function New-TextBox {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 220, [int]$H = 22)
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $Text
    $textBox.SetBounds($X, $Y, $W, $H)
    return $textBox
}

function New-CheckBox {
    param([string]$Text, [int]$X, [int]$Y, [bool]$Checked = $false, [int]$W = 160, [int]$H = 22)
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Text = $Text
    $checkBox.Checked = $Checked
    $checkBox.SetBounds($X, $Y, $W, $H)
    return $checkBox
}

function Add-OutputLine {
    param([string]$Text)
    $txtOutput.AppendText(("{0}`r`n" -f $Text))
    $txtOutput.SelectionStart = $txtOutput.TextLength
    $txtOutput.ScrollToCaret()
}

function Quote-CommandArgument {
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Get-SelectedCampaign {
    $campaigns[$cmbCampaign.SelectedIndex]
}

function Update-CampaignView {
    $campaign = Get-SelectedCampaign
    $lblPath.Text = $campaign.PathLabel
    $txtPath.Text = ''
    $chkFromList.Visible = [bool]$campaign.SupportsFromList
    $chkFromList.Enabled = [bool]$campaign.SupportsFromList
    if (-not $campaign.SupportsFromList) { $chkFromList.Checked = $false }
    $chkSkipConfirmation.Visible = [bool]$campaign.SupportsSkipConfirmation
    $chkSkipConfirmation.Enabled = [bool]$campaign.SupportsSkipConfirmation
    if (-not $campaign.SupportsSkipConfirmation) { $chkSkipConfirmation.Checked = $false }
    $lblScript.Text = $campaign.ScriptPath
    $lblConfig.Text = $campaign.ConfigPath
}

function Test-CampaignFiles {
    $campaign = Get-SelectedCampaign
    $issues = New-Object System.Collections.ArrayList

    foreach ($path in @($campaign.ScriptPath, $campaign.ConfigPath, $campaign.TemplateConfigPath)) {
        if (-not (Test-Path -LiteralPath $path)) { [void]$issues.Add("Missing: $path") }
    }

    try {
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($campaign.ScriptPath, [ref]$null, [ref]$errors) | Out-Null
        if ($errors) {
            foreach ($error in $errors) { [void]$issues.Add("Parser: $($error.Message)") }
        }
    }
    catch {
        [void]$issues.Add("Parser failed: $($_.Exception.Message)")
    }

    foreach ($jsonPath in @($campaign.ConfigPath, $campaign.TemplateConfigPath)) {
        try {
            Get-Content -LiteralPath $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop | Out-Null
        }
        catch {
            [void]$issues.Add("JSON invalid: $jsonPath - $($_.Exception.Message)")
        }
    }

    return @($issues)
}

function Build-CommandArguments {
    $campaign = Get-SelectedCampaign
    $args = New-Object System.Collections.Generic.List[string]
    $args.Add('-NoProfile')
    $args.Add('-ExecutionPolicy')
    $args.Add('Bypass')
    $args.Add('-File')
    $args.Add($campaign.ScriptPath)
    $args.Add('-Tenant')
    $args.Add($txtTenant.Text.Trim())

    $pathValue = $txtPath.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($pathValue)) {
        if ($campaign.SupportsFromList -and $chkFromList.Checked) {
            $args.Add('-FromList')
            $args.Add('-ListCsvPath')
            $args.Add($pathValue)
        }
        elseif ($campaign.PathParameter) {
            $args.Add($campaign.PathParameter)
            $args.Add($pathValue)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($txtForceLanguage.Text.Trim())) {
        $args.Add('-ForceLanguage')
        $args.Add($txtForceLanguage.Text.Trim())
    }
    if ($chkForceSend.Checked) { $args.Add('-ForceSend') }
    if ($chkNoSummary.Checked) { $args.Add('-NoSummaryEmail') }
    if ($chkSkipConfirmation.Visible -and $chkSkipConfirmation.Checked) { $args.Add('-SkipConfirmation') }
    if ($cmbMode.SelectedItem -eq 'Dry Run') { $args.Add('-WhatIf') }

    return $args
}

function Invoke-Campaign {
    $campaign = Get-SelectedCampaign
    $issues = Test-CampaignFiles
    if ($issues.Count -gt 0) {
        Add-OutputLine 'Validation failed:'
        foreach ($issue in $issues) { Add-OutputLine "  $issue" }
        [System.Windows.Forms.MessageBox]::Show("Validation failed. See output.", "SmartM365", 'OK', 'Warning') | Out-Null
        return
    }

    if ($cmbMode.SelectedItem -eq 'Live') {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Run live campaign '$($campaign.Name)' for tenant '$($txtTenant.Text.Trim())'?",
            'Confirm live send',
            'YesNo',
            'Warning'
        )
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Add-OutputLine 'Live run cancelled.'
            return
        }
    }

    $pwsh = if (Get-Command pwsh -ErrorAction SilentlyContinue) { (Get-Command pwsh).Source } else { (Get-Process -Id $PID).Path }
    $args = Build-CommandArguments
    Add-OutputLine ''
    Add-OutputLine ('[{0}] Starting {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $campaign.Name)
    Add-OutputLine ((Quote-CommandArgument $pwsh) + ' ' + (($args | ForEach-Object { Quote-CommandArgument $_ }) -join ' '))

    $btnRun.Enabled = $false
    $btnValidate.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $pwsh
        foreach ($arg in $args) { [void]$psi.ArgumentList.Add($arg) }
        $psi.WorkingDirectory = Split-Path -Path $campaign.ScriptPath -Parent
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        if (-not [string]::IsNullOrWhiteSpace($stdout)) { Add-OutputLine $stdout.TrimEnd() }
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            Add-OutputLine 'STDERR:'
            Add-OutputLine $stderr.TrimEnd()
        }
        Add-OutputLine ('ExitCode={0}' -f $process.ExitCode)
    }
    catch {
        Add-OutputLine ("Run failed: {0}" -f $_.Exception.Message)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRun.Enabled = $true
        $btnValidate.Enabled = $true
    }
}

function Open-PathIfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Invoke-Item -LiteralPath $Path
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Path not found:`r`n$Path", "SmartM365", 'OK', 'Warning') | Out-Null
    }
}

$root = Get-ExchangeUserNotificationsRoot
$campaigns = @(
    [pscustomobject]@{
        Name = 'Exchange Migration'
        ScriptPath = Join-Path $root 'ExchangeMigration\SmartM365-ExchangeMigration-NotifyUsers.ps1'
        ConfigPath = Join-Path $root 'Config\Campaigns\ExchangeMigration.local.json'
        TemplateConfigPath = Join-Path $root 'Config\Campaigns\ExchangeMigration.local.json.template'
        PathLabel = 'Recipients path'
        PathParameter = '-RecipientsPath'
        SupportsFromList = $false
        SupportsSkipConfirmation = $false
    },
    [pscustomobject]@{
        Name = 'Exchange Archive'
        ScriptPath = Join-Path $root 'ExchangeArchive\SmartM365-ExchangeArchive-NotifyUsers.ps1'
        ConfigPath = Join-Path $root 'Config\Campaigns\ExchangeArchive.local.json'
        TemplateConfigPath = Join-Path $root 'Config\Campaigns\ExchangeArchive.local.json.template'
        PathLabel = 'Recipients path'
        PathParameter = '-RecipientsPath'
        SupportsFromList = $false
        SupportsSkipConfirmation = $false
    },
    [pscustomobject]@{
        Name = 'Migration Mailbox Size Reduction'
        ScriptPath = Join-Path $root 'ExchangeMigrationMailboxSizeReduction\SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.ps1'
        ConfigPath = Join-Path $root 'Config\Campaigns\ExchangeMigrationMailboxSizeReduction.local.json'
        TemplateConfigPath = Join-Path $root 'Config\Campaigns\ExchangeMigrationMailboxSizeReduction.local.json.template'
        PathLabel = 'List CSV path'
        PathParameter = ''
        SupportsFromList = $true
        SupportsSkipConfirmation = $true
    }
)

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SmartM365 Exchange User Notifications'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(980, 720)
$form.MinimumSize = New-Object System.Drawing.Size(900, 640)

$lblCampaign = New-Label 'Campaign' 20 20 100
$cmbCampaign = New-Object System.Windows.Forms.ComboBox
$cmbCampaign.DropDownStyle = 'DropDownList'
$cmbCampaign.SetBounds(140, 18, 300, 24)
foreach ($campaign in $campaigns) { [void]$cmbCampaign.Items.Add($campaign.Name) }
$cmbCampaign.SelectedIndex = 0

$lblTenant = New-Label 'Tenant' 470 20 70
$txtTenant = New-TextBox 'test' 540 18 100

$lblMode = New-Label 'Mode' 670 20 50
$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.DropDownStyle = 'DropDownList'
$cmbMode.SetBounds(720, 18, 120, 24)
[void]$cmbMode.Items.Add('Dry Run')
[void]$cmbMode.Items.Add('Live')
$cmbMode.SelectedIndex = 0

$lblForceLanguage = New-Label 'Force language' 20 60 110
$txtForceLanguage = New-TextBox '' 140 58 120

$chkForceSend = New-CheckBox 'Force send' 290 58 $false 120
$chkNoSummary = New-CheckBox 'No summary email' 420 58 $false 150
$chkSkipConfirmation = New-CheckBox 'Skip confirmation' 590 58 $true 160

$lblPath = New-Label 'Recipients path' 20 100 110
$txtPath = New-TextBox '' 140 98 620
$btnBrowse = New-Button 'Browse' 775 95 80
$chkFromList = New-CheckBox 'From list' 865 98 $true 90

$btnRun = New-Button 'Run' 20 140 120 34
$btnValidate = New-Button 'Validate' 155 140 120 34
$btnOpenConfig = New-Button 'Open config' 290 140 120 34
$btnOpenTemplate = New-Button 'Open template' 425 140 120 34
$btnClear = New-Button 'Clear output' 560 140 120 34

$lblScriptTitle = New-Label 'Script' 20 190 80
$lblScript = New-Label '' 140 190 800 36
$lblScript.AutoEllipsis = $true
$lblConfigTitle = New-Label 'Config' 20 230 80
$lblConfig = New-Label '' 140 230 800 36
$lblConfig.AutoEllipsis = $true

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = 'Both'
$txtOutput.WordWrap = $false
$txtOutput.ReadOnly = $true
$txtOutput.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtOutput.SetBounds(20, 280, 920, 360)
$txtOutput.Anchor = 'Top,Bottom,Left,Right'

foreach ($control in @(
    $lblCampaign, $cmbCampaign, $lblTenant, $txtTenant, $lblMode, $cmbMode,
    $lblForceLanguage, $txtForceLanguage, $chkForceSend, $chkNoSummary, $chkSkipConfirmation,
    $lblPath, $txtPath, $btnBrowse, $chkFromList,
    $btnRun, $btnValidate, $btnOpenConfig, $btnOpenTemplate, $btnClear,
    $lblScriptTitle, $lblScript, $lblConfigTitle, $lblConfig, $txtOutput
)) {
    [void]$form.Controls.Add($control)
}

$cmbCampaign.Add_SelectedIndexChanged({ Update-CampaignView })
$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select CSV file'
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtPath.Text = $dialog.FileName }
})
$btnValidate.Add_Click({
    $issues = Test-CampaignFiles
    if ($issues.Count -eq 0) {
        Add-OutputLine ('[{0}] Validation OK for {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), (Get-SelectedCampaign).Name)
    }
    else {
        Add-OutputLine 'Validation issues:'
        foreach ($issue in $issues) { Add-OutputLine "  $issue" }
    }
})
$btnOpenConfig.Add_Click({ Open-PathIfExists -Path (Get-SelectedCampaign).ConfigPath })
$btnOpenTemplate.Add_Click({ Open-PathIfExists -Path (Get-SelectedCampaign).TemplateConfigPath })
$btnClear.Add_Click({ $txtOutput.Clear() })
$btnRun.Add_Click({ Invoke-Campaign })

Update-CampaignView
Add-OutputLine 'SmartM365 Exchange User Notifications GUI ready.'
Add-OutputLine 'Dry Run is selected by default.'

[void]$form.ShowDialog()
