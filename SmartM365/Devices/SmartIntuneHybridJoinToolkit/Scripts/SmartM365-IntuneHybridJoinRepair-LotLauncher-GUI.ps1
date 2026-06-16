<#
.SYNOPSIS
Creates and launches Smart Intune Hybrid Join Toolkit LOT folders from a GUI.

.DESCRIPTION
This operator GUI lists existing LOT-* folders for launch, shows their key
parameters, and offers a simplified new-LOT creation flow.

Operational LOT-* folders can contain real computer names and are ignored by Git.
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$ValidateOnly
)

$ErrorActionPreference = "Stop"

function Get-ToolkitRoot {
    $scriptDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    }
    else {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }

    return (Get-Item -LiteralPath (Split-Path -Parent $scriptDir) -ErrorAction Stop).FullName
}

function Get-SafeLotName {
    param([Parameter(Mandatory = $true)][string]$LotName)

    $safeName = [regex]::Replace($LotName.Trim(), "[^A-Za-z0-9._-]+", "-").Trim("-._")
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw "Enter a LOT name."
    }

    if ($safeName -notmatch "^(?i)LOT-") {
        $safeName = "LOT-$safeName"
    }

    return $safeName
}

function Get-ComputerNamesFromFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @()
    }

    $seen = @{}
    $computers = New-Object System.Collections.ArrayList
    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $name = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith("#")) {
            continue
        }

        $key = $name.ToUpperInvariant()
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        [void]$computers.Add($name)
    }

    return @($computers.ToArray())
}

function Get-LotFolders {
    param([Parameter(Mandatory = $true)][string]$RootPath)

    return @(
        Get-ChildItem -LiteralPath $RootPath -Directory -Filter "LOT-*" -ErrorAction Stop |
            Where-Object { $_.Name -ine "LOT-X" } |
            Sort-Object Name
    )
}

function Test-LotWrapperSet {
    param([Parameter(Mandatory = $true)][string]$LotPath)

    $wrapperNames = @(
        "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd",
        "Run-IntuneHybridJoinRepairWithPsExec-Once.cmd",
        "Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd",
        "Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd"
    )

    $missing = @(
        foreach ($wrapperName in $wrapperNames) {
            $path = Join-Path $LotPath $wrapperName
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                $wrapperName
            }
        }
    )

    return [pscustomobject]@{
        Ready = ($missing.Count -eq 0)
        Missing = $missing
    }
}

function Get-AdDomainText {
    param([Parameter(Mandatory = $true)][string]$LotPath)

    $adDomainPath = Join-Path $LotPath "AdDomain.txt"
    if (-not (Test-Path -LiteralPath $adDomainPath -PathType Leaf)) {
        return ""
    }

    foreach ($line in @(Get-Content -LiteralPath $adDomainPath -ErrorAction Stop)) {
        $value = ([string]$line).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return ""
}

function Get-FileFreshnessText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$FreshMinutes = 120
    )

    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $item) {
        return "Missing"
    }

    $age = (Get-Date) - $item.LastWriteTime
    $state = if ($age.TotalMinutes -le $FreshMinutes) { "Recent" } else { "Stale" }
    return ("{0}; {1:N1} min; {2}" -f $state,$age.TotalMinutes,$item.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
}

function Get-LotSummary {
    param(
        [Parameter(Mandatory = $true)][string]$RootPath,
        [Parameter(Mandatory = $true)][string]$LotPath
    )

    $computersPath = Join-Path $LotPath "Computers.txt"
    $adDomainPath = Join-Path $LotPath "AdDomain.txt"
    $lotAdCsvPath = Join-Path $LotPath "DevicesAD.csv"
    $rootAdCsvPath = Join-Path $RootPath "DevicesAD.csv"
    $rootIntuneCsvPath = Join-Path $RootPath "DevicesIntune.csv"
    $rootEntraCsvPath = Join-Path $RootPath "DevicesEntra.csv"
    $computers = @(Get-ComputerNamesFromFile -Path $computersPath)
    $adDomain = Get-AdDomainText -LotPath $LotPath
    $wrappers = Test-LotWrapperSet -LotPath $LotPath

    $adScope = if ([string]::IsNullOrWhiteSpace($adDomain)) {
        "Forest export from root DevicesAD.csv"
    }
    else {
        "Domain export: $adDomain"
    }

    $selectedAdCsv = if ([string]::IsNullOrWhiteSpace($adDomain)) { $rootAdCsvPath } else { $lotAdCsvPath }

    return [pscustomobject]@{
        Name = Split-Path -Leaf $LotPath
        Path = $LotPath
        ComputersPath = $computersPath
        ComputerCount = $computers.Count
        AdDomainPath = $adDomainPath
        AdDomain = $adDomain
        AdScope = $adScope
        SelectedAdCsv = $selectedAdCsv
        RootAdCsvStatus = Get-FileFreshnessText -Path $rootAdCsvPath -FreshMinutes 120
        SelectedAdCsvStatus = Get-FileFreshnessText -Path $selectedAdCsv -FreshMinutes 120
        IntuneCsvStatus = Get-FileFreshnessText -Path $rootIntuneCsvPath -FreshMinutes 120
        EntraCsvStatus = Get-FileFreshnessText -Path $rootEntraCsvPath -FreshMinutes 120
        WrappersReady = $wrappers.Ready
        MissingWrappers = $wrappers.Missing
    }
}

function Invoke-LotWrapperRefresh {
    param([Parameter(Mandatory = $true)][string]$RootPath)

    $updateScript = Join-Path $RootPath "Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1"
    if (-not (Test-Path -LiteralPath $updateScript -PathType Leaf)) {
        throw "LOT wrapper update script not found: $updateScript"
    }

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $updateScript -RootPath $RootPath
    if ($LASTEXITCODE -ne 0) {
        throw "LOT wrapper refresh failed with exit code $LASTEXITCODE."
    }
}

function New-ToolkitLotFolder {
    param(
        [Parameter(Mandatory = $true)][string]$RootPath,
        [Parameter(Mandatory = $true)][string]$LotName
    )

    $rootItem = Get-Item -LiteralPath $RootPath -ErrorAction Stop
    $safeLotName = Get-SafeLotName -LotName $LotName
    $lotPath = Join-Path $rootItem.FullName $safeLotName
    if (Test-Path -LiteralPath $lotPath) {
        throw "LOT folder already exists: $lotPath"
    }

    New-Item -ItemType Directory -Path $lotPath -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType File -Path (Join-Path $lotPath "Computers.txt") -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType File -Path (Join-Path $lotPath "AdDomain.txt") -Force -ErrorAction Stop | Out-Null

    Invoke-LotWrapperRefresh -RootPath $rootItem.FullName

    return [pscustomobject]@{
        LotPath = $lotPath
        ComputersPath = Join-Path $lotPath "Computers.txt"
        AdDomainPath = Join-Path $lotPath "AdDomain.txt"
    }
}

function Start-ToolkitLot {
    param(
        [Parameter(Mandatory = $true)][string]$LotPath,
        [Parameter(Mandatory = $true)][string]$Mode,
        [int]$GlobalConcurrencyLimit = 15
    )

    $wrapperName = switch ($Mode) {
        "Loop" { "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd"; break }
        "Once" { "Run-IntuneHybridJoinRepairWithPsExec-Once.cmd"; break }
        "LoopIgnoreRunGuard" { "Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd"; break }
        "OnceIgnoreRunGuard" { "Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd"; break }
        default { "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd" }
    }

    $wrapperPath = Join-Path $LotPath $wrapperName
    if (-not (Test-Path -LiteralPath $wrapperPath -PathType Leaf)) {
        throw "LOT wrapper not found: $wrapperPath"
    }

    $psexecToolkitPath = Join-Path $toolkitRoot "Scripts\PsExec.exe"
    $psexecSystem32Path = Join-Path $env:WINDIR "System32\PsExec.exe"
    $psexecCommand = Get-Command -Name "PsExec.exe" -CommandType Application -ErrorAction SilentlyContinue
    if (
        -not (Test-Path -LiteralPath $psexecToolkitPath -PathType Leaf) -and
        -not (Test-Path -LiteralPath $psexecSystem32Path -PathType Leaf) -and
        -not $psexecCommand
    ) {
        throw ("PsExec.exe not found. Place it in '{0}', in '{1}', or add PsExec.exe to PATH before launching the LOT." -f (Split-Path -Parent $psexecToolkitPath), (Split-Path -Parent $psexecSystem32Path))
    }

    if ($GlobalConcurrencyLimit -lt 1) { $GlobalConcurrencyLimit = 1 }
    Start-Process -FilePath "cmd.exe" -ArgumentList @("/k", "`"$wrapperPath`"", "-GlobalConcurrencyLimit", [string]$GlobalConcurrencyLimit) -WorkingDirectory $LotPath -Verb RunAs
}

function Open-TextFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        New-Item -ItemType File -Path $Path -Force -ErrorAction Stop | Out-Null
    }

    Start-Process -FilePath "notepad.exe" -ArgumentList @("`"$Path`"")
}

function Open-FolderPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Folder not found: $Path"
    }

    Start-Process -FilePath "explorer.exe" -ArgumentList @("`"$Path`"")
}

$toolkitRoot = Get-ToolkitRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$titleFont = New-Object System.Drawing.Font("Segoe UI Semibold", 19)
$sectionFont = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$smallFont = New-Object System.Drawing.Font("Segoe UI", 8)
$statusFont = New-Object System.Drawing.Font("Consolas", 9)

$colorBackground = [System.Drawing.ColorTranslator]::FromHtml("#EEF3F8")
$colorPanel = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")
$colorPanelSoft = [System.Drawing.ColorTranslator]::FromHtml("#F8FBFE")
$colorHeaderPanel = [System.Drawing.ColorTranslator]::FromHtml("#E8F3FF")
$colorAccent = [System.Drawing.ColorTranslator]::FromHtml("#2563EB")
$colorInk = [System.Drawing.ColorTranslator]::FromHtml("#102A43")
$colorMuted = [System.Drawing.ColorTranslator]::FromHtml("#52677A")
$colorBorder = [System.Drawing.ColorTranslator]::FromHtml("#D7E1EA")
$colorTextBoxBorder = [System.Drawing.ColorTranslator]::FromHtml("#B9C8D7")
$colorSuccess = [System.Drawing.ColorTranslator]::FromHtml("#0F766E")
$colorWarning = [System.Drawing.ColorTranslator]::FromHtml("#B45309")
$colorDisabled = [System.Drawing.ColorTranslator]::FromHtml("#E4EAF1")
$colorDisabledText = [System.Drawing.ColorTranslator]::FromHtml("#7A8A99")

function Resolve-LogoIconPath {
    $candidatePaths = @(
        (Join-Path $PSScriptRoot "SmartM365-logo.ico"),
        (Join-Path $toolkitRoot "SmartM365-logo.ico")
    )

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            return (Get-Item -LiteralPath $candidatePath -ErrorAction Stop).FullName
        }
    }

    return $null
}

function Resolve-LogoImagePath {
    $devicesRoot = Split-Path -Parent $toolkitRoot
    $candidatePaths = @(
        (Join-Path $toolkitRoot "workplacecloudhub-v2.png"),
        (Join-Path $PSScriptRoot "workplacecloudhub-v2.png"),
        (Join-Path $devicesRoot "DeviceRegistrationTool\workplacecloudhub-v2.png")
    )

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            return (Get-Item -LiteralPath $candidatePath -ErrorAction Stop).FullName
        }
    }

    return $null
}

function Set-FlatButtonStyle {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory = $true)][System.Drawing.Color]$BackColor,
        [Parameter(Mandatory = $true)][System.Drawing.Color]$ForeColor,
        [System.Drawing.Color]$BorderColor = $colorBorder
    )

    $Button.FlatStyle = "Flat"
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatAppearance.BorderColor = $BorderColor
    $Button.FlatAppearance.BorderSize = 1
    $Button.Height = 32
    $Button.Margin = New-Object System.Windows.Forms.Padding(6, 6, 0, 6)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Add-SoftBorder {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Control,
        [System.Drawing.Color]$BorderColor = $colorBorder
    )

    $Control.Add_Paint({
        param($sender, $eventArgs)

        [System.Windows.Forms.ControlPaint]::DrawBorder(
            $eventArgs.Graphics,
            $sender.ClientRectangle,
            $BorderColor,
            [System.Windows.Forms.ButtonBorderStyle]::Solid
        )
    }.GetNewClosure())
}

function Add-AccentBar {
    param([Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Control)

    $Control.Add_Paint({
        param($sender, $eventArgs)

        $brush = New-Object System.Drawing.SolidBrush($colorAccent)
        try {
            $eventArgs.Graphics.FillRectangle($brush, 0, 0, 4, $sender.Height)
        }
        finally {
            $brush.Dispose()
        }
    }.GetNewClosure())
}

function New-Label {
    param([string]$Text)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Dock = "Fill"
    $label.TextAlign = "MiddleLeft"
    $label.ForeColor = $colorMuted
    return $label
}

function New-ValueBox {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = "Fill"
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $box.ReadOnly = $true
    $box.BorderStyle = "FixedSingle"
    $box.BackColor = $colorPanelSoft
    $box.ForeColor = $colorInk
    return $box
}

function New-EntryBox {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = "Fill"
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $box.BorderStyle = "FixedSingle"
    $box.BackColor = $colorPanel
    $box.ForeColor = $colorInk
    return $box
}

function New-SectionPanel {
    param([Parameter(Mandatory = $true)][string]$Title)

    $outer = New-Object System.Windows.Forms.Panel
    $outer.Dock = "Fill"
    $outer.BackColor = $colorPanel
    $outer.Padding = New-Object System.Windows.Forms.Padding(14)
    Add-SoftBorder -Control $outer

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = "Fill"
    $layout.ColumnCount = 1
    $layout.RowCount = 2
    $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.Dock = "Fill"
    $titleLabel.Font = $sectionFont
    $titleLabel.ForeColor = $colorInk
    $titleLabel.TextAlign = "MiddleLeft"

    $content = New-Object System.Windows.Forms.Panel
    $content.Dock = "Fill"
    $content.BackColor = $colorPanel

    $layout.Controls.Add($titleLabel, 0, 0)
    $layout.Controls.Add($content, 0, 1)
    $outer.Controls.Add($layout)

    return [pscustomobject]@{
        Panel = $outer
        Content = $content
    }
}

function Set-ButtonEnabledStyle {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Button]$Button,
        [bool]$Enabled,
        [System.Drawing.Color]$EnabledBackColor = $colorAccent,
        [System.Drawing.Color]$EnabledForeColor = ([System.Drawing.Color]::White)
    )

    if ($Enabled) {
        $Button.Enabled = $true
        Set-FlatButtonStyle -Button $Button -BackColor $EnabledBackColor -ForeColor $EnabledForeColor -BorderColor $EnabledBackColor
    }
    else {
        $Button.Enabled = $false
        Set-FlatButtonStyle -Button $Button -BackColor $colorDisabled -ForeColor $colorDisabledText -BorderColor $colorBorder
        $Button.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Smart Intune Hybrid Join Toolkit - LOT Launcher"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(980, 700)
$form.MinimumSize = New-Object System.Drawing.Size(900, 620)
$form.Font = $font
$form.BackColor = $colorBackground

$logoIconPath = Resolve-LogoIconPath
$logoImagePath = Resolve-LogoImagePath
$script:LogoIcon = $null
$script:LogoImage = $null
if (-not [string]::IsNullOrWhiteSpace($logoIconPath)) {
    try {
        $script:LogoIcon = New-Object System.Drawing.Icon($logoIconPath, 48, 48)
        $form.Icon = $script:LogoIcon
    }
    catch {
        $script:LogoIcon = $null
    }
}

if (-not [string]::IsNullOrWhiteSpace($logoImagePath)) {
    try {
        $script:LogoImage = [System.Drawing.Image]::FromFile($logoImagePath)
    }
    catch {
        $script:LogoImage = $null
    }
}

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = "Fill"
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 3
$rootLayout.Padding = New-Object System.Windows.Forms.Padding(14)
$rootLayout.BackColor = $colorBackground
$rootLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 132))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46))) | Out-Null

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = "Fill"
$headerPanel.BackColor = $colorPanel
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(20, 16, 20, 16)
Add-SoftBorder -Control $headerPanel
Add-AccentBar -Control $headerPanel

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = "Fill"
$headerLayout.ColumnCount = 2
$headerLayout.RowCount = 1
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 250))) | Out-Null

$headerTextPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerTextPanel.Dock = "Fill"
$headerTextPanel.ColumnCount = 1
$headerTextPanel.RowCount = 4
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null

$badgeLabel = New-Object System.Windows.Forms.Label
$badgeLabel.Text = "SMARTM365"
$badgeLabel.AutoSize = $true
$badgeLabel.BackColor = $colorHeaderPanel
$badgeLabel.ForeColor = $colorAccent
$badgeLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8)
$badgeLabel.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Intune Hybrid Join LOT Launcher"
$titleLabel.Dock = "Fill"
$titleLabel.ForeColor = $colorInk
$titleLabel.Font = $titleFont
$titleLabel.TextAlign = "MiddleLeft"

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Run an existing LOT or create an empty LOT ready for Computers.txt."
$subtitleLabel.Dock = "Fill"
$subtitleLabel.ForeColor = $colorMuted
$subtitleLabel.TextAlign = "MiddleLeft"

$psexecLabel = New-Object System.Windows.Forms.Label
$psexecLabel.Dock = "Fill"
$psexecLabel.ForeColor = $colorInk
$psexecLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$psexecLabel.TextAlign = "MiddleLeft"

$headerTextPanel.Controls.Add($badgeLabel, 0, 0)
$headerTextPanel.Controls.Add($titleLabel, 0, 1)
$headerTextPanel.Controls.Add($subtitleLabel, 0, 2)
$headerTextPanel.Controls.Add($psexecLabel, 0, 3)

$logoCard = New-Object System.Windows.Forms.Panel
$logoCard.Dock = "Fill"
$logoCard.BackColor = $colorPanel
$logoCard.Padding = New-Object System.Windows.Forms.Padding(10)
$logoCard.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
Add-SoftBorder -Control $logoCard

if ($script:LogoImage) {
    $logoPicture = New-Object System.Windows.Forms.PictureBox
    $logoPicture.Dock = "Fill"
    $logoPicture.SizeMode = "Zoom"
    $logoPicture.Image = $script:LogoImage
    $logoPicture.BackColor = $colorPanel
    $logoCard.Controls.Add($logoPicture)
}
else {
    $logoFallback = New-Object System.Windows.Forms.Label
    $logoFallback.Dock = "Fill"
    $logoFallback.Text = "Workplace`r`nCloudHub"
    $logoFallback.ForeColor = $colorAccent
    $logoFallback.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
    $logoFallback.TextAlign = "MiddleCenter"
    $logoCard.Controls.Add($logoFallback)
}

$psexecToolkitPath = Join-Path $toolkitRoot "Scripts\PsExec.exe"
$psexecSystem32Path = Join-Path $env:WINDIR "System32\PsExec.exe"
if (Test-Path -LiteralPath $psexecToolkitPath -PathType Leaf) {
    $psexecLabel.Text = "PsExec ready in Scripts"
    $psexecLabel.ForeColor = $colorSuccess
}
elseif (Test-Path -LiteralPath $psexecSystem32Path -PathType Leaf) {
    $psexecLabel.Text = "PsExec ready in System32"
    $psexecLabel.ForeColor = $colorSuccess
}
else {
    $psexecLabel.Text = "PsExec not local"
    $psexecLabel.ForeColor = $colorWarning
}

$headerLayout.Controls.Add($headerTextPanel, 0, 0)
$headerLayout.Controls.Add($logoCard, 1, 0)
$headerPanel.Controls.Add($headerLayout)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$tabControl.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 10)
$tabControl.Font = $font
$tabControl.Appearance = "FlatButtons"
$tabControl.SizeMode = "Fixed"
$tabControl.ItemSize = New-Object System.Drawing.Size(132, 34)
$tabControl.Padding = New-Object System.Drawing.Point(12, 4)

$existingTab = New-Object System.Windows.Forms.TabPage
$existingTab.Text = "Existing LOT"
$existingTab.BackColor = $colorBackground

$newTab = New-Object System.Windows.Forms.TabPage
$newTab.Text = "New LOT"
$newTab.BackColor = $colorBackground

$tabControl.TabPages.Add($existingTab) | Out-Null
$tabControl.TabPages.Add($newTab) | Out-Null

$existingLayout = New-Object System.Windows.Forms.TableLayoutPanel
$existingLayout.Dock = "Fill"
$existingLayout.ColumnCount = 1
$existingLayout.RowCount = 2
$existingLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$existingLayout.BackColor = $colorBackground
$existingLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$existingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 232))) | Out-Null
$existingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$existingSection = New-SectionPanel -Title "Available lots"
$activitySection = New-SectionPanel -Title "Activity"

$existingTable = New-Object System.Windows.Forms.TableLayoutPanel
$existingTable.Dock = "Fill"
$existingTable.ColumnCount = 5
$existingTable.RowCount = 5
$existingTable.BackColor = $colorPanel
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 138))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 118))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 118))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 118))) | Out-Null
for ($i = 0; $i -lt 5; $i++) {
    $existingTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34))) | Out-Null
}

$lotCombo = New-Object System.Windows.Forms.ComboBox
$lotCombo.Dock = "Fill"
$lotCombo.DropDownStyle = "DropDownList"
$lotCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$lotCombo.FlatStyle = "Flat"

$refreshLotsButton = New-Object System.Windows.Forms.Button
$refreshLotsButton.Text = "Refresh"
$refreshLotsButton.Dock = "Fill"
Set-FlatButtonStyle -Button $refreshLotsButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$openLotFolderButton = New-Object System.Windows.Forms.Button
$openLotFolderButton.Text = "Folder"
$openLotFolderButton.Dock = "Fill"
Set-FlatButtonStyle -Button $openLotFolderButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$deviceCountBox = New-ValueBox
$adScopeBox = New-ValueBox

$existingModeCombo = New-Object System.Windows.Forms.ComboBox
$existingModeCombo.Dock = "Fill"
$existingModeCombo.DropDownStyle = "DropDownList"
$existingModeCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$existingModeCombo.FlatStyle = "Flat"
$existingModeCombo.Items.AddRange([object[]]@("Loop", "Once", "LoopIgnoreRunGuard", "OnceIgnoreRunGuard"))
$existingModeCombo.SelectedIndex = 0

$globalConcurrencyLimitBox = New-Object System.Windows.Forms.NumericUpDown
$globalConcurrencyLimitBox.Dock = "Fill"
$globalConcurrencyLimitBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$globalConcurrencyLimitBox.Minimum = 1
$globalConcurrencyLimitBox.Maximum = 50
$globalConcurrencyLimitBox.Value = 15

$openComputersButton = New-Object System.Windows.Forms.Button
$openComputersButton.Text = "Computers"
$openComputersButton.Dock = "Fill"
Set-FlatButtonStyle -Button $openComputersButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$openAdDomainButton = New-Object System.Windows.Forms.Button
$openAdDomainButton.Text = "AD domain"
$openAdDomainButton.Dock = "Fill"
Set-FlatButtonStyle -Button $openAdDomainButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$launchExistingButton = New-Object System.Windows.Forms.Button
$launchExistingButton.Text = "Launch"
$launchExistingButton.Dock = "Fill"
Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $false -EnabledBackColor $colorSuccess

$launchAllLotsButton = New-Object System.Windows.Forms.Button
$launchAllLotsButton.Text = "Launch all"
$launchAllLotsButton.Dock = "Fill"
Set-ButtonEnabledStyle -Button $launchAllLotsButton -Enabled $false -EnabledBackColor $colorSuccess

$refreshWrappersButton = New-Object System.Windows.Forms.Button
$refreshWrappersButton.Text = "Wrappers"
$refreshWrappersButton.Dock = "Fill"
Set-FlatButtonStyle -Button $refreshWrappersButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$existingTable.Controls.Add((New-Label "LOT"), 0, 0)
$existingTable.Controls.Add($lotCombo, 1, 0)
$existingTable.SetColumnSpan($lotCombo, 2)
$existingTable.Controls.Add($refreshLotsButton, 3, 0)
$existingTable.Controls.Add($openLotFolderButton, 4, 0)

$existingTable.Controls.Add((New-Label "Devices"), 0, 1)
$existingTable.Controls.Add($deviceCountBox, 1, 1)
$existingTable.SetColumnSpan($deviceCountBox, 2)
$existingTable.Controls.Add($openComputersButton, 3, 1)
$existingTable.Controls.Add($openAdDomainButton, 4, 1)

$existingTable.Controls.Add((New-Label "AD scope"), 0, 2)
$existingTable.Controls.Add($adScopeBox, 1, 2)
$existingTable.SetColumnSpan($adScopeBox, 4)

$existingTable.Controls.Add((New-Label "Limit"), 0, 3)
$existingTable.Controls.Add($globalConcurrencyLimitBox, 1, 3)
$existingTable.SetColumnSpan($globalConcurrencyLimitBox, 4)

$existingTable.Controls.Add((New-Label "Launch"), 0, 4)
$existingTable.Controls.Add($existingModeCombo, 1, 4)
$existingTable.Controls.Add($refreshWrappersButton, 2, 4)
$existingTable.Controls.Add($launchExistingButton, 3, 4)
$existingTable.Controls.Add($launchAllLotsButton, 4, 4)

$existingSection.Content.Controls.Add($existingTable)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Dock = "Fill"
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$statusBox.Font = $statusFont
$statusBox.BackColor = $colorPanelSoft
$statusBox.ForeColor = $colorInk
$statusBox.BorderStyle = "FixedSingle"
$activitySection.Content.Controls.Add($statusBox)

$existingLayout.Controls.Add($existingSection.Panel, 0, 0)
$existingLayout.Controls.Add($activitySection.Panel, 0, 1)
$existingTab.Controls.Add($existingLayout)

$newLayout = New-Object System.Windows.Forms.TableLayoutPanel
$newLayout.Dock = "Fill"
$newLayout.ColumnCount = 1
$newLayout.RowCount = 2
$newLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$newLayout.BackColor = $colorBackground
$newLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$newLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 174))) | Out-Null
$newLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$newSection = New-SectionPanel -Title "Create an empty LOT"
$newTable = New-Object System.Windows.Forms.TableLayoutPanel
$newTable.Dock = "Fill"
$newTable.ColumnCount = 3
$newTable.RowCount = 2
$newTable.BackColor = $colorPanel
$newTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 138))) | Out-Null
$newTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$newTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
for ($i = 0; $i -lt 2; $i++) {
    $newTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
}

$newLotNameBox = New-EntryBox
$newComputersPathBox = New-ValueBox

$createLotButton = New-Object System.Windows.Forms.Button
$createLotButton.Text = "Create"
$createLotButton.Dock = "Fill"
Set-FlatButtonStyle -Button $createLotButton -BackColor $colorAccent -ForeColor ([System.Drawing.Color]::White) -BorderColor $colorAccent

$openNewComputersButton = New-Object System.Windows.Forms.Button
$openNewComputersButton.Text = "Open Computers"
$openNewComputersButton.Dock = "Fill"
Set-ButtonEnabledStyle -Button $openNewComputersButton -Enabled $false

$newTable.Controls.Add((New-Label "LOT name"), 0, 0)
$newTable.Controls.Add($newLotNameBox, 1, 0)
$newTable.Controls.Add($createLotButton, 2, 0)

$newTable.Controls.Add((New-Label "Computers.txt"), 0, 1)
$newTable.Controls.Add($newComputersPathBox, 1, 1)
$newTable.Controls.Add($openNewComputersButton, 2, 1)

$newSection.Content.Controls.Add($newTable)
$newLayout.Controls.Add($newSection.Panel, 0, 0)
$newTab.Controls.Add($newLayout)

$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = "Fill"
$footerPanel.BackColor = $colorBackground

$footerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$footerLayout.Dock = "Fill"
$footerLayout.ColumnCount = 2
$footerLayout.RowCount = 1
$footerLayout.BackColor = $colorBackground
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120))) | Out-Null

$footerStatus = New-Object System.Windows.Forms.Label
$footerStatus.Dock = "Fill"
$footerStatus.Text = "Ready"
$footerStatus.TextAlign = "MiddleLeft"
$footerStatus.ForeColor = $colorMuted
$footerStatus.Font = $smallFont

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Width = 94
Set-FlatButtonStyle -Button $closeButton -BackColor $colorPanel -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = "Fill"
$buttonPanel.FlowDirection = "RightToLeft"
$buttonPanel.BackColor = $colorBackground
$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$buttonPanel.Controls.Add($closeButton)

$footerLayout.Controls.Add($footerStatus, 0, 0)
$footerLayout.Controls.Add($buttonPanel, 1, 0)
$footerPanel.Controls.Add($footerLayout)

$rootLayout.Controls.Add($headerPanel, 0, 0)
$rootLayout.Controls.Add($tabControl, 0, 1)
$rootLayout.Controls.Add($footerPanel, 0, 2)
$form.Controls.Add($rootLayout)

$script:LotList = @()
$script:SelectedLotSummary = $null
$script:CreatedComputersPath = $null

function Add-Status {
    param([string]$Message)

    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    $footerStatus.Text = $Message
    if ([string]::IsNullOrWhiteSpace($statusBox.Text)) {
        $statusBox.Text = $line
    }
    else {
        $statusBox.AppendText([Environment]::NewLine + $line)
    }
}

function Clear-LotDetails {
    $script:SelectedLotSummary = $null
    $deviceCountBox.Text = ""
    $adScopeBox.Text = ""
    Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $false -EnabledBackColor $colorSuccess
}

function Get-LaunchableLotSummaries {
    $summaries = New-Object System.Collections.ArrayList
    foreach ($lot in @($script:LotList)) {
        try {
            $summary = Get-LotSummary -RootPath $toolkitRoot -LotPath $lot.FullName
            if ($summary.ComputerCount -gt 0 -and $summary.WrappersReady) {
                [void]$summaries.Add($summary)
            }
        }
        catch {
            Add-Status ("Skipping {0}: {1}" -f $lot.Name,$_.Exception.Message)
        }
    }

    return @($summaries.ToArray())
}

function Update-ExistingLotControlState {
    param([bool]$HasLots)

    $lotCombo.Enabled = $HasLots
    $existingModeCombo.Enabled = $HasLots
    $globalConcurrencyLimitBox.Enabled = $HasLots
    $openLotFolderButton.Enabled = $HasLots
    $openComputersButton.Enabled = $HasLots
    $openAdDomainButton.Enabled = $HasLots
    Set-ButtonEnabledStyle -Button $launchAllLotsButton -Enabled ($HasLots -and (Get-LaunchableLotSummaries).Count -gt 0) -EnabledBackColor $colorSuccess
}

function Update-LotDetails {
    if ($lotCombo.SelectedIndex -lt 0 -or $lotCombo.SelectedIndex -ge $script:LotList.Count) {
        Clear-LotDetails
        return
    }

    $lot = $script:LotList[$lotCombo.SelectedIndex]
    $summary = Get-LotSummary -RootPath $toolkitRoot -LotPath $lot.FullName
    $script:SelectedLotSummary = $summary

    $deviceCountBox.Text = [string]$summary.ComputerCount
    $adScopeBox.Text = $summary.AdScope

    if (-not $summary.WrappersReady) {
        Add-Status ("Missing LOT wrappers: {0}" -f ($summary.MissingWrappers -join ", "))
    }

    $canLaunch = ($summary.ComputerCount -gt 0 -and $summary.WrappersReady)
    Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $canLaunch -EnabledBackColor $colorSuccess
    Update-ExistingLotControlState -HasLots ($script:LotList.Count -gt 0)
}

function Refresh-LotList {
    param([string]$PreferredName)

    $selectedName = if (-not [string]::IsNullOrWhiteSpace($PreferredName)) {
        $PreferredName
    }
    elseif ($lotCombo.SelectedItem) {
        [string]$lotCombo.SelectedItem
    }
    else {
        ""
    }
    $script:LotList = @(Get-LotFolders -RootPath $toolkitRoot)

    $lotCombo.Items.Clear()
    foreach ($lot in $script:LotList) {
        $lotCombo.Items.Add($lot.Name) | Out-Null
    }

    if ($lotCombo.Items.Count -eq 0) {
        Clear-LotDetails
        Update-ExistingLotControlState -HasLots $false
        Add-Status "No operational LOT-* folder found."
        return
    }

    Update-ExistingLotControlState -HasLots $true

    $selectedIndex = 0
    if (-not [string]::IsNullOrWhiteSpace($selectedName)) {
        for ($i = 0; $i -lt $lotCombo.Items.Count; $i++) {
            if ([string]$lotCombo.Items[$i] -eq $selectedName) {
                $selectedIndex = $i
                break
            }
        }
    }

    $lotCombo.SelectedIndex = $selectedIndex
    Update-LotDetails
    Add-Status ("Loaded {0} LOT folder(s)." -f $lotCombo.Items.Count)
}

$refreshLotsButton.Add_Click({
    try {
        Refresh-LotList
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Refresh failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$lotCombo.Add_SelectedIndexChanged({
    try {
        Update-LotDetails
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$openLotFolderButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        Open-FolderPath -Path $script:SelectedLotSummary.Path
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Open folder failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openComputersButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        Open-TextFile -Path $script:SelectedLotSummary.ComputersPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$openAdDomainButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        Open-TextFile -Path $script:SelectedLotSummary.AdDomainPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$refreshWrappersButton.Add_Click({
    try {
        Invoke-LotWrapperRefresh -RootPath $toolkitRoot
        Add-Status "LOT wrappers refreshed."
        Refresh-LotList
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Wrapper refresh failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchExistingButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        if ($script:SelectedLotSummary.ComputerCount -le 0) { throw "Computers.txt is empty." }

        Add-Status ("Launching {0} in {1} mode. Global limit={2}." -f $script:SelectedLotSummary.Name,$existingModeCombo.SelectedItem,[int]$globalConcurrencyLimitBox.Value)
        Start-ToolkitLot -LotPath $script:SelectedLotSummary.Path -Mode ([string]$existingModeCombo.SelectedItem) -GlobalConcurrencyLimit ([int]$globalConcurrencyLimitBox.Value)
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "LOT launch failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchAllLotsButton.Add_Click({
    try {
        $launchableLots = @(Get-LaunchableLotSummaries)
        if ($launchableLots.Count -eq 0) {
            throw "No launchable LOT found. Check Computers.txt and wrapper files."
        }

        $mode = [string]$existingModeCombo.SelectedItem
        $limit = [int]$globalConcurrencyLimitBox.Value
        Add-Status ("Launching {0}/{1} LOT folder(s) in {2} mode. Global limit={3}." -f $launchableLots.Count,$script:LotList.Count,$mode,$limit)

        foreach ($lotSummary in $launchableLots) {
            Add-Status ("Launching {0}." -f $lotSummary.Name)
            Start-ToolkitLot -LotPath $lotSummary.Path -Mode $mode -GlobalConcurrencyLimit $limit
        }

        $skippedCount = $script:LotList.Count - $launchableLots.Count
        if ($skippedCount -gt 0) {
            Add-Status ("Skipped {0} LOT folder(s) without devices or complete wrappers." -f $skippedCount)
        }
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Launch all failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$createLotButton.Add_Click({
    try {
        $result = New-ToolkitLotFolder -RootPath $toolkitRoot -LotName $newLotNameBox.Text
        $newComputersPathBox.Text = $result.ComputersPath
        $script:CreatedComputersPath = $result.ComputersPath
        Set-ButtonEnabledStyle -Button $openNewComputersButton -Enabled $true

        Add-Status ("Created empty LOT: {0}" -f $result.LotPath)
        Refresh-LotList -PreferredName (Split-Path -Leaf $result.LotPath)
        $tabControl.SelectedTab = $newTab

        $answer = [System.Windows.Forms.MessageBox]::Show(
            $form,
            ("LOT created:`r`n{0}`r`n`r`nOpen Computers.txt now?" -f $result.LotPath),
            "Fill Computers.txt",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Open-TextFile -Path $result.ComputersPath
        }
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "LOT creation failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openNewComputersButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($script:CreatedComputersPath)) {
            throw "Create a LOT first."
        }

        Open-TextFile -Path $script:CreatedComputersPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$closeButton.Add_Click({
    $form.Close()
})

$form.Add_Shown({
    try {
        Refresh-LotList
        $lotCombo.Focus() | Out-Null
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$form.Add_FormClosed({
    if ($script:LogoImage) {
        $script:LogoImage.Dispose()
        $script:LogoImage = $null
    }

    if ($script:LogoIcon) {
        $script:LogoIcon.Dispose()
        $script:LogoIcon = $null
    }
})

if ($ValidateOnly) {
    $folders = @(Get-LotFolders -RootPath $toolkitRoot)
    Write-Host ("Smart Intune Hybrid Join Toolkit LOT Launcher GUI validation completed. Lots={0}" -f $folders.Count)
    return
}

Add-Status "Ready. Select an existing LOT or create a new empty LOT."
[void]$form.ShowDialog()
