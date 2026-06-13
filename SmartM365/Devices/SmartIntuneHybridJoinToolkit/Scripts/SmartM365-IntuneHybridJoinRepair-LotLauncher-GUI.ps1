<#
.SYNOPSIS
Creates and launches Smart Intune Hybrid Join Toolkit LOT folders from a GUI.

.DESCRIPTION
This operator GUI asks for a computer list file, creates a local LOT-* folder,
writes a normalized Computers.txt, refreshes the LOT CMD wrappers, and offers
to launch the selected LOT wrapper.

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
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerFilePath
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ComputerFilePath)
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = "Manual"
    }

    $safeName = [regex]::Replace($baseName.Trim(), "[^A-Za-z0-9._-]+", "-").Trim("-._")
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "Manual"
    }

    if ($safeName -notmatch "^(?i)LOT-") {
        $safeName = "LOT-$safeName"
    }

    return $safeName
}

function Get-UniqueLotPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$LotName
    )

    $safeLotName = [regex]::Replace($LotName.Trim(), "[^A-Za-z0-9._-]+", "-").Trim("-._")
    if ([string]::IsNullOrWhiteSpace($safeLotName)) {
        $safeLotName = "LOT-Manual"
    }

    if ($safeLotName -notmatch "^(?i)LOT-") {
        $safeLotName = "LOT-$safeLotName"
    }

    $lotPath = Join-Path $RootPath $safeLotName
    if (-not (Test-Path -LiteralPath $lotPath)) {
        return $lotPath
    }

    $suffix = Get-Date -Format "yyyyMMdd-HHmmss"
    return (Join-Path $RootPath ("{0}-{1}" -f $safeLotName,$suffix))
}

function Get-NormalizedComputerList {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Values
    )

    $result = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    foreach ($value in $Values) {
        $name = ([string]$value).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith("#")) {
            continue
        }

        if ($seen.ContainsKey($name.ToUpperInvariant())) {
            continue
        }

        $seen[$name.ToUpperInvariant()] = $true
        $result.Add($name)
    }

    return @($result)
}

function Convert-ComputerListFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $sourceItem = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
    $extension = [System.IO.Path]::GetExtension($sourceItem.FullName)
    $rawValues = @()

    if ($extension -and $extension.Equals(".csv", [System.StringComparison]::OrdinalIgnoreCase)) {
        $rows = @(Import-Csv -LiteralPath $sourceItem.FullName -ErrorAction Stop)
        if ($rows.Count -gt 0) {
            $preferredColumns = @("ComputerName", "DeviceName", "Name", "DisplayName")
            $properties = @($rows[0].PSObject.Properties.Name)
            $selectedColumn = $preferredColumns | Where-Object { $properties -contains $_ } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($selectedColumn) -and $properties.Count -gt 0) {
                $selectedColumn = $properties[0]
            }

            if (-not [string]::IsNullOrWhiteSpace($selectedColumn)) {
                $rawValues = @($rows | ForEach-Object { $_.$selectedColumn })
            }
        }
    }
    else {
        $rawValues = @(Get-Content -LiteralPath $sourceItem.FullName -ErrorAction Stop)
    }

    $computers = @(Get-NormalizedComputerList -Values $rawValues)
    if ($computers.Count -eq 0) {
        throw "No computer names found in the selected file."
    }

    Set-Content -LiteralPath $DestinationPath -Value $computers -Encoding ASCII -Force
    return $computers.Count
}

function New-ToolkitLotFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$ComputerFilePath,

        [Parameter(Mandatory = $true)]
        [string]$LotName
    )

    $computerFile = Get-Item -LiteralPath $ComputerFilePath -ErrorAction Stop
    if (-not $computerFile -or $computerFile.PSIsContainer) {
        throw "Select a readable computer list file."
    }

    $rootItem = Get-Item -LiteralPath $RootPath -ErrorAction Stop
    $lotPath = Get-UniqueLotPath -RootPath $rootItem.FullName -LotName $LotName
    New-Item -ItemType Directory -Path $lotPath -Force -ErrorAction Stop | Out-Null

    $targetComputers = Join-Path $lotPath "Computers.txt"
    $computerCount = Convert-ComputerListFile -SourcePath $computerFile.FullName -DestinationPath $targetComputers

    $updateScript = Join-Path $rootItem.FullName "Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1"
    if (-not (Test-Path -LiteralPath $updateScript)) {
        throw "LOT wrapper update script not found: $updateScript"
    }

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $updateScript -RootPath $rootItem.FullName
    if ($LASTEXITCODE -ne 0) {
        throw "LOT wrapper refresh failed with exit code $LASTEXITCODE."
    }

    [pscustomobject]@{
        LotPath       = $lotPath
        ComputersPath = $targetComputers
        ComputerCount = $computerCount
    }
}

function Start-ToolkitLot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LotPath,

        [Parameter(Mandatory = $true)]
        [string]$Mode
    )

    $wrapperName = switch ($Mode) {
        "Loop" { "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd"; break }
        "Once" { "Run-IntuneHybridJoinRepairWithPsExec-Once.cmd"; break }
        "LoopIgnoreRunGuard" { "Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd"; break }
        "OnceIgnoreRunGuard" { "Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd"; break }
        default { "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd" }
    }

    $wrapperPath = Join-Path $LotPath $wrapperName
    if (-not (Test-Path -LiteralPath $wrapperPath)) {
        throw "LOT wrapper not found: $wrapperPath"
    }

    $psexecPath = Join-Path $toolkitRoot "Scripts\PsExec.exe"
    $psexecCommand = Get-Command -Name "PsExec.exe" -CommandType Application -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $psexecPath -PathType Leaf) -and -not $psexecCommand) {
        throw "PsExec.exe not found. Place PsExec.exe in the toolkit Scripts folder or add it to PATH before launching the LOT."
    }

    Start-Process -FilePath "cmd.exe" -ArgumentList @("/k", "`"$wrapperPath`"") -WorkingDirectory $LotPath -Verb RunAs
}

$toolkitRoot = Get-ToolkitRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$titleFont = New-Object System.Drawing.Font("Segoe UI Semibold", 18)
$sectionFont = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$smallFont = New-Object System.Drawing.Font("Segoe UI", 8)
$statusFont = New-Object System.Drawing.Font("Consolas", 9)

$colorBackground = [System.Drawing.ColorTranslator]::FromHtml("#F3F6FA")
$colorPanel = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")
$colorPanelSoft = [System.Drawing.ColorTranslator]::FromHtml("#F8FBFE")
$colorHeader = [System.Drawing.ColorTranslator]::FromHtml("#061F36")
$colorHeaderPanel = [System.Drawing.ColorTranslator]::FromHtml("#143657")
$colorAccent = [System.Drawing.ColorTranslator]::FromHtml("#0F6CBD")
$colorInk = [System.Drawing.ColorTranslator]::FromHtml("#102A43")
$colorMuted = [System.Drawing.ColorTranslator]::FromHtml("#52677A")
$colorBorder = [System.Drawing.ColorTranslator]::FromHtml("#D7E1EA")
$colorTextBoxBorder = [System.Drawing.ColorTranslator]::FromHtml("#B9C8D7")
$colorSuccess = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
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

$form = New-Object System.Windows.Forms.Form
$form.Text = "Smart Intune Hybrid Join Toolkit - LOT Launcher"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(940, 640)
$form.MinimumSize = New-Object System.Drawing.Size(860, 560)
$form.Font = $font
$form.BackColor = $colorBackground

$logoIconPath = Resolve-LogoIconPath
$script:LogoIcon = $null
$script:LogoBitmap = $null
if (-not [string]::IsNullOrWhiteSpace($logoIconPath)) {
    try {
        $script:LogoIcon = New-Object System.Drawing.Icon($logoIconPath, 48, 48)
        $script:LogoBitmap = $script:LogoIcon.ToBitmap()
        $form.Icon = $script:LogoIcon
    }
    catch {
        $script:LogoIcon = $null
        $script:LogoBitmap = $null
    }
}

function Set-FlatButtonStyle {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Button]$Button,

        [Parameter(Mandatory = $true)]
        [System.Drawing.Color]$BackColor,

        [Parameter(Mandatory = $true)]
        [System.Drawing.Color]$ForeColor,

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
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]$Control,

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

function Set-LaunchButtonState {
    param([bool]$CanLaunch)

    if ($CanLaunch) {
        $launchButton.Enabled = $true
        Set-FlatButtonStyle -Button $launchButton -BackColor $colorSuccess -ForeColor ([System.Drawing.Color]::White) -BorderColor $colorSuccess
    }
    else {
        $launchButton.Enabled = $false
        Set-FlatButtonStyle -Button $launchButton -BackColor $colorDisabled -ForeColor $colorDisabledText -BorderColor $colorBorder
        $launchButton.Cursor = [System.Windows.Forms.Cursors]::Default
    }
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

function New-TextBox {
    param([switch]$ReadOnly)
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = "Fill"
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $box.ReadOnly = [bool]$ReadOnly
    $box.BorderStyle = "FixedSingle"
    $box.BackColor = if ($ReadOnly) { $colorPanelSoft } else { $colorPanel }
    $box.ForeColor = $colorInk
    return $box
}

function New-SectionPanel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

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
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null
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
        Panel   = $outer
        Content = $content
    }
}

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = "Fill"
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 3
$rootLayout.Padding = New-Object System.Windows.Forms.Padding(14)
$rootLayout.BackColor = $colorBackground
$rootLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 136))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54))) | Out-Null

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = "Fill"
$headerPanel.BackColor = $colorHeader
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(20, 16, 20, 16)

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = "Fill"
$headerLayout.ColumnCount = 2
$headerLayout.RowCount = 1
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 250))) | Out-Null

$headerTextPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerTextPanel.Dock = "Fill"
$headerTextPanel.ColumnCount = 1
$headerTextPanel.RowCount = 3
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null

$badgeLabel = New-Object System.Windows.Forms.Label
$badgeLabel.Text = "SMARTM365"
$badgeLabel.AutoSize = $true
$badgeLabel.BackColor = $colorAccent
$badgeLabel.ForeColor = [System.Drawing.Color]::White
$badgeLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8)
$badgeLabel.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Intune Hybrid Join LOT Launcher"
$titleLabel.Dock = "Fill"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = $titleFont
$titleLabel.TextAlign = "MiddleLeft"

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Operational LOT launcher"
$subtitleLabel.Dock = "Fill"
$subtitleLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#BFD8EC")
$subtitleLabel.TextAlign = "MiddleLeft"

$headerTextPanel.Controls.Add($badgeLabel, 0, 0)
$headerTextPanel.Controls.Add($titleLabel, 0, 1)
$headerTextPanel.Controls.Add($subtitleLabel, 0, 2)

$headerBrandLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerBrandLayout.Dock = "Fill"
$headerBrandLayout.ColumnCount = 2
$headerBrandLayout.RowCount = 1
$headerBrandLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 58))) | Out-Null
$headerBrandLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$logoPanel = New-Object System.Windows.Forms.Panel
$logoPanel.Width = 46
$logoPanel.Height = 46
$logoPanel.Margin = New-Object System.Windows.Forms.Padding(0, 7, 12, 0)
$logoPanel.BackColor = $colorHeaderPanel
$logoPanel.Padding = New-Object System.Windows.Forms.Padding(5)
Add-SoftBorder -Control $logoPanel -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#24577F"))

if ($script:LogoBitmap) {
    $logoPicture = New-Object System.Windows.Forms.PictureBox
    $logoPicture.Dock = "Fill"
    $logoPicture.SizeMode = "Zoom"
    $logoPicture.Image = $script:LogoBitmap
    $logoPicture.BackColor = $colorHeaderPanel
    $logoPanel.Controls.Add($logoPicture)
}
else {
    $logoFallback = New-Object System.Windows.Forms.Label
    $logoFallback.Dock = "Fill"
    $logoFallback.Text = "SM"
    $logoFallback.ForeColor = [System.Drawing.Color]::White
    $logoFallback.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
    $logoFallback.TextAlign = "MiddleCenter"
    $logoPanel.Controls.Add($logoFallback)
}

$headerBrandLayout.Controls.Add($logoPanel, 0, 0)
$headerBrandLayout.Controls.Add($headerTextPanel, 1, 0)

$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = "Fill"
$statusPanel.BackColor = $colorHeaderPanel
$statusPanel.Padding = New-Object System.Windows.Forms.Padding(12)
Add-SoftBorder -Control $statusPanel -BorderColor ([System.Drawing.ColorTranslator]::FromHtml("#24577F"))

$psexecLabel = New-Object System.Windows.Forms.Label
$psexecLabel.Dock = "Top"
$psexecLabel.Height = 24
$psexecLabel.ForeColor = [System.Drawing.Color]::White
$psexecLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)

$rootHintLabel = New-Object System.Windows.Forms.Label
$rootHintLabel.Dock = "Fill"
$rootHintLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#BFD8EC")
$rootHintLabel.Text = "Toolkit root detected."
$rootHintLabel.TextAlign = "BottomLeft"

$statusPanel.Controls.Add($rootHintLabel)
$statusPanel.Controls.Add($psexecLabel)

$psexecPath = Join-Path $toolkitRoot "Scripts\PsExec.exe"
if ((Test-Path -LiteralPath $psexecPath -PathType Leaf) -or (Get-Command -Name "PsExec.exe" -CommandType Application -ErrorAction SilentlyContinue)) {
    $psexecLabel.Text = "PsExec ready"
}
else {
    $psexecLabel.Text = "PsExec missing"
    $psexecLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FFD166")
}

$headerLayout.Controls.Add($headerBrandLayout, 0, 0)
$headerLayout.Controls.Add($statusPanel, 1, 0)
$headerPanel.Controls.Add($headerLayout)

$contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
$contentLayout.Dock = "Fill"
$contentLayout.ColumnCount = 1
$contentLayout.RowCount = 2
$contentLayout.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 10)
$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 228))) | Out-Null
$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$configSection = New-SectionPanel -Title "LOT configuration"
$activitySection = New-SectionPanel -Title "Activity"

$table = New-Object System.Windows.Forms.TableLayoutPanel
$table.Dock = "Fill"
$table.ColumnCount = 3
$table.RowCount = 5
$table.BackColor = $colorPanel
$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 140))) | Out-Null
$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 126))) | Out-Null
for ($i = 0; $i -lt 5; $i++) {
    $table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 38))) | Out-Null
}

$rootBox = New-TextBox -ReadOnly
$rootBox.Text = $toolkitRoot

$computerFileBox = New-TextBox
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse..."
$browseButton.Dock = "Fill"
Set-FlatButtonStyle -Button $browseButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$lotNameBox = New-TextBox

$modeCombo = New-Object System.Windows.Forms.ComboBox
$modeCombo.Dock = "Fill"
$modeCombo.DropDownStyle = "DropDownList"
$modeCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$modeCombo.FlatStyle = "Flat"
$modeCombo.BackColor = $colorPanel
$modeCombo.ForeColor = $colorInk
$modeCombo.Items.AddRange([object[]]@("Loop", "Once", "LoopIgnoreRunGuard", "OnceIgnoreRunGuard"))
$modeCombo.SelectedIndex = 0

$createdLotBox = New-TextBox -ReadOnly

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Dock = "Fill"
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$statusBox.Font = $statusFont
$statusBox.BackColor = $colorPanelSoft
$statusBox.ForeColor = $colorInk
$statusBox.BorderStyle = "FixedSingle"

$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = "Create LOT"
$createButton.Width = 116
Set-FlatButtonStyle -Button $createButton -BackColor $colorAccent -ForeColor ([System.Drawing.Color]::White) -BorderColor $colorAccent

$launchButton = New-Object System.Windows.Forms.Button
$launchButton.Text = "Launch LOT"
$launchButton.Width = 116
Set-LaunchButtonState -CanLaunch $false

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
$buttonPanel.Controls.Add($launchButton)
$buttonPanel.Controls.Add($createButton)

$table.Controls.Add((New-Label "Toolkit root"), 0, 0)
$table.Controls.Add($rootBox, 1, 0)
$table.SetColumnSpan($rootBox, 2)

$table.Controls.Add((New-Label "Computer file"), 0, 1)
$table.Controls.Add($computerFileBox, 1, 1)
$table.Controls.Add($browseButton, 2, 1)

$table.Controls.Add((New-Label "LOT folder"), 0, 2)
$table.Controls.Add($lotNameBox, 1, 2)
$table.SetColumnSpan($lotNameBox, 2)

$table.Controls.Add((New-Label "Launch mode"), 0, 3)
$table.Controls.Add($modeCombo, 1, 3)
$table.SetColumnSpan($modeCombo, 2)

$table.Controls.Add((New-Label "Created LOT"), 0, 4)
$table.Controls.Add($createdLotBox, 1, 4)
$table.SetColumnSpan($createdLotBox, 2)

$configSection.Content.Controls.Add($table)
$activitySection.Content.Controls.Add($statusBox)

$contentLayout.Controls.Add($configSection.Panel, 0, 0)
$contentLayout.Controls.Add($activitySection.Panel, 0, 1)

$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = "Fill"
$footerPanel.BackColor = $colorBackground

$footerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$footerLayout.Dock = "Fill"
$footerLayout.ColumnCount = 2
$footerLayout.RowCount = 1
$footerLayout.BackColor = $colorBackground
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 360))) | Out-Null

$footerStatus = New-Object System.Windows.Forms.Label
$footerStatus.Dock = "Fill"
$footerStatus.Text = "Ready"
$footerStatus.TextAlign = "MiddleLeft"
$footerStatus.ForeColor = $colorMuted
$footerStatus.Font = $smallFont

$footerLayout.Controls.Add($footerStatus, 0, 0)
$footerLayout.Controls.Add($buttonPanel, 1, 0)
$footerPanel.Controls.Add($footerLayout)

$rootLayout.Controls.Add($headerPanel, 0, 0)
$rootLayout.Controls.Add($contentLayout, 0, 1)
$rootLayout.Controls.Add($footerPanel, 0, 2)

$form.Controls.Add($rootLayout)

$script:CurrentLotPath = $null

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

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select computer list file"
    $dialog.Filter = "Computer list files (*.txt;*.csv)|*.txt;*.csv|Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false

    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $computerFileBox.Text = $dialog.FileName
        $lotNameBox.Text = Get-SafeLotName -ComputerFilePath $dialog.FileName
        Add-Status ("Selected computer file: {0}" -f $dialog.FileName)
    }
})

$createButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($computerFileBox.Text)) {
            throw "Select a computer list file first."
        }

        if ([string]::IsNullOrWhiteSpace($lotNameBox.Text)) {
            $lotNameBox.Text = Get-SafeLotName -ComputerFilePath $computerFileBox.Text
        }

        Add-Status ("Creating LOT folder from: {0}" -f $computerFileBox.Text)
        $result = New-ToolkitLotFolder -RootPath $toolkitRoot -ComputerFilePath $computerFileBox.Text -LotName $lotNameBox.Text
        $script:CurrentLotPath = $result.LotPath
        $createdLotBox.Text = $script:CurrentLotPath
        Set-LaunchButtonState -CanLaunch $true

        Add-Status ("Created {0} with {1} computer(s)." -f $result.LotPath,$result.ComputerCount)
        Add-Status "LOT wrappers refreshed."

        $answer = [System.Windows.Forms.MessageBox]::Show(
            $form,
            ("LOT created:`r`n{0}`r`n`r`nLaunch it now?" -f $result.LotPath),
            "Launch LOT",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Add-Status ("Launching LOT in {0} mode." -f $modeCombo.SelectedItem)
            Start-ToolkitLot -LotPath $script:CurrentLotPath -Mode ([string]$modeCombo.SelectedItem)
        }
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "LOT creation failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($script:CurrentLotPath)) {
            throw "Create a LOT folder first."
        }

        Add-Status ("Launching LOT in {0} mode." -f $modeCombo.SelectedItem)
        Start-ToolkitLot -LotPath $script:CurrentLotPath -Mode ([string]$modeCombo.SelectedItem)
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "LOT launch failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$closeButton.Add_Click({
    $form.Close()
})

$form.Add_Shown({
    $createButton.Focus() | Out-Null
})

$form.Add_FormClosed({
    if ($script:LogoBitmap) {
        $script:LogoBitmap.Dispose()
        $script:LogoBitmap = $null
    }

    if ($script:LogoIcon) {
        $script:LogoIcon.Dispose()
        $script:LogoIcon = $null
    }
})

if ($ValidateOnly) {
    Write-Host "Smart Intune Hybrid Join Toolkit LOT Launcher GUI validation completed."
    return
}

Add-Status "Ready. Select a computer list file to create a local LOT-* folder."
[void]$form.ShowDialog()
