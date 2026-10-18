#Requires -Version 7.0
<#
.SYNOPSIS
    Creates the fictional "Contoso Sample App" MSI used by the example profile.

.DESCRIPTION
    Builds a minimal per-machine MSI with the Windows Installer COM API. The package contains
    no files: it only writes HKEY_LOCAL_MACHINE\SOFTWARE\Contoso\IntunePackagerSample\Version
    (64-bit view) so it installs and uninstalls cleanly on a test device. ProductCode and
    UpgradeCode are fixed fictional GUIDs that match Contoso-SampleApp.profile.json.template;
    the package code is new for each build.

.VERSION
    1.1.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com

.EXAMPLE
    pwsh -NoProfile -ExecutionPolicy Bypass -File .\Examples\SmartM365-IntunePackager-New-SampleMsi.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputFolder = (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'SmartM365\IntunePackager\Samples\Contoso-SampleApp\Source'),
    [ValidatePattern('^\d+\.\d+\.\d+$')][string]$ProductVersion = '1.0.0',
    # Adds an incompressible ContosoSampleApp-Payload.bin of this size next to the MSI (large upload tests).
    [ValidateRange(0, 4096)][int]$PayloadSizeMB = 0,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$productCode = '{5F0C3B2A-6D41-4E8B-9A57-2C1D7E3F9B10}'
$upgradeCode = '{A3E1D6C4-2B7F-4A90-8C35-6E9F1B2D4C78}'
$componentId = '{0B6E2D4F-8A13-4C57-9E21-7D3F5A9C1E64}'
$registryKey = 'SOFTWARE\Contoso\IntunePackagerSample'

function Write-SampleLog {
    param([string]$Message)
    Microsoft.PowerShell.Utility\Write-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Invoke-MsiSql {
    param($Database, [string]$Sql)
    $view = $Database.OpenView($Sql)
    try { [void]$view.Execute(); [void]$view.Close() }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
}

function Add-MsiRows {
    <#
    .SYNOPSIS
        Inserts rows with a parameterized statement. Integers use IntegerData, $null stays null.
    #>
    param($Installer, $Database, [string]$Table, [string[]]$Columns, [object[][]]$Rows)

    $columnList = ($Columns | ForEach-Object { '`{0}`' -f $_ }) -join ', '
    $placeholders = (@('?') * $Columns.Count) -join ', '
    $view = $Database.OpenView(('INSERT INTO `{0}` ({1}) VALUES ({2})' -f $Table, $columnList, $placeholders))
    try {
        foreach ($row in $Rows) {
            $record = $Installer.CreateRecord($Columns.Count)
            try {
                for ($i = 0; $i -lt $Columns.Count; $i++) {
                    $value = $row[$i]
                    if ($null -eq $value) { continue }
                    if ($value -is [int]) { $record.IntegerData($i + 1) = $value }
                    else { $record.StringData($i + 1) = [string]$value }
                }
                [void]$view.Execute($record)
            }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
        }
        [void]$view.Close()
    }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
}

[IO.Directory]::CreateDirectory($OutputFolder) | Out-Null
$msiPath = Join-Path -Path $OutputFolder -ChildPath 'ContosoSampleApp.msi'
if (Test-Path -LiteralPath $msiPath) {
    if (-not $Force) { throw "The sample MSI already exists. Use -Force to replace it: $msiPath" }
    Remove-Item -LiteralPath $msiPath -Force
}

$installer = New-Object -ComObject WindowsInstaller.Installer
$database = $null
try {
    # 3 = msiOpenDatabaseModeCreate
    $database = $installer.OpenDatabase($msiPath, 3)

    Invoke-MsiSql $database 'CREATE TABLE `Property` (`Property` CHAR(72) NOT NULL, `Value` LONGCHAR NOT NULL LOCALIZABLE PRIMARY KEY `Property`)'
    Invoke-MsiSql $database 'CREATE TABLE `Directory` (`Directory` CHAR(72) NOT NULL, `Directory_Parent` CHAR(72), `DefaultDir` CHAR(255) NOT NULL LOCALIZABLE PRIMARY KEY `Directory`)'
    Invoke-MsiSql $database 'CREATE TABLE `Component` (`Component` CHAR(72) NOT NULL, `ComponentId` CHAR(38), `Directory_` CHAR(72) NOT NULL, `Attributes` SHORT NOT NULL, `Condition` CHAR(255), `KeyPath` CHAR(72) PRIMARY KEY `Component`)'
    Invoke-MsiSql $database 'CREATE TABLE `Feature` (`Feature` CHAR(38) NOT NULL, `Feature_Parent` CHAR(38), `Title` CHAR(64) LOCALIZABLE, `Description` CHAR(255) LOCALIZABLE, `Display` SHORT, `Level` SHORT NOT NULL, `Directory_` CHAR(72), `Attributes` SHORT NOT NULL PRIMARY KEY `Feature`)'
    Invoke-MsiSql $database 'CREATE TABLE `FeatureComponents` (`Feature_` CHAR(38) NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Feature_`, `Component_`)'
    Invoke-MsiSql $database 'CREATE TABLE `Registry` (`Registry` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` LONGCHAR LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Registry`)'
    Invoke-MsiSql $database 'CREATE TABLE `InstallExecuteSequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`)'
    Invoke-MsiSql $database 'CREATE TABLE `InstallUISequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`)'

    Add-MsiRows $installer $database 'Property' @('Property', 'Value') @(
        , @('ProductCode', $productCode)
        , @('UpgradeCode', $upgradeCode)
        , @('ProductName', 'Contoso Sample App')
        , @('ProductVersion', $ProductVersion)
        , @('Manufacturer', 'Contoso Ltd')
        , @('ProductLanguage', '1033')
        , @('ALLUSERS', '1')
        , @('ARPNOMODIFY', '1')
        , @('ARPNOREPAIR', '1')
    )
    Add-MsiRows $installer $database 'Directory' @('Directory', 'Directory_Parent', 'DefaultDir') @(, @('TARGETDIR', $null, 'SourceDir'))
    # Attributes 260 = registry key path (4) + 64-bit component (256).
    Add-MsiRows $installer $database 'Component' @('Component', 'ComponentId', 'Directory_', 'Attributes', 'Condition', 'KeyPath') @(, @('ContosoSampleRegistry', $componentId, 'TARGETDIR', 260, $null, 'ContosoSampleVersion'))
    Add-MsiRows $installer $database 'Feature' @('Feature', 'Feature_Parent', 'Title', 'Description', 'Display', 'Level', 'Directory_', 'Attributes') @(, @('Main', $null, 'Contoso Sample App', 'Fictional sample package.', 1, 1, 'TARGETDIR', 0))
    Add-MsiRows $installer $database 'FeatureComponents' @('Feature_', 'Component_') @(, @('Main', 'ContosoSampleRegistry'))
    # Root 2 = HKEY_LOCAL_MACHINE.
    Add-MsiRows $installer $database 'Registry' @('Registry', 'Root', 'Key', 'Name', 'Value', 'Component_') @(, @('ContosoSampleVersion', 2, $registryKey, 'Version', '[ProductVersion]', 'ContosoSampleRegistry'))
    Add-MsiRows $installer $database 'InstallExecuteSequence' @('Action', 'Condition', 'Sequence') @(
        , @('CostInitialize', $null, 800)
        , @('FileCost', $null, 900)
        , @('CostFinalize', $null, 1000)
        , @('InstallValidate', $null, 1400)
        , @('InstallInitialize', $null, 1500)
        , @('ProcessComponents', $null, 1600)
        , @('UnpublishFeatures', $null, 1800)
        , @('RemoveRegistryValues', $null, 2600)
        , @('WriteRegistryValues', $null, 5000)
        , @('RegisterProduct', $null, 6100)
        , @('PublishFeatures', $null, 6300)
        , @('PublishProduct', $null, 6400)
        , @('InstallFinalize', $null, 6600)
    )
    Add-MsiRows $installer $database 'InstallUISequence' @('Action', 'Condition', 'Sequence') @(
        , @('CostInitialize', $null, 800)
        , @('FileCost', $null, 900)
        , @('CostFinalize', $null, 1000)
        , @('ExecuteAction', $null, 1300)
    )

    # 20 = maximum number of summary properties updated.
    $summary = $database.SummaryInformation(20)
    try {
        $summary.Property(2) = 'Installation Database'
        $summary.Property(3) = 'Contoso Sample App'
        $summary.Property(4) = 'Contoso Ltd'
        $summary.Property(7) = 'x64;1033'
        $summary.Property(9) = '{' + [guid]::NewGuid().ToString().ToUpperInvariant() + '}'
        $summary.Property(14) = 200
        $summary.Property(15) = 2
        $summary.Property(19) = 2
        [void]$summary.Persist()
    }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($summary) }

    [void]$database.Commit()
}
finally {
    if ($null -ne $database) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer)
}

$payloadPath = Join-Path -Path $OutputFolder -ChildPath 'ContosoSampleApp-Payload.bin'
if ($PayloadSizeMB -gt 0) {
    # Random bytes cannot be compressed, so the .intunewin stays larger than the requested size.
    $buffer = [byte[]]::new(1MB)
    $stream = [IO.File]::Create($payloadPath)
    try {
        for ($i = 0; $i -lt $PayloadSizeMB; $i++) {
            [Security.Cryptography.RandomNumberGenerator]::Fill($buffer)
            $stream.Write($buffer, 0, $buffer.Length)
        }
    }
    finally { $stream.Dispose() }
    Write-SampleLog ("Created incompressible payload: {0} ({1} MB)" -f $payloadPath, $PayloadSizeMB)
}
elseif (Test-Path -LiteralPath $payloadPath) {
    Remove-Item -LiteralPath $payloadPath -Force
}

Write-SampleLog "Created fictional sample MSI: $msiPath"
Write-SampleLog "ProductCode=$productCode; ProductVersion=$ProductVersion; Registry=HKEY_LOCAL_MACHINE\$registryKey\Version"
