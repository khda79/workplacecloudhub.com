<#
.SYNOPSIS
    Shared fixtures for the Smart IntunePackager Pester tests. All data is synthetic.

.VERSION
    1.0.0
#>

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-TestDetectionXml {
    param(
        [switch]$WithMsi,
        [string[]]$Omit = @(),
        [int64]$UnencryptedSize = 2048,
        [string]$SetupFile = 'ContosoSetup.msi'
    )

    $fields = [ordered]@{
        EncryptionKey        = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA='
        MacKey               = 'BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB='
        InitializationVector = 'CCCCCCCCCCCCCCCCCCCCCC=='
        Mac                  = 'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD='
        ProfileIdentifier    = 'ProfileVersion1'
        FileDigest           = 'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE='
        FileDigestAlgorithm  = 'SHA256'
    }
    $encryption = ($fields.Keys | Where-Object { $_ -notin $Omit } | ForEach-Object { '<{0}>{1}</{0}>' -f $_, $fields[$_] }) -join ''
    $msi = ''
    if ($WithMsi) {
        $msi = '<MsiInfo><MsiProductCode>{11111111-2222-3333-4444-555555555555}</MsiProductCode><MsiProductVersion>2.4.0</MsiProductVersion>' +
            '<MsiUpgradeCode>{66666666-7777-8888-9999-000000000000}</MsiUpgradeCode><MsiExecutionContext>System</MsiExecutionContext>' +
            '<MsiRequiresReboot>false</MsiRequiresReboot><MsiIsMachineInstall>true</MsiIsMachineInstall><MsiIsUserInstall>false</MsiIsUserInstall>' +
            '<MsiPublisher>Contoso Ltd</MsiPublisher></MsiInfo>'
    }
    return ('<ApplicationInfo xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ToolVersion="1.8.7.0">' +
        '<Name>{0}</Name><UnencryptedContentSize>{1}</UnencryptedContentSize><FileName>IntunePackage.intunewin</FileName><SetupFile>{0}</SetupFile>' +
        '<EncryptionInfo>{2}</EncryptionInfo>{3}</ApplicationInfo>') -f $SetupFile, $UnencryptedSize, $encryption, $msi
}

function New-TestIntuneWin {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$XmlText,
        [int]$PayloadSize = 4096
    )

    $directory = Split-Path -Path $Path -Parent
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $archive = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $xmlEntry = $archive.CreateEntry('IntuneWinPackage/Metadata/Detection.xml')
        $writer = [IO.StreamWriter]::new($xmlEntry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write($XmlText) } finally { $writer.Dispose() }

        $payloadEntry = $archive.CreateEntry('IntuneWinPackage/Contents/IntunePackage.intunewin', [IO.Compression.CompressionLevel]::NoCompression)
        $stream = $payloadEntry.Open()
        try {
            $bytes = [byte[]]::new($PayloadSize)
            [Random]::new(42).NextBytes($bytes)
            $stream.Write($bytes, 0, $bytes.Length)
        }
        finally { $stream.Dispose() }
    }
    finally { $archive.Dispose() }
    return $Path
}

function New-TestMsi {
    <#
    .SYNOPSIS
        Creates a minimal MSI database containing only a Property table.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$Properties
    )

    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = $null
    try {
        # 3 = msiOpenDatabaseModeCreate
        $database = $installer.OpenDatabase($Path, 3)
        $view = $database.OpenView('CREATE TABLE `Property` (`Property` CHAR(72) NOT NULL, `Value` LONGCHAR NOT NULL LOCALIZABLE PRIMARY KEY `Property`)')
        try { [void]$view.Execute(); [void]$view.Close() }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }

        # Parameterized insert: MSI SQL literals cannot contain quotes.
        $view = $database.OpenView('INSERT INTO `Property` (`Property`, `Value`) VALUES (?, ?)')
        try {
            foreach ($key in $Properties.Keys) {
                $record = $installer.CreateRecord(2)
                try {
                    $record.StringData(1) = [string]$key
                    $record.StringData(2) = [string]$Properties[$key]
                    [void]$view.Execute($record)
                }
                finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
            }
            [void]$view.Close()
        }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
        [void]$database.Commit()
    }
    finally {
        if ($database) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer)
    }
    return $Path
}

function New-TestDefinition {
    param([string]$DetectionType = 'MsiProductCode')

    $definition = New-IpAppDefinition
    $definition.App.DisplayName = 'Contoso Sample App'
    $definition.App.Publisher = 'Contoso Ltd'
    $definition.App.Version = '2.4.0'
    $definition.App.InformationUrl = 'https://www.contoso.example/sample-app'
    $definition.Program.InstallCommandLine = 'msiexec /i "ContosoSetup.msi" /qn /norestart'
    $definition.Program.UninstallCommandLine = 'msiexec /x {11111111-2222-3333-4444-555555555555} /qn /norestart'
    $definition.Requirements.Architectures = @('x64', 'arm64')
    $definition.Requirements.MinimumWindowsRelease = 'Windows11_23H2'
    $definition.Detection.Type = $DetectionType
    $definition.Detection.Msi.ProductCode = '{11111111-2222-3333-4444-555555555555}'
    $definition.Detection.Msi.ProductVersionOperator = 'greaterThanOrEqual'
    $definition.Detection.Msi.ProductVersion = '2.4.0'
    return $definition
}

function New-TestPngFile {
    param([Parameter(Mandatory = $true)][string]$Path)
    # 1x1 transparent PNG.
    $bytes = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==')
    [IO.File]::WriteAllBytes($Path, $bytes)
    return $Path
}
