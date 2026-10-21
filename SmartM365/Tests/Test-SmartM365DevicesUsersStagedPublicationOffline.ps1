#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged DATA-LAST publication for the Intune devices, Entra devices and active
users producers: with the previous run's different CSVs in DATA-LAST, a running or failed run neither
changes, uploads nor historizes them, and a complete run promotes this run's CSVs with the receipt.
.VERSION
1.0.0
.NOTES
Each producer is replayed through the same Core publication calls as its export stage. Synthetic TEMP
data only; SharePoint upload and WeeklyHistory are replaced by recorders. No tenant, Graph, API, mail
or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Synthetic TEMP fixtures only.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.95' -Force -PassThru
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-DevicesUsersStaged-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $temporary 'DATA-LAST'
$output = Join-Path $temporary 'DATA-ALL/M365'
New-Item -ItemType Directory -Path $latest, $output | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths', 'RetentionMaxCSV', 'csvFilePath1', 'csvFilePath2', 'csvFilePath3')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction Ignore
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'; $global:RetentionMaxCSV = 0
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
function Get-Hash([string]$Path) { if (Test-Path -LiteralPath $Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash } else { '' } }
function Get-LatestHashTable([string[]]$Files) { $result = @{}; foreach ($file in $Files) { $result[$file] = Get-Hash (Join-Path $latest $file) }; return $result }

& $core {
    $script:Recorded = New-Object 'Collections.Generic.List[object]'
    function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add([pscustomobject]@{ Kind = 'Upload'; Path = $LocalFilePath; SHA256 = (Get-FileHash -LiteralPath $LocalFilePath -Algorithm SHA256).Hash }) }
    function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param([string[]]$SourceFiles, [string]$TimestampedPath) $null = $TimestampedPath; $script:Recorded.Add([pscustomobject]@{ Kind = 'History'; Path = $SourceFiles[0]; SHA256 = (Get-FileHash -LiteralPath $SourceFiles[0] -Algorithm SHA256).Hash }) }
    function script:Remove-SmartM365SharePointTimestampedCsvOlderThan { param($TimestampedPath, $RetentionDays) $null = $TimestampedPath, $RetentionDays }
    function script:Assert-SmartM365CsvDataCompleteness { param($Data, $Columns, $BaseFileName, $TimestampedPath, $LatestPath) $null = $Data, $Columns, $BaseFileName, $TimestampedPath, $LatestPath }
    function script:WriteLog { param($Message, $Level) $null = $Message, $Level }
}
function Get-Recorded { & $core { $script:Recorded.ToArray() } }
function Clear-Recording { & $core { $script:Recorded.Clear() } }
$script:stamp = 0
function New-Stamp { $script:stamp++; '2026102{0}_0{0}0000' -f $script:stamp }
function New-Row([string]$Label, [int]$Index) { [pscustomobject]@{ Id = "$Label-$Index"; DisplayName = "$Label item $Index"; CollectedAtUtc = '2026-10-21T00:00:00Z' } }

# Same Core calls as each producer's export stage (see the producer files for the full column lists).
$producers = @(
    [pscustomobject]@{
        Name = 'Devices'; Script = 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1'; Scope = 'CMDB:managed,hardware'
        Files = @('Intune_ManagedDevices_All.csv', 'Intune_DeviceHardware_All.csv', 'Intune_Devices_Inventory.csv')
        Export = {
            param([string]$Label, [int]$Count)
            $rows = @(1..$Count | ForEach-Object { New-Row $Label $_ })
            foreach ($base in @('Intune_ManagedDevices_All', 'Intune_DeviceHardware_All')) {
                Export-SmartM365Csv -Data $rows -Columns @('Id', 'DisplayName', 'CollectedAtUtc') -TimestampedPath (Join-Path $output ('{0}_{1}.csv' -f $base, (New-Stamp))) -LatestPath (Join-Path $latest "$base.csv") -NoWeeklyHistory | Out-Null
            }
            ExportAndCopyCsv -BaseFileName 'Intune_Devices_Inventory' -OutputPath $output -GlobalPath $latest -Data $rows -Encoding 'UTF8' -NoTypeInformation
        }
    },
    [pscustomobject]@{
        Name = 'EntraDevices'; Script = 'SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1'; Scope = 'CMDB:entra_devices'
        Files = @('M365_EntraDevices_All.csv', 'M365_Entra_Devices.csv', 'M365_Entra_Devices_HardwareId_Pending.csv')
        Export = {
            param([string]$Label, [int]$Count)
            $rows = @(1..$Count | ForEach-Object { New-Row $Label $_ })
            Export-SmartM365Csv -Data $rows -Columns @('Id', 'DisplayName', 'CollectedAtUtc') -TimestampedPath (Join-Path $output ('M365_EntraDevices_All_{0}.csv' -f (New-Stamp))) -LatestPath (Join-Path $latest 'M365_EntraDevices_All.csv') -NoWeeklyHistory | Out-Null
            ExportAndCopyCsv -BaseFileName 'M365_Entra_Devices' -OutputPath $output -GlobalPath $latest -Data $rows -Encoding 'UTF8' -NoTypeInformation
            # Header-only pending report: DATA-LAST copy through Core, current copy in DATA-ALL.
            $empty = Export-SmartM365Csv -BaseFileName 'M365_Entra_Devices_HardwareId_Pending' -OutputPath $output -GlobalPath $latest -Data $null -Columns @('Id', $Label) -Encoding 'UTF8'
            Copy-SmartM365FileAtomically -SourcePath $empty.TimestampedPath -DestinationPath (Join-Path $output 'M365_Entra_Devices_HardwareId_Pending.csv')
            Start-Sleep -Milliseconds 1100
        }
    },
    [pscustomobject]@{
        Name = 'ActiveUsers'; Script = 'SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1'; Scope = 'CMDB:users'
        Files = @('M365_Users_Active.csv')
        Export = {
            param([string]$Label, [int]$Count)
            if ($Count -eq 0) {
                $publication = Export-SmartM365Csv -BaseFileName 'M365_Users_Active' -OutputPath $output -GlobalPath $latest -Data @() -Columns @('Display name', 'User principal name', 'Object Id')
                Copy-SmartM365FileAtomically -SourcePath $publication.TimestampedPath -DestinationPath (Join-Path $output 'M365_Users_Active.csv')
                Start-Sleep -Milliseconds 1100
                return
            }
            $rows = @(1..$Count | ForEach-Object { [pscustomobject]@{ 'Display name' = "$Label user $_"; 'User principal name' = "$Label$_@contoso.test"; 'Object Id' = [guid]::NewGuid().ToString() } })
            ExportAndCopyCsvFromConvert -BaseFileName 'M365_Users_Active' -OutputPath $output -GlobalPath $latest -Data $rows -Encoding 'UTF8' -NoTypeInformation -Delimiter ','
            Start-Sleep -Milliseconds 1100
        }
    }
)

try {
    foreach ($producer in $producers) {
        $path = Join-Path $root $producer.Script
        $name = [IO.Path]::GetFileNameWithoutExtension($path)
        $receiptPath = Join-Path $latest "SmartInventory_$name.current.json.txt"
        $runPath = Join-Path $latest "SmartInventory_$name.run.json.txt"
        $files = $producer.Files

        # Static wiring: staging right after the receipt, staging Core, one version, no direct DATA-LAST write.
        $text = [IO.File]::ReadAllText($path)
        Check ($text -match "(?m)^\s*Start-SmartM365CmdbSourceReceipt [^\r\n]*\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") "$($producer.Name) does not enable staging right after its receipt."
        Check ($text -match "-MinimumVersion '1\.0\.(?:9[5-9]|[1-9]\d\d)'") "$($producer.Name) does not require the staging Core."
        $version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
        Check ($version -and $text -match ('\$ScriptVersion = "' + [regex]::Escape($version) + '"')) "$($producer.Name) header and runtime versions differ."
        $ast = [Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$null)
        $direct = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @('Copy-SmartM365FileAtomically', 'Write-SmartM365CsvAtomically', 'Copy-Item', 'Move-Item', 'Export-Csv', 'Set-Content', 'Out-File') -and $n.Extent.Text -match 'LatestCsvFolderPath|GlobalPath|csvFilePath3|\.LatestPath' }, $true))
        Check ($direct.Count -eq 0) ("$($producer.Name) writes DATA-LAST outside Core publication: " + (($direct | ForEach-Object { $_.Extent.StartLineNumber }) -join ','))

        # Previous run, published directly.
        $global:csvGeneratedPaths = $null
        Start-SmartM365CmdbSourceReceipt -ScriptPath $path -SourceRootPath $latest
        & $producer.Export 'yesterday' 2
        Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $producer.Scope
        Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
        $yesterdayReceipt = Read-Json $receiptPath
        $yesterday = Get-LatestHashTable $files
        Check ($yesterdayReceipt.Status -eq 'Completed' -and @($files | Where-Object { -not $yesterday[$_] }).Count -eq 0) "The previous $($producer.Name) run did not publish."

        # Running run with different data (and, for active users, the empty-export path), then failed.
        foreach ($count in @(4, 0)) {
            if ($count -eq 0 -and $producer.Name -ne 'ActiveUsers') { continue }
            Clear-Recording
            $global:csvGeneratedPaths = $null
            Start-SmartM365CmdbSourceReceipt -ScriptPath $path -SourceRootPath $latest
            Enable-SmartM365StagedSourcePublication
            & $producer.Export 'today' $count
            $during = Get-LatestHashTable $files
            Check (@($files | Where-Object { $during[$_] -ne $yesterday[$_] }).Count -eq 0) "A running $($producer.Name) run changed DATA-LAST (rows=$count)."
            Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) "A running $($producer.Name) run changed the receipt (rows=$count)."
            Check ((Read-Json $runPath).StagedPublication -eq $true) "A running $($producer.Name) run is not staged (rows=$count)."
            Check (@(Get-Recorded | Where-Object { $_.Path -like "$latest*" }).Count -eq 0) "A running $($producer.Name) run uploaded or historized a DATA-LAST file (rows=$count)."
            Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
            $after = Get-LatestHashTable $files
            Check (@($files | Where-Object { $after[$_] -ne $yesterday[$_] }).Count -eq 0 -and (Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) "A failed $($producer.Name) run changed DATA-LAST (rows=$count)."
            Check (@(Get-ChildItem -LiteralPath (Join-Path $latest ".staging/$name") -Force).Count -eq 0) "A failed $($producer.Name) run left its staging folder."
        }

        # Complete run: this run's CSVs replace yesterday's with the receipt; uploads follow by canonical path.
        Clear-Recording
        $global:csvGeneratedPaths = $null
        Start-SmartM365CmdbSourceReceipt -ScriptPath $path -SourceRootPath $latest
        Enable-SmartM365StagedSourcePublication
        & $producer.Export 'today' 3
        Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $producer.Scope
        Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
        $receipt = Read-Json $receiptPath
        $final = Get-LatestHashTable $files
        Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $yesterdayReceipt.RunId) "A complete $($producer.Name) run did not publish its receipt."
        foreach ($file in $files) {
            $entry = @($receipt.Files | Where-Object File -eq $file)
            Check ($final[$file] -ne $yesterday[$file] -and ($entry.Count -eq 0 -or $entry[0].SHA256 -eq $final[$file])) "Promoted $($producer.Name) CSV and receipt disagree: $file"
        }
        $uploads = @(Get-Recorded | Where-Object { $_.Kind -eq 'Upload' -and $_.Path -like "$latest*" })
        Check ($uploads.Count -eq $files.Count -and @($uploads | Where-Object { $_.Path -like '*.staging*' -or $final[[IO.Path]::GetFileName($_.Path)] -ne $_.SHA256 }).Count -eq 0) "Promoted $($producer.Name) CSVs were not uploaded by canonical path."
        Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
        $script:checks++
    }
}
finally {
    try { Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction Ignore }
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks devices and users staged publication checks; PowerShell $($PSVersionTable.PSVersion)."
