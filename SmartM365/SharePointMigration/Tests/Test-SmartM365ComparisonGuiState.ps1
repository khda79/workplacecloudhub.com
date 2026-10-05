<#
.SYNOPSIS
    Offline checks for comparison dates and selected inventory availability.
.VERSION
    1.0.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Summary.ps1')
$guiPath = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | ForEach-Object Message) }
foreach ($name in @('Format-RunAgeText','Format-ItemAge','Get-ComparisonBadgeText',
    'Get-SelectedScanFile','Get-ComparisonRunState','Update-ComparisonRunState',
    'Update-ScanFileSelection','Update-HistoryRunState','Update-PermissionHistoryRunState',
    'Get-LatestCsvFile','Get-CsvFileItems','Set-ScanComboItems')) {
    $definition = $ast.Find({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    if (-not $definition) { throw "Missing GUI function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
function Set-Badge { param($Badge,$Label,$Text,$HasRun) $Label.Text = $Text }

$tempParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$root = [IO.Path]::GetFullPath((Join-Path $tempParent ('SharePointMigration-comparison-gui-' + [guid]::NewGuid().ToString('N'))))
if (-not $root.StartsWith($tempParent,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test directory.' }
$originalCulture = [Threading.Thread]::CurrentThread.CurrentCulture
try {
    [void](New-Item -ItemType Directory -Path $root)
    $now = [datetime]::new(2026,10,5,12,0,0)
    foreach ($culture in @('fr-FR','en-US')) {
        [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($culture)
        if ((Format-RunAgeText -Date $now.AddMinutes(-10) -Now $now) -ne '11:50 today (10 min ago)' -or
            (Format-RunAgeText -Date $now.AddHours(-3) -Now $now) -ne '09:00 today' -or
            (Format-RunAgeText -Date $now.AddDays(-1) -Now $now) -ne '12:00 yesterday' -or
            (Format-RunAgeText -Date $now.AddDays(-2) -Now $now) -ne '2026-10-03 12:00') {
            throw "Relative dates are incorrect under $culture."
        }
    }

    $reportDate = (Get-Date).Date.AddDays(-1).AddHours(12)
    $folder = Join-Path $root ('Fixture-files-' + $reportDate.ToString('yyyyMMdd-HHmmss') + '-test')
    [void](New-Item -ItemType Directory -Path $folder)
    [pscustomobject]@{ MatchedKeys=8; SourceUniqueKeys=10; TargetUniqueKeys=10 } |
        Export-Csv -LiteralPath (Join-Path $folder 'Summary.csv') -NoTypeInformation -Delimiter ';'
    # The folder was copied today; the report date must remain yesterday.
    $badge = Get-ComparisonBadgeText -Folder (Get-Item $folder) -MigrationName Fixture -Kind Files
    if ($badge -notmatch '^12:00 yesterday · 80[,.]00 %$') { throw "Copied comparison lost its original date or rate: $badge" }
    if ((Get-ComparisonBadgeText -Folder $null -MigrationName Fixture -Kind Files) -ne 'No run yet') {
        throw 'Absent comparison must not invent a date or rate.'
    }

    foreach ($name in @('cmbScanSrcFile','cmbScanTgtFile','cmbScanSrcPermFile','cmbScanTgtPermFile',
        'cmbHistoryOldFile','cmbHistoryNewFile','cmbPermHistoryOldFile','cmbPermHistoryNewFile')) {
        Set-Variable -Name $name -Value ([Windows.Controls.ComboBox]::new())
    }
    foreach ($name in @('btnRunCmpFiles','btnRunCmpPerms','btnRunHistory','btnRunPermHistory',
        'btnOpenScanSrc','btnOpenScanTgt','btnOpenScanSrcPerm','btnOpenScanTgtPerm')) {
        Set-Variable -Name $name -Value ([Windows.Controls.Button]::new())
    }
    foreach ($name in @('lblCmpFilesAvailability','lblCmpPermsAvailability',
        'lblScanSrcAge','lblScanTgtAge','lblScanSrcPermAge','lblScanTgtPermAge')) {
        Set-Variable -Name $name -Value ([Windows.Controls.TextBlock]::new())
    }
    foreach ($name in @('badgeScanSrc','badgeScanTgt','badgeScanSrcPerm','badgeScanTgtPerm')) {
        Set-Variable -Name $name -Value ([Windows.Controls.Border]::new())
    }
    # Attach the real selection handlers from the GUI to in-memory WPF controls.
    foreach ($line in (Get-Content -LiteralPath $guiPath | Where-Object {
        $_ -match '^\$cmbScan(Src|Tgt)(Perm)?File.Add_SelectionChanged\('
    })) { Invoke-Expression $line }

    Update-ComparisonRunState
    if ($btnRunCmpFiles.IsEnabled -or $btnRunCmpPerms.IsEnabled -or
        $lblCmpFilesAvailability.Text -notmatch 'Source scan unavailable.*Target scan unavailable') {
        throw 'Missing scans left comparison enabled or unexplained.'
    }
    $sourcePath = Join-Path $root 'SP2019-FileInventory-Fixture-20261005-100000.csv'
    $targetPath = Join-Path $root 'SPO-FileInventory-Fixture-20261005-100100.csv'
    'File' | Set-Content -LiteralPath $sourcePath
    'File' | Set-Content -LiteralPath $targetPath
    $source = Get-Item $sourcePath; $target = Get-Item $targetPath
    $sourceItems = @(Get-CsvFileItems -Directory $root -Filter 'SP2019-FileInventory*.csv')
    $targetItems = @(Get-CsvFileItems -Directory $root -Filter 'SPO-FileInventory*.csv')
    Set-ScanComboItems $cmbScanSrcFile $sourceItems $source
    if ($btnRunCmpFiles.IsEnabled -or $lblCmpFilesAvailability.Text -notmatch 'Target scan unavailable') {
        throw 'One missing target scan left comparison enabled.'
    }
    Set-ScanComboItems $cmbScanTgtFile $targetItems $target
    if (-not $btnRunCmpFiles.IsEnabled -or $lblCmpFilesAvailability.Visibility -ne 'Collapsed' -or
        $btnRunCmpPerms.IsEnabled) { throw 'Selected complete scans did not enable only the matching comparison.' }
    $cmbScanSrcFile.SelectedIndex = -1
    if ($btnRunCmpFiles.IsEnabled -or $lblCmpFilesAvailability.Text -notmatch 'Source scan unavailable') {
        throw 'Removing the source selection left comparison enabled.'
    }
    $cmbScanSrcFile.SelectedIndex = 0

    $errorPath = Join-Path $root ($target.BaseName + '-Errors.csv')
    'Scope;Message' | Set-Content -LiteralPath $errorPath
    Update-ComparisonRunState
    if ($btnRunCmpFiles.IsEnabled -or $lblCmpFilesAvailability.Text -notmatch 'Target scan is incomplete' -or
        (Get-ComparisonRunState $source $target).Ready) { throw 'An error sidecar did not block the selected scan.' }
    # Match the launcher: even a header-only error sidecar marks a rejected run.
    if (@(Get-CsvFileItems -Directory $root -Filter 'SPO-FileInventory*.csv').Count -ne 0) {
        throw 'Incomplete scan remained available in the selector.'
    }
    Set-ScanComboItems $cmbScanTgtFile @() $null
    if ($btnRunCmpFiles.IsEnabled) { throw 'Refresh removed the incomplete target but left comparison enabled.' }
    Remove-Item -LiteralPath $errorPath
    Set-ScanComboItems $cmbScanTgtFile $targetItems $target
    Set-ScanComboItems $cmbScanSrcPermFile $sourceItems $source
    Set-ScanComboItems $cmbScanTgtPermFile $targetItems $target
    if (-not $btnRunCmpFiles.IsEnabled -or -not $btnRunCmpPerms.IsEnabled) { throw 'Comparison did not recover after successful scans.' }

    Set-ScanComboItems $cmbHistoryOldFile $sourceItems $source
    Set-ScanComboItems $cmbHistoryNewFile $targetItems $target
    Set-ScanComboItems $cmbPermHistoryOldFile $sourceItems $source
    Set-ScanComboItems $cmbPermHistoryNewFile $targetItems $target
    # Source history can remain usable while the destination selection is absent.
    $cmbScanTgtFile.SelectedIndex = -1
    Update-HistoryRunState; Update-PermissionHistoryRunState
    if (-not $btnRunHistory.IsEnabled -or -not $btnRunPermHistory.IsEnabled) { throw 'History incorrectly depends on the destination selection.' }
    if ((Get-ComparisonRunState $source $source -History).Ready) { throw 'History accepted the same scan twice.' }
    Remove-Item -LiteralPath $targetPath
    Update-HistoryRunState; Update-PermissionHistoryRunState; Update-ComparisonRunState
    if ($btnRunHistory.IsEnabled -or $btnRunPermHistory.IsEnabled -or $btnRunCmpPerms.IsEnabled -or
        $lblCmpPermsAvailability.Text -notmatch 'no longer available') { throw 'Deleted scan still enabled comparison.' }

    # Exercise the launch path, replacing only its modal dialog boundary.
    $invoke = $ast.Find({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-MigrationAction'
    }, $true)
    Add-Type -TypeDefinition @'
public static class SmartM365ComparisonTestDialog {
    public static int WarningCount = 0;
    public static object Show(object message, object title, object buttons, object image) {
        WarningCount++;
        return null;
    }
}
'@
    . ([scriptblock]::Create($invoke.Extent.Text.Replace('[System.Windows.MessageBox]', '[SmartM365ComparisonTestDialog]')))
    function Get-MigrationStatus {
        param($Migration)
        return [pscustomobject]@{
            SourceFileCsvItems=$sourceItems; SourceFileCsv=$source
            TargetFileCsvItems=@(); TargetFileCsv=$null
            SourcePermCsvItems=$sourceItems; SourcePermCsv=$source
            TargetPermCsvItems=@(); TargetPermCsv=$null
        }
    }
    function New-SmartM365GuiActivity { throw 'A blocked comparison tried to create an activity.' }
    function Start-Process { throw 'A blocked comparison tried to launch a worker.' }
    $script:CurrentMigration = [pscustomobject]@{ Name='Fixture' }
    $script:AppName = 'Offline comparison test'
    foreach ($action in @('CompareFiles','ComparePermissions','CompareScanHistory','ComparePermissionScanHistory')) {
        Invoke-MigrationAction $action
    }
    if ([SmartM365ComparisonTestDialog]::WarningCount -ne 4) { throw 'Direct invocation did not stop every unavailable comparison before launch.' }
    Write-Host 'PASS: relative comparison dates, missing/error/deleted scan gates, selection changes, history and blocked worker launches.'
}
finally {
    [Threading.Thread]::CurrentThread.CurrentCulture = $originalCulture
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAseZfpMcdm7xpz
# 0Fk1XBh7Qa2Yd1lxRzxkq85xL0c76aCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCFch/unyDtfIsX92UfLEog
# C7lO7ENoniuYvar6G2konTANBgkqhkiG9w0BAQEFAASCAYBAbAvUz+WtHBaENckZ
# gs2kHy0BPcVl+j5cyCYN2cD1jKV6cUhzesZbHLflMQrunP/x6DJIVeox8cvDCiaI
# mjObwEa6Pcc/f1rm3wHFUGujBngZAdLBuzeTALGysWUcOG1Pgg8qp8D4BWxytjyn
# B+2z0hMw0J9sUA8+R7pZ6rwblwrhd56nJOYBiCMGNMUXbNTHvZKSrmkH2rtk9ZTH
# 6O9OUWiyy0MdjhQ4Zo4NLRoG9+7vXoYIj3x9tKhp9CP4cMvdw1e/H4vVSCr2bIy3
# jSLw3ieAmkIvzFZPL9IZBGP7r7MVmO4LjDLMGhlFlBMCcAdE6X/rxjEuo7kFHiFf
# abBNvhJuGpNxJgP1AvPHo2tP8h8iqEV8BMdKabm2HjB9Euf4TbsMKQL1CiRSW6By
# CoR4lVKm5PprtJdAPcsrzeK/DmZriW8pwk8z0lIrZQ6ZfYp4o47+LHcBogbRhTfi
# OGBHOazF1FOInP5XPIFN+opPYgyDBvpeKsxOgdL8lbGK53I=
# SIG # End signature block
