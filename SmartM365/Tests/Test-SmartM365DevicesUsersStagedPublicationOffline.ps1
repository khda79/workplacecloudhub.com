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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCF0mngOW5n0YJs
# jnss5E31FzLD1TigpMgnwI4CzC0CnKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPVjRPFNmd+haEEDd0Gb9tOdzQJVCoWBmfOFP5Ig86lOMA0GCSqG
# SIb3DQEBAQUABIIBgKIh9UZoxGpDB2dfzBwufaMX8S3W8UJN6Xpq3Mtuu6FH7U9W
# M+kuwIk2D5AC7H5bxTd+Fewfh3VnMrXdPubweH8BaUm4YD3qVFr4mvH1oWZSx16B
# T5su2Ouh4WtisMLoV04iNZFQUZYVICFysqoTuIwFHKffRPjQvEZXmHpnb6zVgLew
# BS4WINtmUnBk4rSzeuUWSR7M0dwZfqXyCEfqp2yLnatPRekyxdKm5KWo+oMstmoD
# 2GFkXX9XHq+76WU9Pt8kn4WAWc08eYDj7BWHd+rwOJ1wpj8sYqp2Q4fJSsE/ld33
# 0KnMC/dR7wE7RRhM2dB21oKPrgOh9D5IOYYr/jHszXgxn/FLFYU9FZyQF5DCMdX4
# GYljMnUD40v9SuaX40C8rX0m0+wXAbRd3l0T3YV+/gW2PDp60NAZSwReJnMFbI8v
# ZzrJuWZCMdTxFXdBN7CcUfjvEhWW3/27X0fFkqK7JAAqDB+C5m/k84f4Z8Ixx0oz
# 8+sC2YbSAA4PoS9Y0qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEwOTIz
# MTRaMC8GCSqGSIb3DQEJBDEiBCBUUA7H56xUbsfPslN3l4ovHloJQFzQ6XAwLW92
# uyk/MjANBgkqhkiG9w0BAQEFAASCAgAj2pzvUuwXV4llVT48gQIPDqTRAN7NuB8+
# 16XcSLkpyg3WllyyP5g6CiGq2fArS4cdmaFu429OBlxvRaylDmlGX+rQUQMXMfuZ
# xd3+rnsIv/FwhT1CJAHESpM+4doEFdpJyHcEe2CP9ealKd8O4ePuxJNBKDBw3eDP
# 49CYVSMSsMguRx1WUaqlCU40vvJCU6z/YdRClGMb/JgbwUgpUpiiiSZyD5s4q88G
# PRrb7DZzlj9+ThP43uvp1PCNOGccC2IC0BKK+OELygtwNEpABcVfm86lGHHhsA0m
# nsDzQP02A1gEY0Vev+YWLUlMJ9W3gFPYADl156MBPeaOAbBTLGzTmvwHOLN+5RT9
# Pm01Cx3HZwDY2KgukwnzWnI7ri+l04o0hK1lmdA8PhxoZyQ9weN2hIdyBxtlEbWg
# dy3NhYOgi/q5c2c12E4z3w0QR3HqcITt5yojLAfOWIUu/Zo2MDBr5GTLCeu/yplh
# WBowNb5VHRRrt/vcswoQ5ZCWBUBIXVo8od7gvz0drGqnaCGuG3dfsgnX4jnPJPz3
# EUvrPZAsrUmNJxnvOHwKxkRV2ZDSBSfc7LgQ5ZgPJqjwZGtLF5rTTHj4hMo/ItIM
# hcHLX9nYOjNg7mYoh4FBLtCe/BrP76vNCoJg2cw/2egfkqzzoSP42EIL10TzaP/F
# fbxFcksbCA==
# SIG # End signature block
