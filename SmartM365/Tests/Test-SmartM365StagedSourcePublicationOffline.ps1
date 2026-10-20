#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged source publication: a running or failed run never changes DATA-LAST,
a complete run promotes its CSVs under the Publishing fence and writes the receipt last.
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No tenant, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.93' -Force -PassThru
Import-Module (Join-Path $root 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1') -Force
$producer = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-IntuneManagedApps-Inventory.ps1' | Select-Object -First 1
$files = @('Intune_ManagedApps.csv', 'Intune_ManagedAppAssignments.csv', 'Intune_ManagedAppInstallSummary.csv', 'Intune_ManagedAppDeviceStates.csv')
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-StagedPublication-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $temporary 'DATA-LAST'
$output = Join-Path $temporary 'DATA-ALL/Intune/WorkplaceScope/ManagedApps'
$shared = Join-Path $temporary 'DATA-ALL/Orchestrator'
New-Item -ItemType Directory -Path $latest, $output, $shared | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$receiptPath = Join-Path $latest 'SmartInventory_SmartM365-IntuneManagedApps-Inventory.current.json.txt'
$runPath = Join-Path $latest 'SmartInventory_SmartM365-IntuneManagedApps-Inventory.run.json.txt'
$stagingProducer = Join-Path $latest '.staging/SmartM365-IntuneManagedApps-Inventory'
function Get-LatestHashTable { $result = @{}; foreach ($file in $files) { $path = Join-Path $latest $file; $result[$file] = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } else { '' } }; return $result }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
function Publish-Set([string]$Label) {
    foreach ($file in $files) {
        $rows = @([pscustomobject]@{ AppId = [guid]::NewGuid().ToString(); Name = "$Label $file" })
        $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss') + [guid]::NewGuid().ToString('N').Substring(0, 4)
        Publish-SmartM365Csv -Data $rows -Columns @('AppId', 'Name') -TimestampedPath (Join-Path $output ("{0}_{1}.csv" -f [IO.Path]::GetFileNameWithoutExtension($file), $stamp)) `
            -LatestPath (Join-Path $latest $file) -RetentionMaxCsv 0 -SingleSerialization | Out-Null
    }
}
try {
    # Recorders replace SharePoint upload and WeeklyHistory inside the module; they also capture
    # the publication state seen at each promotion step.
    & $core {
        $script:Recorded = New-Object 'Collections.Generic.List[object]'
        function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add([pscustomobject]@{ Kind = 'Upload'; Path = $LocalFilePath }) }
        function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param([string[]]$SourceFiles, [string]$TimestampedPath) $script:Recorded.Add([pscustomobject]@{ Kind = 'History'; Path = $SourceFiles[0]; Timestamped = $TimestampedPath }) }
        function script:WriteLog {
            param($Message, $Level)
            $null = $Level
            if ([string]$Message -match 'promoted from source staging') {
                $run = Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName(($Message -replace '^CSV latest copy written to: (.+) \(promoted.*$', '$1'))) -Filter '*.run.json.txt' | Select-Object -First 1
                $script:Recorded.Add([pscustomobject]@{ Kind = 'Promoted'; Path = $Message; RunStatus = (Get-Content -LiteralPath $run.FullName -Raw | ConvertFrom-Json).Status })
            }
        }
    }
    function Get-Recorded { & $core { $script:Recorded.ToArray() } }
    function Clear-Recorded { & $core { $script:Recorded.Clear() } }

    # Baseline: a direct, unstaged complete run publishes the previous good set.
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Publish-Set 'baseline'
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $baselineReceipt = Read-Json $receiptPath
    $baselineHashes = Get-LatestHashTable
    Check ($baselineReceipt.Status -eq 'Completed' -and @($baselineReceipt.Files).Count -eq 4) 'Baseline receipt not completed.'
    Check (-not (Read-Json $runPath).PSObject.Properties['StagedPublication']) 'An unstaged run declared staging.'
    Clear-Recorded

    # Resolve leaves paths unchanged without staging.
    Check ((Resolve-SmartM365SourcePublicationPath -Path (Join-Path $latest 'X.csv')) -eq (Join-Path $latest 'X.csv')) 'Resolve changed a path without an active staging.'

    # Staged run, interrupted before completion: an orphan from a crashed run is cleaned first.
    New-Item -ItemType Directory -Path (Join-Path $stagingProducer 'crashedrun') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $stagingProducer 'crashedrun/Intune_ManagedApps.csv') -Value 'TenantKey,Name'
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Check (-not (Test-Path -LiteralPath (Join-Path $stagingProducer 'crashedrun'))) 'An orphaned staging folder survived the next run start.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Collecting' -and $run.StagedPublication -eq $true) 'Run state does not declare staged collection.'
    $state = Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job ([pscustomobject]@{ ScriptPath = $producer.FullName })
    Check ($state.Known -and $state.StagedPublication -and $state.Status -eq 'Collecting') 'Orchestrator does not read the staged run state.'
    Check ((Resolve-SmartM365SourcePublicationPath -Path (Join-Path $latest 'X.txt')) -eq (Join-Path $latest 'X.txt')) 'A non-CSV path was staged.'
    Check ((Resolve-SmartM365SourcePublicationPath -Path (Join-Path $output 'X.csv')) -eq (Join-Path $output 'X.csv')) 'A CSV outside DATA-LAST was staged.'
    Publish-Set 'running'
    $stagedFolder = @(Get-ChildItem -LiteralPath $stagingProducer -Directory)
    Check ($stagedFolder.Count -eq 1 -and @(Get-ChildItem -LiteralPath $stagedFolder[0].FullName -Filter '*.csv').Count -eq 4) 'Run CSVs were not written to its staging folder.'
    $during = Get-LatestHashTable
    Check (@($files | Where-Object { $during[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A running staged run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A running staged run changed the receipt.'
    Check ((Read-Json $runPath).Status -eq 'Collecting') 'A staged write switched the run to Publishing.'
    Check (@(Get-Recorded | Where-Object Kind -in @('Upload', 'History') | Where-Object { $_.Path -like "$latest*" -and $_.Path -notlike '*DATA-ALL*' }).Count -eq 0) 'A staged DATA-LAST copy was uploaded or historized before promotion.'
    Check (@(Get-Recorded | Where-Object { $_.Path -like '*.staging*' }).Count -eq 0) 'A staging path reached SharePoint or WeeklyHistory.'

    # Failed run: DATA-LAST, receipt and previous qualification are untouched; staging is removed.
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A failed staged run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A failed staged run replaced the receipt.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Failed' -and -not $run.UnqualifiedPublication -and -not $run.PublicationStarted) 'A failed staged run disqualified the previous publication.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A failed run left its staging folder.'
    Check (@(Get-Recorded | Where-Object { $_.Kind -eq 'Promoted' -or $_.Path -like "$latest\Intune*" }).Count -eq 0) 'A failed run promoted or uploaded DATA-LAST copies.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath)
    $script:checks++

    # Complete run: promotion under Publishing, receipt last, follow-ups after the receipt.
    Clear-Recorded
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Publish-Set 'complete'
    $stagedHashes = @{}
    foreach ($file in $files) { $stagedHashes[$file] = (Get-FileHash -LiteralPath (Join-Path (Get-ChildItem -LiteralPath $stagingProducer -Directory)[0].FullName $file) -Algorithm SHA256).Hash }
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $baselineReceipt.RunId) 'A complete staged run did not publish its receipt.'
    foreach ($file in $files) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -eq $stagedHashes[$file] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final[$file]) "Promoted CSV and receipt disagree: $file"
    }
    $promoted = @(Get-Recorded | Where-Object Kind -eq 'Promoted')
    Check ($promoted.Count -eq 4 -and @($promoted | Where-Object RunStatus -ne 'Publishing').Count -eq 0) 'A CSV was promoted outside the Publishing fence.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication -and $run.StagedPublication) 'Completed staged run state is not qualified.'
    $uploads = @(Get-Recorded | Where-Object { $_.Kind -eq 'Upload' -and $_.Path -like "$latest*" })
    $history = @(Get-Recorded | Where-Object Kind -eq 'History')
    Check ($uploads.Count -eq 4 -and @($uploads | Where-Object { $_.Path -like '*.staging*' }).Count -eq 0) 'Promoted DATA-LAST copies were not uploaded by canonical path.'
    Check ($history.Count -eq 4 -and @($history | Where-Object { $_.Path -like '*.staging*' }).Count -eq 0) 'WeeklyHistory did not follow the promoted canonical copies.'
    $order = @(Get-Recorded | ForEach-Object Kind)
    Check ([array]::LastIndexOf($order, 'Promoted') -lt [array]::IndexOf($order, 'History')) 'WeeklyHistory ran before the promotion.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A completed run left its staging folder.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
    $script:checks++

    # Promotion failure: some CSVs replaced, receipt unchanged, publication disqualified (fail-closed).
    $previousReceipt = $receipt
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Publish-Set 'promotion-failure'
    $lockedFile = (@($files | Sort-Object))[-1]
    $lock = [IO.File]::Open((Join-Path $latest $lockedFile), [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    try { Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null }
    finally { $lock.Dispose() }
    Check ((Read-Json $receiptPath).RunId -eq $previousReceipt.RunId) 'A failed promotion published a receipt.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Failed' -and $run.UnqualifiedPublication -and $run.Error -match 'promotion') 'A failed promotion was not disqualified.'
    $rejected = $false
    try { Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath) } catch { $rejected = $true }
    Check $rejected 'Readers accepted a partially promoted set.'
    $state = Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job ([pscustomobject]@{ ScriptPath = $producer.FullName })
    Check ($state.Known -and $state.Status -eq 'Failed') 'Orchestrator does not see the failed promotion.'

    # A direct DATA-LAST write voids the staging declaration for the rest of the run.
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Write-SmartM365CsvAtomically -Data @([pscustomobject]@{ TenantKey = 'synthetic'; Name = 'direct' }) -Path (Join-Path $latest 'Intune_ManagedApps.csv') -NoTenantKey
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Publishing' -and -not $run.PSObject.Properties['StagedPublication']) 'A direct write kept the staging declaration.'
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null

    # Staging cannot start after a direct publication and is a no-op without an active receipt.
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Write-SmartM365CsvAtomically -Data @([pscustomobject]@{ TenantKey = 'synthetic'; Name = 'direct' }) -Path (Join-Path $latest 'Intune_ManagedApps.csv') -NoTenantKey
    $refused = $false
    try { Enable-SmartM365StagedSourcePublication } catch { $refused = $true }
    Check $refused 'Staging was enabled after a direct DATA-LAST write.'
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    Enable-SmartM365StagedSourcePublication
    $script:checks++
}
finally {
    # Releases the producer lock if a check failed while a receipt was active.
    try { Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks staged source publication checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOj2YTgGTBQ0ZO
# HfusKMEjay/7vy1UF4Z73zIH912Qf6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBvyjF1tFtasLvpJOEzpyG8FjyBDxb6BuIvxPhWssKVdMA0GCSqG
# SIb3DQEBAQUABIIBgKNvCL7tZf9ZVznFpryABC9Z7+8Nq0TlRcTdIeZkyDJ+HU0q
# OyR95cAfhXY/PPf5/lOfy1MvGMuDbJ6JbDLZH4Ix+eQlTIEZJNqif9zjHgs4i/eh
# ZIN/iWyAqRSpbM1MhrAmphehpFp8VJZ/PyOMLGqUx3uiaCkREGb29cenA/ONA88s
# GxrL/0SMWDcKIwSEVZClNfHhPy9JF6VpiM9K6I/1plw/S/der1lZ4V8C4FbrmbMI
# Ph9NtX5RbOnAOu9/mxLG4iRRVTBms09RX1cGmSwWD6n8Rdr6bM+02RK9HXFXPgnQ
# brsZBtpy49/C0hCPbCOAHuO+Fk1J5Fgn44+2WS7jvTfoTj2HllWHSzlsPII34bZV
# oBuuhRvxk+cwqoVU7nHYReHCt9BTDvzub6pzO36jIQFZoKR3eyySqFZd5pjTkWUs
# tBlLib2TNqDVZt6TjYOgqwFyhHBU51FBM4DpdyNI7W0VdTrtrq+cxBrRdyfQaGkI
# cXChVlEqS3xGhNgnG6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAwNjIx
# NDlaMC8GCSqGSIb3DQEJBDEiBCB2ijiUUlVNTc1J/Dksl2we3c7UnEfnk0IhGRUy
# RIhufzANBgkqhkiG9w0BAQEFAASCAgB25Mzsma0oNl3v6IdOsjoW2raC0P/P9C4I
# 6GIxFv79jKzVdfpknvTbfB9D5RnnzJIQrJZLD8xM8gNwS//TYG17Ifu9ED4WmY+8
# 9cjbDbL4oRP+SZLKKAWuZpxyhB8UanAYeNzJq6kq1jiev+hk8q+L7TVD2DL0H+Io
# 5Mwt4iLIT/Tm8WGyRkzABSWeVW9o/+ceE6COVbNnzKXEKU/Cq6QdEb2vLXce83Xz
# 5YdMM6PEhtJxGdnieq5Bav1G1Q0GP1haTIdw4nLblPvOnGL7DzUj/yBIAJMnxX1U
# 6uD2QdD6quc2omv1bRQByW+uRPBHexsR9Z+WIW3DxR5F0mruQEnwxeHBEEWN4Fjd
# PedffLV+MC9dIJCglPAGRyuT5MPS/E9vIHbeYcElO8neEIOcn9ScNgz0HcqJSMJ+
# 8/ZX2YE6EM3vmN2WzT+2hOf1lwhmn78QjM0oHnX78HdLt5tpob+vGKKZCaxjtWtm
# GFkUvycyn3rN2GWfNiCFor1kK/+4nYmHgqeUpNAMxPAvaXFROTNbOV8WiSHWBkDL
# lU9IwzaG/gm88eJXYikXHcL+ygJDQB89lyYCdRKDFM5F8LflF7Fs4pYjaXQGkT/U
# 5xKSoN+6+TLs9i9VM5fE79N4weXWXgzYArop1fpy8GuVZXVL4wZlvLW9NRm3Q58D
# 0iWcxUeqxw==
# SIG # End signature block
