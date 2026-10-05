<#
.SYNOPSIS
    Verify automatic ImportExcel preparation without downloading modules.
.VERSION
    1.0.0
#>
#requires -Version 7.4
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-ImportExcel.ps1')

function Assert-Excel { param([bool]$Condition,[string]$Message) if(-not $Condition){throw $Message} }
function Get-Module {
    param([switch]$ListAvailable,[string]$Name)
    if($Name -ne 'ImportExcel'){throw 'Unexpected module lookup.'}
    if($script:Present){[pscustomobject]@{Name='ImportExcel';Version=[version]'7.8.10';Path='synthetic-ImportExcel.psd1'}}
}
function Get-Command {
    [CmdletBinding()]param([string]$Name,[string]$Module)
    if($Name -eq 'Install-PSResource') {if($script:Resource){[pscustomobject]@{Parameters=@{Quiet=$true;AcceptLicense=$true}}};return}
    if($Name -eq 'Install-Module'){[pscustomobject]@{Parameters=@{AcceptLicense=$true}};return}
    if($Name -in @('Import-Excel','Get-ExcelSheetInfo') -and $script:CommandsReady){[pscustomobject]@{Name=$Name;ModuleName='ImportExcel'}}
}
function Get-PSResourceRepository {
    [CmdletBinding()]param([string]$Name)
    if($script:RepositoryPresent){[pscustomobject]@{Name=$Name;Uri=[uri]$script:RepositoryUri}}
}
function Get-PSRepository {
    [CmdletBinding()]param([string]$Name)
    if($script:RepositoryPresent){[pscustomobject]@{Name=$Name;SourceLocation=$script:RepositoryUri}}
}
function Register-PSResourceRepository {
    [CmdletBinding()]param([switch]$PSGallery)
    Assert-Excel $PSGallery 'Repository registration did not use the official default.'
    $script:Registered++;$script:RepositoryPresent=$true
}
function Register-PSRepository {
    [CmdletBinding()]param([switch]$Default)
    Assert-Excel $Default 'Legacy repository registration did not use the default.'
    $script:Registered++;$script:RepositoryPresent=$true
}
function Install-PSResource {
    [CmdletBinding()]param([string]$Name,[string]$Scope,[string]$Repository,[switch]$TrustRepository,[switch]$AcceptLicense,[switch]$Quiet)
    Assert-Excel ($Name -eq 'ImportExcel' -and $Scope -eq 'CurrentUser' -and $Repository -eq 'PSGallery' -and $TrustRepository -and $AcceptLicense -and $Quiet) 'Incorrect PSResource installation options.'
    $script:Installs++
    if($script:NetworkFailure){throw 'Synthetic network failure.'}
    if(-not $script:MissingAfterInstall){$script:Present=$true}
}
function Install-Module {
    [CmdletBinding()]param([string]$Name,[string]$Scope,[string]$Repository,[switch]$Force,[switch]$AcceptLicense)
    Assert-Excel ($Name -eq 'ImportExcel' -and $Scope -eq 'CurrentUser' -and $Repository -eq 'PSGallery' -and $Force -and $AcceptLicense -and -not $PSBoundParameters.Confirm) 'Incorrect legacy installation options.'
    $script:Installs++
    if($script:NetworkFailure){throw 'Synthetic network failure.'}
    $script:Present=$true
}
function Import-Module {
    [CmdletBinding()]param([string]$Name,[switch]$Global)
    Assert-Excel ($Name -eq 'synthetic-ImportExcel.psd1' -and $Global) 'Wrong module imported.'
    $script:Imports++
    if($script:ImportFailure){throw 'Synthetic import policy failure.'}
}
function Set-PSRepository {throw 'Repository trust must not be changed permanently.'}
function Set-ExecutionPolicy {throw 'Execution policy must not be changed.'}
function Reset-ExcelScenario {
    $script:Present=$false;$script:Resource=$true;$script:RepositoryPresent=$true
    $script:RepositoryUri='https://www.powershellgallery.com/api/v2'
    $script:CommandsReady=$true;$script:NetworkFailure=$false;$script:MissingAfterInstall=$false;$script:ImportFailure=$false
    $script:Installs=0;$script:Imports=0;$script:Registered=0;$script:Phases=[Collections.Generic.List[string]]::new()
}
$progress={param($State,$Message) $script:Phases.Add($State)}
Reset-ExcelScenario
$module=Initialize-SmartM365ImportExcel -DryRun -Progress $progress
Assert-Excel (-not $module -and $script:Installs -eq 0 -and $script:Imports -eq 0 -and $script:Registered -eq 0 -and $script:Phases.Count -eq 0) 'DryRun installed, imported or registered a dependency.'

Reset-ExcelScenario;$script:Present=$true
$module=Initialize-SmartM365ImportExcel -Progress $progress
Assert-Excel ($module.Version -eq [version]'7.8.10' -and $script:Installs -eq 0 -and $script:Imports -eq 1) 'Existing module was reinstalled or not imported.'

foreach($resource in @($true,$false)){
    Reset-ExcelScenario;$script:Resource=$resource;$script:RepositoryPresent=$false
    $module=Initialize-SmartM365ImportExcel -Progress $progress
    Assert-Excel ($module -and $script:Installs -eq 1 -and $script:Registered -eq 1 -and $script:Imports -eq 1) 'Missing module/repository preparation failed.'
    Assert-Excel (($script:Phases -join ',') -eq 'InstallingImportExcel,ImportExcelReady') 'Preparation phases missing or reordered.'
    $null=Initialize-SmartM365ImportExcel -Progress $progress
    Assert-Excel ($script:Installs -eq 1) 'Retry reinstalled an already available module.'
}
foreach($scenario in @('Network','MissingAfterInstall','Import','Commands','Repository')){
    Reset-ExcelScenario
    switch($scenario){
        'Network'{$script:NetworkFailure=$true}
        'MissingAfterInstall'{$script:MissingAfterInstall=$true}
        'Import'{$script:Present=$true;$script:ImportFailure=$true}
        'Commands'{$script:Present=$true;$script:CommandsReady=$false}
        'Repository'{$script:RepositoryUri='https://example.invalid/gallery'}
    }
    $failed=$false
    try{$null=Initialize-SmartM365ImportExcel -Progress $progress}catch{$failed=($_.Exception.Message -match 'ImportExcel preparation failed for the current user' -and $_.Exception.Message -match 'retry Analyze latest report')}
    Assert-Excel $failed "Failure was hidden for $scenario."
    Assert-Excel (-not ($script:Phases -contains 'ImportExcelReady')) "Failed preparation was reported ready for $scenario."
    if($scenario -eq 'Repository'){Assert-Excel ($script:Installs -eq 0) 'Unexpected repository was used.'}
}
Write-Output 'ImportExcel offline tests passed: reuse, both installers, repository setup, DryRun, progress and failure propagation.'

# Remove only the mocks defined above. Integration reuses a locally installed
# ImportExcel; it never downloads a real dependency or reads customer workbooks.
foreach($name in @('Get-Module','Get-Command','Get-PSResourceRepository','Get-PSRepository','Register-PSResourceRepository','Register-PSRepository','Install-PSResource','Install-Module','Import-Module','Set-PSRepository','Set-ExecutionPolicy')) {
    Remove-Item -LiteralPath "Function:$name"
}
$localExcel=Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
if(-not $localExcel){Write-Output 'Synthetic workbook integration skipped: ImportExcel is not installed locally.';return}
Import-Module -Name $localExcel.Path -ErrorAction Stop
$temporary=Join-Path $env:TEMP ('SmartM365-ExcelAnalysisTest-'+[guid]::NewGuid().ToString('N'))
$project=Join-Path $temporary 'Synthetic'
$reports=Join-Path $project 'ShareGate\MigrationReport'
[void](New-Item -ItemType Directory -Path $reports -Force)
try {
    $workbook=Join-Path $reports 'Synthetic.xlsx'
    $rows=@(
        [pscustomobject]@{'Session ID'='260105-1';ID='1';Date='2026-01-05T10:00:00Z';Status='Success';Type='File';Title='a.txt';'Source site address'='https://source.example/site';'Source list title'='Documents';'Source ID'='1';Messages='Copied.'},
        [pscustomobject]@{'Session ID'='260105-1';ID='2';Date='2026-01-05T10:01:00Z';Status='Warning';Type='File';Title='b.txt';'Source site address'='https://source.example/site';'Source list title'='Documents';'Source ID'='2';Messages='Synthetic warning.'}
    )
    $rows | Export-Excel -Path $workbook -WorksheetName Data
    $wrapper=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-Diagnostics.ps1'))
    $output=Join-Path $project 'ShareGate\Diagnostics\Integration'
    & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $wrapper -ProjectRoot $project -InputPath $workbook -OutputDirectory $output | Out-Null
    Assert-Excel ($LASTEXITCODE -eq 0) 'Synthetic workbook analysis failed.'
    $summary=Get-Content (Join-Path $output 'Summary.json.txt') -Raw | ConvertFrom-Json
    Assert-Excel ($summary.Lines -eq 2 -and $summary.LineStatus.Success -eq 1 -and $summary.LineStatus.Warning -eq 1 -and (Test-Path -LiteralPath $summary.ReportPath)) 'Synthetic workbook rows or HTML report missing.'
    $phase=Get-Content (Join-Path $output 'analysis.phase.json.txt') -Raw | ConvertFrom-Json
    Assert-Excel ($phase.State -eq 'Completed') 'Completion phase missing.'
    Assert-Excel (-not @(Get-ChildItem -LiteralPath $output -Filter '.converted-*.csv').Count) 'Temporary converted CSV was retained.'

    # A missing module in DryRun must not trigger either real installer.
    $driver=Join-Path $temporary 'DryRunMissingExcel.ps1'
    @'
param($Wrapper,$Project,$Workbook,$Output)
function Get-Module {param($Name,[switch]$ListAvailable) if($Name -eq 'ImportExcel'){return}; Microsoft.PowerShell.Core\Get-Module @PSBoundParameters}
function Install-Module {throw 'DryRun must not install a module.'}
function Install-PSResource {throw 'DryRun must not install a resource.'}
& $Wrapper -ProjectRoot $Project -InputPath $Workbook -OutputDirectory $Output -DryRun
'@ | Set-Content -LiteralPath $driver -Encoding utf8
    $dryOutput=Join-Path $temporary 'DryRunOutput'
    $plan=@(& (Join-Path $PSHOME 'pwsh.exe') -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $driver -Wrapper $wrapper -Project $project -Workbook $workbook -Output $dryOutput)
    Assert-Excel ($LASTEXITCODE -eq 0 -and ($plan -join "`n") -match 'Requires ImportExcel' -and -not (Test-Path -LiteralPath $dryOutput)) 'XLSX-only DryRun installed a dependency, failed or created analysis output.'
    Write-Output 'Synthetic XLSX integration passed: two rows, HTML, completion phase, conversion cleanup and missing-module DryRun.'
}finally{
    if([IO.Path]::GetFullPath($temporary).StartsWith([IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) {Remove-Item -LiteralPath $temporary -Recurse -Force}
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCGJvbzQ5pGAC9t
# x4g2LT0FqeadCW52oxcKKNYXgE9idKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAH6R8rFX4pTn8BrWPkExb5
# PwE8fvCiV5/feiJfL4chLTANBgkqhkiG9w0BAQEFAASCAYBzSNbbpC8hFgIP5uvV
# E9Zy7uwxWm/90g2ybXVGBGscfFkk3WWnOunhrigTl5rSGC87iCO9JVOluWJw5t6H
# N1iqYyAIQcigwdMJ1SQzJZaavuDnaL1cl9y3qmoRGoKpOpqU/74IEOESAFT+Mrru
# HQpAqXQJ6vtk7ALtdl81rdYaoYQslr5uzNIylUVrEnvGobqoae2PhrkpkzO4SNsi
# CeiAYIvl1bJDibw0NJ7HyHhPa354ajNkNosru9UGH+Y3rsZ5+MwfdavI6JJJ12ei
# w/RzgVHM2xCbSFcO4dpe9uHawnZT5ivPct5L8l/UkZM5jaDTLV8nBjQZWIuqXK+M
# 4eThJ7Mgkn59+WmRgDXj7HnGKUonP2AVzItWMGkkNc+4hhchGmYDUxANOEGOLiLQ
# X9rOO9QrMjBbNu6J9KMzjEONoNKlsk+hbOhpbx1qREF4Gj0qWkCgZEkLmTFRoPIJ
# A7MocS19LKvFhcg8Bs4JwjJtk+mwpetfcYM7TiFCIjQjwXM=
# SIG # End signature block
