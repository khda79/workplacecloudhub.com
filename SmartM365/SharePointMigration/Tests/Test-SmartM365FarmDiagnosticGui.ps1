<#
.SYNOPSIS
    Offline Migration Diagnostics farm-result and command-generation test.
.VERSION
    1.0.2
#>
#Requires -Version 7.4
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$guiPath = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors | ForEach-Object Message) }
$function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Refresh-FarmDiagnostics' },$true)
if (-not $function) { throw 'Farm refresh function is missing from the GUI.' }
. ([scriptblock]::Create($function.Extent.Text))
$preflight = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-FarmDiagnosticsPrerequisites' },$true)
if (-not $preflight) { throw 'Farm prerequisite check is missing from the GUI.' }
. ([scriptblock]::Create($preflight.Extent.Text))
function Refresh-DiagnosticReportState { param([switch]$Force) }
$root = Join-Path $PSScriptRoot ('.farm-gui-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }
try {
    $project = Join-Path $root 'Synthetic'
    $diagnostics = Join-Path $project 'ShareGate\Diagnostics'
    $farm = Join-Path $diagnostics 'Farm-20261003-000000'
    $analysis = Join-Path $diagnostics 'Analysis-20261002'
    [void](New-Item -ItemType Directory -Path $farm,$analysis -Force)
    $report = Join-Path $farm 'Farm-Report.html'
    '<html></html>' | Out-File -LiteralPath $report -Encoding utf8
    [pscustomobject]@{ SchemaVersion=1; Project='Synthetic'; Status='Partial'; GeneratedAtUtc='2026-10-03T00:00:00Z'; Coverage=@([pscustomobject]@{ Server='WFE02'; Source='IIS logs'; Status='Missing'; Detail='Share unavailable' }); ReportPath=$report } | ConvertTo-Json -Depth 6 | Out-File -LiteralPath (Join-Path $farm 'Farm-Summary.json.txt') -Encoding utf8
    @'
WindowUtc,Lines,Items,Source,Destination,Undetermined
2026-10-02 22:00 UTC,9,4,9,0,0
2026-10-02 22:05 UTC,4,2,4,0,0
'@ | Out-File -LiteralPath (Join-Path $analysis 'AccessFailures-5min.csv') -Encoding utf8
    $script:CurrentMigration = [pscustomobject]@{ Name='Synthetic'; Root=$project }
    $script:DiagLoadedDirectory = $analysis
    $script:DiagAnalysisVerified = $true
    $script:FarmReportPath = ''
    $script:FarmToolkitRoot = '\\fileserver.example\toolkit\SharePointMigration'
    $lblFarmResult = [pscustomobject]@{ Text='' }
    $lblFarmPeaks = [pscustomobject]@{ Text='' }
    $btnFarmOpenReport = [pscustomobject]@{ IsEnabled=$false }
    $btnFarmCheck = [pscustomobject]@{ IsEnabled=$false }
    $btnFarmRun = [pscustomobject]@{ IsEnabled=$false }
    $lblFarmPrerequisites = [pscustomobject]@{ Text='' }
    $txtFarmDryRun = [pscustomobject]@{ Text='' }
    $txtFarmRun = [pscustomobject]@{ Text='' }
    Set-StrictMode -Version Latest
    Refresh-FarmDiagnostics
    if ($lblFarmResult.Text -notmatch 'Partial; 1 missing') { throw 'Farm coverage was not displayed.' }
    if (-not $btnFarmOpenReport.IsEnabled) { throw 'Farm HTML report was not detected.' }
    if (-not $btnFarmCheck.IsEnabled) { throw 'Farm prerequisite check was not offered after current analysis.' }
    if ($btnFarmRun.IsEnabled -or $lblFarmPrerequisites.Text -notmatch 'Check prerequisites') { throw 'Farm Run was enabled before prerequisite checks.' }
    if ($lblFarmPeaks.Text -notmatch '2 in') { throw 'ShareGate peak count was not displayed.' }
    foreach($command in @($txtFarmDryRun.Text,$txtFarmRun.Text)) {
        if ($command -notmatch 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File') { throw 'Windows PowerShell command flags are missing.' }
        if ($command -notmatch '-Project "Synthetic"') { throw 'Project was not included.' }
        if ($command -notmatch '-Around "2026-10-02T22:00:00Z"') { throw 'Largest UTC peak was not included.' }
        if (-not $command.Contains('-ShareGatePeaksCsv "' + $script:FarmToolkitRoot)) { throw 'UNC peak CSV was not included.' }
        if (-not $command.Contains('-ToolkitRoot "' + $script:FarmToolkitRoot + '"')) { throw 'UNC toolkit root was not included.' }
        if ($command -match "`n") { throw 'A command spans multiple lines.' }
    }
    if ($txtFarmDryRun.Text -notmatch ' -DryRun$' -or $txtFarmRun.Text -match ' -DryRun$') { throw 'DryRun and real command separation failed.' }
    $script:FarmInvocation = $null
    if (Test-FarmDiagnosticsPrerequisites) { throw 'Farm Run passed without an analysis.' }
    if ($btnFarmRun.IsEnabled -or $lblFarmPrerequisites.Text -notmatch 'Run unavailable') { throw 'Farm Run did not remain disabled.' }
    Write-Output 'Farm diagnostics GUI offline test passed.'
}
finally {
    if ((Test-Path -LiteralPath $root) -and $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAYE/q6m/1x5r7h
# mnjVnxQz0epQWvjkU+qnxZErkemUl6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAB5EhM1Jtb0ZcR23bi73k8
# 7pu37yMC8LqIWY0D+as/bjANBgkqhkiG9w0BAQEFAASCAYAzO/XSH64LaLaT3Bmv
# UlU8urMv7I1vCeRLfqRswe8tUWfm1a2KU2XH5OhLIXODKXY4LJylwrYIHWHYlSfm
# vPrRf5D0zLMMlFDe98yAn27obkRj/6H64ajR7j9eVzpI7Xce6IGyCvBaWV3Vh5CR
# 0RRHZBlPheZmiCof7odc8CZjYAQXr/yjNNGDB/fTdBg80AF7bnSHQp7brYTg5lKB
# 3o/ueb5X8TiZpXnpo+44rVEQyOJo9Rdg8aCdZELSyztIkKZ2FpCNtp7QOXj+HFCz
# yNIyMpEzfCqYdqXPs2ULM4s5Ppj0SFmB83tYcj1Rq2majvVHYptKmnImQNi4GT+G
# okaJbqxnMWAE+r1rh8rQd2ivJxXbRhkTkstW/Nat7xvCYAZ1gDRHSsMZ7QocgBft
# khOpEGHsPzPZTPksqAeM76u4JLNtRjs2Ff4LOj/nVmzxKT+63QaRykOpjKQ5vuNa
# p4kL6LaEAMNsk2qHfWMvix3jVvW3sD/h/wZcWETFuGKof/4=
# SIG # End signature block
