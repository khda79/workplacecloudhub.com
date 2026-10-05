<#
.SYNOPSIS
    Check farm-side file inventory metrics under Windows PowerShell 5.1.
.VERSION
    1.0.1
#>

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -ne 5) { throw 'Run this test with Windows PowerShell 5.1.' }
$path = Join-Path $PSScriptRoot '..\Scripts\Inventory\SmartM365-SharePointSource-FileInventory.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors | ForEach-Object Message) }
foreach ($name in @('Add-FileInventoryMetric','Write-FileInventoryMetrics')) {
    $definition = $ast.Find({
        param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    },$true)
    if (-not $definition) { throw "Missing function: $name" }
    $body = $definition.Body.Extent.Text
    Set-Item -Path ("function:global:{0}" -f $name) -Value ([scriptblock]::Create($body.Substring(1,$body.Length - 2)))
}
$script:MetricFileSizes = @{}
$script:MetricFolders = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$script:MetricRows = [long]0
$script:MetricKnownSizeBytes = [long]0
$script:MetricMissingPathRows = [long]0
$script:MetricInvalidLibraryRows = [long]0
$script:MetricDuplicateRows = [long]0
$script:MetricConflictingSizeRows = [long]0
$root = Join-Path $PSScriptRoot ('.file-inventory-ps5-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
    throw 'Unsafe test path.'
}
try {
    [void](New-Item -ItemType Directory -Path $root)
    $csv = Join-Path $root 'SP2019-FileInventory-Fixture.csv'
    'ServerRelativeUrl;LibraryUrl;SizeBytes' | Set-Content -LiteralPath $csv -Encoding utf8
    Add-FileInventoryMetric -Row ([pscustomobject]@{ ServerRelativeUrl='/sites/a/Docs/one.txt'; LibraryUrl='/sites/a/Docs'; SizeBytes=1024 })
    Add-FileInventoryMetric -Row ([pscustomobject]@{ ServerRelativeUrl='/sites/a/Docs/Nested/two.txt'; LibraryUrl='/sites/a/Docs'; SizeBytes=2048 })
    Write-FileInventoryMetrics -CsvPath $csv
    $result = Get-Content -LiteralPath "$csv.metrics.json.txt" -Raw | ConvertFrom-Json
    if ($result.Files -ne 2 -or $result.FoldersWithFiles -ne 1 -or $result.KnownSizeBytes -ne 3072 -or
        $result.MissingSizeFiles -ne 0 -or $result.CsvSha256 -ne (Get-FileHash -LiteralPath $csv).Hash) { throw 'Incorrect farm-side inventory metrics.' }
    'Windows PowerShell 5.1 file inventory metrics test passed.'
}
finally {
    if (Test-Path -LiteralPath $root -PathType Container) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcXSp1TcnytEaq
# qs7aArIBzK9vmOis+Tc6mMFV4ASMc6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBHA2yeeoE7z5N8SGBQxVJ8
# 8A5cY8ANVFuz8Ai7kTxCVzANBgkqhkiG9w0BAQEFAASCAYBh1UFrpiuvW45T7u9Y
# k8uAO80vxRs0/dNBkyA3tYIDmONPDlKIrUgAeT/rY+hDruYGAGwXo9RkNaJUKoXX
# +oAN1CEXzOf+1p3uN8dU863QA6EmcL+RE8mansyGaF4efhoneKEkfDZGCzDloUKJ
# sEEwYiZB3+Hp8YV0RI8DO7PG19GIuTg15TRElteNMDUJ0vt1YEsYWRxPLLPC3Pzg
# bEc9EF+wbsVD03ERxEgVE0m9JY1KFbkYpiRKw1YXVL48/FiTFJDIfvB9HJC2tyq1
# hR0JpTKj+h0fKkrvpDaF73/kmUi2e8CCZENF1p+7VfgTcgIRoeFHYgxAIH42R0Zm
# m90XBll8qaq2ygBx4UtIrgk6RccXN++DeC3eFJjQxg+bG3GJMrsjpm5Zyw3G2txe
# 2txgfL0ZeH7ieI7RZ8ihBmRG24Q5vkDNgzTTYXErOyoC4i486v/EIgF+5Dyy4ujo
# S6KitQ9FfFekSnvhRgjHy/vamMJQC7AbhmjuQyjzdgDzNGw=
# SIG # End signature block
