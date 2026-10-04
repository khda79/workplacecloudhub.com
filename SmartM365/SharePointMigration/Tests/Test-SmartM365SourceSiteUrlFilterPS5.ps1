<#
.SYNOPSIS
    Verify source scan URL filters under Windows PowerShell 5.1.
.VERSION
    1.0.0
#>

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
if ($PSVersionTable.PSVersion.Major -ne 5) { throw 'Run this test with Windows PowerShell 5.1.' }
$testRoot = Join-Path $env:TEMP ('SmartM365-SourceUrlFilter-' + [guid]::NewGuid().ToString('N'))
$resolvedRoot = [IO.Path]::GetFullPath($testRoot)
if (-not $resolvedRoot.StartsWith([IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Unsafe test path.'
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -ErrorAction Stop)
    $urlsFile = Join-Path $testRoot 'source-urls.txt'
    foreach ($scriptName in @('SmartM365-SharePointSource-FileInventory.ps1', 'SmartM365-SharePointSource-PermissionInventory.ps1')) {
        $path = Join-Path $PSScriptRoot ("..\Scripts\Inventory\{0}" -f $scriptName)
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        if (@($errors).Count -ne 0) { throw "Parser errors in $scriptName" }
        foreach ($name in @('Get-NormalizedUrlPath', 'Import-SiteUrlFilter', 'Test-PathMatchesSiteUrlFilter')) {
            $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
            if (-not $definition) { throw "Missing function $name in $scriptName" }
            Invoke-Expression $definition.Extent.Text
        }

        [IO.File]::WriteAllLines($urlsFile, [string[]]@('https://source.example.test/ProjectA/'))
        $script:SiteUrlFilterPaths = Import-SiteUrlFilter -Path $urlsFile
        if ($script:SiteUrlFilterPaths.Count -ne 1 -or $script:SiteUrlFilterPaths -isnot [Collections.Generic.HashSet[string]]) {
            throw "Single URL filter lost its collection type in $scriptName"
        }
        if (-not (Test-PathMatchesSiteUrlFilter -WebPath '/projecta') -or
            -not (Test-PathMatchesSiteUrlFilter -WebPath '/PROJECTA/child') -or
            (Test-PathMatchesSiteUrlFilter -WebPath '/projecta-other')) {
            throw "Single URL filter scope failed in $scriptName"
        }

        [IO.File]::WriteAllLines($urlsFile, [string[]]@('# roots', 'https://source.example.test/Alpha', 'https://source.example.test/alpha/', '', 'https://source.example.test/Beta'))
        $script:SiteUrlFilterPaths = Import-SiteUrlFilter -Path $urlsFile
        if ($script:SiteUrlFilterPaths.Count -ne 2 -or -not $script:SiteUrlFilterPaths.Contains('/alpha') -or
            -not (Test-PathMatchesSiteUrlFilter -WebPath '/beta/child') -or
            (Test-PathMatchesSiteUrlFilter -WebPath '/outside')) {
            throw "Multiple URL filter scope or deduplication failed in $scriptName"
        }

        [IO.File]::WriteAllLines($urlsFile, [string[]]@('', '# no URLs'))
        $emptyRejected = $false
        try { [void](Import-SiteUrlFilter -Path $urlsFile) }
        catch { $emptyRejected = $_.Exception.Message -like '*contains no usable site URLs*' }
        if (-not $emptyRejected) { throw "Empty filter was not rejected in $scriptName" }

        $missingRejected = $false
        try { [void](Import-SiteUrlFilter -Path (Join-Path $testRoot 'missing.txt')) }
        catch { $missingRejected = $_.Exception.Message -like '*filter file not found*' }
        if (-not $missingRejected) { throw "Missing filter was not rejected in $scriptName" }
    }
    'Windows PowerShell 5.1 source URL filter tests passed for files and permissions.'
}
finally {
    if (Test-Path -LiteralPath $resolvedRoot -PathType Container) { Remove-Item -LiteralPath $resolvedRoot -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAWuM6g3t9H3lVf
# xTsuAz1ev3vepEFqv2WZ2gotA2BD66CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAWfSDPfzP0w/qAf7e7rV/f
# oMeurn53mzEodB15MN7k+zANBgkqhkiG9w0BAQEFAASCAYApQHOgFRF74UvvPaWS
# yzdCV/Nz/ihrT1rAXLaAIULJlD0KBmwagT/asm+ArwAR1rRc3VwRlGZKvoPS/X+e
# qfRQjJqkVrdd7LQOdUCiikCVn/OybVv8MVE0+OwBvLMSVJMe/dUG9VzIwBIXJE51
# YOEbKJ17TCZ8xC+EQyi5qfIZv1WhYqm9VgKQv7fC+BIA8+/KKIcx1N3t8luuJ9iH
# tkSkpb5cgkZcnXx0NGeLwhXDWNzLu5RLEOgnWKUmLORX0EKqahAsWF+/2I5EldAM
# ASevBaLlNudiiF9M31S2jvdXCDJt8SrV8cXpHfpSwWs8cTWjZku+JFUsRBaI3pPD
# BSGWR5TNnbW9qIh1xJIXkFWnlIbbKJv1ojRkmj0e3/Hx7Q8pK/89KElUFI6HdRUQ
# UIcsWNqVpKfJNDryHqzD7MA4h7P5zyMRld1nOFYevgLmCsLOXgXt2bNcsvGetYdh
# 6sbFP2FM0dGSfcvzfqhNHaahl3vruZ1mpMU9jkLwQnOW0Us=
# SIG # End signature block
