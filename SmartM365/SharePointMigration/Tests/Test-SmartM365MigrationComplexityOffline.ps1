<#
.SYNOPSIS
    Verify the migration complexity request and presentation without SharePoint access.
.VERSION
    1.0.3
#>

#Requires -Version 7.4
[CmdletBinding()]
param([string]$PythonPath)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $projectRoot 'SmartM365-SharePointMigration-Summary.ps1')

$python = if ($PythonPath) { $PythonPath } else { Join-Path $projectRoot 'Tools\Python\python.exe' }
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { $python = (Get-Command python -ErrorAction Stop).Source }

$safeRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$testRoot = [IO.Path]::GetFullPath((Join-Path $safeRoot ('SharePointMigration-complexity-tests-' + [guid]::NewGuid().ToString('N'))))
if (-not $testRoot.StartsWith($safeRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Test directory is outside the temporary folder.' }

function Assert-Complexity {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Complexity test failed: $Message" }
}

try {
    $fileDir = Join-Path $testRoot 'scans\source\files'
    $permissionDir = Join-Path $testRoot 'scans\source\permissions'
    [void](New-Item -ItemType Directory -Path $fileDir, $permissionDir -Force)
    $fileScan = Join-Path $fileDir 'SP2019-FileInventory-Fixture-20261001-120000.csv'
    $permissionScan = Join-Path $permissionDir 'SP2019-PermissionInventory-Fixture-20261001-120000.csv'
    @(
        '"WebUrl";"LibraryUrl";"FileName";"ServerRelativeUrl";"SizeBytes";"VersionsCount";"CheckedOutBy"'
        '"https://source.test";"/Docs";"a.docx";"/Docs/a.docx";"100";"1";""'
    ) | Set-Content -LiteralPath $fileScan -Encoding utf8
    @(
        '"ObjectScope";"ObjectId";"HasUniqueRoleAssignments";"PrincipalType";"PrincipalLoginName"'
        '"Item";"1";"True";"User";"i:0e.t|adfs|user@contoso.test"'
    ) | Set-Content -LiteralPath $permissionScan -Encoding utf8

    $migration = [pscustomobject]@{
        Name = 'Fixture'; Root = $testRoot
        Config = @{ Name = 'Fixture'; Target = @{ SiteUrl = 'https://target.test/sites/fixture' }
            Comparison = @{ EntraUsersCachePath = 'missing-entra-cache.csv' }
            Output = @{ SourceFileScans = 'scans\source\files'; SourcePermissionScans = 'scans\source\permissions' } }
    }
    $inputs = Get-SmartM365ComplexityScanInputs -Migration $migration -SourceType 'SP2019'
    Assert-Complexity ($inputs.FileScan.Name -eq (Split-Path $fileScan -Leaf) -and
        $inputs.PermissionScan.Name -eq (Split-Path $permissionScan -Leaf)) 'latest source scans must be selected'
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $null).State -eq 'Unavailable') 'no source scan must be unavailable'

    $request = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ($null -eq $request.EntraCache) 'a configured but missing Entra cache must not be used'
    $before = Get-SmartM365ComplexityPresentation -Request $request
    Assert-Complexity ($before.State -eq 'Missing' -and $before.Text -eq 'Computing') 'missing result must request a calculation'
    $readOnly = Get-SmartM365ComplexityPresentation -Request $request -ReadOnly
    Assert-Complexity ($readOnly.Caption -eq 'Not evaluated') 'read-only views must not announce a calculation'

    # Same path as the source scan launchers: the stored result must satisfy the GUI without recalculation.
    $scanResult = Invoke-SmartM365MigrationComplexity -Migration $migration -ProjectRoot $projectRoot `
        -SourceType 'SP2019' -PythonExecutable $python
    Assert-Complexity ($scanResult.Status -eq 'Succeeded') "scan-side calculation failed: $($scanResult.Message) $($scanResult.Output -join ' ')"
    $broken = Invoke-SmartM365MigrationComplexity -Migration $migration -ProjectRoot $projectRoot `
        -SourceType 'SP2019' -PythonExecutable $python -ScriptPath (Join-Path $testRoot 'missing.py')
    Assert-Complexity ($broken.Status -eq 'Failed') 'a failed calculation must be reported without throwing'
    $none = Invoke-SmartM365MigrationComplexity -Migration $migration -ProjectRoot $projectRoot `
        -SourceType 'SP2016' -PythonExecutable $python
    Assert-Complexity ($none.Status -eq 'Skipped') 'no source file scan must skip the calculation'
    $request = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    $current = Get-SmartM365ComplexityPresentation -Request $request
    Assert-Complexity ($current.State -eq 'Current' -and $current.Text -eq 'Low') 'calculated result must be current'
    Assert-Complexity ($current.Caption -like '*/100*partial*80% coverage') 'missing evidence must be shown as partial with coverage'
    Assert-Complexity (@($current.Criteria).Count -eq 13) 'all criteria must be listed'
    Assert-Complexity ((@($current.Criteria) | Where-Object Criterion -eq 'Disabled or unresolved accounts').Points -eq 'N/A') 'unevaluated criterion must show N/A'
    Assert-Complexity ((@($current.Criteria) | Where-Object Criterion -eq 'Used custom permission levels').Points -eq 'N/A') 'custom roles without evidence must remain N/A'
    Assert-Complexity ((@($current.Criteria) | Where-Object Criterion -eq 'Workflows, forms and customizations').Points -eq 'N/A') 'functional scope without evidence must remain N/A'
    $assessment = Join-Path $testRoot 'source-complexity-assessment.json.txt'
    Copy-Item -LiteralPath (Join-Path $projectRoot 'Config\source-complexity-assessment.json.template') -Destination $assessment
    $assessedRequest = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ($assessedRequest.Signature -ne $request.Signature) 'assessment addition must change the signature'
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $assessedRequest).State -eq 'Stale') 'assessment addition must invalidate stored result'
    Assert-Complexity ($assessedRequest.Arguments -contains '--source-assessment') 'assessment must reach the calculator'
    $reviewed = Get-Content -LiteralPath $assessment -Raw | ConvertFrom-Json
    $reviewed.Inputs.FileInventorySha256 = (Get-FileHash -LiteralPath $fileScan -Algorithm SHA256).Hash
    $reviewed.Inputs.PermissionInventorySha256 = (Get-FileHash -LiteralPath $permissionScan -Algorithm SHA256).Hash
    foreach ($section in @($reviewed.FunctionalComponents, $reviewed.CustomPermissions)) {
        $section.Complete = $true; $section.ReviewedAtUtc = '2026-01-01T00:00:00Z'; $section.Evidence = 'DEMO - complete synthetic source assessment'
    }
    $reviewed | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $assessment -Encoding utf8
    $assessmentRun = Invoke-SmartM365MigrationComplexity -Migration $migration -ProjectRoot $projectRoot -SourceType 'SP2019' -PythonExecutable $python
    Assert-Complexity ($assessmentRun.Status -eq 'Succeeded') 'complete synthetic source assessment must calculate'
    $reviewRequest = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    $reviewPresentation = Get-SmartM365ComplexityPresentation -Request $reviewRequest
    Assert-Complexity ($reviewPresentation.State -eq 'Current' -and $reviewPresentation.Caption -like '*95% coverage') 'complete source assessment must increase coverage to 95%, leaving Entra N/A'
    $reviewed.FunctionalComponents.Evidence = 'DEMO - modified assessment reference'
    $reviewed | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $assessment -Encoding utf8
    $changedAssessment = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $changedAssessment).State -eq 'Stale') 'edited assessment must invalidate cached result'
    $autoPath = $fileScan + '.assessment.json.txt'
    $detailPath = Join-Path $fileDir ('AssessmentDetails-'+[IO.Path]::GetFileNameWithoutExtension($fileScan)+'.csv')
    [IO.File]::WriteAllText($detailPath,'DEMO automatic details')
    $auto = @{
        SchemaVersion=1;TargetPlatform='SPO';Inputs=@{FileInventorySha256=(Get-FileHash $fileScan -Algorithm SHA256).Hash}
        DetailsFile=[IO.Path]::GetFileName($detailPath);DetailsSha256=(Get-FileHash $detailPath -Algorithm SHA256).Hash
        FunctionalComponents=@{Mode='Automatic';Complete=$false;CollectionComplete=$true;QualificationComplete=$false
            Items=@(@{Id='wf';Category='Workflow';Technology='Workflow2013';Effort=3;Classification='Estimated';Evidence='DEMO known risk'},@{Id='form';Category='Form';Effort=$null;Classification='Unqualified';Evidence='DEMO unknown form'});Errors=@();Detectors=@()}
    }
    $auto | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $autoPath -Encoding utf8
    $autoRun = Invoke-SmartM365MigrationComplexity -Migration $migration -ProjectRoot $projectRoot -SourceType 'SP2019' -PythonExecutable $python
    Assert-Complexity ($autoRun.Status -eq 'Succeeded') ('Automatic evidence failed: '+($autoRun.Output -join ' '))
    $autoRequest = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    $autoPresentation = Get-SmartM365ComplexityPresentation -Request $autoRequest
    Assert-Complexity ($autoPresentation.State -eq 'Current') ('automatic evidence must satisfy cache matching: '+($autoRequest | Select-Object ConfigHash,AssessmentHash,AutomaticAssessmentHashes | ConvertTo-Json -Depth 4 -Compress)+'; result='+($autoPresentation.Result.Inputs|ConvertTo-Json -Depth 4 -Compress))
    Assert-Complexity ((@($autoPresentation.Criteria)|Where-Object Criterion -eq 'Workflows, forms and customizations').Points -eq 'N/A') 'partial automatic evidence must override complete manual fallback'
    Assert-Complexity ($autoPresentation.Tooltip -like '*Workflow2013*effort 3*') 'known risk must remain visible despite N/A score'
    [IO.File]::WriteAllText($detailPath,'DEMO corrupted automatic details')
    $tamperedRequest = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ($tamperedRequest.Signature -ne $autoRequest.Signature) 'detail CSV edits must invalidate cache'
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $tamperedRequest).State -eq 'Stale') 'changed detail evidence must be marked stale'

    Add-Content -LiteralPath $fileScan -Value '"https://source.test";"/Docs";"b.docx";"/Docs/b.docx";"100";"0";""' -Encoding utf8
    $newer = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ($newer.Signature -ne $request.Signature) 'a changed scan must change the signature'
    $stale = Get-SmartM365ComplexityPresentation -Request $newer
    Assert-Complexity ($stale.State -eq 'Stale' -and $stale.Caption -eq 'Updating' -and $stale.Text -eq 'Low') 'stale result must keep the previous level'
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $newer -ReadOnly).Caption -eq 'Outdated') 'read-only stale result must be outdated'
    # A rewritten mapping with the same name and length must still make the result stale.
    $mapping = Join-Path $testRoot 'migration.mapping.txt'
    Set-Content -LiteralPath $mapping -Value '/ /sites/fixture' -Encoding ascii
    $migration.Config.Comparison.PathMappingsFile = 'migration.mapping.txt'
    $mapped = Invoke-SmartM365MigrationComplexity -Migration $migration -ProjectRoot $projectRoot -SourceType 'SP2019' -PythonExecutable $python
    Assert-Complexity ($mapped.Status -eq 'Succeeded') "mapped calculation failed: $($mapped.Message)"
    $mappedRequest = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $mappedRequest).State -eq 'Current') 'mapped result must be current'
    (Get-Item $mapping).LastWriteTimeUtc = (Get-Item $mapping).LastWriteTimeUtc.AddMinutes(5)
    $touched = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ($touched.Signature -ne $mappedRequest.Signature) 'a rewritten mapping must change the signature'
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $touched).State -eq 'Stale') 'a rewritten mapping must make the result stale'

    # Pending markers tell the GUI that a scan will calculate complexity itself.
    $marker = Set-SmartM365ComplexityPending -MigrationRoot $testRoot -Reason 'Test'
    Assert-Complexity ($null -ne (Get-SmartM365ComplexityPending -MigrationRoot $testRoot)) 'a fresh marker must be pending'
    (Get-Item $marker).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddHours(-30)
    Assert-Complexity ($null -eq (Get-SmartM365ComplexityPending -MigrationRoot $testRoot)) 'an old marker must be ignored'
    Clear-SmartM365ComplexityPending -Path $marker
    Assert-Complexity (-not (Test-Path -LiteralPath $marker)) 'the marker must be removed'
    Assert-Complexity ([bool](Get-SmartM365PythonCommand -ProjectRoot $projectRoot).Executable) 'Python must be resolved'

    $failed = Get-SmartM365ComplexityPresentation -Request $newer -ErrorMessage 'boom'
    Assert-Complexity ($failed.State -eq 'Error' -and $failed.Tooltip -like '*boom*') 'failure must be visible'

    Write-Host ('{0} Migration complexity offline checks passed.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
}
finally {
    if ((Split-Path $testRoot -Leaf) -like 'SharePointMigration-complexity-tests-*' -and (Test-Path -LiteralPath $testRoot)) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCQrns29njoIdw8
# bC/jx86IrVlmHlJhuU3xKyxbT/nStKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAwoUQ2eubsuQiehUHRJ0op
# aPHutnIhQEi1myLesgmGJTANBgkqhkiG9w0BAQEFAASCAYABTCslJticPm6/875G
# OAZ/1GPTdafLGz6FWlCb0wbrgZWop3XJ56jPgwLMDAjTaKpdonD2OquH3+/z1GQs
# +bA2+vHqQ5SMPVWYBVyST06NCA4MqeJCfv/JKUcn9wOifjuTiRjKIHSBfCZSnAsk
# P61gc4tmWLh/XIfAW+eSQIOuUeSqic3E4F0keiAaopr55Ia2wMMyR1zfHXCu1aqt
# 8aAKEuMAS0hiA5LUfvrpdBgYnSdgChcNcWWUDxqOiq4doM4VTZwOncIYCaSiebbR
# fNf5oV9A9ZEOdbZTk+yN4pHkPOGpgPqS5DULesZUh7BEkY4pBOAYMbrpYaLmpGbF
# HnK9TiP9fOdUMJcvpDSUOLp/4V3C74e+cplCuzYxzM9GHWcswVIW4RivprSJJFb1
# LtqcC/nebmmypaU2RIauxdUqtq0ypGR8kFCi3wD3OoEyWb9+QYbMTtV4l+yddYfU
# 3IorJ3WnGShgUr6frO/EP1umZeDfrsNoSlJ7czuYyR1CKEs=
# SIG # End signature block
