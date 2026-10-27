<#
.SYNOPSIS
    Verifies the GUI archive routing using synthetic migration folders.
.VERSION
    1.0.2
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$guiPath = Join-Path $project 'SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
foreach ($name in @('Get-MigrationFolders', 'Move-MigrationToArchive',
    'Get-ArchivedOverviewRows', 'Get-MigrationEndpointType')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if (-not $definition) { throw "Missing GUI function: $name" }
    Set-Item -Path "function:script:$name" -Value $definition.Body.GetScriptBlock()
}
# GUI helpers resolve roots through Config\storage.json.txt.
. (Join-Path $project 'Scripts\Launchers\SmartM365-SharePointMigration-Storage.ps1')
function Assert-Archive { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message } }
$guiText = Get-Content -LiteralPath $guiPath -Raw
$activeGrid = [regex]::Match($guiText, '(?s)<DataGrid x:Name="gridSummary".*?</DataGrid>').Value
$archivedGrid = [regex]::Match($guiText, '(?s)<DataGrid x:Name="gridArchivedOverview".*?</DataGrid>').Value
$columnsPattern = '<DataGrid(?:Text|Template)Column Header="([^"]+)"'
$activeColumns = @([regex]::Matches($activeGrid, $columnsPattern) | ForEach-Object { $_.Groups[1].Value })
$archivedColumns = @([regex]::Matches($archivedGrid, $columnsPattern) | ForEach-Object { $_.Groups[1].Value })
Assert-Archive ($activeColumns.Count -eq 13 -and $activeColumns -contains 'Complexity' -and
    [string]::Join('|', $activeColumns) -eq [string]::Join('|', $archivedColumns)) `
    'ArchivedOverview columns no longer match Overview.'
foreach ($control in @('lblArchivedMigrations', 'lblArchivedFilesCompared',
    'lblArchivedPermissionsCompared', 'lblArchivedGlobalCompared', 'lblArchivedRefreshErrors')) {
    Assert-Archive ($guiText.Contains(('x:Name="{0}"' -f $control))) "Missing archived overview card: $control"
}
function Get-MigrationScope {
    param($Migration, [string]$Side)
    return [pscustomobject]@{ Text = "$Side scope"; Tooltip = '' }
}
function Get-SmartM365PortfolioRow {
    param($Migration, $SourceScope, $SourceTooltip, $TargetScope, $TargetTooltip, $SourceType, $TargetType)
    return [pscustomobject]@{
        Migration = $Migration.Name
        Source = $SourceScope; SourceTooltip = $SourceTooltip
        Destination = $TargetScope; DestinationTooltip = $TargetTooltip
        SourceScansDisplay = 'Files 2026-01-01'; TargetScansDisplay = 'Files 2026-01-02'
        ComparisonDisplay = '2026-01-03'; ComparisonDate = '2026-01-03'
        ComparisonRate = 99.0; ComparisonVisual = $null
        PermissionComparisonDisplay = '2026-01-04'; PermissionComparisonDate = '2026-01-04'
        PermissionComparisonRate = 98.0; PermissionComparisonVisual = $null
        GlobalComparisonRate = 98.5; GlobalComparisonVisual = $null
        Status = 'Review needed'; StatusTooltip = 'Review the comparison.'
        SourceInventoryDisplay = '—'; TargetInventoryDisplay = '—'
    }
}
function Update-PortfolioInventoryRow {
    param($Row)
    $Row.SourceInventoryDisplay = 'Files 10 · folders 2'
    $Row.TargetInventoryDisplay = 'Files 10 · folders 2'
}
function Update-PortfolioComplexityRow {
    param($Row, $Migration, [switch]$ReadOnly)
    # Archived rows must only read stored complexity results, never queue a calculation.
    $Row | Add-Member -NotePropertyName ComplexityText -NotePropertyValue $(if ($ReadOnly) { 'ReadOnly' } else { 'Queued' }) -Force
}
function Get-SmartM365ComplexityVisual { [pscustomobject]@{ Color = '#64748B'; Background = '#F0F4F8' } }
function Get-SmartM365PortfolioGapVisual { param($Days) [pscustomobject]@{ Band = 'Unavailable' } }
function Get-SmartM365PortfolioRateVisual {
    param($Rate, $Date, [string]$Notice = '', [string]$Caption = '')
    [pscustomobject]@{ Percent = if ($null -ne $Rate) { "$Rate %" } else { '—' }; Notice = $Notice; Caption = $Caption }
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
$script:ScriptRoot = Join-Path $tempRoot ('SharePointMigration-archive-test-' + [guid]::NewGuid().ToString('N'))
try {
    $activeRoot = New-Item -ItemType Directory -Path (Join-Path $script:ScriptRoot 'Migrations') -Force
    $archiveRoot = New-Item -ItemType Directory -Path (Join-Path $script:ScriptRoot 'ArchivedMigrations') -Force
    foreach ($name in @('Alpha', 'Beta')) {
        $folder = New-Item -ItemType Directory -Path (Join-Path $activeRoot.FullName $name)
        [IO.File]::WriteAllText((Join-Path $folder.FullName 'migration.config.psd1'),
            "@{ Name='$name'; Source=@{Type='SP2019'}; Target=@{Type='SPO'} }")
        [IO.File]::WriteAllText((Join-Path $folder.FullName 'evidence.txt'), "Synthetic $name")
    }
    $active = @(Get-MigrationFolders)
    Assert-Archive ($active.Count -eq 2) 'Both synthetic active migrations must be discovered.'
    $alpha = @($active | Where-Object Name -eq 'Alpha')[0]
    $destination = Move-MigrationToArchive -Migration $alpha -ProjectRoot $script:ScriptRoot
    Assert-Archive ((Test-Path -LiteralPath (Join-Path $destination 'evidence.txt') -PathType Leaf) -and
        -not (Test-Path -LiteralPath $alpha.Root)) 'The complete selected folder was not moved.'
    Assert-Archive (@(Get-MigrationFolders).Count -eq 1 -and @(Get-MigrationFolders)[0].Name -eq 'Beta') `
        'An archived migration remains in active discovery.'
    $archivedRows = @(Get-ArchivedOverviewRows)
    Assert-Archive ($archivedRows.Count -eq 1 -and $archivedRows[0].Migration -eq 'Alpha' -and
        $archivedRows[0].Status -eq 'Archived' -and $archivedRows[0].ComparisonDisplay -eq '2026-01-03' -and
        $archivedRows[0].GlobalComparisonRate -eq 98.5 -and
        $archivedRows[0].SourceInventoryDisplay -eq 'Files 10 · folders 2' -and
        $archivedRows[0].ComparisonVisual.Notice -eq '' -and
        $archivedRows[0].ComplexityText -eq 'ReadOnly' -and
        $archivedRows[0].FolderPath -eq $destination) `
        'ArchivedOverview did not display the archived migration and its evidence.'

    $invalid = New-Item -ItemType Directory -Path (Join-Path $archiveRoot.FullName 'Unconfigured')
    $archivedRows = @(Get-ArchivedOverviewRows)
    Assert-Archive ($archivedRows.Count -eq 2 -and
        @($archivedRows | Where-Object { $_.Migration -eq 'Unconfigured' -and $_.Status -eq 'Config unavailable' }).Count -eq 1) `
        'An archived folder without configuration was hidden.'

    [IO.File]::WriteAllText((Join-Path $archiveRoot.FullName 'Beta'), 'Keep this destination')
    $beta = @(Get-MigrationFolders)[0]
    $blocked = $false
    try { Move-MigrationToArchive -Migration $beta -ProjectRoot $script:ScriptRoot | Out-Null }
    catch { $blocked = $_.Exception.Message -like '*already exists*' }
    Assert-Archive ($blocked -and (Test-Path -LiteralPath $beta.Root -PathType Container) -and
        [IO.File]::ReadAllText((Join-Path $archiveRoot.FullName 'Beta')) -eq 'Keep this destination') `
        'An existing archive destination was overwritten.'

    $outside = [pscustomobject]@{ Name = 'Beta'; Root = $archiveRoot.FullName }
    $blocked = $false
    try { Move-MigrationToArchive -Migration $outside -ProjectRoot $script:ScriptRoot | Out-Null }
    catch { $blocked = $_.Exception.Message -like '*direct child*' }
    Assert-Archive ($blocked -and (Test-Path -LiteralPath $beta.Root -PathType Container)) `
        'A folder outside active migrations was accepted.'
    'Archive GUI offline test passed: move, active exclusion, overview, collision and path guard.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($script:ScriptRoot).TrimEnd('\', '/')
    if ($resolved.StartsWith($tempRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDN2srbTMW5pKng
# FAnLDr2tYpfC4Akiko0Lbah+qE6i7aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILJW55Hn35luN9BylLmMniD2cQWjo1oWHBBOc291aedtMA0GCSqG
# SIb3DQEBAQUABIIBgHarIqW6uhNNpgi5/CC8QPCBauPQgLduD03qBxalF71zICc1
# XK1y5p2LyQJyBVXmb9SwSFehxLuudRgNO2JljE1Uv7c8qrg1dSKKlMcM6wOhaMOU
# ew2yp111XsU2VnqwyWPIq7sg/fD0fCBi8PSjjvb8KvxDBMFligDTJkgOw0zr4ZYx
# IusuTK8yVTi9S+R0XjimughqGgfeMPNNKk7rjgTC1t36CWEz+2CeDWmfPry/BWUf
# J51AW7M2F4JKOLROOYvFcQJ2OwDc5UrUl4suQNT8I+BRN0i7C1da5URcT3Jcwmfs
# gid6Ttk792h9A2GcgGiYIazW9n5Va4YaljHie9zsDXtsz8GLXRIJPIScmhw4Dh3X
# Uv7sLqYR6zKGjGIQPgjjUWkt+sFpZErQKZxvxfoW+KBBb96JbW8TlMvTWV40z3Eg
# 6dx05GksV4OEZ0UnpubFwpdyjYkqwzDdkTBHH1TwKlfe/d9fCxT5GbZAu2D8Ce1L
# 4fnD+bjvjoeR7eLV8aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTcwNzE0
# NTZaMC8GCSqGSIb3DQEJBDEiBCCXSPsXukq8iRpFYCPm8dPw1zbAWLZg7SWFOPvO
# ehZe3DANBgkqhkiG9w0BAQEFAASCAgBcMiYKICYKYuZLS7+itx+ubdRffQ9VOdLc
# RmLfMnDrl9atbECT4A86ZbBuZJNvHgEE2Ome+VPj1hajhGk7v2cSiDoQFYbs6pOX
# CxMQ8ECaKcrNOg42qAxHQ5qtjZHTQQ0MSZnbpyRQYRkjBb4KFPFckjb4R9u0J/gD
# aduwCgqQ/mvlUBVglT2fk/YKg7vQO0rm686AMXRs9QZBh5feOrXYONRXi6hzii+i
# fOCjqZp0VNw9Apzf/pa/bJoZnsChbyt2v1DSDOIXMXKq1wx05um+ov4wZ8ZVkqQp
# DpDKY8Px6EJNDwNebFZOl9RKciJ8RN5OL9uLcYGyY4QQH3nyGz50Bu+UpnzQ8ncr
# +whNBVX5w3d93uFtj/KQC1Q3P89p3OMHP6uMARrMkZfLUUJ/pC2B75Jw6f4oY/Dt
# ItLxhQxTZ6ZNF1UXE6687BAOCK5e4soO/FbuRw9rn/5K7QSa2VPXEnJbnQm9FwuM
# eHS8XzACe2diKPlQ4wroKouJ5hGvGwHaTqoGQrxa9cPwIew40OkY0hJq25gswWe4
# DDPUxRF/e1L7H+OF/TfYf4HE7EI1ZiujZA3KiVcxBdMgA2vikm1ULD5ofnMh6oPg
# bMVXE1oPkwYxlsvV+6imjPMISvYfRc9uNmk/5UVDDHjQsYxc8DWq566dHrUtXU/r
# E5c6Ys4gig==
# SIG # End signature block
