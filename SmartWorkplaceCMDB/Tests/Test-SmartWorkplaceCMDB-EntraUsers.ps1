<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB Entra users collector and normalizer tests.

.VERSION
0.2.1
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.2.1'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$script:Passed = 0
$script:Failed = 0

function Invoke-SmartWorkplaceCMDBEntraUserTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$Test
    )

    try {
        & $Test
        $script:Passed++
        Write-Information "PASS $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Information "FAIL $Name - $($_.Exception.Message)" -InformationAction Continue
    }
}

function Assert-SmartWorkplaceCMDBEntraUserTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$Condition,

        [Parameter(Mandatory)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBEntraUserThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [Parameter(Mandatory)]
        [string]$MessagePattern
    )

    $caught = $null
    try {
        & $Action
    }
    catch {
        $caught = $_
    }
    if ($null -eq $caught) {
        throw "Expected an exception matching '$MessagePattern'."
    }
    if ($caught.Exception.Message -notmatch $MessagePattern) {
        throw "Unexpected exception: $($caught.Exception.Message)"
    }
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$collectorPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1'
$normalizerPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Normalize.ps1'
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$fixturePath = Join-Path $PSScriptRoot 'Fixtures\EntraUsers.sample.json'
$tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase ('SmartWorkplaceCMDB-EntraUsers-Tests-{0}' -f [guid]::NewGuid().ToString('N'))

Import-Module $modulePath -Force

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    $identityParameters = @{
        Tenant          = 'audit'
        OrganizationKey = 'contoso'
        EnvironmentKey  = 'prod'
        TenantKey       = 'contoso-prod'
        TenantId        = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        NoConfigWrite   = $true
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Validate offline collection without output writes' -Test {
        $validateRoot = Join-Path $tempRoot 'ValidateOnly'
        & $collectorPath @identityParameters `
            -DataRootPath $validateRoot `
            -InputJsonPath $fixturePath `
            -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition (-not (Test-Path -LiteralPath $validateRoot)) `
            -Message 'Collector ValidateOnly created a runtime output folder.'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Bind live Graph settings without an unsupported default argument' -Test {
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $collectorPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $settingCalls = @($ast.FindAll({
                    param($node)
                    $node -is [System.Management.Automation.Language.CommandAst] -and
                    $node.GetCommandName() -eq 'Get-SmartWorkplaceCMDBGraphSetting'
                }, $true))
        foreach ($call in $settingCalls) {
            $positionalArguments = @($call.CommandElements |
                Select-Object -Skip 1 |
                Where-Object { $_ -isnot [System.Management.Automation.Language.CommandParameterAst] })
            Assert-SmartWorkplaceCMDBEntraUserTrue `
                -Condition ($positionalArguments.Count -le 2) `
                -Message "Graph setting call passes an unsupported positional argument: $($call.Extent.Text)"
        }
    }

    $runtimeRoot = Join-Path $tempRoot 'Runtime'
    $script:CollectionResult = $null
    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Collect two offline Entra users into raw outputs' -Test {
        $script:CollectionResult = & $collectorPath @identityParameters `
            -DataRootPath $runtimeRoot `
            -InputJsonPath $fixturePath
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($script:CollectionResult.UserCount -eq 2) `
            -Message "Expected two collected users, got $($script:CollectionResult.UserCount)."
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition (Test-Path -LiteralPath $script:CollectionResult.RawLatestOutputPath -PathType Leaf) `
            -Message 'Latest raw Entra users CSV was not created.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition (Test-Path -LiteralPath $script:CollectionResult.HistoryPath -PathType Leaf) `
            -Message 'Historical raw Entra users CSV was not created.'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Validate the raw Entra users CSV contract' -Test {
        $latestRoot = Join-Path $runtimeRoot 'DATA-LAST'
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $latestRoot `
            -ContractPath $rawContractPath)
        $userResult = @($results | Where-Object Name -eq 'Entra_Users.csv')
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($userResult.Count -eq 1 -and $userResult[0].Status -eq 'Valid') `
            -Message 'Latest raw Entra users CSV is not contract-valid.'
        $rows = @(Import-Csv -LiteralPath $script:CollectionResult.RawLatestOutputPath)
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($rows.Count -eq 2) `
            -Message "Expected two latest raw rows, got $($rows.Count)."
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition (@($rows | Where-Object TenantKey -ne 'contoso-prod').Count -eq 0) `
            -Message 'Raw Entra users rows do not carry the expected tenant identity.'
        $franceRow = @($rows | Where-Object SourceUserId -eq '11111111-1111-1111-1111-111111111111')[0]
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($franceRow.UsageLocation -eq 'fr') `
            -Message 'Raw Entra users did not preserve usageLocation from the source.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($franceRow.LastSignInDateTime -match '^2026-09-10T08:30:00' -and
                $franceRow.LastNonInteractiveSignInDateTime -match '^2026-09-11T07:00:00' -and
                $franceRow.LastSuccessfulSignInDateTime -match '^2026-09-10T08:30:00') `
            -Message 'Raw Entra users did not preserve sign-in activity.'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Honor MaxItems without touching canonical output' -Test {
        $boundedRoot = Join-Path $tempRoot 'Bounded'
        $boundedResult = & $collectorPath @identityParameters `
            -DataRootPath $boundedRoot `
            -InputJsonPath $fixturePath `
            -MaxItems 1
        $boundedRows = @(Import-Csv -LiteralPath $boundedResult.RawLatestOutputPath)
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($boundedRows.Count -eq 1) `
            -Message "Expected one bounded raw row, got $($boundedRows.Count)."
    }

    $script:NormalizationResult = $null
    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Normalize raw users into CMDB_Users and DimUser' -Test {
        $script:NormalizationResult = & $normalizerPath @identityParameters -DataRootPath $runtimeRoot
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($script:NormalizationResult.UserCount -eq 2) `
            -Message "Expected two normalized users, got $($script:NormalizationResult.UserCount)."
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition (Test-Path -LiteralPath $script:NormalizationResult.CmdbUserOutputPath -PathType Leaf) `
            -Message 'CMDB_Users.csv was not created.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition (Test-Path -LiteralPath $script:NormalizationResult.DimUserOutputPath -PathType Leaf) `
            -Message 'DimUser.csv was not created.'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Validate curated user contracts and stable keys' -Test {
        $cmdbRows = @(Import-Csv -LiteralPath $script:NormalizationResult.CmdbUserOutputPath)
        $dimRows = @(Import-Csv -LiteralPath $script:NormalizationResult.DimUserOutputPath)
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($cmdbRows.Count -eq 2 -and $dimRows.Count -eq 2) `
            -Message 'Curated user tables do not contain two rows.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($cmdbRows[0].CmdbUserId -match '^contoso-prod\\|entra-user\\|') `
            -Message 'CMDB user key is not tenant-scoped and source-stable.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($dimRows[0].TenantUserKey -eq $cmdbRows[0].CmdbUserId) `
            -Message 'DimUser does not preserve the CMDB user key.'
        $franceUser = @($dimRows | Where-Object CountryCode -eq 'FR')[0]
        $unknownUser = @($dimRows | Where-Object CountryStatus -eq 'Not provided')[0]
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($franceUser.CountryLabel -eq 'FR' -and $franceUser.CountryStatus -eq 'Reported') `
            -Message 'DimUser did not normalize the reported usageLocation.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($unknownUser.CountryCode -eq '' -and $unknownUser.CountryLabel -eq 'Unknown / unassigned') `
            -Message 'DimUser did not preserve a missing usageLocation as unknown.'
        $franceCmdbUser = @($cmdbRows | Where-Object UsageLocation -eq 'FR')[0]
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($franceCmdbUser.UsageLocationStatus -eq 'Reported') `
            -Message 'CMDB_Users did not retain the normalized usageLocation status.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($franceCmdbUser.LastSuccessfulSignInDateTime -match '^2026-09-10T08:30:00') `
            -Message 'CMDB_Users did not retain the observed successful sign-in.'
        $disabledUser = @($dimRows | Where-Object AccountEnabled -eq 'False')[0]
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($disabledUser.ActivityState -eq 'Disabled') `
            -Message 'DimUser did not classify the disabled account explicitly.'
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition (-not [string]::IsNullOrWhiteSpace([string]$franceUser.ActivityState)) `
            -Message 'DimUser did not derive an activity state for the enabled account.'

        $contractResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
            -ContractPath $curatedContractPath)
        $userResults = @($contractResults | Where-Object Name -in @('CMDB_Users.csv', 'DimUser.csv'))
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($userResults.Count -eq 2 -and @($userResults | Where-Object Status -ne 'Valid').Count -eq 0) `
            -Message 'Curated user CSV headers are not contract-valid.'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Reject raw data from another tenant identity' -Test {
        $mismatchPath = Join-Path $tempRoot 'Entra_Users_Mismatch.csv'
        $mismatchRows = @(Import-Csv -LiteralPath $script:CollectionResult.RawLatestOutputPath)
        $mismatchRows[0].TenantKey = 'other-prod'
        $mismatchRows | Export-Csv -LiteralPath $mismatchPath -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBEntraUserThrow -Action {
            & $normalizerPath @identityParameters `
                -DataRootPath (Join-Path $tempRoot 'MismatchOutput') `
                -RawInputPath $mismatchPath | Out-Null
        } -MessagePattern 'identity mismatch'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Reject duplicate Entra source user identifiers' -Test {
        $duplicatePath = Join-Path $tempRoot 'Entra_Users_Duplicate.csv'
        $duplicateRows = @(Import-Csv -LiteralPath $script:CollectionResult.RawLatestOutputPath)
        $duplicateRows[1].SourceUserId = $duplicateRows[0].SourceUserId
        $duplicateRows | Export-Csv -LiteralPath $duplicatePath -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBEntraUserThrow -Action {
            & $normalizerPath @identityParameters `
                -DataRootPath (Join-Path $tempRoot 'DuplicateOutput') `
                -RawInputPath $duplicatePath | Out-Null
        } -MessagePattern 'duplicate SourceUserId'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Classify an invalid usageLocation without fabricating a country' -Test {
        $invalidPath = Join-Path $tempRoot 'Entra_Users_InvalidUsageLocation.csv'
        $invalidRows = @(Import-Csv -LiteralPath $script:CollectionResult.RawLatestOutputPath)
        $invalidRows[0].UsageLocation = 'France'
        $invalidRows | Export-Csv -LiteralPath $invalidPath -NoTypeInformation -Encoding UTF8
        $invalidResult = & $normalizerPath @identityParameters `
            -DataRootPath (Join-Path $tempRoot 'InvalidUsageLocationOutput') `
            -RawInputPath $invalidPath
        $invalidUser = @(Import-Csv -LiteralPath $invalidResult.DimUserOutputPath | Where-Object CountryStatus -eq 'Invalid')[0]
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($invalidUser.CountryCode -eq '' -and $invalidUser.CountryLabel -eq 'Unknown / unassigned') `
            -Message 'An invalid usageLocation was converted into a country.'
    }

    Invoke-SmartWorkplaceCMDBEntraUserTest -Name 'Validate normalizer without modifying curated outputs' -Test {
        $beforeCmdbHash = (Get-FileHash -LiteralPath $script:NormalizationResult.CmdbUserOutputPath -Algorithm SHA256).Hash
        $beforeDimHash = (Get-FileHash -LiteralPath $script:NormalizationResult.DimUserOutputPath -Algorithm SHA256).Hash
        & $normalizerPath @identityParameters `
            -DataRootPath $runtimeRoot `
            -ValidateOnly | Out-Null
        $afterCmdbHash = (Get-FileHash -LiteralPath $script:NormalizationResult.CmdbUserOutputPath -Algorithm SHA256).Hash
        $afterDimHash = (Get-FileHash -LiteralPath $script:NormalizationResult.DimUserOutputPath -Algorithm SHA256).Hash
        Assert-SmartWorkplaceCMDBEntraUserTrue `
            -Condition ($beforeCmdbHash -eq $afterCmdbHash -and $beforeDimHash -eq $afterDimHash) `
            -Message 'Normalizer ValidateOnly modified a curated user table.'
    }
}
finally {
    $resolvedTempRoot = [System.IO.Path]::GetFullPath($tempRoot)
    if ($resolvedTempRoot.StartsWith($tempBase, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedTempRoot) -like 'SmartWorkplaceCMDB-EntraUsers-Tests-*' -and
        (Test-Path -LiteralPath $resolvedTempRoot)) {
        Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB Entra users tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion,
    $script:Passed,
    $script:Failed
) -InformationAction Continue

if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPK2+yHQYjXzmF
# g78Tn3gY9erYFr9VnMgDyNXth3TElqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP4r1kYnMDwW8Nyn33rQploKvELwAr1HENJRfw7yPePBMA0GCSqG
# SIb3DQEBAQUABIIBgB7tz6zE+i9T7M/UK70qBThtLrZINN0Z/TwyPSmolnpYItyd
# Njyy++D4btxAVZRJiosQf8AuPrFI1l3iyCL1lMx5pzetL8gPaN+pclpZkW5r0GcI
# AwMRlUVVpWQtXjJw5Klli5qYSJrHRbFn5xbgio1gtQOy/UAv6lkQnlAf0qehPj6E
# /33w3Hn13B2Ob2KWyeTVtC/HDMw3IDnuUQhg+RJgkmNla4dmUXsdnmRd/bjfq8x3
# 6T02tp0HjRcfD2B/HDgjeRAsMMn5Csa5a07ATYOKBx3/+FLkMcTbSPNRyG0CX182
# 6rDODCGJKSCBb9VXKEegNYtURS3awvEFNn4fLz4cQfirGT3q7z8yxF4wUlhLXrNU
# xhpwZPuSEk0XsMZ995xkYAWg6QBlfMLHmpoRdXbzJWoMaVarVjlfhft80hl5zNcE
# 6EVMhY0QQbJBpPQIMBz/OA+2bmQsGZh3YzgnjoQIgwxallvn5ax5FNDcMAauhIHc
# dmT/uGAK5wm9VBAwbqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIyMDI1
# MjNaMC8GCSqGSIb3DQEJBDEiBCD9lQD8T86vLzt2bYMVfWv6hLzYN/7aGlQh+N6l
# 2aKvlzANBgkqhkiG9w0BAQEFAASCAgBZi88QWfdwyRJ3jdFCSqfZrv0K7fm69/Cc
# evhRUxIYnvxRxGSO1PhSFZUH6RPjDaemhogByIxOqI8tHhuZVGUUv7Ha6qCevPce
# CHKFenH7NL9HkdD9rdSAPXQ8J0IkX95pNhWTJgzRCC6kXflkv4+D0C2G89+lr3Wm
# EMtllAfbUjzmRQENQINoQwEUgXaE+TVkc5anOF2LvXvTZG4ZJGLyJTJzhm8X5iqa
# 8TGDXAwzRHNnzrP+tdXKcp3X4y9OsYiUOIvVbZdEBYaH7R7KYmzpGblO2wZrHmqQ
# DYmJ3UnM9aaJqK+oYv+srO3LeKfv7kHlADAyOuYVrrKEL7MagAHvm97MA90MmMcM
# W2I7bot4EswjfgzzAAZc9nd+uhNjeTJ8akDL/n3POZb+XjV3eYfHxkweGwp7czz0
# fefV5U2Qn9jpoayE9929qsrtJfAQToOk6hc2w8hUFHbbNJX3nibmGpfV6AuoCoIU
# qmvZo59my6GjYgoOFMUS8V3vL/v9rmkYsu0YeQGgfvBuyYHm5J3XaJ8Q5gVgg7qT
# 6eSgC5KEXBmypLOfOhqB1mmlAy+8wgNlBVrsvQGkVVZ4k6eOQjG2e0zFSYBlzBZN
# bD7nJ93vDRh1VrnvWnzNnduIPCwZ/PX3BiVaxfy+tZX436c6fDBC15iJAAuwdwpi
# 2HQQADTUeQ==
# SIG # End signature block
