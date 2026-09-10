<#
.SYNOPSIS
Synthetic regression tests for shared inventory identity and atomic persistence.
.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = New-Object 'System.Collections.Generic.List[object]'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-Offline-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}
function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try { & $Body; $results.Add([pscustomobject]@{ Name = $Name; Passed = $true; Error = '' }) }
    catch { $results.Add([pscustomobject]@{ Name = $Name; Passed = $false; Error = $_.Exception.Message }) }
}
function Import-OfflineFunctions {
    param([string]$Path, [string[]]$Names)
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "Source parse failed: $Path" }
    $definitions = foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if ($null -eq $node) { throw "Function not found: $name" }
        $node.Extent.Text
    }
    # Only the named function definitions run. No module initializers, tenant loader,
    # collector entry point, credentials, task scheduler, mail or network calls.
    New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
}

try {
    foreach ($variant in @('Core', 'WindowsPowerShell5')) {
        $relative = if ($variant -eq 'Core') { 'Modules/SmartM365.Core/SmartM365.Core.psm1' } else { 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1' }
        $module = Import-OfflineFunctions (Join-Path $SourceRoot $relative) @(
            'Add-SmartM365TenantKeyToCsvData', 'Write-SmartM365CsvAtomically',
            'Get-SmartM365CoreContextValue', 'Assert-SmartM365CsvDataCompleteness',
            'Get-SmartM365CsvValidationRule', 'Get-SmartM365CsvValidationBaseName',
            'Publish-SmartM365Csv', 'Get-SmartM365MaxItemsValue', 'Test-SmartM365MaxItemsMode',
            'Get-SmartM365MaxItemsSuffix', 'Add-SmartM365MaxItemsSuffixToCsvPath', 'Limit-SmartM365RowsForMaxItems'
        )
        & $module {
            function script:WriteLog { param($Message, $Level) }
            function script:Invoke-SmartM365SharePointCsvUpload { throw 'Unexpected SharePoint call in offline test.' }
            function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param($SourceFiles,$TimestampedPath) }
            $script:SmartM365CoreTenantKey = 'synthetic-a'
            $script:SmartM365CoreOrganizationKey = 'synthetic-org'
            $script:SmartM365CoreEnvironmentKey = 'test'
            $script:SmartM365CoreTenantId = '00000000-0000-0000-0000-000000000001'
        }
        Test-OfflineCase "$variant identity-first CSV roundtrip" {
            $path = Join-Path $testRoot "$variant-roundtrip.csv"
            $label = 'quoted "value", ' + [char]0x00e9 + "`nsecond line"
            & $module { param($p,$s) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{ Id = '001'; Label = $s }) -Columns @('Label','Id') } $path $label
            $row = @(Import-Csv -LiteralPath $path -Encoding UTF8)
            Assert-Offline ($row.Count -eq 1 -and $row[0].Label -ceq $label -and $row[0].Id -ceq '001') 'CSV multiline, Unicode or string identity changed.'
            Assert-Offline (($row[0].PSObject.Properties.Name -join ',') -eq 'TenantKey,OrganizationKey,EnvironmentKey,TenantId,Label,Id') 'Column contract changed.'
        }
        Test-OfflineCase "$variant empty explicit schema" {
            $path = Join-Path $testRoot "$variant-empty.csv"
            & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @() -Columns @('Id','Value') } $path
            Assert-Offline ((Get-Content -LiteralPath $path -TotalCount 1) -eq '"TenantKey","OrganizationKey","EnvironmentKey","TenantId","Id","Value"') 'Empty header changed.'
        }
        foreach ($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
            Test-OfflineCase "$variant reject conflicting $field and preserve last" {
                $path = Join-Path $testRoot "$variant-$field.csv"
                [IO.File]::WriteAllText($path, 'LAST VALID SYNTHETIC EXPORT')
                $hash = (Get-FileHash -LiteralPath $path).Hash
                $row = [pscustomobject]@{ Id = 'fixture' }
                $row | Add-Member -NotePropertyName $field -NotePropertyValue 'synthetic-other'
                $caught = $false
                try { & $module { param($p,$r) Write-SmartM365CsvAtomically -Path $p -Data @($r) } $path $row }
                catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
                Assert-Offline $caught 'Conflicting identity was silently relabelled.'
                Assert-Offline ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Last valid file changed.'
            }
        }
        Test-OfflineCase "$variant mixed rows without context" {
            $caught = $false
            try {
                & $module {
                    Add-SmartM365TenantKeyToCsvData -TenantKey '' -OrganizationKey '' -EnvironmentKey '' -TenantId '' -Data @(
                        [pscustomobject]@{TenantKey='synthetic-a'; OrganizationKey='org'; EnvironmentKey='test'; Id='1'},
                        [pscustomobject]@{TenantKey='synthetic-b'; OrganizationKey='org'; EnvironmentKey='test'; Id='2'}
                    )
                } | Out-Null
            } catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
            Assert-Offline $caught 'Mixed tenants accepted when deriving identity from rows.'
        }
        Test-OfflineCase "$variant dictionary identity inference" {
            $value = & $module { Add-SmartM365TenantKeyToCsvData -TenantKey '' -OrganizationKey '' -EnvironmentKey '' -TenantId '' -Data @(@{TenantKey='synthetic-a';OrganizationKey='org';EnvironmentKey='test';Id='1'}) }
            Assert-Offline ($value.Data[0].TenantKey -eq 'synthetic-a' -and $value.Data[0].Id -eq '1') 'Dictionary identity was not recognized.'
        }
        Test-OfflineCase "$variant matching and missing identity accepted" {
            $value = & $module { Add-SmartM365TenantKeyToCsvData -Data @([pscustomobject]@{TenantKey='SYNTHETIC-A';Id='1'}, [pscustomobject]@{TenantKey='';Id='2'}) }
            Assert-Offline ($value.Data.Count -eq 2 -and $value.Data[1].TenantKey -eq 'synthetic-a') 'Valid identity inheritance changed.'
        }
        Test-OfflineCase "$variant dictionary conflicting identity rejected" {
            $caught = $false
            try { & $module { Add-SmartM365TenantKeyToCsvData -Data @(@{TenantKey='synthetic-other'; Id='1'}) } | Out-Null }
            catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
            Assert-Offline $caught 'Dictionary conflict accepted.'
        }
        Test-OfflineCase "$variant failed serializer preserves last under Continue" {
            $path = Join-Path $testRoot "$variant-write.csv"
            [IO.File]::WriteAllText($path, 'LAST VALID SYNTHETIC EXPORT')
            $hash = (Get-FileHash -LiteralPath $path).Hash
            # A partial staging file followed by a non-terminating cmdlet error.
            & $module {
                function script:Export-Csv {
                    [CmdletBinding()]param([Parameter(ValueFromPipeline)]$InputObject, $Path, $Encoding, $Delimiter, [switch]$NoTypeInformation)
                    process { [IO.File]::WriteAllText($Path, 'PARTIAL'); Write-Error 'Synthetic disk failure' }
                }
            }
            try {
                $caught = $false
                try { & $module { param($p) $ErrorActionPreference='Continue'; Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Id='1'}) } $path 2>$null }
                catch { $caught = $_.Exception.Message -match 'Synthetic disk failure' }
                Assert-Offline $caught 'Serializer failure did not terminate publication.'
                Assert-Offline ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Partial staging file replaced last valid CSV.'
            } finally { & $module { Remove-Item Function:Export-Csv } }
        }
        Test-OfflineCase "$variant locked destination preserves last" {
            $path = Join-Path $testRoot "$variant-locked.csv"
            [IO.File]::WriteAllText($path, 'LAST VALID SYNTHETIC EXPORT')
            $lock = [IO.File]::Open($path, 'Open', 'Read', 'Read')
            try {
                $caught = $false
                try { & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Id='1'}) } $path }
                catch { $caught = $true }
                Assert-Offline $caught 'Locked replacement unexpectedly succeeded.'
            } finally { $lock.Dispose() }
            Assert-Offline ([IO.File]::ReadAllText($path) -eq 'LAST VALID SYNTHETIC EXPORT') 'Locked file changed.'
        }
        Test-OfflineCase "$variant successful DATA-ALL DATA-LAST byte parity" {
            $history = Join-Path $testRoot "$variant-history.csv"
            $latest = Join-Path $testRoot "$variant-latest.csv"
            & $module { param($h,$l) Publish-SmartM365Csv -TimestampedPath $h -LatestPath $l -Data @([pscustomobject]@{Id='001'; Date='2026-01-01T12:00:00Z'; Missing=$null}) -RetentionMaxCsv 0 -NoSharePointUpload } $history $latest | Out-Null
            Assert-Offline ((Get-FileHash $history).Hash -eq (Get-FileHash $latest).Hash) 'Latest and historical bytes differ.'
            $row = Import-Csv $latest
            Assert-Offline ($row.Date -ceq '2026-01-01T12:00:00Z' -and $row.Missing -ceq '') 'Timestamp or missing value changed.'
        }
        Test-OfflineCase "$variant publication conflict preserves both copies" {
            $history = Join-Path $testRoot "$variant-pair-history.csv"
            $latest = Join-Path $testRoot "$variant-pair-latest.csv"
            [IO.File]::WriteAllText($history, 'HISTORY'); [IO.File]::WriteAllText($latest, 'LATEST')
            $caught = $false
            try { & $module { param($h,$l) Publish-SmartM365Csv -TimestampedPath $h -LatestPath $l -Data @([pscustomobject]@{TenantKey='synthetic-other';Id='001'}) -RetentionMaxCsv 0 -NoSharePointUpload } $history $latest | Out-Null }
            catch { $caught = $_.Exception.Message -match 'identity.*conflict|conflict.*identity' }
            Assert-Offline $caught 'Publication identity conflict accepted.'
            Assert-Offline ([IO.File]::ReadAllText($history) -eq 'HISTORY' -and [IO.File]::ReadAllText($latest) -eq 'LATEST') 'Publication changed a prior copy.'
        }
        Test-OfflineCase "$variant MAXITEMS protects canonical files" {
            $history = Join-Path $testRoot "$variant-max-history.csv"
            $latest = Join-Path $testRoot "$variant-max-latest.csv"
            [IO.File]::WriteAllText($latest, 'CANONICAL')
            $global:SmartM365MaxItems = 1
            try {
                $published = & $module { param($h,$l) Publish-SmartM365Csv -TimestampedPath $h -LatestPath $l -Data @([pscustomobject]@{Id='1'},[pscustomobject]@{Id='2'}) -RetentionMaxCsv 0 -NoSharePointUpload } $history $latest
                Assert-Offline ($published.LatestPath -like '*_MAXITEMS-1.csv') 'Limited run used canonical name.'
                Assert-Offline ([IO.File]::ReadAllText($latest) -eq 'CANONICAL') 'Canonical latest changed.'
                Assert-Offline (@(Import-Csv $published.LatestPath).Count -eq 1) 'MaxItems row limit changed.'
            } finally { $global:SmartM365MaxItems = 0 }
        }
        Test-OfflineCase "$variant tenant-neutral semicolon contract" {
            $path = Join-Path $testRoot "$variant-neutral.csv"
            & $module { param($p) Write-SmartM365CsvAtomically -Path $p -Data @([pscustomobject]@{Date='2026-01-01';Value='1.25'}) -NoTenantKey -Delimiter ';' } $path
            $row = Import-Csv $path -Delimiter ';'
            Assert-Offline (($row.PSObject.Properties.Name -join ',') -eq 'Date,Value' -and $row.Value -ceq '1.25') 'Neutral/delimiter contract changed.'
        }
        Remove-Module $module
    }
    $distributed = Import-OfflineFunctions (Join-Path $SourceRoot 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1') @(
        'Write-JsonAtomically', 'ConvertTo-SafeFileName', 'Enter-SmartM365OrchestratorConcurrencyLease',
        'Set-SmartM365OrchestratorConcurrencyLease', 'Exit-SmartM365OrchestratorConcurrencyLease',
        'Enter-SmartM365OrchestratorOccurrenceClaim', 'Set-SmartM365OrchestratorOccurrenceClaim'
    )
    Test-OfflineCase 'Orchestrator locked lease update must not report success' {
        $lease = & $distributed { param($p) Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $p -ConcurrencyKey 'Synthetic' -JobName 'JobA' -Occurrence ([datetime]'2026-01-01') -OwnerServer 'SERVER-A' } (Join-Path $testRoot 'leases')
        $hash = (Get-FileHash -LiteralPath $lease.LeasePath).Hash
        $lock = [IO.File]::Open($lease.LeasePath, 'Open', 'Read', 'Read')
        try {
            $caught = $false
            try { & $distributed { param($l) $ErrorActionPreference='Continue'; Set-SmartM365OrchestratorConcurrencyLease -LeasePath $l.LeasePath -LeaseId $l.Lease.LeaseId -OwnerServer 'SERVER-A' -SafeUntilUtc ([datetime]::UtcNow.AddHours(5)) } $lease 2>$null | Out-Null }
            catch { $caught = $true }
            Assert-Offline $caught 'Lease update acknowledged although persistence failed.'
        } finally { $lock.Dispose() }
        Assert-Offline ((Get-FileHash -LiteralPath $lease.LeasePath).Hash -eq $hash) 'Locked lease changed.'
    }
    Test-OfflineCase 'Orchestrator locked claim update must fail' {
        $claim = & $distributed { param($p) Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $p -JobName 'JobA' -Occurrence ([datetime]'2026-01-01') -OwnerServer 'SERVER-A' -PlanId 'synthetic' } (Join-Path $testRoot 'claims')
        $lock = [IO.File]::Open($claim.ClaimPath, 'Open', 'Read', 'Read')
        try {
            $caught = $false
            try { & $distributed { param($c) $ErrorActionPreference='Continue'; Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $c.ClaimPath -OwnerServer 'SERVER-A' -Status Running } $claim 2>$null | Out-Null }
            catch { $caught = $true }
            Assert-Offline $caught 'Claim update acknowledged although persistence failed.'
        } finally { $lock.Dispose() }
    }
    Remove-Module $distributed
}
finally {
    # Delete only this run's verified synthetic temporary root.
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $expectedParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($expectedParent, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartInventory-Offline-*') { throw 'Unsafe cleanup root.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
if ($ResultPath) { $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ResultPath -Encoding UTF8 }
$results | Format-Table Name, Passed, Error -AutoSize
$failed = @($results | Where-Object { -not $_.Passed }).Count
Write-Output ("Cases={0}; Passed={1}; Failed={2}; PowerShell={3}" -f $results.Count, ($results.Count-$failed), $failed, $PSVersionTable.PSVersion)
if ($failed) { exit 1 }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAAMqxVxMssUzyA
# 9XfBv2qAMm/4vqzbBgED0L26ziDhh6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBFET33jWc3ichPJwmu8DuJ
# sxLgV2SD+e0Iq6aPH/bfcjANBgkqhkiG9w0BAQEFAASCAYAyExaZ2KAaWjHuaHsb
# 7uSWcx+AXEP2SM7/zfTWoEnshPwPqpZkfGEPE/0lkm/uiaJzl9qaGDEj+vOQ1QlY
# /A4fDSOSfxPZFI9+hSVS24teC62XcJ2A7eRBb8VzKXrsijFGUglDdmvtHfRY1nMp
# HYc8EArFEA0OH1/Jx2i0m4//KJPl4IK6MlBqPL5VqJnPsCVX+8IGNZVqG41t9vQb
# i0vMK0lnOgL0DJMTNFYG5wlckEbTvV8rsA5zUevSyVB+bl9JQdAga4ohb+Ax5uhZ
# hekBZ9hyGumi2WCElYNUCxLVV07kDpjQbZCmyHVSxUleO8a2JxotI18B0NteicPQ
# NEhD1zAWxJWKb5/3kAY9T554BYUES9MAMydYN1eoK9hsuPlXIAvGzEHq4UTWir9N
# C+Tjg+HhlH4kzrmF/k5mqCx1DHJGgcbPOrfNjeNR3GbmhHrC8n05BpkO9sBYhJKD
# DfS+VYmKQZ8s2DRyzjQ+V8k40dzHm0oUtWNoP4zhtZX1Flw=
# SIG # End signature block
