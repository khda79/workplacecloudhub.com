# V1 stable. Synthetic offline tests only; no Graph authentication or request.
[CmdletBinding()]
param()
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:passed = 0; $script:failed = 0
function Test-Case {
    param([string]$Name, [scriptblock]$Action)
    try { & $Action; $script:passed++; Write-Output "PASS $Name" }
    catch { $script:failed++; Write-Output "FAIL $Name : $($_.Exception.Message)" }
}
function Assert-True {
    param([bool]$Value, [string]$Message)
    if (-not $Value) { throw $Message }
}
function Assert-Throw {
    param([scriptblock]$Action, [string]$Pattern)
    $caught = $null
    try { & $Action | Out-Null } catch { $caught = $_.Exception.Message }
    if (-not $caught -or $caught -notmatch $Pattern) { throw "Expected failure '$Pattern'; got '$caught'." }
}
function Write-Fixture {
    param([AllowEmptyCollection()][object[]]$Rows)
    @{value = @($Rows)} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $fixture -Encoding UTF8
}
$project = Split-Path $PSScriptRoot -Parent
$collector = Join-Path $project 'Collectors\Intune\SmartWorkplaceCMDB-IntuneHardware-Collect.ps1'
$contract = Get-Content -Raw (Join-Path $project 'Schema\SmartWorkplaceCMDB.hardware.tables.json') | ConvertFrom-Json
$base = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
$temp = Join-Path $base ('SmartWorkplaceCMDB-Hardware-Tests-' + [guid]::NewGuid().ToString('N'))
$identity = @{Tenant='audit'; OrganizationKey='example'; EnvironmentKey='test'; TenantKey='example-test'; TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'}
$fixture = Join-Path $temp 'synthetic.json'
$runtime = Join-Path $temp 'Runtime'
$protected = @(
    'Collectors\Intune\SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1',
    'Collectors\Intune\SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1',
    'Schema\SmartWorkplaceCMDB.raw.tables.json', 'Schema\SmartWorkplaceCMDB.tables.json'
)
$before = @{}
foreach ($file in $protected) { $before[$file] = (Get-FileHash -LiteralPath (Join-Path $project $file)).Hash }
try {
    New-Item -ItemType Directory -Path $temp | Out-Null
    $valid = @(
        @{id='managed-1';azureADDeviceId='entra-1';serialNumber='SERIAL-SYNTHETIC';manufacturer='Example vendor';model='Example model';totalStorageSpaceInBytes=512110190592;imei='MUST-NOT-EXPORT';userPrincipalName='fake@example.invalid';TenantKey='foreign-test'},
        @{id='managed-2';serialNumber='SERIAL-SYNTHETIC';manufacturer="  Example`nVendor  ";totalStorageSpaceInBytes=0},
        @{id='managed-3';serialNumber=$null;model='';totalStorageSpaceInBytes=$null}
    )
    Write-Fixture $valid
    Test-Case 'Default Graph mode validates without writes or authentication' {
        $out = & $collector @identity -DataRootPath (Join-Path $temp 'NoWrites')
        Assert-True ($out.Status -eq 'ValidatedOffline' -and $out.SourceMode -eq 'GraphNotExecuted' -and -not $out.AuthenticationValidated) 'Unexpected default behavior.'
        Assert-True (-not (Test-Path (Join-Path $temp 'NoWrites'))) 'Default created files.'
        Assert-True ($out.RequestUri -ceq ('https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=' + $contract.graphSelect + '&$top=999')) 'Unexpected request.'
    }
    Test-Case 'Collect plus ValidateOnly never runs Graph' {
        $out = & $collector @identity -DataRootPath (Join-Path $temp 'NoWrites2') -Collect -ValidateOnly
        Assert-True ($out.SourceMode -eq 'GraphNotExecuted' -and -not (Test-Path (Join-Path $temp 'NoWrites2'))) 'Validation wrote output.'
    }
    Test-Case 'Fixture ValidateOnly validates rows without files' {
        $out = & $collector @identity -DataRootPath (Join-Path $temp 'NoWrites3') -InputJsonPath $fixture -ValidateOnly
        Assert-True ($out.DeviceCount -eq 3 -and -not (Test-Path (Join-Path $temp 'NoWrites3'))) 'Fixture validation mismatch.'
    }
    Test-Case 'Explicit false Collect remains offline' {
        $out = & $collector @identity -DataRootPath (Join-Path $temp 'NoWrites4') -Collect:$false
        Assert-True ($out.SourceMode -ceq 'GraphNotExecuted' -and -not (Test-Path (Join-Path $temp 'NoWrites4'))) 'False Collect enabled execution.'
    }
    Test-Case 'Fixture collection preserves attributes and separate source identities' {
        $script:result = & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture
        $script:rows = @(Import-Csv -LiteralPath $script:result.RawLatestOutputPath)
        Assert-True ($script:result.DeviceCount -eq 3 -and $script:result.DuplicateSerialValueCount -eq 1) 'Repeated serial merged or lost.'
        Assert-True ($script:rows[0].TotalStorageSpaceInBytes -ceq '512110190592' -and $script:rows[0].StorageStatus -ceq 'Reported') 'Capacity precision lost.'
        Assert-True ($script:rows[0].Manufacturer -ceq 'Example vendor' -and $script:rows[0].Model -ceq 'Example model') 'Text values lost.'
        Assert-True ($script:rows[1].Manufacturer -ceq 'Example Vendor') 'Multiline text was not cleaned.'
    }
    Test-Case 'Missing and zero values remain distinct' {
        Assert-True ($script:rows[1].StorageStatus -ceq 'ZeroReported' -and $script:rows[1].TotalStorageSpaceInBytes -ceq '0') 'Zero was inferred or erased.'
        Assert-True ($script:rows[2].StorageStatus -ceq 'Missing' -and $script:rows[2].TotalStorageSpaceInBytes -ceq '' -and $script:rows[2].SerialNumberStatus -ceq 'Missing') 'Missing value fabricated.'
    }
    Test-Case 'Exact contract excludes unrelated sensitive fields and foreign fixture tenant' {
        Assert-True (($script:rows[0].PSObject.Properties.Name -join ',') -ceq ($contract.tables[0].columns -join ',')) 'Header drift.'
        Assert-True (@($script:rows | Where-Object { $_.TenantKey -cne 'example-test' -or $_.TenantId -cne $identity.TenantId }).Count -eq 0) 'Tenant projection mismatch.'
        Assert-True ((Get-Content -Raw $script:result.RawLatestOutputPath) -notmatch 'MUST-NOT-EXPORT|fake@example.invalid|foreign-test') 'Unselected data leaked.'
    }
    Test-Case 'Source sidecar, history hash and UTC retrieval time agree' {
        $state = Get-Content -Raw ($script:result.RawLatestOutputPath + '.status.json.txt') | ConvertFrom-Json
        $hash = (Get-FileHash $script:result.RawLatestOutputPath).Hash
        Assert-True ($state.Status -ceq 'Completed' -and $state.RowCount -eq 3 -and $state.SHA256 -ceq $hash -and $state.Coverage -ceq 'Fixture' -and $state.Mode -ceq 'Fixture') 'Invalid source evidence.'
        Assert-True ((Get-FileHash $script:result.HistoryPath).Hash -ceq $hash) 'History differs.'
        Assert-True (@($script:rows | Where-Object { $_.SourceCollectedDateTime -cne $script:result.CollectedDateTime -or $_.SourceCollectedDateTime -notmatch 'Z$' }).Count -eq 0) 'Retrieval time mismatch.'
        Assert-True ($script:result.HistoryPath -match '\\DeviceHardware\\\d{4}\\\d{2}\\') 'Incorrect history folders.'
    }
    Test-Case 'Bounded collection is isolated and labelled Bounded' {
        $out = & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture -MaxItems 1
        $state = Get-Content -Raw ($out.RawLatestOutputPath + '.status.json.txt') | ConvertFrom-Json
        Assert-True ($out.DeviceCount -eq 1 -and $state.Coverage -ceq 'Bounded' -and $state.MaxItems -eq 1 -and $out.RawLatestOutputPath -ne $script:result.RawLatestOutputPath) 'Bounded snapshot replaced complete fixture.'
    }
    Test-Case 'Foreign tenant root is isolated, original snapshot preserved' {
        $other = $identity.Clone(); $other.OrganizationKey='other'; $other.TenantKey='other-test'; $other.TenantId='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
        $hash = (Get-FileHash $script:result.RawLatestOutputPath).Hash
        $out = & $collector @other -DataRootPath $runtime -InputJsonPath $fixture
        Assert-True ($out.RawLatestOutputPath -ne $script:result.RawLatestOutputPath -and (Get-FileHash $script:result.RawLatestOutputPath).Hash -eq $hash) 'Cross-tenant overwrite.'
        Assert-True (@(Import-Csv $out.RawLatestOutputPath | Where-Object TenantKey -cne 'other-test').Count -eq 0) 'Wrong output tenant.'
    }
    Test-Case 'Duplicate native IDs fail before changing existing evidence' {
        $hash = (Get-FileHash $script:result.RawLatestOutputPath).Hash
        $stateHash = (Get-FileHash ($script:result.RawLatestOutputPath + '.status.json.txt')).Hash
        Write-Fixture @(@{id='duplicate'}, @{id='DUPLICATE'})
        Assert-Throw { & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture } 'Duplicate managed'
        Assert-True ((Get-FileHash $script:result.RawLatestOutputPath).Hash -eq $hash -and (Get-FileHash ($script:result.RawLatestOutputPath + '.status.json.txt')).Hash -eq $stateHash) 'Rejected fixture damaged prior evidence.'
    }
    Test-Case 'Missing native ID is rejected in validation' {
        Write-Fixture @(@{serialNumber='UNKEYED'})
        Assert-Throw { & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture -ValidateOnly } 'missing id'
    }
    foreach ($invalid in @(-1, 1.5, '', '9223372036854775808', $true)) {
        Test-Case "Invalid storage rejected ($invalid)" {
            Write-Fixture @(@{id='bad-storage';totalStorageSpaceInBytes=$invalid})
            Assert-Throw { & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture -ValidateOnly } 'nonnegative Int64'
        }
    }
    Test-Case 'Non-scalar text rejected without exposing raw content' {
        Write-Fixture @(@{id='bad-text';serialNumber=@{secret='NOT-FOR-LOGS'}})
        Assert-Throw { & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture -ValidateOnly } 'strings or null'
    }
    Test-Case 'Int64 maximum remains exact' {
        Write-Fixture @(@{id='max-storage';totalStorageSpaceInBytes='9223372036854775807'})
        $out = & $collector @identity -DataRootPath (Join-Path $temp 'MaxStorage') -InputJsonPath $fixture
        Assert-True ((Import-Csv $out.RawLatestOutputPath).TotalStorageSpaceInBytes -ceq '9223372036854775807') 'Int64 precision loss.'
    }
    Test-Case 'Existing inventory fixture stays compatible with explicitly missing hardware' {
        $legacyFixture = Join-Path $PSScriptRoot 'Fixtures\IntuneManagedDevices.sample.json'
        $out = & $collector @identity -DataRootPath (Join-Path $temp 'LegacyFixture') -InputJsonPath $legacyFixture
        $legacyRows = @(Import-Csv $out.RawLatestOutputPath)
        Assert-True ($legacyRows.Count -eq 3 -and @($legacyRows | Where-Object {
            $_.SerialNumberStatus -cne 'Missing' -or $_.ManufacturerStatus -cne 'Missing' -or $_.ModelStatus -cne 'Missing' -or $_.StorageStatus -cne 'Missing'
        }).Count -eq 0) 'Legacy data became fabricated hardware.'
    }
    Test-Case 'Case-sensitive Graph dictionary attributes are projected by their actual API names' {
        Import-Module (Join-Path $project 'Modules\SmartWorkplaceCMDB.Graph\SmartWorkplaceCMDB.Graph.psd1') -Force
        # Load only the pure projection functions, never the collector entrypoint.
        $ast = [Management.Automation.Language.Parser]::ParseFile($collector, [ref]$null, [ref]$null)
        foreach ($name in @('Get-HardwareText', 'ConvertTo-HardwareRow')) {
            $node = $ast.Find({ param($item) $item -is [Management.Automation.Language.FunctionDefinitionAst] -and $item.Name -eq $name }, $true)
            . ([scriptblock]::Create($node.Extent.Text))
        }
        $device = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $device.Add('id','dictionary-id'); $device.Add('serialNumber','CASE-SERIAL')
        $device.Add('manufacturer','CASE-MANUFACTURER'); $device.Add('model','CASE-MODEL')
        $device.Add('totalStorageSpaceInBytes',[long]1024)
        $projected = @(ConvertTo-HardwareRow -Devices @($device) -CollectedDateTime '2026-01-01T00:00:00Z')
        Assert-True ($projected.Count -eq 1 -and $projected[0].SerialNumber -ceq 'CASE-SERIAL' -and $projected[0].Manufacturer -ceq 'CASE-MANUFACTURER' -and $projected[0].Model -ceq 'CASE-MODEL') 'API casing lost hardware values.'
    }
    Test-Case 'Empty response produces a header-only completed snapshot' {
        Write-Fixture @()
        $out = & $collector @identity -DataRootPath (Join-Path $temp 'Empty') -InputJsonPath $fixture
        $state = Get-Content -Raw ($out.RawLatestOutputPath + '.status.json.txt') | ConvertFrom-Json
        Assert-True ($out.DeviceCount -eq 0 -and @(Import-Csv $out.RawLatestOutputPath).Count -eq 0 -and $state.Status -ceq 'Completed' -and $state.RowCount -eq 0) 'Empty response handling failed.'
        Assert-True ((Get-Content $out.RawLatestOutputPath -TotalCount 1).Replace('"','') -ceq ($contract.tables[0].columns -join ',')) 'Empty header drift.'
    }
    Test-Case 'Write failure preserves last-valid evidence, releases lock, and retry recovers' {
        Write-Fixture $valid
        $recovery = Join-Path $temp 'Recovery'
        $out = & $collector @identity -DataRootPath $recovery -InputJsonPath $fixture
        $latest = $out.RawLatestOutputPath
        # Lock the existing latest file exclusively to simulate an unavailable destination.
        $handle = [IO.File]::Open($latest, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        try { Assert-Throw { & $collector @identity -DataRootPath $recovery -InputJsonPath $fixture } '.' }
        finally { $handle.Dispose() }
        $state = Get-Content -Raw ($latest + '.status.json.txt') | ConvertFrom-Json
        Assert-True ($state.Status -ceq 'Completed') 'Failed write did not preserve last-valid evidence.'
        $retry = & $collector @identity -DataRootPath $recovery -InputJsonPath $fixture
        $state = Get-Content -Raw ($latest + '.status.json.txt') | ConvertFrom-Json
        Assert-True ($retry.DeviceCount -eq 3 -and $state.Status -ceq 'Completed') 'Retry did not recover.'
    }
    Test-Case 'Invalid JSON envelope cannot become an empty snapshot' {
        '{"value":null}' | Set-Content $fixture -Encoding UTF8
        Assert-Throw { & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture -ValidateOnly } 'value array'
    }
    Test-Case 'Existing signed collectors and original contracts unchanged' {
        foreach ($file in $protected) { Assert-True ((Get-FileHash (Join-Path $project $file)).Hash -ceq $before[$file]) 'Existing pipeline changed.' }
    }
} finally {
    $resolved = [IO.Path]::GetFullPath($temp)
    if ($resolved.StartsWith($base + '\SmartWorkplaceCMDB-Hardware-Tests-', [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
Write-Output "Hardware offline tests: $script:passed passed; $script:failed failed."
if ($script:failed -gt 0) { throw 'Hardware offline tests failed.' }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOdDqvzxCAVc0A
# K/Vo/FZXI37ntxJe0CRjMHDlFElZWKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBFZnN8gpzVVMZO1kDJqmqu
# jZmTR7VknXsU4PFWbofxojANBgkqhkiG9w0BAQEFAASCAYCjhPEcALDi8d4tv3PI
# vi1K2HX4vLw/6c7BUr8FOWvNlJcPImL59ApBRdXiLw0DWaaZtzQc/4wMIiOaMyvl
# kzl6/Ch165/ZYsx52n74xDP4Ch1SeZb6AQwJNBtFcAtUSVQ+DSL1hbs6JmpBuUwN
# k3ld753TVhCjzUG8exjs6Ez/gOKlzhJdw4YshK4kR6g9tRNATZGlJJgSQRkpNagc
# y/RNtDgT+DWE2TPVv9r4MQjCMF5BwK5GU70h8nw+rOW6KSNpqfGkz4zVK3PROuLF
# kOcg9Rc0HmWC/ox/8kM4F9qM21/zk9DHx6P8mGLCQTG/gkxvibFjhP77dJ2aPP6s
# C9nzIqO6mgFF1CTcGrAxdzI+ycsoh79lB6Ja87pDKVDtDGPWXp7l6MXlOwQC220+
# IoJ4tAZnZaGQgM6tnR707Et+uWF3cCQdC3IX4cwOXXlLLrl+Tw1HaDKvmoeRl/+C
# S5q7tgTBWWgrUP0vyi0lhkDQw75RHYAsWpzdxB+gf+8GWe4=
# SIG # End signature block
