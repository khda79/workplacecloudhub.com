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
        $state = Get-Content -Raw ($script:result.RawLatestOutputPath + '.status.json') | ConvertFrom-Json
        $hash = (Get-FileHash $script:result.RawLatestOutputPath).Hash
        Assert-True ($state.Status -ceq 'Completed' -and $state.RowCount -eq 3 -and $state.SHA256 -ceq $hash -and $state.Coverage -ceq 'Fixture' -and $state.Mode -ceq 'Fixture') 'Invalid source evidence.'
        Assert-True ((Get-FileHash $script:result.HistoryPath).Hash -ceq $hash) 'History differs.'
        Assert-True (@($script:rows | Where-Object { $_.SourceCollectedDateTime -cne $script:result.CollectedDateTime -or $_.SourceCollectedDateTime -notmatch 'Z$' }).Count -eq 0) 'Retrieval time mismatch.'
        Assert-True ($script:result.HistoryPath -match '\\DeviceHardware\\\d{4}\\\d{2}\\') 'Incorrect history folders.'
    }
    Test-Case 'Bounded collection is isolated and labelled Bounded' {
        $out = & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture -MaxItems 1
        $state = Get-Content -Raw ($out.RawLatestOutputPath + '.status.json') | ConvertFrom-Json
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
        $stateHash = (Get-FileHash ($script:result.RawLatestOutputPath + '.status.json')).Hash
        Write-Fixture @(@{id='duplicate'}, @{id='DUPLICATE'})
        Assert-Throw { & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture } 'Duplicate managed'
        Assert-True ((Get-FileHash $script:result.RawLatestOutputPath).Hash -eq $hash -and (Get-FileHash ($script:result.RawLatestOutputPath + '.status.json')).Hash -eq $stateHash) 'Rejected fixture damaged prior evidence.'
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
        $state = Get-Content -Raw ($out.RawLatestOutputPath + '.status.json') | ConvertFrom-Json
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
        $state = Get-Content -Raw ($latest + '.status.json') | ConvertFrom-Json
        Assert-True ($state.Status -ceq 'Completed') 'Failed write did not preserve last-valid evidence.'
        $retry = & $collector @identity -DataRootPath $recovery -InputJsonPath $fixture
        $state = Get-Content -Raw ($latest + '.status.json') | ConvertFrom-Json
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAgQJL3yM7F7a8p
# bgH/Y2htZt6TB7UIQFCgqOI71rXd+aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB2HG9yE5X5WFCJi8dTW/X2hArpV7f0sisj2foNvWmV5MA0GCSqG
# SIb3DQEBAQUABIIBgAd5zW755Dk+0YE7vPasev0wEG1tIy2f4QJ5DgsotShhn/Eu
# zn+qgw/K79Kzj8FmJ5g02dRjYzTx9CURQKFXFAFMCAyS9y4MVk/jEcNkWYSQ+P4b
# tS2Rlo9F9whi0pY1FprAdFme8pvG6seyz6qveGYhp0pTUIJtsZsdHgKV8wCwobZ2
# 7j5NH3m8faEqwH9QXuSjqxjiYxvz7pKIKORcXhbZMBALXVcuF1J2/iwJms+IwV3J
# GMe4Vo6tBuNBeIN+V8GHh67z1wsSC6sFDY+nVVGbygVG7pDJPEfsPdRehQU9Z1FP
# E74xBVRUQkDu0mroEVBEdZ67a78f21Q5Fr5aCX4z+7ifQIxo7ylsMEKOTZwQUzcK
# Tg6PT+/TvTzoUBLuYD1owlIPDIv66bKYchinEBYw3jb28bwXb8z1uuJAkvZaE80G
# eTaMJORix/+bg8HYasmJOcIyTDadMJwJLC26Yuac4/a3/4wW3kCz0Lu78EZ5l1ie
# TIt91Ip6sAT0XpH6dKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzE3
# MjVaMC8GCSqGSIb3DQEJBDEiBCBzXx0yvTp6w3bNxKnL19P9zKEC+kJj+zB/fkV5
# qtl1QDANBgkqhkiG9w0BAQEFAASCAgAwBMCz8kE85QvU2/5KuPZYgcLOr87cxidi
# w2plbI8rPHBM866nJ3RH0g5t/NPdGk3uQPZbSmesVhJ+dZKgAGVUaIx06t6zDJYy
# fHKG1Gm/G3+NgnZZcsTvNZ54V6z6G0KfCHNwy18kydI6IoIEScOU16iKf3rbx/4z
# S1cZfDePB1Y+aM/WXRJmg0WIX26852gF5I5Z2bmkM2DitNk895yKCmgKxF/n4PlO
# Lo3XSguXcgITvnXr8DL4DJYINlowGgQoM1CtXJNBUV/9hRlbNXB6nb4boFBIe1HJ
# j0QZOZ2WwPrxr55cBb+bCuT26/RXsZP1LoTscL5Yd6LE30gZb0ESPyOamJCxn9NV
# xQTJbyq+h2iGX1x9GlUUhs7uXIzRCgOmct4665wMguQVarBuzjfnIheAkOgL60aP
# E3BwF44mKPU3T0KbY602F11OZFzvznMyuRVbW6+PApiVlElLWEwttk7If//RfO9C
# RQDX//yN+iNXYPrvNAihcCX30wtQX84M6YTR59wzg1swwylILjf7pjvXPhrAuAxY
# nzj9pvKaufFXBD9zEFf749Tg2ZM0hkHrCVcrOVYxYppuFm19w/9E01Emo5REHVt2
# VITLbcVqrmlR/7hjfAi32ZnafmFxWXJxyeMC+SrMuOvUZAkIKukHjspga5JOEYtv
# rqVD9XGNIA==
# SIG # End signature block
