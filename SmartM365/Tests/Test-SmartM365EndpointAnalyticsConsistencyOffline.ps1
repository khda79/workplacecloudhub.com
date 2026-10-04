#Requires -Version 7.0
<#
.SYNOPSIS
Offline Endpoint Analytics fresh-export consistency and publication boundary tests.
.VERSION
1.0.3
.NOTES
Loads AST functions only; all export jobs, downloads, delays and publication are mocked.
Does not import Core, read tenant configuration or execute a collector entrypoint.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Start-Sleep is mocked in an isolated module so tests never wait.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='State-changing names mirror production helpers but only mutate synthetic in-memory lists.')]
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$path = Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartInventory/M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Endpoint Analytics parse failed.' }
$definitions = foreach ($name in @('Get-EAReportCatalog','Get-EAOutputSchemas','Get-EAStatusCode','Get-EARawValue','New-EANormalizedRow','Assert-EADeviceReportGrain','Get-EARejectedRowDiagnostic','Save-EARejectedRowDiagnostic','Get-EAScoreExclusion','Invoke-EAConsistentReport','Invoke-EAReport','Publish-EAOutputs')) {
    $node = $ast.Find({ param($item) $item -is [Management.Automation.Language.FunctionDefinitionAst] -and $item.Name -eq $name }, $true)
    if (-not $node) { throw "Missing production function: $name" }
    $node.Extent.Text
}
$module = New-Module -Name SyntheticEndpointConsistency -ScriptBlock {
    param($Definitions)
    Set-StrictMode -Version Latest
    foreach ($definition in $Definitions) { . ([scriptblock]::Create($definition)) }
    $script:Checks = 0
    function Assert-Test {
        param([bool]$Condition,[string]$Message)
        $script:Checks++
        if (-not $Condition) { throw $Message }
    }
    function Get-TestFailure {
        param([scriptblock]$Action)
        try { & $Action | Out-Null } catch { return $_ }
        throw 'Expected synthetic failure was not raised.'
    }
    function Reset-TestState {
        param([object[]]$Exports)
        $script:Exports = $Exports
        $script:JobNames = [Collections.Generic.List[string]]::new()
        $script:JobIds = [Collections.Generic.List[string]]::new()
        $script:Delays = [Collections.Generic.List[int]]::new()
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:Published = [Collections.Generic.List[object]]::new()
        $script:ImportCount = 0
        $script:FailureStatus = 0
        $script:CoreImported = $false
        $script:ReportConsistencyAttempts = 3
        $script:ReportConsistencyRetryDelaySeconds = 15
        $script:RunId = 'synthetic-run'
        $script:OutputPath = 'synthetic-history'
        $script:LatestCsvFolderPath = 'synthetic-current'
        $script:IncludeStartupProcesses = $true
    }
    function Start-EAExportJob {
        param($Report,$EffectiveReportName)
        $script:JobNames.Add($EffectiveReportName)
        if ($script:FailureStatus -and $EffectiveReportName -eq $Report.Name) {
            $failure = [Exception]::new('Synthetic HTTP failure')
            $failure.Data['StatusCode'] = $script:FailureStatus
            throw $failure
        }
        $id = 'new-export-{0}' -f $script:JobNames.Count
        $script:JobIds.Add($id)
        [pscustomobject]@{id=$id}
    }
    function Wait-EAExportJob {
        param($JobId,$ApiVersion)
        if ($ApiVersion -ne 'beta') { throw 'Unexpected API version.' }
        [pscustomobject]@{id=$JobId;status='completed';url='synthetic-only'}
    }
    function Import-EAExportedCsv {
        param($CompletedJob,$ReportName)
        if ($ReportName -ne $script:JobNames[$script:JobNames.Count-1]) { throw 'Download report did not match the new job.' }
        $script:ImportCount++
        if ($CompletedJob.id -ne $script:JobIds[$script:ImportCount-1]) { throw 'Export download did not match the fresh job.' }
        return @($script:Exports[[Math]::Min($script:ImportCount-1,$script:Exports.Count-1)].Rows)
    }
    function Write-EALog {
        param($Message,$Level)
        if ($Level -ne 'WARNING') { throw 'Expected a warning for a rejected export.' }
        $script:Logs.Add([string]$Message)
    }
    function Start-Sleep { param([int]$Seconds) $script:Delays.Add($Seconds) }
    function Publish-CoreSmartM365Csv {
        param($Data,$TimestampedPath,$LatestPath,$Columns)
        $script:Published.Add([pscustomobject]@{Data=@($Data);Columns=@($Columns);Latest=$LatestPath;History=$TimestampedPath})
    }
    $catalog = @(Get-EAReportCatalog)
    $scores = @($catalog | Where-Object Name -eq 'EADeviceScoresV2')[0]
    Assert-Test (@(@('DeviceScopeIds','HealthStatus','PartnerFeaturesBitmask') | Where-Object { $_ -notin $scores.Select }).Count -eq 0) 'Raw score diagnostics omitted scope/health/feature metadata.'
    $schemas = Get-EAOutputSchemas
    Assert-Test (($schemas.DevicePerformance -join ',') -ceq 'RunId,ReportName,ReportRefreshDate,DeviceId,DeviceName,Manufacturer,Model,EndpointAnalyticsScore,StartupScore,AppReliabilityScore,WorkFromAnywhereScore,CrashCount,MeanTimeToFailure') 'Diagnostic metadata changed the canonical device score schema.'
    Reset-TestState @()
    $normalized = New-EANormalizedRow EADeviceScoresV2 ([pscustomobject]@{DeviceId='synthetic';EndpointAnalyticsScore=80;DeviceScopeIds='scope';HealthStatus='health';PartnerFeaturesBitmask=3})
    Assert-Test (@(@('DeviceScopeIds','HealthStatus','PartnerFeaturesBitmask') | Where-Object { $_ -in $normalized.PSObject.Properties.Name }).Count -eq 0) 'Raw diagnostic metadata entered canonical rows.'
    $wfa = @($catalog | Where-Object Name -eq 'EAWFADeviceList')[0]
    $models = @($catalog | Where-Object Name -eq 'EAModelScoresV2')[0]
    $conflict = @(
        [pscustomobject]@{DeviceId='synthetic-secret-device';EndpointAnalyticsScore=71},
        [pscustomobject]@{DeviceId='synthetic-secret-device';EndpointAnalyticsScore=88}
    )
    $clean = @(
        [pscustomobject]@{DeviceId='synthetic-clean-one';EndpointAnalyticsScore=73},
        [pscustomobject]@{DeviceId='synthetic-clean-two';EndpointAnalyticsScore=92}
    )
    Reset-TestState @([pscustomobject]@{Rows=$conflict},[pscustomobject]@{Rows=$clean})
    $result = Invoke-EAReport $scores
    Assert-Test ($script:JobIds.Count -eq 2 -and @($script:JobIds | Sort-Object -Unique).Count -eq 2) 'Retry did not create a fresh export job.'
    Assert-Test ($script:ImportCount -eq 2 -and $script:Delays.Count -eq 1 -and $script:Delays[0] -eq 15) 'Retry delay/download contract changed.'
    Assert-Test ($result.Rows.Count -eq 2 -and $result.Rows[0].EndpointAnalyticsScore -eq 73 -and $result.Rows[1].EndpointAnalyticsScore -eq 92) 'Clean export rows were changed or mixed with the rejected attempt.'
    Assert-Test (($script:Logs -join ' ') -notmatch 'synthetic-secret-device|\b71\b|\b88\b') 'Conflict diagnostics exposed row values.'

    Reset-TestState @([pscustomobject]@{Rows=$conflict})
    $failure = Get-TestFailure { Invoke-EAReport $wfa }
    Assert-Test ([bool]$failure.Exception.Data['EndpointAnalyticsGrain']) 'Persistent grain failure lost its classification.'
    Assert-Test ($script:JobNames.Count -eq 3 -and @($script:JobNames | Where-Object { $_ -ne $wfa.Name }).Count -eq 0) 'A persistent conflict escaped through alias fallback or exceeded the bound.'
    Assert-Test (($script:Delays -join ',') -eq '15,30') 'Expected bounded progressive delays.'
    Assert-Test ($script:Published.Count -eq 0) 'A conflicting export was published.'
    foreach ($count in @(400,404)) {
        $repeated=@(1..($count+1) | ForEach-Object { New-EANormalizedRow EADeviceScoresV2 $conflict[0] })
        $failure=Get-TestFailure { Assert-EADeviceReportGrain -Rows $repeated -ReportName EADeviceScoresV2 }
        Assert-Test ((Get-EAStatusCode $failure) -eq 0) 'A duplicate count was confused with HTTP 400/404 and could bypass the terminal failure boundary.'
    }
    Reset-TestState @([pscustomobject]@{Rows=@([pscustomobject]@{DeviceId=' ';EndpointAnalyticsScore=73})})
    $script:ReportConsistencyAttempts = 1
    $failure = Get-TestFailure { Invoke-EAReport $scores }
    Assert-Test ($script:JobNames.Count -eq 1 -and $script:Delays.Count -eq 0) 'Invalid identity was accepted or ignored the one-attempt setting.'
    foreach ($rows in @(
        [pscustomobject]@{Rows=@($conflict[0],$conflict[0])},
        [pscustomobject]@{Rows=@($conflict[0],[pscustomobject]@{DeviceId=' SYNTHETIC-SECRET-DEVICE ';EndpointAnalyticsScore=73})}
    )) {
        Reset-TestState @($rows)
        $script:ReportConsistencyAttempts = 1
        $result = Invoke-EAReport $scores
        Assert-Test ($result.Rows.Count -eq 0 -and $result.Exclusion.ExcludedRowCount -eq 2 -and $result.Exclusion.ExcludedDeviceCount -eq 1) 'All repeated score rows were not excluded.'
    }
    Reset-TestState @([pscustomobject]@{Rows=@($conflict + $clean)})
    $result = Invoke-EAReport $scores
    Assert-Test ($script:JobNames.Count -eq 3 -and $result.RawRowCount -eq 4 -and $result.Rows.Count -eq 2) 'Persistent score conflict did not exhaust fresh retries or lost clean rows.'
    Assert-Test ($result.Exclusion.ExcludedRowCount -eq 2 -and $result.Exclusion.Diagnostic.DuplicateGroups[0].RawRows.Count -eq 2) 'Exclusion counts or raw evidence changed.'
    Assert-Test (($result.Rows.DeviceId -join ',') -notmatch 'synthetic-secret-device' -and ($script:Logs -join ' ') -notmatch 'synthetic-secret-device|\b71\b|\b88\b') 'Excluded identity leaked into output or console.'
    $multiple=@($conflict+$clean+@(
        [pscustomobject]@{DeviceId='other-conflict';EndpointAnalyticsScore=80},
        [pscustomobject]@{DeviceId=' OTHER-CONFLICT ';EndpointAnalyticsScore=80}
    ))
    Reset-TestState @([pscustomobject]@{Rows=$multiple})
    $result=Invoke-EAReport $scores
    Assert-Test ($result.RawRowCount -eq 6 -and $result.Exclusion.ExcludedRowCount -eq 4 -and $result.Exclusion.ExcludedDeviceCount -eq 2 -and $result.Rows.Count -eq 2) 'Multiple conflicting/identical duplicate groups were not fully excluded and reconciled.'
    Assert-Test (($result.Rows.DeviceId -join ',') -eq 'synthetic-clean-one,synthetic-clean-two') 'Multiple-group exclusion changed valid device identity/order.'
    Reset-TestState @([pscustomobject]@{Rows=@($conflict+@([pscustomobject]@{DeviceId='';EndpointAnalyticsScore=80}))})
    $failure=Get-TestFailure { Invoke-EAReport $scores }
    Assert-Test ($script:JobIds.Count -eq 3 -and $script:Published.Count -eq 0) 'Duplicate-score exclusion bypassed an invalid identity.'
    Reset-TestState @([pscustomobject]@{Rows=$conflict},[pscustomobject]@{Rows=$conflict},[pscustomobject]@{Rows=$clean})
    $script:ReportConsistencyRetryDelaySeconds = 300
    $result = Invoke-EAReport $scores
    Assert-Test (($script:Delays -join ',') -eq '300,300' -and $result.Rows.Count -eq 2) 'Delay cap or third-attempt success failed.'

    Reset-TestState @([pscustomobject]@{Rows=@()})
    $result = Invoke-EAReport $scores
    Assert-Test ($result.Rows.Count -eq 0 -and $script:ImportCount -eq 1) 'A complete empty export was rejected.'
    Reset-TestState @([pscustomobject]@{Rows=$conflict})
    $result = Invoke-EAReport $scores -AvailabilityOnly
    Assert-Test ($result.Rows.Count -eq 0 -and $script:ImportCount -eq 0 -and $script:JobNames.Count -eq 1) 'Availability check downloaded or qualified CSV contents.'
    Reset-TestState @([pscustomobject]@{Rows=@([pscustomobject]@{Model='test';EndpointAnalyticsScore=80})})
    $result = Invoke-EAReport $models
    Assert-Test ($result.Rows.Count -eq 1 -and $script:JobNames.Count -eq 1) 'Device identity requirement was incorrectly applied to a model report.'
    Reset-TestState @([pscustomobject]@{Rows=$clean})
    $script:FailureStatus = 403
    $failure = Get-TestFailure { Invoke-EAReport $wfa }
    Assert-Test ((Get-EAStatusCode $failure) -eq 403 -and $script:JobNames.Count -eq 1 -and $script:ImportCount -eq 0) 'Permission failure entered consistency retries or alias fallback.'
    Reset-TestState @([pscustomobject]@{Rows=$clean})
    $script:FailureStatus = 400
    $result = Invoke-EAReport $wfa
    Assert-Test ($result.AliasUsed -and $script:JobNames.Count -eq 2 -and $result.Rows.Count -eq 2 -and $script:Delays.Count -eq 0) 'Documented alias fallback was degraded.'

    Reset-TestState @()
    $goodPerformance = @(New-EANormalizedRow EADeviceScoresV2 $clean[0]; New-EANormalizedRow EADevicePerformanceV2 $clean[0])
    $badStartup = @(New-EANormalizedRow EAStartupPerfDevicePerformanceV2 $conflict[0]; New-EANormalizedRow EAStartupPerfDevicePerformanceV2 $conflict[1])
    $failure = Get-TestFailure { Publish-EAOutputs -OutputRows ([ordered]@{DevicePerformance=$goodPerformance;StartupDevices=$badStartup}) -Schemas (Get-EAOutputSchemas) }
    Assert-Test ($script:Published.Count -eq 0 -and [bool]$failure.Exception.Data['EndpointAnalyticsGrain']) 'The final guard wrote an earlier canonical CSV before detecting the later conflict.'
    $goodWfa = @(New-EANormalizedRow EAWFADeviceList $clean[0]; New-EANormalizedRow EAWFAPerDevicePerformance $clean[0]; New-EANormalizedRow EAWFAModelPerformance ([pscustomobject]@{Model='model';WorkFromAnywhereScore=85}))
    Publish-EAOutputs -OutputRows ([ordered]@{DevicePerformance=$goodPerformance;WorkFromAnywhere=$goodWfa}) -Schemas (Get-EAOutputSchemas)
    Assert-Test ($script:Published.Count -eq 9) 'Healthy output publication contract changed.'
    Assert-Test ($script:Published[0].Data.Count -eq 2 -and $script:Published[7].Data.Count -eq 3) 'Different reports or WFA model evidence were removed.'
    Reset-TestState @()
    $failure = Get-TestFailure { Publish-EAOutputs -OutputRows ([ordered]@{DevicePerformance=@([pscustomobject]@{ReportName='unknown';DeviceId='synthetic'})}) -Schemas (Get-EAOutputSchemas) }
    Assert-Test ($script:Published.Count -eq 0) 'Unknown report output was published.'
} -ArgumentList (,$definitions)
try { & $module { [pscustomobject]@{Status='Passed';Checks=$script:Checks;LiveCalls=0;RealDelays=0;CanonicalWrites=0} } }
finally { Remove-Module $module -ErrorAction SilentlyContinue }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB/dFhOmOASdheh
# EMsF0Pw1CkEgH6NULKUMHcMPMfY2zqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIO3XBbq0EFjOPbq5OxsX9L9kXcf/BehcsxR2JPXNBLRjMA0GCSqG
# SIb3DQEBAQUABIIBgEMUIBJifNluAiD6cNiyPAlknzXBTUjyMeaPyvIHcFIItMVd
# XZAC61G0zscUmMMT1zkoluf8rNXVabUWWGGDa7itJGIX8YppkAsVn8cSXfdpCUa4
# 3u8SQE7CCnk78/vAnTJpjdnpixOeqZdQFfNJGrmj3ngFzwoOG7f2pGTSIp1TyaxT
# hjkzrv1FaYJv5sdAv8dja9feC4rGKq9/7vrgHFvJvSnIElD9WUUNs4bPvaL3IZEC
# CY1+naEvDLTeR8WSlJ5JUMmGYgX6bYv7y1m54Wn88FSQKgkxyEHy5+Ole8kkhViU
# QkLQeh1aSctB8PNQJLQ7Sm2QcgBs03ZRqi/Z+TQU2BpqCNSxdJXP9Sr+Z/XFw87I
# D4IswpzdiB7fImW06FooPCrpbngQdoILyrZFjNwG9mlvnVWqRWmL4qd87ASyOKqO
# XwGb8EGmwXTF+SP/zWTPaTf5oMqHGtLejwobuEW36eSUqXXaSjYH566R7CXRoDfB
# VWrFJKoEi4yMdQE50KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQyMTAx
# NTZaMC8GCSqGSIb3DQEJBDEiBCDeYgqTRszanmFaL22+xhJc2hnzPJ31xDz6FAeC
# pMuQvzANBgkqhkiG9w0BAQEFAASCAgBF1FV6i6yKALIGssjKeR15IJWRUp/+Jc+z
# rS8pGwac4uUgHKeVqSgV8wTEj5cNxJG/V4z4sFrdOJlWtMoSxBMoqp33S+Eov5zq
# kKjyzrG2zYsXp+fXG+XBHP7ICz37Uks6vzMPIbzDbITe1MyLDfblQJOWXDr15f+D
# EyXR8fLfFo3S0UTJH6feuZ6cgLzrJz77IKtzgHXaGr1pY27NtayiKek6Dfdi1DnF
# yD5W7j9IoxjwBG102ekKRWGxUiZRAAqp/cjHrpNj4v11qQ7WCuAawdDAAIG1CXCW
# QzQKd4rvwqCSpt2k7P6/oMlwIvNuJtPzawFqF2JyL3psFYlscnWu6RnX6ifQ2gSN
# XC04cx+9McM5O3W7vlmkoh8F0R4BOCv6/8ijCg9ibVYhpLYkxOMNQD5X9Z5gQULZ
# skQPfIK1+kPXOP6+pkLwRx4e3KkDbZf5/YDozYl1m6ovZB+DDqnpkyNHcErdJDDg
# A5z5NMCmXcMic4Rf5DcospeY4/LWQ1iX/6CsjeztTJUUlALVYMso+mhnQpdkOBgk
# ZXXEUsfhEiCoUqX3RwEI5yOGbDs70o/TLI1YvWDc68Cl7gEC9i7C9xBBUjvCq50v
# NpxW5NdDG5wrdXamfjW/fWGnW8Vw8DKv4B2eR7dgLBeFPhGm41j9Kmf3nLSKz/99
# 5lXedNwijg==
# SIG # End signature block
