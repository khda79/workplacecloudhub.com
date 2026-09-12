<#
.SYNOPSIS
Offline V1 regressions for direct collection isolation and source evidence.
.VERSION
1.0.0
#>
[CmdletBinding()]param()
$ScriptVersion = '1.0.0'
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$root = Join-Path ([IO.Path]::GetTempPath()) ('CMDB-Evidence-' + [guid]::NewGuid().ToString('N'))
$identity = @{ Tenant='test'; OrganizationKey='example'; EnvironmentKey='test'; TenantKey='example-test'; TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; NoConfigWrite=$true }
$script:passed=0; $script:failed=0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Case([string]$Name,[scriptblock]$Body) { try { & $Body; $script:passed++; Write-Host "PASS $Name" } catch { $script:failed++; Write-Host "FAIL $Name : $($_.Exception.Message)" } }
function Reject([scriptblock]$Body,[string]$Pattern) { try { & $Body | Out-Null } catch { if ($_.Exception.Message -notmatch $Pattern) {throw}; return }; throw "Expected rejection: $Pattern" }
$sources = @(
 @('Entra/SmartWorkplaceCMDB-EntraUsers-Collect.ps1','EntraUsers.sample.json','Raw/Entra/Entra_Users.csv'),
 @('Entra/SmartWorkplaceCMDB-EntraGroups-Collect.ps1','EntraGroups.sample.json','Raw/Entra/Entra_Groups.csv'),
 @('Entra/SmartWorkplaceCMDB-EntraDevices-Collect.ps1','EntraDevices.sample.json','Raw/Entra/Entra_Devices.csv'),
 @('Intune/SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1','IntuneManagedDevices.sample.json','Raw/Intune/Intune_ManagedDevices.csv'),
 @('M365/SmartWorkplaceCMDB-M365SubscribedSkus-Collect.ps1','M365SubscribedSkus.sample.json','Raw/M365/M365_SubscribedSkus.csv'),
 @('M365/SmartWorkplaceCMDB-M365UserLicenseAssignments-Collect.ps1','M365UserLicenseAssignments.sample.json','Raw/M365/M365_UserLicenseAssignments.csv'),
 @('ExchangeOnline/SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Collect.ps1','ExchangeOnlineMailboxes.sample.json','Raw/ExchangeOnline/ExchangeOnline_Mailboxes.csv'),
 @('ActiveDirectory/SmartWorkplaceCMDB-ActiveDirectory-Collect.ps1','ActiveDirectory.sample.json','Raw/ActiveDirectory/ActiveDirectory_Users.csv')
)
try {
 New-Item -ItemType Directory -Path $root | Out-Null
 foreach ($source in $sources) {
  Case ("Isolate direct fixture and bounded collector: " + $source[0]) {
   $data = Join-Path $root ([IO.Path]::GetFileNameWithoutExtension($source[0]))
   $latest = Join-Path $data 'DATA-LAST'
   $canonical = Join-Path $latest $source[2]
   New-Item -ItemType Directory -Path (Split-Path $canonical) -Force | Out-Null
   'canonical-sentinel' | Set-Content $canonical
   $hash = (Get-FileHash $canonical).Hash
   $r = & (Join-Path $project ('Collectors/' + $source[0])) @identity -DataRootPath $data -LatestOutputRootPath $latest -InputJsonPath (Join-Path $PSScriptRoot ('Fixtures/' + $source[1])) -MaxItems 1
   Check ((Get-FileHash $canonical).Hash -eq $hash) 'Canonical raw snapshot was overwritten.'
   Check ($r.Status -eq 'Completed') 'Isolated collection failed.'
  }
 }
 $runtime = Join-Path $root 'Evidence'
 $collector = Join-Path $project 'Collectors/Entra/SmartWorkplaceCMDB-EntraUsers-Collect.ps1'
 $normalizer = Join-Path $project 'Collectors/Entra/SmartWorkplaceCMDB-EntraUsers-Normalize.ps1'
 Case 'Empty fixture records successful non-live coverage and raw hash' {
  $empty = Join-Path $root 'empty.json'; '[]' | Set-Content $empty
  $r = & $collector @identity -DataRootPath $runtime -InputJsonPath $empty
  $script:raw = $r.RawLatestOutputPath
  $status = Get-Content ($script:raw + '.status.json') -Raw | ConvertFrom-Json
  Check ($status.Status -eq 'Completed' -and $status.Coverage -eq 'Fixture' -and $status.RowCount -eq 0) 'Empty fixture was not distinguished from live complete collection.'
  Check ($status.SHA256 -eq (Get-FileHash $script:raw).Hash) 'Source evidence does not match CSV bytes.'
 }
 Case 'Failed attempt preserves previous CSV and prevents normalization of it' {
  $r = & $collector @identity -DataRootPath $runtime -InputJsonPath (Join-Path $PSScriptRoot 'Fixtures/EntraUsers.sample.json')
  $script:raw = $r.RawLatestOutputPath
  $before = (Get-FileHash $script:raw).Hash
  $bad = Join-Path $root 'bad.json'; '{broken' | Set-Content $bad
  Reject { & $collector @identity -DataRootPath $runtime -InputJsonPath $bad } '.'
  Check ((Get-FileHash $script:raw).Hash -eq $before) 'Failure replaced the previous CSV.'
  $status = Get-Content ($script:raw + '.status.json') -Raw | ConvertFrom-Json
  Check ($status.Status -eq 'Failed') 'Failed attempt is not visible.'
  Reject { & $normalizer @identity -DataRootPath $runtime } 'source snapshot'
  Reject { & $normalizer @identity -DataRootPath $runtime -ValidateOnly } 'source snapshot'
 }
 Case 'Reject missing or malformed pages and preserve true empty arrays' {
  Assert-SmartWorkplaceCMDBCollectionPage -Response @{value=@()}
  foreach ($badPage in @(@{}, @{value=$null}, @{value='unavailable'}, @{value=@{id='wrong-shape'}})) {
   Reject { Assert-SmartWorkplaceCMDBCollectionPage -Response $badPage } 'Invalid collection page'
  }
  foreach ($json in @('[]','{"value":[]}','[{"id":"synthetic"}]','{"value":[{"id":"synthetic"}]}')) {
   $path=Join-Path $root 'page.json'; $json | Set-Content $path
   $items=@(Read-SmartWorkplaceCMDBCollectionFixture $path)
   Check ($items.Count -eq $(if ($json.Contains('synthetic')) {1} else {0})) 'Fixture array cardinality changed.'
  }
 }
 Case 'Unbounded live mode cannot reuse a fixture root' {
  $paths=Resolve-SmartWorkplaceCMDBTenantPath -Tenant test -OrganizationKey example -EnvironmentKey test -TenantId $identity.TenantId -DataRootPath $runtime
  Reject { Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $paths -ExplicitDataRoot -NoWrite } 'mode mismatch'
 }
 Case 'Raw latest overrides cannot escape fixture isolation' {
  $outside=Join-Path $root 'outside.csv'; 'keep' | Set-Content $outside
  $before=(Get-FileHash $outside).Hash
  Reject { & $collector @identity -DataRootPath (Join-Path $root 'Override') -InputJsonPath (Join-Path $PSScriptRoot 'Fixtures/EntraUsers.sample.json') -RawLatestOutputPath $outside } 'isolated latest output root'
  Check ((Get-FileHash $outside).Hash -eq $before) 'Explicit override changed another file.'
 }
 Case 'Source health distinguishes complete empty, stale, tampered and unknown' {
  $paths=Resolve-SmartWorkplaceCMDBTenantPath -Tenant test -OrganizationKey example -EnvironmentKey test -TenantId $identity.TenantId -DataRootPath (Join-Path $root 'MockLive')
  $raw=Join-Path $paths.LatestOutputRootPath 'Raw/Entra/Entra_Users.csv'
  # Simulated collector completion only: these helpers never connect to Graph.
  $run=Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($raw)
  'TenantKey,OrganizationKey,EnvironmentKey,TenantId' | Set-Content $raw
  Complete-SmartWorkplaceCMDBSourceCollection -Run $run
  $health=@(Get-SmartWorkplaceCMDBSourceHealth -Paths $paths)
  $user=$health | Where-Object SourceName -eq 'Entra_Users.csv'
  Check ($user.Status -eq 'Complete' -and $user.RowCount -eq 0) 'Mock empty completed source was lost.'
  Check (@($health | Where-Object Status -eq 'Unknown').Count -eq 11) 'Missing evidence is not unknown.'
  $state=Get-Content ($raw+'.status.json') -Raw | ConvertFrom-Json
  $state.CompletedUtc='2026-01-01T00:00:00Z'; $state | ConvertTo-Json | Set-Content ($raw+'.status.json')
  $user=Get-SmartWorkplaceCMDBSourceHealth -Paths $paths -ReferenceDateTime '2026-09-09T01:00:00Z' | Where-Object SourceName -eq 'Entra_Users.csv'
  Check ($user.Status -eq 'Stale' -and $user.Severity -eq 'Critical') 'Empty source freshness was not checked.'
  'tampered' | Add-Content $raw
  $user=Get-SmartWorkplaceCMDBSourceHealth -Paths $paths | Where-Object SourceName -eq 'Entra_Users.csv'
  Check ($user.Status -eq 'SnapshotMismatch') 'Old timestamp masked a tampered snapshot.'
  Reject { Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $raw -Paths $paths } 'source snapshot'
 }
 # JSON date materialization differs between PS5.1 and PS7. Exercise equivalent
 # UTC instants through the real sidecar reader, with exact tick boundaries.
 $datePaths=Resolve-SmartWorkplaceCMDBTenantPath -Tenant test -OrganizationKey example -EnvironmentKey test -TenantId $identity.TenantId -DataRootPath (Join-Path $root 'DateFormats')
 $dateRaw=Join-Path $datePaths.LatestOutputRootPath 'Raw/Entra/Entra_Users.csv'
 $dateRun=Start-SmartWorkplaceCMDBSourceCollection -Paths $datePaths -RawPath @($dateRaw) -Fixture -MaxItems 500
 'TenantKey,OrganizationKey,EnvironmentKey,TenantId' | Set-Content $dateRaw
 Complete-SmartWorkplaceCMDBSourceCollection -Run $dateRun
 $dateState=Get-Content ($dateRaw+'.status.json') -Raw | ConvertFrom-Json
 $reference=[datetimeoffset]::Parse('2026-09-09T20:00:00.1234567Z',[Globalization.CultureInfo]::InvariantCulture)
 $tickCases=@(
  @{Name='fresh';Ticks=[timespan]::FromHours(24).Ticks;Status='Bounded';Severity='Warning'},
  @{Name='warning-exact';Ticks=[timespan]::FromHours(48).Ticks;Status='Bounded';Severity='Warning'},
  @{Name='warning-plus-tick';Ticks=([timespan]::FromHours(48).Ticks+1);Status='Stale';Severity='Warning'},
  @{Name='critical-exact';Ticks=[timespan]::FromHours(168).Ticks;Status='Stale';Severity='Warning'},
  @{Name='critical-plus-tick';Ticks=([timespan]::FromHours(168).Ticks+1);Status='Stale';Severity='Critical'},
  @{Name='169-hours';Ticks=[timespan]::FromHours(169).Ticks;Status='Stale';Severity='Critical'},
  @{Name='future-four-minutes';Ticks=-[timespan]::FromMinutes(4).Ticks;Status='Bounded';Severity='Warning'},
  @{Name='future-exact';Ticks=-[timespan]::FromMinutes(5).Ticks;Status='Bounded';Severity='Warning'},
  @{Name='future-plus-tick';Ticks=(-[timespan]::FromMinutes(5).Ticks-1);Status='FutureDate';Severity='Warning'},
  @{Name='future-six-minutes';Ticks=-[timespan]::FromMinutes(6).Ticks;Status='FutureDate';Severity='Warning'}
 )
 $savedCulture=[Threading.Thread]::CurrentThread.CurrentCulture
 try {
  foreach($cultureName in @('en-US','fr-FR')) {
   [Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo($cultureName)
   foreach($dateFormat in @('Z','Offset00','Offset02')) {
    foreach($point in $tickCases) {
     Case ("Exact source date: $cultureName / $dateFormat / $($point.Name)") {
      $instant=$reference.AddTicks(-[long]$point.Ticks)
      $dateState.CompletedUtc=switch($dateFormat) {
       'Z' {$instant.UtcDateTime.ToString('o',[Globalization.CultureInfo]::InvariantCulture)}
       'Offset00' {$instant.ToUniversalTime().ToString('o',[Globalization.CultureInfo]::InvariantCulture)}
       'Offset02' {$instant.ToOffset([timespan]::FromHours(2)).ToString('o',[Globalization.CultureInfo]::InvariantCulture)}
      }
      $dateState | ConvertTo-Json -Depth 8 | Set-Content ($dateRaw+'.status.json')
      $h=Get-SmartWorkplaceCMDBSourceHealth -Paths $datePaths -ReferenceDateTime $reference | Where-Object SourceName -eq 'Entra_Users.csv'
      Check ($h.Status -eq $point.Status -and $h.Severity -eq $point.Severity) ("Expected $($point.Status)/$($point.Severity), got $($h.Status)/$($h.Severity)")
      $roundTrip=[datetimeoffset]::Parse($h.CompletedUtc,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal)
      Check ($roundTrip.UtcTicks -eq $instant.UtcTicks) 'Source completion instant or subsecond precision changed.'
     }
    }
   }
  }
 } finally { [Threading.Thread]::CurrentThread.CurrentCulture=$savedCulture }
 Case 'Overlapping collection attempts cannot overwrite source status' {
  $paths=Resolve-SmartWorkplaceCMDBTenantPath -DataRootPath (Join-Path $root 'Lock')
  $raw=Join-Path $paths.LatestOutputRootPath 'Raw/Entra/Entra_Users.csv'
  $run=Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($raw) -Fixture
  try {
   $before=(Get-FileHash ($raw+'.status.json')).Hash
   Reject { Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($raw) -Fixture } '.'
   Check ((Get-FileHash ($raw+'.status.json')).Hash -eq $before) 'Overlapping attempt changed source state.'
  } finally { Complete-SmartWorkplaceCMDBSourceCollection -Run $run -Failed }
 }
 Case 'Intune enrichment retains the oldest contributing date and missing date' {
  $data = Join-Path $root 'Devices'
  & (Join-Path $project 'Collectors/Entra/SmartWorkplaceCMDB-EntraDevices-Collect.ps1') @identity -DataRootPath $data -InputJsonPath (Join-Path $PSScriptRoot 'Fixtures/EntraDevices.sample.json') | Out-Null
  & (Join-Path $project 'Collectors/Intune/SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1') @identity -DataRootPath $data -InputJsonPath (Join-Path $PSScriptRoot 'Fixtures/IntuneManagedDevices.sample.json') | Out-Null
  $entra = @(Import-Csv (Join-Path $data 'DATA-LAST/Raw/Entra/Entra_Devices.csv'))
  $intune = @(Import-Csv (Join-Path $data 'DATA-LAST/Raw/Intune/Intune_ManagedDevices.csv'))
  $old = '2026-01-01T00:00:00Z'; $entra[0].SourceCollectedDateTime=$old
  $intune[0].SourceCollectedDateTime='2026-09-09T00:00:00Z'
  $ep=Join-Path $root 'entra-dates.csv'; $ip=Join-Path $root 'intune-dates.csv'
  $entra | Export-Csv $ep -NoTypeInformation; $intune | Export-Csv $ip -NoTypeInformation
  $n = Join-Path $project 'Collectors/Intune/SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1'
  & $n @identity -DataRootPath $data -EntraRawInputPath $ep -IntuneRawInputPath $ip | Out-Null
  $row = Import-Csv (Join-Path $data 'DATA-LAST/CMDB/CMDB_Devices.csv') | Where-Object SourceDeviceId -eq $entra[0].SourceDeviceId
  Check ([datetimeoffset]$row.SourceCollectedDateTime -eq [datetimeoffset]$old) 'Entra age was hidden by newer Intune data.'
  $entra[0].SourceCollectedDateTime=''; $entra | Export-Csv $ep -NoTypeInformation
  & $n @identity -DataRootPath $data -EntraRawInputPath $ep -IntuneRawInputPath $ip | Out-Null
  $row = Import-Csv (Join-Path $data 'DATA-LAST/CMDB/CMDB_Devices.csv') | Where-Object SourceDeviceId -eq $entra[0].SourceDeviceId
  Check ([string]::IsNullOrWhiteSpace($row.SourceCollectedDateTime)) 'Missing contributing date was concealed.'
 }
} finally {
 $resolved=[IO.Path]::GetFullPath($root); $allowed=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
 if ($resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -and (Test-Path $resolved)) {Remove-Item -LiteralPath $resolved -Recurse -Force}
}
Write-Host "Source evidence $ScriptVersion : Passed=$script:passed; Failed=$script:failed"
if ($script:failed) {exit 1}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA8vjIBD+F6Jr0B
# H47didDgY3SXHh0hNbvslVB42kVpAaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDgrt2ct9efywtXgCTQpiohmmniGUaCTR20StV1MKDqWMA0GCSqG
# SIb3DQEBAQUABIIBgF/RbCKwwL6nFqJxT38PxEEf2CQBm2cJQNkgjv6B8zxFxTxK
# /QtmNbo8MJigk5KcVtFo1CcwlnQLfgpgTkcGkjRtOSBaodDOsZ1KeOf6t9Y8ZYWV
# 8WpnaYIQkaunVvf99lRHja+anpobrgyl5/GnOpt6dvS9XVS/B5MLIFR20v31oczo
# rXUhamZwdgR9nMUNvh0oCa+I5n1gZlOW0b0PlfXrvY5s9ys9TmD2XYIQq/2LBkpH
# aOkSrkAV1I6s76+Wl8PbTAqMPLJBK2ZdlCcDEvQC2ptpi1Z88oZpAnxp5bzkltxh
# gjtUjdgWnwzJDlu014qIkAJYaYwzPvq19j2fkzTexmCfIP6aWatMM2FBQXxbD/ns
# PENX/hJHG/wNMM4evqdb2RjscgLxTZNSVHYZufmjo9Dirz7IpfN6NaGaNczVJLcc
# dLTlRcgfkZTQylw0Q/aUMJOCcI4FC644usmUHslHouZpYxVExQaTubuLeXrLdHIm
# lGA3lMeUTpo0HO+CiaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MjFaMC8GCSqGSIb3DQEJBDEiBCBrPqMlVGXb7D3XFa0a2gqpyl1+9RXhHr2L48FC
# l7ry9DANBgkqhkiG9w0BAQEFAASCAgBmsKSqxT6LcD7nLrd9GX0N3j8xws9KaPf/
# mNH6PRjH0SYXefMQxrqNDMtb0+pqg3nXINtAsb02lk4Isg+q2+goXH8nOBTIEz/y
# BexYLD6DcHBZlu+6cyW0J9d1SvaiyWCvc6kO7s1bngy+5T54aUZsak75463EFYCh
# AwLU67l3eS6IEdzFN8rAtLfac3ma7aEmLwUF8RxD2fFm4Xp9kZnjB8i793axNHVQ
# 7HkP7P+WbvZkEBFr2pHCXugn+rgTwU0FNfB6fUwxkTzHA5NcAL8Wjs8OATzudq+p
# StasykypTHDVfao505GMqFBa90DZdPUa1JWY/d96GJXmd/hAUDdY8KBk0/SC1Qj9
# 16lC6HqBJijQB89dJXR3Sni5G4g6KlIN2Q8MSBK0BiUHp83IVZuwhaDOl0MqTEqs
# K7ESj5XW5OlYsh2Rt6M4sGpREte6Y1qA1KzPZiAvfkdSePUAawHY9A7/2WznMbqH
# CUP1y/g1Lo1QCSiByTA07vdV7r5ymO7Y9sv1vTrlhr5XIeHdsWfcG1OEge69if2N
# XL7VtLqXoZiYrcVt3ESpPKuU/IFXupjtJ4hn/UblXFzK1FxMtb1sNKi9akEoTj8N
# 9R7gQONuLYSpueWKHbCKEkSeGcxTqyDr7f5/wTgmRw0gaaUu0EvmXLjvTjLztNMV
# +EH9Kofkiw==
# SIG # End signature block
