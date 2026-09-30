<#
.SYNOPSIS
Offline V1 regressions for direct collection isolation and source evidence.
.VERSION
1.0.1
#>
[CmdletBinding()]param()
$ScriptVersion = '1.0.1'
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
  $status = Get-Content ($script:raw + '.status.json.txt') -Raw | ConvertFrom-Json
  Check ($status.Status -eq 'Completed' -and $status.Coverage -eq 'Fixture' -and $status.RowCount -eq 0) 'Empty fixture was not distinguished from live complete collection.'
  Check ($status.SHA256 -eq (Get-FileHash $script:raw).Hash) 'Source evidence does not match CSV bytes.'
 }
 Case 'Failed attempt preserves the previous usable snapshot and evidence' {
  $r = & $collector @identity -DataRootPath $runtime -InputJsonPath (Join-Path $PSScriptRoot 'Fixtures/EntraUsers.sample.json')
  $script:raw = $r.RawLatestOutputPath
  $before = (Get-FileHash $script:raw).Hash
  $statusBefore = (Get-FileHash ($script:raw + '.status.json.txt')).Hash
  $bad = Join-Path $root 'bad.json'; '{broken' | Set-Content $bad
  Reject { & $collector @identity -DataRootPath $runtime -InputJsonPath $bad } '.'
  Check ((Get-FileHash $script:raw).Hash -eq $before) 'Failure replaced the previous CSV.'
  $status = Get-Content ($script:raw + '.status.json.txt') -Raw | ConvertFrom-Json
  Check ($status.Status -eq 'Completed') 'Previous completed evidence was not restored.'
  Check ((Get-FileHash ($script:raw + '.status.json.txt')).Hash -eq $statusBefore) 'Failure replaced the previous completed evidence.'
  & $normalizer @identity -DataRootPath $runtime | Out-Null
  & $normalizer @identity -DataRootPath $runtime -ValidateOnly | Out-Null
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
  Check (@($health | Where-Object Status -eq 'Unknown').Count -eq ($health.Count - 1)) 'Missing evidence is not unknown.'
  $state=Get-Content ($raw+'.status.json.txt') -Raw | ConvertFrom-Json
  $state.CompletedUtc='2026-01-01T00:00:00Z'; $state | ConvertTo-Json | Set-Content ($raw+'.status.json.txt')
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
 $dateState=Get-Content ($dateRaw+'.status.json.txt') -Raw | ConvertFrom-Json
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
      $dateState | ConvertTo-Json -Depth 8 | Set-Content ($dateRaw+'.status.json.txt')
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
   $before=(Get-FileHash ($raw+'.status.json.txt')).Hash
   Reject { Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($raw) -Fixture } '.'
   Check ((Get-FileHash ($raw+'.status.json.txt')).Hash -eq $before) 'Overlapping attempt changed source state.'
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAbE7pMmq/nJlcR
# ut9WMmYCp9MM9a37TSKT1yLuxZSiGaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCALp0uRdiBzR7Qq9Oy5ElBa
# mkGLMNog85fLuYKPLesvxDANBgkqhkiG9w0BAQEFAASCAYCJYhciRd3ep+KXWuRA
# n+t4ml8iHcMVfFMBPQe+/JJDAqHozQZdSadT0qZE9Rj/aM65PhlwJGw7SY7udYIk
# aUnOGYoacO/qff2sZkJfQm/SsyWPRDlvGBoffowandDbAWmKQpF17X0tx1Zs0r+i
# hLhzakAUXTPj548IY7PRKoFqcB7xYVbYCuuDbPn6p6UKp3y52dcXPDcKMEMGXeiw
# oK5qNWpQff9RDidhO/AM5U+i5movfabaH9OHP9jNvcHbts/S24tjN+V/hYwxFBxe
# zFROisN6tGB4f98XFvDEH0/+Ivwl97b5RA89kXtUELCkyQVxijHf4f6BpY3jNiG2
# BNMMgcdy4o24qC0AZv01+AzcRc8K4mwrtTwYrFXLBcAYGDeLQjtOFvRK7BsZV9Xr
# +zbe+H1Aru5T3U6ihzdUa9TsoGWcUb7RZdntyRfkJoQkZaurxDsVSN8jpV7zHLlm
# FyfaY1jF29aV03pUJDSPhZxiLhYQ2NFst0jMZSRDxIeWWzg=
# SIG # End signature block
