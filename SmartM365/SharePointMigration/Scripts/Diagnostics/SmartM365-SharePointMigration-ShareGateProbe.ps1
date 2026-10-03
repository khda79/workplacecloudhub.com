<#
.SYNOPSIS
    Read-only ShareGate capability, session and access probe. DryRun is the default.
.VERSION
    1.0.5
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [string]$AnalysisDirectory = '',
    [string]$SessionId = '',
    [ValidateSet('Default','Browser','ModernAuth')][string]$SourceAuthMode = 'Default',
    [switch]$ProbeSource,
    [switch]$DryRun,
    [switch]$Run
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Microsoft.PowerShell.Utility\Write-Host ('{0} Script  : {1} v{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MyInvocation.MyCommand.Name, '1.0.5') -ForegroundColor Cyan
if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
$project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
if (-not (Test-Path -LiteralPath $project -PathType Container)) { throw 'ProjectRoot must be a directory.' }
$diagnosticsRoot = Join-Path $project 'ShareGate\Diagnostics'
if (-not $AnalysisDirectory) {
    if (-not (Test-Path -LiteralPath $diagnosticsRoot -PathType Container)) {
        throw "No phase 2a analysis exists at '$diagnosticsRoot'. Run SmartM365-SharePointMigration-Diagnostics.ps1 for this ProjectRoot first; generated Diagnostics folders are local and are not transferred by Git."
    }
    $latest = Get-ChildItem -LiteralPath $diagnosticsRoot -Directory -ErrorAction Stop |
        Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'ClassifiedRows.csv') -PathType Leaf } |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    if (-not $latest) { throw "No phase 2a ClassifiedRows.csv was found under '$diagnosticsRoot'. Run SmartM365-SharePointMigration-Diagnostics.ps1 for this ProjectRoot first." }
    $AnalysisDirectory = $latest.FullName
}
$analysis = (Resolve-Path -LiteralPath $AnalysisDirectory -ErrorAction Stop).ProviderPath
$classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
if (-not (Test-Path -LiteralPath $classifiedPath -PathType Leaf)) { throw 'ClassifiedRows.csv is missing.' }
$runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$output = Join-Path $diagnosticsRoot ('Probe-' + $runId)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$logPath = Join-Path $output 'Probe.log'
$results = New-Object 'System.Collections.Generic.List[object]'
$parameters = New-Object 'System.Collections.Generic.List[object]'
$properties = New-Object 'System.Collections.Generic.List[object]'
$access = New-Object 'System.Collections.Generic.List[object]'
$siteCache = @{}

function Write-ProbeLog {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Output $line
}
function Add-ProbeResult {
    param([string]$Area, [string]$Name, [string]$Status, [string]$Detail)
    $results.Add([pscustomobject]@{ Area=$Area; Name=$Name; Status=$Status; Detail=$Detail })
    Write-ProbeLog "$Area | $Name | $Status | $Detail"
}
function Export-AtomicCsv {
    param([string]$Name, [object[]]$Rows, [string[]]$Columns)
    $path = Join-Path $output $Name
    $temp = Join-Path $output ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        if ($Rows.Count) { $Rows | Select-Object -Property $Columns | Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding UTF8 }
        else {
            $header = ($Columns | ForEach-Object { '"' + $_.Replace('"','""') + '"' }) -join ','
            Set-Content -LiteralPath $temp -Value $header -Encoding UTF8
        }
        Move-Item -LiteralPath $temp -Destination $path -ErrorAction Stop
    }
    finally { if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force } }
}
function Get-EndpointKey {
    param([string]$Side, [string]$Url, [string]$List)
    return ($Side + '|' + $Url.TrimEnd('/').ToLowerInvariant() + '|' + $List.ToLowerInvariant())
}
function Test-Endpoint {
    param([string]$Side, [string]$Url, [string]$List)
    if (-not $Url) { return [pscustomobject]@{ Status='Missing URL'; Error='' } }
    try {
        $authMode = if ($Side -eq 'Source') { $SourceAuthMode } else { 'Browser' }
        $arguments = @{ Url=$Url; ErrorAction='Stop' }
        if ($authMode -eq 'Browser') { $arguments.Browser = $true }
        if ($authMode -eq 'ModernAuth') { $arguments.ModernAuth = $true }
        $siteKey = $Side + '|' + $Url.TrimEnd('/').ToLowerInvariant()
        if ($siteCache.ContainsKey($siteKey)) {
            $cached = $siteCache[$siteKey]
            if ($cached.Error) { return $cached.Error }
            $site = $cached.Site
        }
        else {
            try {
                $site = Connect-Site @arguments
                if (-not $site) { throw 'Connect-Site returned no site.' }
                $siteCache[$siteKey] = [pscustomobject]@{ Site=$site; Error=$null }
            }
            catch {
                $message = $_.Exception.Message
                $status = if ($message -match '(?i)access.denied|unauthori[sz]ed|forbidden|403|refus.d.acc.s|not authorized') { 'Access denied' } else { 'Read error' }
                $failure = [pscustomobject]@{ Status=$status; Error=$message }
                $siteCache[$siteKey] = [pscustomobject]@{ Site=$null; Error=$failure }
                return $failure
            }
        }
        if ($List) {
            $matches = @(Get-List -Site $site -Name $List -ErrorAction Stop)
            $exact = @($matches | Where-Object {
                ($_.PSObject.Properties['Title'] -and $_.Title -eq $List) -or
                ($_.PSObject.Properties['Name'] -and $_.Name -eq $List)
            })
            if ($exact.Count -eq 0) { return [pscustomobject]@{ Status='List not found'; Error="Get-List did not return an exact match for '$List'." } }
        }
        return [pscustomobject]@{ Status='Readable'; Error='' }
    }
    catch {
        $message = $_.Exception.Message
        $status = if ($message -match '(?i)access.denied|unauthori[sz]ed|forbidden|403|refus.d.acc.s|not authorized') { 'Access denied' } else { 'Read error' }
        return [pscustomobject]@{ Status=$status; Error=$message }
    }
}

try {
    Write-ProbeLog ('Mode={0}; Actor={1}\{2}; Machine={3}; Session={4}; Analysis={5}; SourceProbe={6}; DestinationAuth=Browser' -f $(if ($Run) { 'Run' } else { 'DryRun' }), $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME, $(if ($SessionId) { $SessionId } else { 'all' }), $analysis, $(if ($ProbeSource) { 'Enabled' } else { 'Skipped' }))
    $allRows = @(Import-Csv -LiteralPath $classifiedPath -Encoding UTF8)
    $accessRows = @($allRows | Where-Object {
        $_.RuleId -in @('SG-ACCESS-SOURCE','SG-ACCESS-TARGET','SG-ACCESS-UNKNOWN') -and
        (-not $SessionId -or $_.SessionId -eq $SessionId)
    })
    $sessionIds = @($allRows | ForEach-Object SessionId | Where-Object { $_ } | Sort-Object -Unique)
    Add-ProbeResult -Area 'Input' -Name 'Session IDs' -Status 'Detected' -Detail ($sessionIds -join ', ')
    if ($SessionId -and $SessionId -notin $sessionIds) { throw "Session ID '$SessionId' is absent from the selected phase 2a analysis." }
    Add-ProbeResult -Area 'Input' -Name 'Access rows' -Status 'Detected' -Detail ([string]$accessRows.Count)
    $endpoints = @{}
    foreach ($row in $accessRows) {
        foreach ($side in @('Source','Destination')) {
            $url = if ($side -eq 'Source') { [string]$row.SourceUrl } else { [string]$row.DestinationUrl }
            $list = if ($side -eq 'Source') { [string]$row.SourceList } else { [string]$row.DestinationList }
            $key = Get-EndpointKey -Side $side -Url $url -List $list
            if (-not $endpoints.ContainsKey($key)) { $endpoints[$key] = [pscustomobject]@{ Side=$side; Url=$url; List=$list; Result=$null } }
        }
    }
    Add-ProbeResult -Area 'Input' -Name 'Distinct site/list endpoints' -Status 'Detected' -Detail ([string]$endpoints.Count)
    if (-not $ProbeSource) { Add-ProbeResult -Area 'Scope' -Name 'Source site checks' -Status 'Skipped' -Detail 'Use -ProbeSource on a host with access to the source farm to enable read-only source checks.' }
    else { Add-ProbeResult -Area 'Scope' -Name 'Source authentication' -Status $SourceAuthMode -Detail 'Default uses the current Windows identity; no username or password is supplied.' }
    Add-ProbeResult -Area 'Scope' -Name 'Destination authentication' -Status 'Browser only' -Detail 'Connect-Site uses -Browser without username, password, or SaveConnection.'
    $installedApps = @(foreach ($registryRoot in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
        Get-ItemProperty -Path (Join-Path $registryRoot '*') -ErrorAction SilentlyContinue |
            Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.PSObject.Properties['DisplayVersion'] -and $_.DisplayName -match '^ShareGate( Migrate)?$' -and $_.DisplayVersion }
    })
    if ($installedApps.Count) { Add-ProbeResult -Area 'ShareGate' -Name 'Application version' -Status 'Detected' -Detail (($installedApps | ForEach-Object DisplayVersion | Sort-Object -Unique) -join ', ') }
    else { Add-ProbeResult -Area 'ShareGate' -Name 'Application version' -Status 'Undetermined' -Detail 'No ShareGate Migrate entry found in the machine uninstall registry.' }
    $moduleInfo = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
    if (-not $moduleInfo.Count) {
        foreach ($app in $installedApps) {
            if (-not $app.PSObject.Properties['InstallLocation'] -or -not $app.InstallLocation) { continue }
            $candidate = Join-Path ([string]$app.InstallLocation) 'ShareGate\Sharegate.psd1'
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                try {
                    $manifest = Test-ModuleManifest -Path $candidate -ErrorAction Stop -WarningAction SilentlyContinue
                    $moduleInfo = @([pscustomobject]@{ Version=$manifest.Version; Path=$candidate })
                    break
                }
                catch { Add-ProbeResult -Area 'ShareGate' -Name 'Module manifest' -Status 'Error' -Detail $_.Exception.Message }
            }
        }
    }
    if ($moduleInfo.Count) { Add-ProbeResult -Area 'ShareGate' -Name 'Module' -Status 'Present' -Detail ("Version={0}; Path={1}" -f $moduleInfo[0].Version, $moduleInfo[0].Path) }
    else { Add-ProbeResult -Area 'ShareGate' -Name 'Module' -Status 'Missing' -Detail 'ShareGate module is not installed or discoverable by Windows PowerShell 5.1.' }
    if (-not $Run) {
        Add-ProbeResult -Area 'ShareGate' -Name 'License' -Status 'Not tested' -Detail 'DryRun does not import ShareGate or invoke any ShareGate command.'
        foreach ($row in $accessRows) {
            $access.Add([pscustomobject]@{ SessionId=$row.SessionId; RowId=$row.RowId; SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; DestinationUrl=$row.DestinationUrl; DestinationList=$row.DestinationList; SourceStatus=$(if ($ProbeSource) { 'Planned' } else { 'Skipped - source unreachable' }); DestinationStatus='Planned - Browser'; Attribution='Undetermined'; Reason='DryRun: no site connection'; SourceError=''; DestinationError='' })
        }
    }
    elseif (-not $moduleInfo.Count) {
        Add-ProbeResult -Area 'ShareGate' -Name 'License' -Status 'Undetermined' -Detail 'Module missing; live checks skipped.'
        throw 'ShareGate module is unavailable for the real probe.'
    }
    else {
        $shareGateImported = $false
        try {
            Import-Module -Name $moduleInfo[0].Path -ErrorAction Stop
            $shareGateImported = $true
            Add-ProbeResult -Area 'ShareGate' -Name 'Module import' -Status 'Succeeded' -Detail ("Imported version {0}" -f (Get-Module ShareGate | Select-Object -First 1).Version)
        }
        catch { Add-ProbeResult -Area 'ShareGate' -Name 'Module import' -Status 'Error' -Detail $_.Exception.Message }
        if (-not $shareGateImported) { throw 'ShareGate module import failed; live checks cannot continue.' }
        $cmdletNames = @('Find-CopySessions','Export-Report','Connect-Site','Get-List','Copy-Content','New-CopySettings','Wait-ImportCompletion')
        if ($shareGateImported) {
            $installedCommands = @(Get-Command -Module ShareGate -ErrorAction SilentlyContinue)
            Add-ProbeResult -Area 'ShareGate' -Name 'Installed cmdlets' -Status 'Detected' -Detail ([string]$installedCommands.Count)
            $cmdletNames += @($installedCommands | Where-Object { $_.Name -match 'Mapping' } | Select-Object -ExpandProperty Name)
            foreach ($name in @($cmdletNames | Sort-Object -Unique)) {
                $command = Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue | Select-Object -First 1
                if (-not $command) { Add-ProbeResult -Area 'Cmdlet' -Name $name -Status 'Missing' -Detail 'Not installed in this ShareGate module.'; continue }
                Add-ProbeResult -Area 'Cmdlet' -Name $name -Status 'Present' -Detail ([string]$command.CommandType)
                foreach ($set in $command.ParameterSets) {
                    foreach ($parameter in $set.Parameters) {
                        $parameters.Add([pscustomobject]@{ Cmdlet=$name; ParameterSet=$set.Name; Parameter=$parameter.Name; Type=$parameter.ParameterType.FullName; Mandatory=$parameter.IsMandatory })
                    }
                }
            }
            $find = Get-Command Find-CopySessions -Module ShareGate -ErrorAction SilentlyContinue
            $licenseConfirmed = $false
            if ($find) {
                try {
                    $query = @{}
                    if ($SessionId -match '^(\d{2})(\d{2})(\d{2})-') {
                        $date = [datetime]::ParseExact($Matches[1] + $Matches[2] + $Matches[3], 'yyMMdd', [Globalization.CultureInfo]::InvariantCulture)
                        $query.From = $date.AddDays(-1); $query.To = $date.AddDays(2)
                    }
                    $sessionObjects = @(Find-CopySessions @query -ErrorAction Stop)
                    $licenseConfirmed = $true
                    Add-ProbeResult -Area 'ShareGate' -Name 'License' -Status 'Pro or Enterprise inferred' -Detail 'Find-CopySessions succeeded. Exact subscription tier is not exposed by this probe.'
                    if (-not $sessionObjects.Count) {
                        Add-ProbeResult -Area 'Session' -Name $SessionId -Status 'No local history' -Detail 'Find-CopySessions succeeded but this machine has no session objects in the requested period.'
                    }
                    else {
                        $selected = @($sessionObjects | Where-Object {
                            $candidate = $_
                            @($candidate.PSObject.Properties | Where-Object { $_.Name -in @('SessionId','SessionID','Id','ID') -and [string]$_.Value -eq $SessionId }).Count -gt 0
                        } | Select-Object -First 1)
                        if (-not $SessionId) { $selected = @($sessionObjects | Select-Object -First 1) }
                        if ($selected.Count) {
                            foreach ($property in $selected[0].PSObject.Properties) {
                                $valueType = if ($null -eq $property.Value) { '(null)' } else { $property.Value.GetType().FullName }
                                $properties.Add([pscustomobject]@{ RequestedSession=$SessionId; Property=$property.Name; Type=$valueType; IsNull=($null -eq $property.Value) })
                            }
                            Add-ProbeResult -Area 'Session' -Name $SessionId -Status 'Found' -Detail ("{0} properties recorded by name and type." -f $properties.Count)
                        }
                        else { Add-ProbeResult -Area 'Session' -Name $SessionId -Status 'Absent locally' -Detail 'This ShareGate installation has no matching local session object.' }
                    }
                }
                catch { Add-ProbeResult -Area 'ShareGate' -Name 'License/session query' -Status 'Error' -Detail $_.Exception.Message }
            }
            if ($licenseConfirmed -and (Get-Command Connect-Site -Module ShareGate -ErrorAction SilentlyContinue) -and (Get-Command Get-List -Module ShareGate -ErrorAction SilentlyContinue)) {
                foreach ($key in @($endpoints.Keys | Sort-Object)) {
                    $entry = $endpoints[$key]
                    if ($entry.Side -eq 'Source' -and -not $ProbeSource) {
                        $entry.Result = [pscustomobject]@{ Status='Skipped - source unreachable'; Error='' }
                        continue
                    }
                    $entry.Result = Test-Endpoint -Side $entry.Side -Url $entry.Url -List $entry.List
                    Add-ProbeResult -Area 'Endpoint' -Name $key -Status $entry.Result.Status -Detail $entry.Result.Error
                }
            }
            elseif (-not $licenseConfirmed) { Add-ProbeResult -Area 'Endpoint' -Name 'Site/list reads' -Status 'Skipped' -Detail 'ShareGate PowerShell license was not confirmed by Find-CopySessions.' }
        }
    }
    if ($Run) {
        foreach ($row in $accessRows) {
            $source = $endpoints[(Get-EndpointKey -Side 'Source' -Url ([string]$row.SourceUrl) -List ([string]$row.SourceList))]
            $destination = $endpoints[(Get-EndpointKey -Side 'Destination' -Url ([string]$row.DestinationUrl) -List ([string]$row.DestinationList))]
            $sourceResult = if ($source.Result) { $source.Result } elseif (-not $ProbeSource) { [pscustomobject]@{ Status='Skipped - source unreachable'; Error='' } } else { [pscustomobject]@{ Status='Not tested'; Error='' } }
            $destinationResult = if ($destination.Result) { $destination.Result } else { [pscustomobject]@{ Status='Not tested'; Error='' } }
            $attribution = 'Undetermined'
            if ($sourceResult.Status -eq 'Access denied' -and $destinationResult.Status -eq 'Readable') { $attribution = 'Source' }
            elseif ($destinationResult.Status -eq 'Access denied' -and $sourceResult.Status -eq 'Readable') { $attribution = 'Destination' }
            $access.Add([pscustomobject]@{ SessionId=$row.SessionId; RowId=$row.RowId; SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; DestinationUrl=$row.DestinationUrl; DestinationList=$row.DestinationList; SourceStatus=$sourceResult.Status; DestinationStatus=$destinationResult.Status; Attribution=$attribution; Reason=$(if ($attribution -eq 'Undetermined') { 'Site/list reads cannot isolate the failing side; item-level access remains untested.' } else { 'One endpoint denied access while the other was readable.' }); SourceError=$sourceResult.Error; DestinationError=$destinationResult.Error })
        }
    }
    Write-ProbeLog "Probe completed. Output=$output"
}
catch {
    Add-ProbeResult -Area 'Probe' -Name 'Execution' -Status 'Failed' -Detail $_.Exception.Message
    throw
}
finally {
    Export-AtomicCsv -Name 'Probe-Results.csv' -Rows $results.ToArray() -Columns @('Area','Name','Status','Detail')
    Export-AtomicCsv -Name 'Probe-CmdletParameters.csv' -Rows $parameters.ToArray() -Columns @('Cmdlet','ParameterSet','Parameter','Type','Mandatory')
    Export-AtomicCsv -Name 'Probe-SessionProperties.csv' -Rows $properties.ToArray() -Columns @('RequestedSession','Property','Type','IsNull')
    Export-AtomicCsv -Name 'Probe-Access.csv' -Rows $access.ToArray() -Columns @('SessionId','RowId','SourceUrl','SourceList','DestinationUrl','DestinationList','SourceStatus','DestinationStatus','Attribution','Reason','SourceError','DestinationError')
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAl5LYI1U1BgRxN
# 6PfmPMD2bjGJL52JzBOUbIKfKPt6VqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL3mk+l5WxVAAgoHk4Hy/JKpO+TUibla2nXVMQ+AxcT3MA0GCSqG
# SIb3DQEBAQUABIIBgLEX0e/h0n3vKcokl++4HQnDfTXse5rPXkdLnMEKVvdVdi+6
# EunspDucTjJjofHFfbeIK5A8gl41bzZk3W77mH7R7EGBDIKUWK7USoQHVYpCqyw6
# ppbMo0mqn9bN+c2NGG6LFvEJUKcp22Szm0E7N4BwesTZRpYZepRQBBWLa2Q9nbDH
# KbKA8fnQVYfnsDQo3fpOxRaka6PqoKgv8e0hPB/OjIX/a42V9LPFWUiiD126jLV6
# 0pik9Ug43Hn9WN5GPQx0w2HwwlduWs29q2SpPqTvyxgaXaB8BWVMmxZDaInw9TAG
# 6rU9cx+WGr2J58YA0psPmyMa9YOujw8NvCUMVXl9r/cq+hlUVpUdDjHr0T9Hh8SO
# WOY50k7VCmwRqcKWuK495CRElZF0L64E+r++eK8MslD2PIo4yS7YV8w5AtfdLRUV
# /ShZOo7OXqGIgBdBcNWH9SIZVNospbKzx4BBCWyuDu7qCw1JaHVnGYOfziHAdnbG
# uxXgP11W+ie/pAsWl6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMwMDQw
# MjBaMC8GCSqGSIb3DQEJBDEiBCBEWwVolXm2mSYNbNhEFoa00JStHqd8b94K6JDq
# 257hdTANBgkqhkiG9w0BAQEFAASCAgC1HWOIVZ+h1+kwwDIUjTVniQ6OikfgQHde
# heMOsJp55gmhzeYurm91XhJIElIydVBMPEgWjDcN5rY97cv/bfQ52Cu8Uhg8TANR
# uMuNBTvy66tjrous9EGnQtfJpGBY5a/b1LlkFM0bR3SsJE39bXnbtPnO7EtBeGyR
# +D777L3AbSVf3Yhoo2rvHc1odDLr7MwDLkxXZ857qfs6m0ErOJrtClsFLf5dpVfU
# Pw9GkGs3cQEi4395PXrmIXZF6z3NuwIywqn6q5QZ/HaeuBvmqaU9zc8RpPFkOpaW
# QzblV6GMZjstAKqmit4tT6JFJYkp3aWt0FGyG6DI5XitHl2jBKrwdG9lx/GT1IQq
# n2u85cKeAYGVMDQSwiCEMCmdb8GWwG5JanckucFoc5iDL2pX1W0t9M1LgvNgx+eF
# HcnO31SXgqfbaWAbEEtAvnujBy1cbsyJ4ktsOIq1Zf2ks5sQlpUhLk9dn1JW75YM
# sjQvC3M8RasXM0s2nbiRFPbPbzFSUwpGXTXwmARnR35VC03ahHU+eAUxFA37lVXM
# o07dNUX8+qmFABfqWroHZjZShsF8aib+Pxw8K+/ZNz5Bu5Koiuy0rvbNnDyFGxZn
# Im26lc4ewZMyC7Ac4+pYqaZOTn9pZgcf+YE2g/F4BlFqc1DYJOAKfiQlwAL0X0Cz
# Ltjpbc4DGA==
# SIG # End signature block
