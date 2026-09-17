Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-SmartM365EvidenceProperty {
    param([AllowNull()]$InputObject,[Parameter(Mandatory)][string[]]$Names,[AllowNull()]$DefaultValue='')
    if ($null -eq $InputObject) { return $DefaultValue }
    foreach ($name in $Names) {
        if ($InputObject -is [System.Collections.IDictionary] -and $InputObject.Contains($name)) { return $InputObject[$name] }
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $DefaultValue
}

function ConvertTo-SmartM365EvidenceText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return (@($Value) | ForEach-Object { [string]$_ }) -join ';'
    }
    return [string]$Value
}

function ConvertTo-SmartM365EvidenceJson {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ($Value | ConvertTo-Json -Depth 20 -Compress)
}

function Resolve-SmartM365EvidenceToken {
    param([AllowNull()]$Value,[Parameter(Mandatory)]$EffectiveConfig)
    if ($Value -isnot [string]) { return $Value }
    $resolved = [string]$Value
    for ($iteration=0; $iteration -lt 10; $iteration++) {
        $tokenMatches = [regex]::Matches($resolved,'\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $property = $EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $resolved = $resolved.Replace($match.Value,[string]$property.Value)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-SmartM365EvidenceConfig {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)]$EffectiveConfig
    )
    $base = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $folder = Split-Path -Path $ScriptPath -Parent
    $localPath = Join-Path $folder "$base.local.json"
    $templatePath = "$localPath.template"
    $sourcePath = if (Test-Path -LiteralPath $localPath -PathType Leaf) { $localPath } else { $templatePath }
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "Collector configuration template not found: $templatePath" }
    $local = Get-Content -LiteralPath $sourcePath -Raw | ConvertFrom-Json -ErrorAction Stop
    $result = [ordered]@{}
    foreach ($property in $EffectiveConfig.PSObject.Properties) { $result[$property.Name] = $property.Value }
    foreach ($property in $local.PSObject.Properties) {
        if ($property.Value -is [string] -and $property.Value.Trim() -in @('__USE_GLOBAL__','USE_GLOBAL')) { continue }
        $result[$property.Name] = Resolve-SmartM365EvidenceToken -Value $property.Value -EffectiveConfig ([pscustomobject]$result)
    }
    foreach ($key in @($result.Keys)) { $result[$key] = Resolve-SmartM365EvidenceToken -Value $result[$key] -EffectiveConfig ([pscustomobject]$result) }
    return [pscustomobject]$result
}

function Initialize-SmartM365EvidenceRuntime {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)]$EffectiveConfig,
        [Parameter(Mandatory)][string]$DefaultOutputRelativePath,
        [string]$OutputPath,
        [string]$LatestCsvFolderPath,
        [switch]$ValidateOnly,
        [switch]$EnableConfiguredExternalActions
    )
    $config = Get-SmartM365EvidenceConfig -ScriptPath $ScriptPath -EffectiveConfig $EffectiveConfig
    $root = [string](Get-SmartM365EvidenceProperty $EffectiveConfig @('SmartM365RootPath'))
    $coreManifest = Join-Path $root 'Modules\SmartM365.Core\SmartM365.Core.psd1'
    Import-Module -Name $coreManifest -MinimumVersion '1.0.24' -Force -ErrorAction Stop
    $dataAllRoot = [string](Get-SmartM365EvidenceProperty $config @('DataAllRootPath'))
    $resolvedOutput = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath } elseif (-not [string]::IsNullOrWhiteSpace([string](Get-SmartM365EvidenceProperty $config @('ScriptCsvLogFolderPath')))) { [string](Get-SmartM365EvidenceProperty $config @('ScriptCsvLogFolderPath')) } else { Join-Path $dataAllRoot $DefaultOutputRelativePath }
    $resolvedLatest = if (-not [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath } else { [string](Get-SmartM365EvidenceProperty $config @('LatestCsvFolderPath')) }
    if ([string]::IsNullOrWhiteSpace($resolvedOutput) -or [string]::IsNullOrWhiteSpace($resolvedLatest)) { throw 'Collector output paths are not configured.' }

    $externalActionsEnabled = $EnableConfiguredExternalActions -and -not $ValidateOnly
    $global:EnableSharePointUpload = $externalActionsEnabled -and [bool](Get-SmartM365EvidenceProperty $config @('EnableSharePointUpload') $false)
    $global:EnableTeamsNotifications = $externalActionsEnabled -and [bool](Get-SmartM365EvidenceProperty $config @('EnableTeamsNotifications') $false)
    foreach ($name in @('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath','AppId','TenantId','Thumb','Thumbprint','SmtpServer','From','To','Cc','ErrorMailTo')) {
        Set-Variable -Name $name -Scope Global -Value ([string](Get-SmartM365EvidenceProperty $config @($name)))
    }
    $global:RetentionMaxCSV = [int](Get-SmartM365EvidenceProperty $config @('RetentionMaxCSV') 30)
    $global:RetentionMaxLogs = [int](Get-SmartM365EvidenceProperty $config @('RetentionMaxLogs') 30)
    InitializeScriptEnvironment -OutputPathInit $resolvedOutput -LogFileName ([System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)) -CallerScriptPath $ScriptPath | Out-Null
    return [pscustomobject]@{ Config=$config; OutputPath=$resolvedOutput; LatestCsvFolderPath=$resolvedLatest; RunId=[guid]::NewGuid().Guid; CollectedAtUtc=[datetime]::UtcNow.ToString('o') }
}

function Connect-SmartM365EvidenceGraph {
    param([Parameter(Mandatory)]$Runtime,[Parameter(Mandatory)][string[]]$Scopes,[switch]$InteractiveAuth)
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    try { if (Get-MgContext -ErrorAction SilentlyContinue) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } } catch {}
    if ($InteractiveAuth) { Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null; return }
    $appId = [string](Get-SmartM365EvidenceProperty $Runtime.Config @('AppId'))
    $tenantId = [string](Get-SmartM365EvidenceProperty $Runtime.Config @('TenantId'))
    $thumb = [string](Get-SmartM365EvidenceProperty $Runtime.Config @('Thumbprint','Thumb'))
    if ([string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($tenantId) -or [string]::IsNullOrWhiteSpace($thumb)) { throw 'AppId, TenantId, and Thumbprint are required for app-only Graph authentication.' }
    Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $thumb -NoWelcome -ErrorAction Stop | Out-Null
}

function Get-SmartM365EvidenceStatusCode {
    param([Parameter(Mandatory)]$ErrorRecord)
    try { if ($ErrorRecord.Exception.Response.StatusCode) { return [int]$ErrorRecord.Exception.Response.StatusCode } } catch {}
    if ($ErrorRecord.Exception.Message -match '(?<!\d)(?<Code>400|401|403|404|408|409|429|500|502|503|504)(?!\d)') { return [int]$Matches.Code }
    return 0
}

function Invoke-SmartM365EvidenceGraphRequest {
    param([ValidateSet('GET','POST')][string]$Method='GET',[Parameter(Mandatory)][string]$Uri,[AllowNull()][string]$Body,[int]$MaxRetryCount=6,[scriptblock]$Invoker,[switch]$NoSleep)
    for ($attempt=1; $attempt -le $MaxRetryCount; $attempt++) {
        try {
            if ($Invoker) { return & $Invoker $Method $Uri $Body }
            $parameters = @{Method=$Method;Uri=$Uri;OutputType='PSObject';ErrorAction='Stop'}
            if ($Method -eq 'POST') { $parameters.Body=$Body; $parameters.ContentType='application/json' }
            return Invoke-MgGraphRequest @parameters
        }
        catch {
            $status = Get-SmartM365EvidenceStatusCode $_
            if ($status -notin @(408,409,429,500,502,503,504) -or $attempt -eq $MaxRetryCount) { throw }
            if (-not $NoSleep) { Start-Sleep -Seconds ([Math]::Min(120,[int](5*[Math]::Pow(2,$attempt-1)))) }
        }
    }
}

function Get-SmartM365EvidenceGraphCollection {
    param([Parameter(Mandatory)][string]$Uri,[ValidateRange(1,100000)][int]$MaxPages=2000,[int]$MaxItems=0,[scriptblock]$Invoker)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    $page = 0
    while (-not [string]::IsNullOrWhiteSpace($next) -and $page -lt $MaxPages) {
        $page++
        $response = Invoke-SmartM365EvidenceGraphRequest -Uri $next -Invoker $Invoker -NoSleep:$($null-ne$Invoker)
        $values = @(Get-SmartM365EvidenceProperty $response @('value') @())
        foreach ($value in $values) {
            $items.Add($value)
            if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { return @($items) }
        }
        $next = [string](Get-SmartM365EvidenceProperty $response @('@odata.nextLink') '')
    }
    if (-not [string]::IsNullOrWhiteSpace($next)) { throw "Graph pagination exceeded MaxPages=$MaxPages." }
    return @($items)
}

function Invoke-SmartM365EvidenceExportReport {
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [string[]]$Select=@(),
        [ValidateSet('v1.0','beta')][string]$ApiVersion='v1.0',
        [int]$TimeoutSeconds=600,
        [int]$PollSeconds=5,
        [int]$MaxItems=0
    )
    $base = 'https://graph.microsoft.com'
    $uri = "$base/$ApiVersion/deviceManagement/reports/exportJobs"
    $requestBody = [ordered]@{reportName=$ReportName;format='csv';localizationType='replaceLocalizableValues'}
    if (@($Select).Count -gt 0) { $requestBody.select=@($Select) }
    $body = $requestBody | ConvertTo-Json -Depth 8 -Compress
    $job = Invoke-SmartM365EvidenceGraphRequest -Method POST -Uri $uri -Body $body
    $jobId = [string](Get-SmartM365EvidenceProperty $job @('id'))
    if ([string]::IsNullOrWhiteSpace($jobId)) { throw "No export job id returned for $ReportName." }
    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $job = Invoke-SmartM365EvidenceGraphRequest -Uri "$uri/$([uri]::EscapeDataString($jobId))"
        $status = [string](Get-SmartM365EvidenceProperty $job @('status'))
        if ($status -ieq 'completed') { break }
        if ($status -ieq 'failed') { throw "Intune export job failed for $ReportName." }
        Start-Sleep -Seconds $PollSeconds
    } while ([datetime]::UtcNow -lt $deadline)
    if ($status -ine 'completed') { throw "Intune export job timed out for $ReportName." }
    $downloadUrl = [string](Get-SmartM365EvidenceProperty $job @('url'))
    if ([string]::IsNullOrWhiteSpace($downloadUrl)) { throw "Completed export job returned no URL for $ReportName." }
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "SmartM365-Evidence-$([guid]::NewGuid().Guid)"
    try {
        $extract = Join-Path $tempRoot 'content'; $zip = Join-Path $tempRoot 'export.zip'
        New-Item -Path $extract -ItemType Directory -Force | Out-Null
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zip -ErrorAction Stop
        Expand-Archive -LiteralPath $zip -DestinationPath $extract -Force
        $csv = Get-ChildItem -LiteralPath $extract -Filter '*.csv' -File -Recurse | Select-Object -First 1
        if (-not $csv) { throw "No CSV found in the Intune export for $ReportName." }
        $rows = @(Import-Csv -LiteralPath $csv.FullName)
        if ($MaxItems -gt 0) { $rows = @($rows | Select-Object -First $MaxItems) }
        return $rows
    }
    finally { if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } }
}

function Export-SmartM365EvidenceDataset {
    param([Parameter(Mandatory)]$Runtime,[Parameter(Mandatory)][string]$BaseFileName,[AllowEmptyCollection()][object[]]$Rows,[Parameter(Mandatory)][string[]]$Columns,[switch]$NoWeeklyHistory)
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $timestamped = Join-Path $Runtime.OutputPath "${BaseFileName}_${stamp}.csv"
    $latest = Join-Path $Runtime.LatestCsvFolderPath "${BaseFileName}.csv"
    return Export-SmartM365Csv -Data $Rows -TimestampedPath $timestamped -LatestPath $latest -Columns $Columns -NoWeeklyHistory:$NoWeeklyHistory
}

function Write-SmartM365EvidenceLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR','DEBUG')][string]$Level='INFO'
    )
    WriteLog -Message $Message -Level $Level
}

function Complete-SmartM365EvidenceRuntime {
    param(
        [ValidateSet('Auto','Success','CompletedWithWarnings','Failed')][string]$Status='Auto',
        [AllowNull()]$ErrorRecord,
        [string]$FailureStage=''
    )
    Complete-SmartM365ExecutionContext -Status $Status -ErrorRecord $ErrorRecord -FailureStage $FailureStage
}

Export-ModuleMember -Function Get-SmartM365EvidenceProperty,ConvertTo-SmartM365EvidenceText,ConvertTo-SmartM365EvidenceJson,Get-SmartM365EvidenceConfig,Initialize-SmartM365EvidenceRuntime,Connect-SmartM365EvidenceGraph,Get-SmartM365EvidenceStatusCode,Invoke-SmartM365EvidenceGraphRequest,Get-SmartM365EvidenceGraphCollection,Invoke-SmartM365EvidenceExportReport,Export-SmartM365EvidenceDataset,Write-SmartM365EvidenceLog,Complete-SmartM365EvidenceRuntime

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCALJELKBlZQeMlQ
# UzAkCv6QE47chnjUCtlwORfPNDU6y6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILnDvPmAFqqz33anoZKoq123lvK+zG2YwtEGYZkmAsjrMA0GCSqG
# SIb3DQEBAQUABIIBgAKv55AYxzoBp+w/8Ja54Z9Uu3hr6LIIAGIiE5EWzepqbZDp
# LqBoTWr8IcDQj2KMQOurpbczinMPKPn9gho2XlVfDwHsfx0+ewq068JqW7Dwzm9u
# IHnGgq4z8QxiuED9sEev8q6lHjjj0PdC/NADN+P4Ph/t45CJMq2yNHrJFqew++K4
# OMWR5cPrw+aEXqtoYDfy2b2XYFFtQP+GmtmBkSPhjLeVSFeuvLVrJtuELSU56SzP
# CHXt4Mv7/2V5kwk3Dxw2qY2w/7BB6q3Z5oYliK6j//6rk7Z561bdtZoQFt5zGdyj
# j5Chbp1NjFgGebQNtxLzgjj14H0eFmVl2SteeQ3nFctW79qFWk8jeArwYV0lDAiX
# +APOZieS0J4zM9Mo4hErpzwqADjPKjKIiWwtqT7LTw9V4B5uHlobedvsIVOymeoL
# duKWtvmJ9+bDVGjFPWTf26uct+aSyUzhamB35scUbIrG4lbIW3F81oXc7eUAVb4y
# kL7unnbmbkI9/uaGY6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTcxMjQy
# NTdaMC8GCSqGSIb3DQEJBDEiBCDjxNP83T29+kEF0EqLY17klO5TcKZVWMA173rp
# X8najDANBgkqhkiG9w0BAQEFAASCAgANiWhXiWc88zvw+c+Z7H15Re9tyfsTsjP5
# NAuyfmvV1bCvUjgsrlXp4o+Duk2xGwxgCt2qlkwgrEdx6RYeknEDWtg4vueIMGCF
# sr4qPggWo6vpbXixxoGodfW8oJyHi5jHZ3ebxLEmkzeADnNTDVHV3xpycIy7rP4B
# DAzQWE5yoxR2cPZHYjLb70LErcVievVzZvOTGb9AsB8/Ko3W2pRQB9SS4e57x064
# egaSTp7GGti4qGC1stADVj/8n01K1bm215nY6Lv26mfS4ykfwETUONeEwI01BS+5
# etvcGzf6NQ7IIutIiEhUEOONMruLE/1cdOh8CcUf1taoleCxszIJzr/GZErt38BK
# uPbKdqer0uUR/rqUn0BpEWv/pSvBLhgWDvLXpwXUYzVxMMCdpH/hjqBm45Tp6Pve
# OHiHT5NvRrtDPfqzu57Z+gNTwBtIsqsmLArqmmcuC+qprPRxrvDNkmRHBuZFnBtF
# tjDI5ipFTUwublQ9gp1XNsz2rKojZJ8b/XroBZww8rHzf1qsDRr/nHnDYa+pCw8J
# MgWNSPprPJJK2yBX40Ce5rcAG/Q37+2G97ulMsyEYKB9EZMUj78MYVSpQLvDOql8
# YXB7C5amFW428w7r64NGkJvb+dG+G5GChR4963eMKCvSpUEmwMYUz7uH6IDuQ826
# SwTFudRXlg==
# SIG # End signature block
