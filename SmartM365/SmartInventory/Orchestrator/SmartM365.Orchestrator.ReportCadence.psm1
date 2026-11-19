# Script report cadence editor for the Orchestrator GUI.
# Version: 1.3.0
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ConfigView.psm1') -ErrorAction Stop

function Get-ReportCadenceDefinitions {
    param([string]$InventoryRoot)
    $manifest = Join-Path $InventoryRoot 'Orchestrator/Orchestrator-Jobs.json.template'
    $jobs = (Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json).Jobs
    $paths = @('Orchestrator/SmartM365-Inventory-Orchestrator.ps1') + @($jobs | ForEach-Object { $_.ScriptPath })
    foreach ($scriptPath in @($paths | Where-Object { $_ } | Sort-Object -Unique)) {
        if ($scriptPath -notmatch '^([a-zA-Z0-9_-]+[\\/])+[a-zA-Z0-9_-]+\.ps1$') { continue }
        $path = Join-Path $InventoryRoot ([IO.Path]::ChangeExtension($scriptPath, '.local.json'))
        $templatePath = Get-SmartM365JsonTemplateName -Path $path
        if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) { continue }
        # Catalog entries are shipped jobs; only templates that declare this setting are editable.
        $templateText = Get-Content -LiteralPath $templatePath -Raw
        if ($templateText -notmatch '"ReportMailCadence"\s*:') { continue }
        $template = $templateText | ConvertFrom-Json
        $property = $template.PSObject.Properties['ReportMailCadence']
        if (-not $property) { continue }
        [pscustomobject]@{ Script=[IO.Path]::GetFileNameWithoutExtension($scriptPath); ScriptPath=$scriptPath.Replace([char]92,[char]47); Path=$path; TemplatePath=$templatePath; Default=$property.Value }
    }
}

function Resolve-ReportCadencePath {
    param([string]$Path, [string]$ConfigRoot)
    # Mirror the runtime relocation without creating directories during a read.
    if (-not $configRoot) { return $Path }
    $relative = [IO.Path]::GetRelativePath((Join-Path $PSScriptRoot '..'), $Path)
    if ($relative -match '^\.\.[\\/]' -or [IO.Path]::IsPathRooted($relative)) {
        throw 'The script catalog must come from this SmartInventory installation.'
    }
    $target = Join-Path $configRoot $relative
    if (-not (Get-SmartM365JsonReadPath $target -Optional) -and (Get-SmartM365JsonReadPath $Path -Optional)) {
        throw 'Move the script configuration to the configured inventory configuration root before editing.'
    }
    return $target
}

function ConvertTo-ReportCadenceSetting {
    param([string]$Mode, [string[]]$Days)
    $names = [Enum]::GetNames([DayOfWeek])
    if ($Mode -eq 'Daily') { return @{ Mode='Daily' } }
    if ($Mode -ne 'SelectedDays' -or -not $Days.Count -or @($Days | Where-Object { $_ -notin $names }).Count -or $Days.Count -ne @($Days | Sort-Object -Unique).Count) {
        throw 'Select Daily or choose at least one valid weekday for SelectedDays.'
    }
    return @{ Mode='SelectedDays'; Days=@(($names[1..6] + $names[0]) | Where-Object { $_ -in $Days }) }
}

function Get-SmartM365OrchestratorReportCadence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$InventoryRoot)
    $configRoot = Get-SmartM365InventoryConfigRoot
    $rows = foreach ($definition in Get-ReportCadenceDefinitions $InventoryRoot) {
        $value = $definition.Default
        $source = 'Template default'
        $status = 'Ready'
        $configuredPath = ''
        $hash = ''
        $canEdit = $false
        try {
            $path = Resolve-ReportCadencePath $definition.Path $configRoot
            $names = Get-SmartM365JsonNames $path
            $file = Read-SmartM365OrchestratorConfigFile -Path $path -LockTimeoutSeconds 0
            $configuredPath = $file.Path
            if ($file.Status -notin @('OK','Missing')) { throw "Configuration unavailable ($($file.Status))." }
            $hash = if ($file.Status -eq 'Missing') { 'ABSENT' } else { $file.SHA256 }
            if ($file.Status -eq 'OK') {
                $property = $file.Document.PSObject.Properties['ReportMailCadence']
                if ($property) { $value = $property.Value; $source = 'Script configuration' }
            }
            if ((Test-Path -LiteralPath $names.Legacy) -and (Test-Path -LiteralPath $names.Preferred)) {
                throw 'Both JSON variants exist. Resolve the pair before editing.'
            }
            $canEdit = $true
        } catch { $status = $_.Exception.Message }
        $mode = ''
        $days = @()
        if ($value -and $value.PSObject.Properties['Mode']) { $mode = [string]$value.Mode }
        if ($mode -eq 'SelectedDays' -and $value.PSObject.Properties['Days']) { $days = @($value.Days) }
        if ($canEdit) {
            try { $null = ConvertTo-ReportCadenceSetting -Mode $mode -Days $days }
            catch { $status = 'Invalid cadence. Select a valid mode and days to repair it.' }
        }
        [pscustomobject]@{
            Script=$definition.Script; Mode=$mode; Days=$days; Weekdays=($days -join ', ')
            Source=$source; Status=$status; CanEdit=$canEdit; SHA256=$hash; Path=$configuredPath
        }
    }
    return @($rows | Sort-Object Script)
}

function Set-SmartM365OrchestratorReportCadence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InventoryRoot, [Parameter(Mandatory)][string]$Script,
        [Parameter(Mandatory)][string]$Mode, [string[]]$Days=@(),
        [Parameter(Mandatory)][string]$ExpectedPath, [Parameter(Mandatory)][string]$ExpectedSHA256
    )
    $cadence = ConvertTo-ReportCadenceSetting -Mode $Mode -Days $Days
    $matches = @(Get-ReportCadenceDefinitions $InventoryRoot | Where-Object Script -EQ $Script)
    if ($matches.Count -ne 1) { throw 'Select a script that supports report cadence.' }
    $definition = $matches[0]
    $path = Resolve-ReportCadencePath $definition.Path (Get-SmartM365InventoryConfigRoot)
    $names = Get-SmartM365JsonNames $path
    $file = Read-SmartM365OrchestratorConfigFile -Path $path -LockTimeoutSeconds 0
    if ((Test-Path -LiteralPath $names.Legacy) -and (Test-Path -LiteralPath $names.Preferred)) {
        throw 'Both JSON variants exist. Resolve the pair before editing.'
    }
    $hash = if ($file.Status -eq 'Missing') { 'ABSENT' } else { $file.SHA256 }
    if ($file.Status -notin @('OK','Missing') -or $file.Path -ne $ExpectedPath -or $hash -cne $ExpectedSHA256) {
        throw 'The script configuration changed or is unavailable. Refresh and try again.'
    }
    $target = $file.Path
    $source = $file.Path
    if ($file.Status -eq 'Missing') {
        $source = $definition.TemplatePath
        $target = if ((Get-SmartM365JsonTransportPolicy).Mode -eq 'JsonText') { $names.Preferred } else { $names.Legacy }
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $target))
    }
    # Preserve data types of unrelated settings, including JSON strings that look like dates.
    $reader = [Newtonsoft.Json.JsonTextReader]::new([IO.StringReader]::new([IO.File]::ReadAllText($source)))
    try {
        $reader.DateParseHandling = [Newtonsoft.Json.DateParseHandling]::None
        $document = [Newtonsoft.Json.Linq.JObject]::Load($reader)
    } finally { $reader.Close() }
    $cadenceProperty = $document.Property('ReportMailCadence', [StringComparison]::OrdinalIgnoreCase)
    if ($cadenceProperty) { $cadenceProperty.Remove() }
    $document.Add('ReportMailCadence', [Newtonsoft.Json.Linq.JObject]::Parse(($cadence | ConvertTo-Json -Compress)))
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($document.ToString() + [Environment]::NewLine)
    $null = Write-SmartM365JsonBytesAtomically -Path $target -Bytes $bytes -ExpectedSHA256 $hash -LockTimeoutSeconds 1 -Validate {
        param($doc)
        if ($doc -isnot [pscustomobject] -or -not $doc.PSObject.Properties['ReportMailCadence']) { throw 'Invalid report cadence document.' }
    }
}

function Get-SmartM365OrchestratorReportMailProducers {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$InventoryRoot)
    $producers = @{}
    foreach ($definition in Get-ReportCadenceDefinitions $InventoryRoot) {
        $producers[$definition.ScriptPath] = $definition.Script
    }
    return $producers
}

function Get-SmartM365OrchestratorMailOnlySource {
    param([string]$InventoryRoot,[string]$Producer)
    $matches = @(Get-ReportCadenceDefinitions $InventoryRoot | Where-Object Script -EQ $Producer)
    if ($matches.Count -ne 1) { throw 'The selected script does not support report replay.' }
    $file = Read-SmartM365OrchestratorConfigFile -Path (Resolve-ReportCadencePath $matches[0].Path (Get-SmartM365InventoryConfigRoot))
    if ($file.Status -ne 'OK') { throw 'Current script mail configuration is unavailable.' }
    $scriptPath = $matches[0].Path -replace '\.local\.json$','.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { throw 'The selected inventory script is unavailable.' }
    return [pscustomobject]@{Config=$file.Document;ScriptPath=$scriptPath}
}

Export-ModuleMember -Function Get-SmartM365OrchestratorReportCadence, Set-SmartM365OrchestratorReportCadence, Get-SmartM365OrchestratorMailOnlySource, Get-SmartM365OrchestratorReportMailProducers

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDRHixL9w0pdmv4
# ohJlYiJhtXj0v0Ruae7BlNHK7cYvcaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgoDWQj4HAn6G+
# 8gQnn5n1opk/LF/zsiS7+vDmGw32058wDQYJKoZIhvcNAQEBBQAEggGART7TA6Qh
# yTshYGYu+H5tadveZ35D8YFP1y9SxBz1Ap2F1WtlSdBbVoCe0kM1eMbulnregy58
# lunPMs41lzUNMWlMB3zKIS6139zNQ4t46Drlb0xVDIiQBwwHv73i0iQoTfmhPVga
# CQ9qZluUiyTeyJVmbvS2xRIIZxO4XysXE0AJvqFfyLPN+ECr3hpJUPG7cg9QSFGT
# ovfqGTc5Ksz/dzTDcFGT5hbTKBuLGR9xgfSCOIHjUswu/pV4nllCWT6/WpcBZ5iR
# RQFSXbXnGOFlkB6f0NpGzheBXeAJcIffOVt3p963j0gznZQl9QE0x4hqB4rLbgXN
# /jegL8j56gwrXQhBmvF41/aM2o06rWYpYpEKAzXka8Ivh+zXuU4jc4W4AzqB7Th4
# 5ntghPygKaTU4JwigwQv2ruoNtT06H7sc2EC1mCNmKeOzmV+3Kvg3t+IZIMKJQUo
# oEvkCImzFm0hrPHpyegxePfEGcL+Ecqkp+GTUsMr5dDKqasDjy3+bHSroYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDk0MzAxWjA/BgkqhkiG9w0BCQQxMgQwmUcGkvcw7PrkqeNntd/zblW/
# nL6ZBgM0t8UrXaASJKf8B8+5m2it9LRZn/99NBBMMA0GCSqGSIb3DQEBAQUABIIC
# AITWinzj2fnloTJ8B4iAptFIxFbbYPR6yIotW1hbrWn0OSYw4bkuhoyURxIyXVAa
# InDPB5DLTzi381oPUSlVrrikVq4DbceZ6KjKcaBgQPEPMHG4r5hCX1uTLLrMW2wK
# N3lctyFNKcEItc1kVBohtFGWAHWq7YpvS9d/BrVXEYigAGFFLTBpN6/3uusK+TIl
# 5aJ1FCVJ4UYRx0suCH2kg259IDsTj71Mf40Mm2mJa7Z4DiTgKDNvfgo/ThVmAVf4
# 8OBwqIqXub900gukqqAh5AjoC+BmSIfBhNGBYiF/3H6DGj4CIoe0fjbG2JNuiz0v
# 3snT8wKomyhMkJuGCcdPqK5X9Rjnk5fxfv1zV9c4tljQCROuHTMBvGYVPUH5T3yT
# Li/7NsLBwc72DYmR1SS5IbOmqhwWxXmaDdpJrWz4dTjwWdZFI3xsvAn9vxm6dZmz
# ZPtsGKVy0efJA8XwkZKYqoFFRN9FRZDbTvdD2D7chevp6P8o+dgBt+K5qDaROatN
# oyGXf5S6DGEXkBmnMhb5x3GrAjiqqY8/sOstPClFR4q5b9fyDG/9nYQOGx1i/VOu
# 1iyIvIkFQdNOLmyYKI21HZ9RMSB2sTmfyTsM74EgG91lnI7WCsy9E9eKmYxLbxZB
# OZBcuwQjGREZlgw7H336w6N6KcKZeX/RKrrhtskLfiH8
# SIG # End signature block
