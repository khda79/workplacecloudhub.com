function Set-SmartAzureCoreContext {
    [CmdletBinding()]
    param(
        [string]$RunId,
        [string]$RunOutputRoot,
        [string]$LatestOutputRoot,
        [string]$LogPath,
        [int]$RetentionMaxCsv = 30
    )

    $script:SmartAzureCoreRunId = $RunId
    $script:SmartAzureCoreRunOutputRoot = $RunOutputRoot
    $script:SmartAzureCoreLatestOutputRoot = $LatestOutputRoot
    $script:SmartAzureCoreLogPath = $LogPath
    $script:SmartAzureCoreRetentionMaxCsv = $RetentionMaxCsv
}

function Get-SmartAzureCoreContextValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    $variableName = "SmartAzureCore$Name"
    $variable = Get-Variable -Name $variableName -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $variable -and $null -ne $variable.Value) {
        return $variable.Value
    }

    return $DefaultValue
}

function Write-SmartAzureLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line

    $logPath = Get-SmartAzureCoreContextValue -Name 'LogPath' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($logPath)) { return }

    $logFolder = Split-Path -Path $logPath -Parent
    if (-not (Test-Path -LiteralPath $logFolder)) {
        New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
    }
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Import-RequiredModule {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw ("Required PowerShell module '{0}' is not installed. Install it with: Install-Module {0} -Scope CurrentUser" -f $Name)
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function Test-SmartAzureFileLocked {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }

    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Close()
        return $false
    }
    catch {
        return $true
    }
}

function Remove-SmartAzureOldFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$FilePattern,
        [int]$MaxFiles = 30
    )

    if ($MaxFiles -le 0 -or -not (Test-Path -LiteralPath $FolderPath)) { return }

    $files = @(Get-ChildItem -LiteralPath $FolderPath -Filter $FilePattern -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($files.Count -le $MaxFiles) { return }

    foreach ($file in @($files | Select-Object -Skip $MaxFiles)) {
        try {
            if (-not (Test-SmartAzureFileLocked -Path $file.FullName)) {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Write-SmartAzureLog -Message ("Removed old file: {0}" -f $file.FullName)
            }
        }
        catch {
            Write-SmartAzureLog -Level WARN -Message ("Failed to remove old file '{0}': {1}" -f $file.FullName, $_.Exception.Message)
        }
    }
}

function ConvertTo-CompactJson {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [int]$Depth = 30
    )

    if ($null -eq $Value) { return '' }

    try {
        return ($Value | ConvertTo-Json -Depth $Depth -Compress -ErrorAction Stop)
    }
    catch {
        return [string]$Value
    }
}

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$PropertyName
    )

    if ($null -eq $InputObject) { return $null }

    foreach ($name in $PropertyName) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }

    return $null
}

function Get-SmartAzureNestedPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$Path
    )

    $current = $InputObject
    foreach ($segment in $Path) {
        if ($null -eq $current) { return $null }
        $property = $current.PSObject.Properties[$segment]
        if ($null -eq $property) { return $null }
        $current = $property.Value
    }
    return $current
}

function Invoke-SmartAzureArmGetPaged {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$Operation
    )

    $rows = @()
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        try {
            $response = Invoke-AzRestMethod -Method GET -Uri $nextUri -ErrorAction Stop
            $content = if ([string]::IsNullOrWhiteSpace($response.Content)) { $null } else { $response.Content | ConvertFrom-Json -Depth 100 }
            if ($null -eq $content) { break }

            $value = Get-ObjectPropertyValue -InputObject $content -PropertyName @('value')
            if ($null -ne $value) {
                $rows += @($value)
            }
            else {
                $rows += $content
            }

            $nextLink = Get-ObjectPropertyValue -InputObject $content -PropertyName @('nextLink', '@odata.nextLink')
            $nextUri = if ([string]::IsNullOrWhiteSpace([string]$nextLink)) { $null } else { [string]$nextLink }
        }
        catch {
            Write-SmartAzureLog -Level WARN -Message ("{0} failed: {1}" -f $Operation, $_.Exception.Message)
            break
        }
    }

    return @($rows)
}

function Invoke-SafeInventoryBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
    )

    try {
        return & $ScriptBlock
    }
    catch {
        Write-SmartAzureLog -Level WARN -Message ("{0} skipped or partially failed: {1}" -f $Name, $_.Exception.Message)
        return @()
    }
}

function Export-SmartAzureCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [string]$RunOutputRoot = $(Get-SmartAzureCoreContextValue -Name 'RunOutputRoot' -DefaultValue ''),
        [string]$LatestOutputRoot = $(Get-SmartAzureCoreContextValue -Name 'LatestOutputRoot' -DefaultValue ''),
        [string]$RunId = $(Get-SmartAzureCoreContextValue -Name 'RunId' -DefaultValue (Get-Date -Format 'yyyyMMdd_HHmmss')),
        [ValidateSet('ASCII', 'BigEndianUnicode', 'Default', 'OEM', 'Unicode', 'UTF7', 'UTF8', 'UTF32')]
        [string]$Encoding = 'UTF8',
        [char]$Delimiter
    )

    if ([string]::IsNullOrWhiteSpace($RunOutputRoot)) { throw 'Export-SmartAzureCsv: RunOutputRoot is missing.' }
    if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { throw 'Export-SmartAzureCsv: LatestOutputRoot is missing.' }

    $historyPath = Join-Path -Path $RunOutputRoot -ChildPath ("{0}_{1}.csv" -f $Name, $RunId)
    $latestPath = Join-Path -Path $LatestOutputRoot -ChildPath ("{0}.csv" -f $Name)

    foreach ($folder in @((Split-Path -Path $historyPath -Parent), (Split-Path -Path $latestPath -Parent))) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }
    }

    $exportParams = @{
        LiteralPath       = $historyPath
        NoTypeInformation = $true
        Encoding          = $Encoding
    }
    if ($PSBoundParameters.ContainsKey('Delimiter')) {
        $exportParams['Delimiter'] = $Delimiter
    }

    $Rows | Export-Csv @exportParams
    Write-SmartAzureLog -Message ("Exported {0} row(s) to {1}" -f @($Rows).Count, $historyPath)

    $maxRetries = 3
    $retryDelaySec = 10
    $attempt = 0
    $latestCopyDone = $false

    while (-not $latestCopyDone -and $attempt -lt $maxRetries) {
        $attempt++
        try {
            Copy-Item -LiteralPath $historyPath -Destination $latestPath -Force -ErrorAction Stop
            $latestCopyDone = $true
            Write-SmartAzureLog -Message ("Latest CSV updated: {0}" -f $latestPath)
        }
        catch {
            if ($attempt -lt $maxRetries) {
                Write-SmartAzureLog -Level WARN -Message ("Latest CSV copy failed (attempt {0}/{1}), retrying in {2}s: {3}" -f $attempt, $maxRetries, $retryDelaySec, $_.Exception.Message)
                Start-Sleep -Seconds $retryDelaySec
            }
            else {
                Write-SmartAzureLog -Level WARN -Message ("Latest CSV copy failed after {0} attempts: {1}" -f $maxRetries, $_.Exception.Message)
            }
        }
    }

    $retentionMaxCsv = [int](Get-SmartAzureCoreContextValue -Name 'RetentionMaxCsv' -DefaultValue 30)
    if ($retentionMaxCsv -gt 0) {
        Remove-SmartAzureOldFiles -FolderPath $RunOutputRoot -FilePattern ("{0}_*.csv" -f $Name) -MaxFiles $retentionMaxCsv
    }

    $uploadPath = if ($latestCopyDone) { $latestPath } else { $historyPath }
    $uploadCommand = Get-Command -Name Invoke-SmartAzureSharePointCsvUpload -ErrorAction SilentlyContinue
    if ($null -ne $uploadCommand) {
        & $uploadCommand -LocalFilePath $uploadPath
    }
}

function Export-SmartAzureCsvFromConvert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [string]$RunOutputRoot = $(Get-SmartAzureCoreContextValue -Name 'RunOutputRoot' -DefaultValue ''),
        [string]$LatestOutputRoot = $(Get-SmartAzureCoreContextValue -Name 'LatestOutputRoot' -DefaultValue ''),
        [string]$RunId = $(Get-SmartAzureCoreContextValue -Name 'RunId' -DefaultValue (Get-Date -Format 'yyyyMMdd_HHmmss')),
        [ValidateSet('ASCII', 'BigEndianUnicode', 'Default', 'OEM', 'Unicode', 'UTF7', 'UTF8', 'UTF32')]
        [string]$Encoding = 'UTF8',
        [string]$Delimiter = ','
    )

    if ([string]::IsNullOrWhiteSpace($RunOutputRoot)) { throw 'Export-SmartAzureCsvFromConvert: RunOutputRoot is missing.' }
    if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { throw 'Export-SmartAzureCsvFromConvert: LatestOutputRoot is missing.' }

    $historyPath = Join-Path -Path $RunOutputRoot -ChildPath ("{0}_{1}.csv" -f $Name, $RunId)
    $latestPath = Join-Path -Path $LatestOutputRoot -ChildPath ("{0}.csv" -f $Name)

    foreach ($folder in @((Split-Path -Path $historyPath -Parent), (Split-Path -Path $latestPath -Parent))) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }
    }

    $Rows | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | Out-File -LiteralPath $historyPath -Encoding $Encoding
    Copy-Item -LiteralPath $historyPath -Destination $latestPath -Force
    Write-SmartAzureLog -Message ("Exported {0} row(s) to {1}" -f @($Rows).Count, $historyPath)

    $uploadCommand = Get-Command -Name Invoke-SmartAzureSharePointCsvUpload -ErrorAction SilentlyContinue
    if ($null -ne $uploadCommand) {
        & $uploadCommand -LocalFilePath $latestPath
    }
}

function Invoke-SmartAzurePreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [string[]]$RequiredModules = @(),
        [string[]]$OutputPaths = @(),
        [switch]$RequireAzContext
    )

    foreach ($moduleName in @($RequiredModules | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        Import-RequiredModule -Name $moduleName
    }

    foreach ($path in @($OutputPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        try {
            New-Item -Path $path -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            throw ("{0} preflight failed: output path '{1}' is not writable: {2}" -f $ScriptName, $path, $_.Exception.Message)
        }
    }

    if ($RequireAzContext -and -not (Get-AzContext -ErrorAction SilentlyContinue)) {
        throw ("{0} preflight failed: no Azure context is connected." -f $ScriptName)
    }

    Write-SmartAzureLog -Message ("Preflight completed for {0}" -f $ScriptName)
}

function Connect-SmartAzureCloudSession {
    [CmdletBinding()]
    param(
        [string]$TenantId,
        [switch]$Connect,
        [switch]$UseDeviceCode
    )

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if ($Connect -or $null -eq $context) {
        Write-SmartAzureLog -Message 'Connecting to Azure.'
        $connectParams = @{}
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connectParams.Tenant = $TenantId }
        if ($UseDeviceCode) { $connectParams.UseDeviceAuthentication = $true }
        Connect-AzAccount @connectParams | Out-Null
        $context = Get-AzContext -ErrorAction SilentlyContinue
    }

    return $context
}

Export-ModuleMember -Function `
    Set-SmartAzureCoreContext, `
    Write-SmartAzureLog, `
    Import-RequiredModule, `
    Test-SmartAzureFileLocked, `
    Remove-SmartAzureOldFiles, `
    ConvertTo-CompactJson, `
    Get-ObjectPropertyValue, `
    Get-SmartAzureNestedPropertyValue, `
    Invoke-SmartAzureArmGetPaged, `
    Invoke-SafeInventoryBlock, `
    Export-SmartAzureCsv, `
    Export-SmartAzureCsvFromConvert, `
    Invoke-SmartAzurePreflight, `
    Connect-SmartAzureCloudSession

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC/HLyR8vQt8hcE
# mtK3KFx33Nb6cw88L/jZAkCNvPHuNqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEINM75xYP7VGGGq5AaFYfGHjsTSL+xksdKUBxoGbDQhQUMA0GCSqG
# SIb3DQEBAQUABIIBgC46YDjjzVhhGvugOcr+HTmMAWCKpdVbRZa41IUGKHvkUJzB
# tXh56mYJYW8Vd921FgW1Zc+86icEahP5ZHIVtBGSroMpEqxb7RO2Bft0HiXrieMd
# 9x4uYgUjfSkDaHB5gEiLNLUzBkNOg1cN1UqoMPmmDo2iFBUuDNe+cyjvTERkh1ql
# JwGEXpNODs7wEl80ho499uaH0T+NbdToQ/LdQpRY0gziFHXDPPplxidgBlxcwNbs
# byjL+JEKgs6t/XAOOg5cBWAr/XEiiv/tVQfj0Mg/3tuwooFRGSpRN9sb4gWI6bTu
# cXFt9IVwlG3LtKiakMYceXE8RXmwtUbXXIS8IPkAcYdKRtvKR0SgOYUDpfok/fpk
# W8taQyQDNXHdjL/GQVNfdHJnt8AqmBC5SNNzyTJPyIfmsDXn0OWw9C9nteOONAbd
# 11bSC7E4yfF72xNB9+kJY4UWbJjsz+PPu+FNdcgO6cIcYB8JhRvNNTm99+4sPRwW
# OFNSteR5Nc845AKkLaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MDdaMC8GCSqGSIb3DQEJBDEiBCBZFIiOL5mYKLjdP2amiy7RnQ/wZCoZ8WTFORFv
# crXD+DANBgkqhkiG9w0BAQEFAASCAgA27+giQJ3KPtJNxHTD4LMzWMSHVWVz1daZ
# /tbnyutYulkQsvt1z/f0xF/cYT1Ww+zHmRlJ6X/mArDCSaX+4E14t/SIMUF2q8dd
# ZzUSp+/Gn+Nkf6CaQA5MRsftLRsmEZSsN7IWhleALKNv7UyOUhGOM6yriRtv4k9g
# P4kqqVj1VK0WDfhiVmbqlwt/V1aCTKJ3M/CSzEfvLmFiorxkvqmcKoKB9pucjrJP
# qmISqVI/P3Ibr6qO090bOkPuiSpu3rtLlhwAwPFZ8ZwQWXx4sYFSXY2ZfI3jxV+d
# JnkOcFvxPjKfzPmhN10BU7lvA6Ru4KeVbQRP8FWF1Z2InYQjLxg0DiYi1z4EzJVv
# durPEEvVD56u+6mTUidVSiLmIyMPKVXaJemVGhRPAMZ4PhCiB6SPKbsCAn8oi84m
# Rc8Aqpam5O6sbJFRjEDcgABsiIbEUSkzGzDqS2zM4TBT7esgEaMYAYE1VMYN8MOW
# 9hAe/HOedX7Cazu/5NvVCMI7NOkcW/F7yIuP60It/OheSXedQ1cPAF15THFHbAG9
# kdMyhyygl+3w3u8glOf2ALcKeWTwDSqvARULD4wwMfty0dd3OSE1QvAFz4wxvwdw
# E3cDo23d7+E/2FpSJw/RmTR0oz4TLrVPC5oSTJSQoBGIuO1Syk/J6+AQ3qkBhSG9
# yxsWtiiIxQ==
# SIG # End signature block
