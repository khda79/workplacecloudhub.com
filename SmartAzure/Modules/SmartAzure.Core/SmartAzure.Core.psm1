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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDLPW4YDXA99EMB
# P/xh/KpuffxpP05VEg9LH/xOMktxAKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCBba20KIWez0lmKQObOMo4xWZRGGy+inj3js7NDSQY4bjANBgkqhkiG9w0B
# AQEFAASCAYA+y7djvIha4rDimxdpBh8pSPBrsQFC62nTk4Zq3tdZkAB0MLxxxXGg
# ALF2oMLcCtP0r+go62WSKobSB6EK1Sm2kEu8+5oXe71WJKOlVZ0yNXEQbcJbIdHT
# QlQ3+Q8ll1klHNAdFCnridfNNc2RU4yy7E98maclHq6xsODXsHB5ISLCbWNwdBNT
# zufGMRABCxu0VmurOD8cX3XnZUNV9+6HkXM9/MwiYpGwpw93JucDDHl6BXWkK9v2
# dJwBFxzKFYbtmJ8jYHB7ZOBifCiZYvRBsPxacAeiW0kpC1pTAh67jCKLgRPuAqOY
# wiKuQgpy5DXgo9hvXzro1JozamoE5FuvCz0w9hbHf8QhNhfp3u9WvmQhoQyo5kv7
# 2xolOWKXMAFXNR0M0+Iqn0Er+yLQXK6RO0CmxxkTEiXTFA7QpLEJ8zX22nqpa++p
# D4L1vP/jf+cfideICigVsmQj5mQVbHzm1NAqiJdGYUqcl4N31H1mihS7YT4yBnFt
# DAzEHlBPJxc=
# SIG # End signature block
