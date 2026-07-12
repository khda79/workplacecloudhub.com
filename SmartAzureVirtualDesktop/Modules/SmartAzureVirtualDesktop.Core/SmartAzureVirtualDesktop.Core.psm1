function Set-SmartAvdCoreContext {
    [CmdletBinding()]
    param(
        [string]$RunId,
        [string]$RunOutputRoot,
        [string]$LatestOutputRoot,
        [string]$LogPath,
        [int]$RetentionMaxCsv = 30
    )

    $script:SmartAvdCoreRunId = $RunId
    $script:SmartAvdCoreRunOutputRoot = $RunOutputRoot
    $script:SmartAvdCoreLatestOutputRoot = $LatestOutputRoot
    $script:SmartAvdCoreLogPath = $LogPath
    $script:SmartAvdCoreRetentionMaxCsv = $RetentionMaxCsv
}

function Get-SmartAvdCoreContextValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    $variableName = "SmartAvdCore$Name"
    $variable = Get-Variable -Name $variableName -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $variable -and $null -ne $variable.Value) { return $variable.Value }
    return $DefaultValue
}

function Write-SmartAvdLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line

    $logPath = Get-SmartAvdCoreContextValue -Name 'LogPath' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($logPath)) { return }

    $logFolder = Split-Path -Path $logPath -Parent
    if (-not (Test-Path -LiteralPath $logFolder)) {
        New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
    }
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Import-SmartAvdRequiredModule {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw ("Required PowerShell module '{0}' is not installed. Install it with: Install-Module {0} -Scope CurrentUser" -f $Name)
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function ConvertTo-SmartAvdCompactJson {
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

function Get-SmartAvdObjectPropertyValue {
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

function Get-SmartAvdNestedPropertyValue {
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

function Invoke-SmartAvdArmGetPaged {
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

            $value = Get-SmartAvdObjectPropertyValue -InputObject $content -PropertyName @('value')
            if ($null -ne $value) { $rows += @($value) } else { $rows += $content }

            $nextLink = Get-SmartAvdObjectPropertyValue -InputObject $content -PropertyName @('nextLink', '@odata.nextLink')
            $nextUri = if ([string]::IsNullOrWhiteSpace([string]$nextLink)) { $null } else { [string]$nextLink }
        }
        catch {
            Write-SmartAvdLog -Level WARN -Message ("{0} failed: {1}" -f $Operation, $_.Exception.Message)
            break
        }
    }

    return @($rows)
}

function Invoke-SmartAvdSafeInventoryBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
    )

    try {
        return & $ScriptBlock
    }
    catch {
        Write-SmartAvdLog -Level WARN -Message ("{0} skipped or partially failed: {1}" -f $Name, $_.Exception.Message)
        return @()
    }
}

function Test-SmartAvdFileLocked {
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

function Remove-SmartAvdOldFiles {
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
            if (-not (Test-SmartAvdFileLocked -Path $file.FullName)) {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Write-SmartAvdLog -Message ("Removed old file: {0}" -f $file.FullName)
            }
        }
        catch {
            Write-SmartAvdLog -Level WARN -Message ("Failed to remove old file '{0}': {1}" -f $file.FullName, $_.Exception.Message)
        }
    }
}

function Export-SmartAvdCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [string]$RunOutputRoot = $(Get-SmartAvdCoreContextValue -Name 'RunOutputRoot' -DefaultValue ''),
        [string]$LatestOutputRoot = $(Get-SmartAvdCoreContextValue -Name 'LatestOutputRoot' -DefaultValue ''),
        [string]$RunId = $(Get-SmartAvdCoreContextValue -Name 'RunId' -DefaultValue (Get-Date -Format 'yyyyMMdd_HHmmss')),
        [ValidateSet('ASCII', 'BigEndianUnicode', 'Default', 'OEM', 'Unicode', 'UTF7', 'UTF8', 'UTF32')]
        [string]$Encoding = 'UTF8'
    )

    if ([string]::IsNullOrWhiteSpace($RunOutputRoot)) { throw 'Export-SmartAvdCsv: RunOutputRoot is missing.' }
    if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { throw 'Export-SmartAvdCsv: LatestOutputRoot is missing.' }

    $historyPath = Join-Path -Path $RunOutputRoot -ChildPath ("{0}_{1}.csv" -f $Name, $RunId)
    $latestPath = Join-Path -Path $LatestOutputRoot -ChildPath ("{0}.csv" -f $Name)

    foreach ($folder in @((Split-Path -Path $historyPath -Parent), (Split-Path -Path $latestPath -Parent))) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }
    }

    $Rows | Export-Csv -LiteralPath $historyPath -NoTypeInformation -Encoding $Encoding
    Write-SmartAvdLog -Message ("Exported {0} row(s) to {1}" -f @($Rows).Count, $historyPath)

    try {
        Copy-Item -LiteralPath $historyPath -Destination $latestPath -Force -ErrorAction Stop
        Write-SmartAvdLog -Message ("Latest CSV updated: {0}" -f $latestPath)
    }
    catch {
        Write-SmartAvdLog -Level WARN -Message ("Latest CSV copy failed: {0}" -f $_.Exception.Message)
    }

    $retentionMaxCsv = [int](Get-SmartAvdCoreContextValue -Name 'RetentionMaxCsv' -DefaultValue 30)
    if ($retentionMaxCsv -gt 0) {
        Remove-SmartAvdOldFiles -FolderPath $RunOutputRoot -FilePattern ("{0}_*.csv" -f $Name) -MaxFiles $retentionMaxCsv
    }
}

function Connect-SmartAvdCloudSession {
    [CmdletBinding()]
    param(
        [string]$TenantId,
        [switch]$Connect,
        [switch]$UseDeviceCode
    )

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if ($Connect -or $null -eq $context) {
        Write-SmartAvdLog -Message 'Connecting to Azure.'
        $connectParams = @{}
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connectParams.Tenant = $TenantId }
        if ($UseDeviceCode) { $connectParams.UseDeviceAuthentication = $true }
        Connect-AzAccount @connectParams | Out-Null
        $context = Get-AzContext -ErrorAction SilentlyContinue
    }
    return $context
}

function Get-SmartAvdResourceGroupNameFromId {
    [CmdletBinding()]
    param([string]$ResourceId)

    if ([string]::IsNullOrWhiteSpace($ResourceId)) { return '' }
    $match = [regex]::Match($ResourceId, '/resourceGroups/([^/]+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) { return [uri]::UnescapeDataString($match.Groups[1].Value) }
    return ''
}

function Get-SmartAvdResourceNameFromId {
    [CmdletBinding()]
    param([string]$ResourceId)

    if ([string]::IsNullOrWhiteSpace($ResourceId)) { return '' }
    return [uri]::UnescapeDataString(($ResourceId.TrimEnd('/') -split '/')[-1])
}

function Get-SmartAvdSubscriptionResources {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceType,
        [Parameter(Mandatory)][string]$ApiVersion
    )

    $encodedSubscriptionId = [uri]::EscapeDataString($SubscriptionId)
    $encodedFilter = [uri]::EscapeDataString("resourceType eq '$ResourceType'")
    $uri = "https://management.azure.com/subscriptions/{0}/resources?`$filter={1}&api-version={2}" -f $encodedSubscriptionId, $encodedFilter, $ApiVersion
    return Invoke-SmartAvdArmGetPaged -Uri $uri -Operation ("Resources {0}" -f $ResourceType)
}

function Get-SmartAvdResourceById {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResourceId,
        [Parameter(Mandatory)][string]$ApiVersion,
        [string]$Operation = 'ARM resource'
    )

    $separator = if ($ResourceId.Contains('?')) { '&' } else { '?' }
    $uri = "https://management.azure.com{0}{1}api-version={2}" -f $ResourceId, $separator, $ApiVersion
    try {
        $response = Invoke-AzRestMethod -Method GET -Uri $uri -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($response.Content)) { return $null }
        return $response.Content | ConvertFrom-Json -Depth 100
    }
    catch {
        Write-SmartAvdLog -Level WARN -Message ("{0} failed for '{1}': {2}" -f $Operation, $ResourceId, $_.Exception.Message)
        return $null
    }
}

function Get-SmartAvdChildResources {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ParentResourceId,
        [Parameter(Mandatory)][string]$ChildPath,
        [Parameter(Mandatory)][string]$ApiVersion,
        [Parameter(Mandatory)][string]$Operation
    )

    $uri = "https://management.azure.com{0}/{1}?api-version={2}" -f $ParentResourceId.TrimEnd('/'), $ChildPath.Trim('/'), $ApiVersion
    return Invoke-SmartAvdArmGetPaged -Uri $uri -Operation $Operation
}

function Get-SmartAvdAgeInDays {
    [CmdletBinding()]
    param([AllowNull()]$DateTimeValue)

    if ($null -eq $DateTimeValue -or [string]::IsNullOrWhiteSpace([string]$DateTimeValue)) { return $null }
    try {
        $parsed = [datetime]$DateTimeValue
        return [math]::Round(((Get-Date).ToUniversalTime() - $parsed.ToUniversalTime()).TotalDays, 2)
    }
    catch {
        return $null
    }
}

function ConvertTo-SmartAvdBoolString {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { return $Value.ToString() }
    return [string]$Value
}

Export-ModuleMember -Function `
    Set-SmartAvdCoreContext, `
    Write-SmartAvdLog, `
    Import-SmartAvdRequiredModule, `
    ConvertTo-SmartAvdCompactJson, `
    Get-SmartAvdObjectPropertyValue, `
    Get-SmartAvdNestedPropertyValue, `
    Invoke-SmartAvdArmGetPaged, `
    Invoke-SmartAvdSafeInventoryBlock, `
    Export-SmartAvdCsv, `
    Connect-SmartAvdCloudSession, `
    Get-SmartAvdResourceGroupNameFromId, `
    Get-SmartAvdResourceNameFromId, `
    Get-SmartAvdSubscriptionResources, `
    Get-SmartAvdResourceById, `
    Get-SmartAvdChildResources, `
    Get-SmartAvdAgeInDays, `
    ConvertTo-SmartAvdBoolString

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCUmh7MVHv5+06t
# p2UoDwkt4Qj/37rwN8tZ2+G4eGVdPqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDMK1wpv4ua1AZNp3sYVBbh8GCEd/NBct3nXTAy3ZhVmzANBgkqhkiG9w0B
# AQEFAASCAYBBxPT4h0GNEM+CxUVaoeFI5nY5E4VgJXQYHrDmTnNIA091ZVukj9gx
# vnmbEqHqCarov6AxRbtE5LXe8PA7juWeMD1uEykn3vMfy1WBnL40p6BXhr5Ni9Cg
# 4INVb/1WW3B0f4lRaQc9ZR/zHPUfqDBho34s2LnixniGLjLsuNEiJiMO4Zk7dFVn
# WCXYZNEralYQFvUiR0ywjOsjheHA5unE/u3pWEjybaExEDiSCBxxo2zmn43aAPtP
# 72HuS/7xcWzjwZ26YB/Z72d8N2fnqoOim2ql1q0UEvmbNv3liaSHGa88AdHBm8ou
# 6+sx0AkuPNhqX+FlwmgiaDTCpyCXsNnu6NX6M2xzJ3VaQ04vg4LvXrGhnFCaRLaY
# x8iWYSrMewScub3JnAhw0z3PRJnggPsU1awgfVHhiuZUx5O7ibN1M89RBW2Xk5Q4
# F9Oshmdn63qun3U6Vc3cU8GdyXredEPGMiZe3dhtIE59kBOOPmmdJt4hn/LSpJGF
# 8WobgU3+k+I=
# SIG # End signature block
