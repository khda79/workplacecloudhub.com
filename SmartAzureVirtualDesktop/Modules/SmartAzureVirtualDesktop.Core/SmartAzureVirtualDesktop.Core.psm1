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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCUmh7MVHv5+06t
# p2UoDwkt4Qj/37rwN8tZ2+G4eGVdPqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDMK1wpv4ua1AZNp3sY
# VBbh8GCEd/NBct3nXTAy3ZhVmzANBgkqhkiG9w0BAQEFAASCAYBuivtOY+wz2rcX
# B/i4R7JrB201ENflwOx1VgPqHtmIhXBewZS9F7aR1MyGwdxUI+eEkSzMku8bjipi
# blal7EkND7CYc/JAbwu0O0a0ZAf6lR/IwS4y2GP/1fnbjXe7LsvtHhylkKR+6N/N
# cqfgAl6U9S7nH5FZYyj/99hHhsQrOb9UFGkIifr9rVIX8rZyGc9bRwNwCp7hQHsR
# UeNgo3ABN30k1GfKp3gSG59CvANCJtVd1hY6JZAxMku21BE0iF4Gyl912L5/1PkE
# ZCf37luekZi+ZV19ybRhXweRsEqOXsblfS6n4Tcu1xuwXnLZl4NX+5uUZxf+mqNr
# eFFAvUO+eWzxiYIUvimO8Q23DO4GxRKDSPBDH5I3onvVED5bFP5vW8pHqMz9RZ4C
# 25qp8zY1GPX0OviKfZjXpJ15Eo0oGWdtiT8j0rDyW/p4OGOZ1eHZvq5nLyp95eOf
# ApcImv4AqmPblrj5lJZGSoenBPvVc2AOhx2r7CInstWFKHH/ym4=
# SIG # End signature block
