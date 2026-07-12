function Set-SmartCitrixCoreContext {
    [CmdletBinding()]
    param(
        [string]$RunId,
        [string]$RunOutputRoot,
        [string]$LatestOutputRoot,
        [string]$LogPath,
        [int]$RetentionMaxCsv = 30
    )

    $script:SmartCitrixCoreRunId = $RunId
    $script:SmartCitrixCoreRunOutputRoot = $RunOutputRoot
    $script:SmartCitrixCoreLatestOutputRoot = $LatestOutputRoot
    $script:SmartCitrixCoreLogPath = $LogPath
    $script:SmartCitrixCoreRetentionMaxCsv = $RetentionMaxCsv
}

function Get-SmartCitrixCoreContextValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    $variableName = "SmartCitrixCore$Name"
    $variable = Get-Variable -Name $variableName -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $variable -and $null -ne $variable.Value) { return $variable.Value }
    return $DefaultValue
}

function Write-SmartCitrixLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line

    $logPath = Get-SmartCitrixCoreContextValue -Name 'LogPath' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($logPath)) { return }

    $logFolder = Split-Path -Path $logPath -Parent
    if (-not (Test-Path -LiteralPath $logFolder)) {
        New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
    }
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Import-SmartCitrixPowerShellComponent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$Name,
        [switch]$Required
    )

    foreach ($componentName in $Name) {
        $getSnapinCommand = Get-Command -Name Get-PSSnapin -ErrorAction SilentlyContinue
        $addSnapinCommand = Get-Command -Name Add-PSSnapin -ErrorAction SilentlyContinue

        if ($null -ne $getSnapinCommand -and (Get-PSSnapin -Name $componentName -ErrorAction SilentlyContinue)) { continue }
        if (Get-Module -Name $componentName -ErrorAction SilentlyContinue) { continue }

        $loaded = $false
        if ($null -ne $getSnapinCommand -and $null -ne $addSnapinCommand -and (Get-PSSnapin -Registered -Name $componentName -ErrorAction SilentlyContinue)) {
            try {
                Add-PSSnapin -Name $componentName -ErrorAction Stop
                Write-SmartCitrixLog -Message ("Loaded PowerShell snap-in: {0}" -f $componentName)
                $loaded = $true
            }
            catch {
                if ($Required) { throw }
                Write-SmartCitrixLog -Level WARN -Message ("Failed to load snap-in '{0}': {1}" -f $componentName, $_.Exception.Message)
            }
        }

        if (-not $loaded -and (Get-Module -ListAvailable -Name $componentName -ErrorAction SilentlyContinue)) {
            try {
                Import-Module -Name $componentName -ErrorAction Stop
                Write-SmartCitrixLog -Message ("Loaded PowerShell module: {0}" -f $componentName)
                $loaded = $true
            }
            catch {
                if ($Required) { throw }
                Write-SmartCitrixLog -Level WARN -Message ("Failed to load module '{0}': {1}" -f $componentName, $_.Exception.Message)
            }
        }

        if (-not $loaded -and $Required) {
            throw ("Required Citrix PowerShell component '{0}' is not installed or could not be loaded." -f $componentName)
        }
        elseif (-not $loaded) {
            Write-SmartCitrixLog -Level WARN -Message ("Citrix PowerShell component not available: {0}" -f $componentName)
        }
    }
}

function ConvertTo-SmartCitrixCompactJson {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [int]$Depth = 20
    )

    if ($null -eq $Value) { return '' }
    try {
        return ($Value | ConvertTo-Json -Depth $Depth -Compress -ErrorAction Stop)
    }
    catch {
        return [string]$Value
    }
}

function Get-SmartCitrixObjectPropertyValue {
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

function Invoke-SmartCitrixSafeInventoryBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
    )

    try {
        return & $ScriptBlock
    }
    catch {
        Write-SmartCitrixLog -Level WARN -Message ("{0} skipped or partially failed: {1}" -f $Name, $_.Exception.Message)
        return @()
    }
}

function Invoke-SmartCitrixSdkCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CommandName,
        [string]$AdminAddress,
        [int]$MaxRecordCount = 250000,
        [hashtable]$Parameters
    )

    $command = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        Write-SmartCitrixLog -Level WARN -Message ("Citrix SDK command not found: {0}" -f $CommandName)
        return @()
    }

    $invokeParams = @{}
    if ($Parameters) {
        foreach ($key in $Parameters.Keys) { $invokeParams[$key] = $Parameters[$key] }
    }
    if (-not [string]::IsNullOrWhiteSpace($AdminAddress) -and $command.Parameters.ContainsKey('AdminAddress')) {
        $invokeParams['AdminAddress'] = $AdminAddress
    }
    if ($MaxRecordCount -gt 0 -and $command.Parameters.ContainsKey('MaxRecordCount')) {
        $invokeParams['MaxRecordCount'] = $MaxRecordCount
    }
    if ($command.Parameters.ContainsKey('ErrorAction')) {
        $invokeParams['ErrorAction'] = 'Stop'
    }

    Write-SmartCitrixLog -Message ("Running Citrix SDK command: {0}" -f $CommandName)
    return @(& $command @invokeParams)
}

function Test-SmartCitrixSimpleValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $true }
    return ($Value -is [string] -or
        $Value -is [char] -or
        $Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int] -or
        $Value -is [long] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal] -or
        $Value -is [datetime] -or
        $Value -is [guid] -or
        $Value -is [timespan])
}

function ConvertTo-SmartCitrixFlatRow {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory = $true)][string]$SourceObject,
        [string]$RunId = $(Get-SmartCitrixCoreContextValue -Name 'RunId' -DefaultValue (Get-Date -Format 'yyyyMMdd_HHmmss'))
    )

    $row = [ordered]@{
        RunId        = $RunId
        SourceObject = $SourceObject
    }

    if ($null -eq $InputObject) { return [pscustomobject]$row }

    foreach ($property in $InputObject.PSObject.Properties) {
        if ($property.Name -in @('RunId', 'SourceObject')) { continue }
        if (Test-SmartCitrixSimpleValue -Value $property.Value) {
            $row[$property.Name] = $property.Value
        }
        else {
            $row[$property.Name] = ConvertTo-SmartCitrixCompactJson -Value $property.Value
        }
    }

    return [pscustomobject]$row
}

function ConvertTo-SmartCitrixFlatRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$SourceObject,
        [string]$RunId = $(Get-SmartCitrixCoreContextValue -Name 'RunId' -DefaultValue (Get-Date -Format 'yyyyMMdd_HHmmss'))
    )

    $flatRows = @()
    foreach ($row in @($Rows)) {
        $flatRows += ConvertTo-SmartCitrixFlatRow -InputObject $row -SourceObject $SourceObject -RunId $RunId
    }
    return @($flatRows)
}

function Test-SmartCitrixFileLocked {
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

function Remove-SmartCitrixOldFiles {
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
            if (-not (Test-SmartCitrixFileLocked -Path $file.FullName)) {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Write-SmartCitrixLog -Message ("Removed old file: {0}" -f $file.FullName)
            }
        }
        catch {
            Write-SmartCitrixLog -Level WARN -Message ("Failed to remove old file '{0}': {1}" -f $file.FullName, $_.Exception.Message)
        }
    }
}

function Export-SmartCitrixCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [string]$RunOutputRoot = $(Get-SmartCitrixCoreContextValue -Name 'RunOutputRoot' -DefaultValue ''),
        [string]$LatestOutputRoot = $(Get-SmartCitrixCoreContextValue -Name 'LatestOutputRoot' -DefaultValue ''),
        [string]$RunId = $(Get-SmartCitrixCoreContextValue -Name 'RunId' -DefaultValue (Get-Date -Format 'yyyyMMdd_HHmmss')),
        [string]$Encoding = 'UTF8'
    )

    if ([string]::IsNullOrWhiteSpace($RunOutputRoot)) { throw 'Export-SmartCitrixCsv: RunOutputRoot is missing.' }
    if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { throw 'Export-SmartCitrixCsv: LatestOutputRoot is missing.' }

    $historyPath = Join-Path -Path $RunOutputRoot -ChildPath ("{0}_{1}.csv" -f $Name, $RunId)
    $latestPath = Join-Path -Path $LatestOutputRoot -ChildPath ("{0}.csv" -f $Name)

    foreach ($folder in @((Split-Path -Path $historyPath -Parent), (Split-Path -Path $latestPath -Parent))) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }
    }

    $Rows | Export-Csv -LiteralPath $historyPath -NoTypeInformation -Encoding $Encoding
    Write-SmartCitrixLog -Message ("Exported {0} row(s) to {1}" -f @($Rows).Count, $historyPath)

    try {
        Copy-Item -LiteralPath $historyPath -Destination $latestPath -Force -ErrorAction Stop
        Write-SmartCitrixLog -Message ("Latest CSV updated: {0}" -f $latestPath)
    }
    catch {
        Write-SmartCitrixLog -Level WARN -Message ("Latest CSV copy failed: {0}" -f $_.Exception.Message)
    }

    $retentionMaxCsv = [int](Get-SmartCitrixCoreContextValue -Name 'RetentionMaxCsv' -DefaultValue 30)
    if ($retentionMaxCsv -gt 0) {
        Remove-SmartCitrixOldFiles -FolderPath $RunOutputRoot -FilePattern ("{0}_*.csv" -f $Name) -MaxFiles $retentionMaxCsv
    }
}

function New-SmartCitrixSummaryRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][int]$Count,
        [string]$RunId = $(Get-SmartCitrixCoreContextValue -Name 'RunId' -DefaultValue (Get-Date -Format 'yyyyMMdd_HHmmss')),
        [string]$Details = ''
    )

    [pscustomobject]@{
        RunId   = $RunId
        Name    = $Name
        Count   = $Count
        Details = $Details
    }
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBqNnv+tgcswqXf
# Jp0Nlu48kk2LLL88nr3HsrV155gSy6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDZPwRFP7+3OdAgBRuj8dvcrXrRBmbbq41ZfQkQp3EQfzANBgkqhkiG9w0B
# AQEFAASCAYCvN1c3H6Y2R4IZne8AYbdyhJTkNyze4p1zeNKYm+9Oa1bQyC2DqZna
# +pMOi1UrLsUSsV0EJCmVThNPFE0Pr1KAKJ48tRf9oDGCFDHIEG9H5RaqgtwHAZVJ
# gyPp3nyieARmnb78uICTUKK3+moCBfS6rTF5sexJmKEIT9yhABtAF2bn2sOfP0LQ
# maPzhwU9vjNDtgenIq004OlpsY4psYbcfixC06rAzsqZchk4b18SXNZSIN845TCq
# d5G6B/PUg3igmTWZOUVoSncHKbNeWjLuQRQUtlEwAbGQcjn4AN6GuwugE4+Eynhe
# USYij9cPLlQMZJgExOgJw7ujgp//aSgJt4TXt52el0JZhePHSOq+HARvCFQTG24r
# afVwfGE6G5WC0ruRQ1AGQdmPDShgXo4emSPqiEIpIiVu0kq7EgqAVzqOYurqVAVr
# g4xfife7FAsKzZ2N3XjN5XlqMq0d24hdMg+IBzkLCNxMw6KAiLTa6KTh5L21a6jW
# 3aJNc7h+4Zw=
# SIG # End signature block
