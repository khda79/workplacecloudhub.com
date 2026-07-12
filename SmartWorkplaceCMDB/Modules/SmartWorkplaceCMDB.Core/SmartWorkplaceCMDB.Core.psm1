# SmartWorkplaceCMDB.Core
# Version: 0.1.1

$script:SmartWorkplaceCMDBCoreVersion = '0.1.1'

function Get-SmartWorkplaceCMDBProjectRoot {
    [CmdletBinding()]
    param()

    $moduleRoot = Split-Path -Parent $PSScriptRoot
    return (Split-Path -Parent $moduleRoot)
}

function Read-SmartWorkplaceCMDBJsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "JSON file not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "JSON file is empty: $Path"
    }

    return ($raw | ConvertFrom-Json -ErrorAction Stop)
}

function Resolve-SmartWorkplaceCMDBTenantPath {
    [CmdletBinding()]
    param(
        [string]$Tenant = 'Default',
        [string]$DataRootPath,
        [string]$DataAllRootPath,
        [string]$LatestOutputRootPath,
        [string]$LogRootPath
    )

    $projectRoot = Get-SmartWorkplaceCMDBProjectRoot
    $tenantKey = if ([string]::IsNullOrWhiteSpace($Tenant)) { 'Default' } else { $Tenant }

    if ([string]::IsNullOrWhiteSpace($DataRootPath)) {
        $DataRootPath = Join-Path -Path $projectRoot -ChildPath ("Data\Tenants\{0}" -f $tenantKey)
    }

    if ([string]::IsNullOrWhiteSpace($DataAllRootPath)) {
        $DataAllRootPath = Join-Path -Path $DataRootPath -ChildPath 'DATA-ALL'
    }

    if ([string]::IsNullOrWhiteSpace($LatestOutputRootPath)) {
        $LatestOutputRootPath = Join-Path -Path $DataRootPath -ChildPath 'DATA-LAST'
    }

    if ([string]::IsNullOrWhiteSpace($LogRootPath)) {
        $LogRootPath = Join-Path -Path $DataRootPath -ChildPath 'LOG-ALL'
    }

    [pscustomobject]@{
        TenantKey            = $tenantKey
        ProjectRootPath      = $projectRoot
        DataRootPath         = $DataRootPath
        DataAllRootPath      = $DataAllRootPath
        LatestOutputRootPath = $LatestOutputRootPath
        LogRootPath          = $LogRootPath
        CmdbLatestPath       = Join-Path -Path $LatestOutputRootPath -ChildPath 'CMDB'
        PowerBILatestPath    = Join-Path -Path $LatestOutputRootPath -ChildPath 'PowerBI'
    }
}

function Initialize-SmartWorkplaceCMDBTenantFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Paths
    )

    foreach ($path in @(
            $Paths.DataRootPath,
            $Paths.DataAllRootPath,
            $Paths.LatestOutputRootPath,
            $Paths.LogRootPath,
            $Paths.CmdbLatestPath,
            $Paths.PowerBILatestPath
        )) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Export-SmartWorkplaceCMDBCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string]$Path,

        [string[]]$Columns
    )

    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $tempPath = "{0}.tmp.{1}.csv" -f $Path, ([guid]::NewGuid().ToString('N'))

    try {
        if ($InputObject.Count -gt 0) {
            $InputObject | Export-Csv -LiteralPath $tempPath -NoTypeInformation -Encoding UTF8 -Force
        }
        elseif ($Columns -and $Columns.Count -gt 0) {
            ($Columns -join ',') | Set-Content -LiteralPath $tempPath -Encoding UTF8 -Force
        }
        else {
            '' | Set-Content -LiteralPath $tempPath -Encoding UTF8 -Force
        }

        Move-Item -LiteralPath $tempPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

Export-ModuleMember -Function @(
    'Get-SmartWorkplaceCMDBProjectRoot',
    'Read-SmartWorkplaceCMDBJsonFile',
    'Resolve-SmartWorkplaceCMDBTenantPath',
    'Initialize-SmartWorkplaceCMDBTenantFolder',
    'Export-SmartWorkplaceCMDBCsv'
)



# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDeo1pgF7/TJ5Dr
# xHoyJYbx1JSTN/ET8uJ4t7IFka0uaKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDUJeh+5UFn7ONNQFYL4wnRBIjcEKdIQByOyzeCFZViIDANBgkqhkiG9w0B
# AQEFAASCAYAu7T5Pa8sleJQIv9HzgzGOqlh8wct8Sk9JFkm4dzf7R0kQNOXGSTO+
# nox+C/zkZCWey0W3whpYs6OQcNGoB7lzsZUQnQn7RyVqKVDF+bn9nQdMD/0bJsv7
# Z2uj68iVEqj0qe4W/13kxlTXg6I5zXcL5iB3yvnBJQ4z1iKmCG1yiRY3ov3Y+Px6
# GQgfLLIWEk5ogAhgmCr3oBfvs7UENzQ0S76fcgIWXfI1vH4fU4H/09HpDKy3YdRx
# 3obcwtoIUm9LtZBL+km6ulHnNxXzz+7HMU9nlOqmHLxyLCBVsuZVsvViSFds8cS0
# gLjTolX1QyAPwUa34w+Z6FGFqysQd5OLz5LgjfnompW+AwElA4l2zlddMIpjOFTh
# 6SHHFeTM78VHfYoXpAsBuOiPSMU9QRWXFoxJmKStgWLCh2QTp/3pmnWsnAzhUq19
# 3cMxOXtDNoxBSOq4MypB33srh7kFYXa61rIvUzJfzTamYbJ83bDkv4r6p0xu8FVu
# RZd9sZacoXc=
# SIG # End signature block
