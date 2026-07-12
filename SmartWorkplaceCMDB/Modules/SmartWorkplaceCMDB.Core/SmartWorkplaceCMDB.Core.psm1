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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDeo1pgF7/TJ5Dr
# xHoyJYbx1JSTN/ET8uJ4t7IFka0uaKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDUJeh+5UFn7ONNQFYL
# 4wnRBIjcEKdIQByOyzeCFZViIDANBgkqhkiG9w0BAQEFAASCAYB8+D83UIHSV9Gg
# TmqfnmxSlcaaFM4C9Q6rZf/rWCQbPtVSUsPh3P2YwzwaKFWJ/cyEdvOElJaI6l8W
# 8eDZyCJfEa1eY98nhKpv91gdnSnliMNXqlxryoUHuUFPvWyMaWmPjGF60/9YTsBl
# HE3sGzBbbeFhrKcwb2ilkwdqjTBeGAGDSJXPOPSX7cBuChopoqUmnVpN+KCgRyx7
# obG52Qkf78hL19IStFMpw0J9jfM2I+3/eUNhoMethC/gcQklkVcvrfgQ90EN5MvN
# 8C6Z8vbdelGcioRFiORiGmj8jctt4KevcWKNlQ3q6wygJL07uMMVCiYl9+EFvxZD
# Gm7DIuAZV06jKAb8hXX9dwiAvjey72egswCfMmf4XQK7dQDcBjaidHK5u7mjD05X
# tv5Lq6vDfcHOolX5in5m/zpfRGPv1wFbdBDXqnI00V2/RxCWE9U47PvtTSnERZPB
# fDRQJeNWP9QtY77rIGPo1QtsOuIog4012DHVq9Bx6nFegodBR0o=
# SIG # End signature block
