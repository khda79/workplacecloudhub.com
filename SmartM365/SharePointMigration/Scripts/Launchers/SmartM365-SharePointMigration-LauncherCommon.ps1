function Get-LauncherProjectRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LauncherScriptRoot
    )

    return (Resolve-Path -LiteralPath (Join-Path -Path $LauncherScriptRoot -ChildPath '..\..\..')).ProviderPath
}

function Import-LauncherMigrationConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$MigrationName,

        [string]$ConfigFileName = 'migration.config.psd1'
    )

    $migrationRoot = Join-Path -Path $ProjectRoot -ChildPath ("Migrations\{0}" -f $MigrationName)
    $configPath = Join-Path -Path $migrationRoot -ChildPath $ConfigFileName
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw "Migration config not found: $configPath"
    }

    [pscustomobject]@{
        Root = $migrationRoot
        Path = $configPath
        Data = Import-PowerShellDataFile -LiteralPath $configPath
    }
}

function Resolve-LauncherMigrationPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MigrationRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path -Path $MigrationRoot -ChildPath $Path
}

function Use-DefaultParameterValue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BoundParameters,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        $Value
    )

    if ($BoundParameters.ContainsKey($Name)) {
        return $null
    }

    return $Value
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDQr8ONgmloA6b9
# 0Rdcs5ln/vVSKZPo7Zn8rzOpJDhfQKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCXkS61dVz50MtJOVXj
# d1Rys+MOxIFOQaKdTe3YlGs5QDANBgkqhkiG9w0BAQEFAASCAYBpoS1ZwdfjrdNC
# iBqZfhofPYvHP/CfMVD+O6TTDTUCleo20uiipmNMnE/dOpiWox3jzZV046ZyZUiR
# x93dhakiY7My9sbYTdamY88Rrr4W84F9F5oV3KlaY9p9Yq1a4b0DyxLFYdssPVhi
# xdAWpcJGwYmQO4M0ANcyiPb5TpyLF47HoXkS8gdLea5Ql2FFLQ8/PB0xjFIMROkf
# ZP6iKf/wpV4/9Z2477nyHTSY1JT/yVbc6yq+sE8jP8TbxDC/4uA57aCdQf45nVyB
# NQi3I2YDygmGYLpjk8ooNuR4IsjH/0z88cSJyRsS5wBpYacZ2Tp/86BQpYKQwdzM
# kGwkALHAg3pSBbcctM7ND7R/Wb9XSaU8CeDFn0lFQZkAJwx/Y1CEO+2I7SrkBbPE
# tOb4VTVAuNIcWp5UpcOvFQ2pdu25Imlpg9Zo2mq2g8ukGWQ2AeHNiFFyDi9DdIH1
# FE4Bgm3/o8dHfTwme0WJEojiVRsBkUzMkpCDmRFS/oikp/x2Zv0=
# SIG # End signature block
