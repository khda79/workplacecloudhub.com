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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBqNnv+tgcswqXf
# Jp0Nlu48kk2LLL88nr3HsrV155gSy6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDZPwRFP7+3OdAgBRuj
# 8dvcrXrRBmbbq41ZfQkQp3EQfzANBgkqhkiG9w0BAQEFAASCAYBJ75UsFL2uvBz9
# 7kOIM6Fq2Ar2+f9cLQn6qWLFlBNwiw6xNMOl82t9jex6CsZPyeHy/pBoViQGMWps
# 4hhnudzwbWYez7G8ZjBXPbmiVOLznHIOuSC0xH6JjCev5E6KG4FU2LvMXhLfp7KE
# DBLu9sfX3OneJ8mmEvSARHKY+uARAWWpUzYgZ/FLTQsl53iChhNA+2+uK2mM+r2s
# C4svU3DjF5OmvaLcS4N3NUIBr1+VmAgZTcWonXGIBeU5a16hlU8jbky46raYVtnk
# 7nvkf2PhdlFmguKakjnbG2NquaCVTWL0DjlKOced9PoFBw1/Bu+8CdBIGcA/Onoo
# kz+zYlGLOtiu8PHEcYXqpRBJLfGjy3asasiGp0FC8WrByyGK3fnGElBSugtRSHV3
# k6AowEUobZWHSxQ0UfOMmuyCOEdk8Hm88jtigsUxltsWnIrVd6dhrynPf2H94HpS
# DbWVRlabjpBunmbvXGFyHAraIYKJczWXU276xY5yNyhE15+vL8E=
# SIG # End signature block
