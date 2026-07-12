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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDLPW4YDXA99EMB
# P/xh/KpuffxpP05VEg9LH/xOMktxAKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBba20KIWez0lmKQObO
# Mo4xWZRGGy+inj3js7NDSQY4bjANBgkqhkiG9w0BAQEFAASCAYCFc6Cw12HYMxH5
# pmO/bK4d3k+R1OySLGr/8OyJaaqYQY7NLaHJ5frmKHqYTOUQr2U28wDJqgRtJpwu
# Hchzpuws2gOM7f+rcEeoQk+bXA5jsrRlKYP2urTX2B1iWrulo2v9p1sO+N4PmY5W
# 4G+Ntn6XdgMSnFnWTJ5hl2ZXX5hXeuTMyEir/ibK2QI8ILl6pgRZLl7avvMuMnBU
# 6PpJsaiO6KMljS/oneQMlXRXn3vzocohvqJPr2/irIq6kRASXf5mvmN8p0wreOLe
# PVzVjq42Ij+l63VNacVY4w9JrW2Froq/HDfPZJZA5sN704+E15wykjfa/XwgKqTC
# 3vPgUilfw4Fg3itZC0VnZRGLmwJh43U+8BwyPx83GRL5dSgaB5m52c9Xq9K+B3c0
# jAaxFl5wV1FJ7UPk5ni6SbAjU7d6/cY0nGHA38FPQPsIV23zgl2F+t19eUzsafvE
# FJqJOBB0RVc45YJb1HCRQ2dkPQWzzZZWLbrSOlujyZ3aYkdivRo=
# SIG # End signature block
