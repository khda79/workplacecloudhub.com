<#
.SYNOPSIS
    Check a source file using a local in-memory PnP identity bridge.
.VERSION
    1.0.2
#>
#Requires -Version 5.1

function Test-SmartM365SpoSourceHost {
    param([Parameter(Mandatory)][string]$Url)
    return ([uri]$Url).DnsSafeHost -match '(?i)(^|\.)sharepoint\.(com|cn|de|us)$'
}

function Get-SmartM365BridgedSpoSourceIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidatePattern('^SmartM365SpoIdentity-[0-9a-f]{32}$')][string]$PipeName,
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$ListName,
        [Parameter(Mandatory)][string]$SourceFilePath,
        [Parameter(Mandatory)][ValidateRange(1,2147483647)][int]$ExpectedItemId
    )
    $pipe = [IO.Pipes.NamedPipeClientStream]::new('.', $PipeName, [IO.Pipes.PipeDirection]::InOut)
    $reader = $null; $writer = $null
    try {
        try { $pipe.Connect(5000) }
        catch { throw 'SPO source identity bridge is unavailable or timed out.' }
        $reader = [IO.StreamReader]::new($pipe, [Text.Encoding]::UTF8, $false, 4096, $true)
        $writer = [IO.StreamWriter]::new($pipe, [Text.UTF8Encoding]::new($false), 4096, $true)
        $writer.AutoFlush = $true
        $nonce = [guid]::NewGuid().ToString('N')
        $request = @{ Nonce=$nonce; SiteUrl=$SiteUrl; ListName=$ListName; SourceFilePath=$SourceFilePath; ExpectedItemId=$ExpectedItemId }
        $writer.WriteLine(($request | ConvertTo-Json -Compress))
        $read = $reader.ReadLineAsync()
        if (-not $read.Wait(60000)) { throw 'SPO source identity verification timed out.' }
        $response = $read.Result | ConvertFrom-Json -ErrorAction Stop
        if (-not $response -or $response.Nonce -cne $nonce -or $response.Success -ne $true -or -not $response.Identity) {
            throw 'SPO source identity verification failed or returned an invalid response.'
        }
        $identity = $response.Identity
        $decodedPath = @($SourceFilePath.Replace('\','/').Split('/') | ForEach-Object { [uri]::UnescapeDataString($_) }) -join '/'
        if ([int]$identity.SourceItemId -ne $ExpectedItemId -or
            -not [string]::Equals(([string]$identity.SiteUrl).TrimEnd('/'), $SiteUrl.TrimEnd('/'), [StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$identity.LibraryTitle, $ListName, [StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$identity.SourceFilePath, $decodedPath, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'SPO source identity response does not match the requested site, library, path and ID.'
        }
        return $identity
    }
    finally {
        if ($writer) { $writer.Dispose() }; if ($reader) { $reader.Dispose() }; $pipe.Dispose()
    }
}
