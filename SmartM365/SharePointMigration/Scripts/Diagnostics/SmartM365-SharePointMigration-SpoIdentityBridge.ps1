<#
.SYNOPSIS
    Reuse one existing PnP connection for local ShareGate source identity checks.
.VERSION
    1.0.1
#>
#Requires -Version 7.4

function Start-SmartM365SpoIdentityBridge {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Connection)
    $name = 'SmartM365SpoIdentity-' + [guid]::NewGuid().ToString('N')
    $helper = Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-SpoSourceIdentity.ps1'
    $worker = [powershell]::Create()
    $cancel = [Threading.CancellationTokenSource]::new()
    $script = {
        param($pipeName, $connection, $helperPath, $cancel)
        $ErrorActionPreference = 'Stop'
        Import-Module PnP.PowerShell -ErrorAction Stop
        . $helperPath
        while (-not $cancel.IsCancellationRequested) {
            $pipe = [IO.Pipes.NamedPipeServerStream]::new($pipeName, [IO.Pipes.PipeDirection]::InOut, 1,
                [IO.Pipes.PipeTransmissionMode]::Byte,
                ([IO.Pipes.PipeOptions]::CurrentUserOnly -bor [IO.Pipes.PipeOptions]::Asynchronous))
            try {
                $pipe.WaitForConnectionAsync($cancel.Token).GetAwaiter().GetResult()
                $reader = [IO.StreamReader]::new($pipe, [Text.Encoding]::UTF8, $false, 4096, $true)
                $writer = [IO.StreamWriter]::new($pipe, [Text.UTF8Encoding]::new($false), 4096, $true)
                $writer.AutoFlush = $true
                try {
                    $request = $reader.ReadLineAsync($cancel.Token).GetAwaiter().GetResult() | ConvertFrom-Json -ErrorAction Stop
                    if (-not $request -or $request.Nonce -notmatch '^[0-9a-f]{32}$') { throw 'Invalid identity request.' }
                    $identity = Get-SmartM365SpoSourceIdentity -Connection $connection -SiteUrl $request.SiteUrl `
                        -ListName $request.ListName -SourceFilePath $request.SourceFilePath -ExpectedItemId $request.ExpectedItemId
                    $response = @{ Nonce=$request.Nonce; Success=$true; Identity=$identity }
                }
                catch { $response = @{ Nonce=if ($request) { [string]$request.Nonce } else { '' }; Success=$false; Error=$_.Exception.Message } }
                if (-not $cancel.IsCancellationRequested) { $writer.WriteLine(($response | ConvertTo-Json -Depth 5 -Compress)) }
            }
            catch { if (-not $cancel.IsCancellationRequested) { Write-Warning 'A local identity request failed.' } }
            finally {
                if ($writer) { $writer.Dispose() }; if ($reader) { $reader.Dispose() }
                $pipe.Dispose()
                $writer = $null; $reader = $null; $request = $null
            }
        }
    }
    [void]$worker.AddScript($script.ToString()).AddArgument($name).AddArgument($Connection).AddArgument($helper).AddArgument($cancel)
    $handle = $worker.BeginInvoke()
    [pscustomobject]@{ PipeName=$name; Worker=$worker; Handle=$handle; Cancellation=$cancel }
}

function Stop-SmartM365SpoIdentityBridge {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Bridge)
    $Bridge.Cancellation.Cancel()
    if (-not $Bridge.Handle.AsyncWaitHandle.WaitOne(5000)) { throw 'Identity bridge is still completing a request; retry stopping it later.' }
    $null = $Bridge.Worker.EndInvoke($Bridge.Handle)
    $Bridge.Worker.Dispose()
    $Bridge.Cancellation.Dispose()
}
