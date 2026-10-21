# Shared local lifecycle for SmartInventory entry points. Compatible with PowerShell 5.1.
<#
.VERSION
1.0.2
#>
function Start-SmartM365InventoryRun {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ScriptPath,[switch]$RequireReturn)
    $previousVariable = Get-Variable SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue
    $previous = if ($previousVariable) { @{Value=$previousVariable.Value} } else { $null }
    $state = [pscustomobject]@{
        ScriptPath=$ScriptPath; ScriptName=[IO.Path]::GetFileNameWithoutExtension($ScriptPath)
        StartedAt=Get-Date; ReachedTerminal=$false; InCleanup=$false; ErrorRecord=$null
        Initialized=$false; TranscriptOwned=$false; Finalized=$false; Finalizing=$false; ExitCode=0
        PendingCompletion=$null; PendingBanner=$null
        RequireReturn=[bool]$RequireReturn; ReturnedNormally=$false; ExplicitExit=$false
        GraphOwned=$false; ExchangeOwned=$false; TeamsOwned=$false
        Previous=$previous; PublicationMask=@{}; CleanupWarnings=New-Object 'Collections.Generic.List[string]'
    }
    $global:SmartM365InventoryRun = $state
    return $state
}

function Disable-SmartM365InterruptedPublication {
    param([Parameter(Mandatory)]$State)
    foreach ($name in @('EnableSharePointUpload','EnableTeamsNotifications')) {
        if (-not $State.PublicationMask.ContainsKey($name)) {
            $value = Get-Variable $name -Scope Global -ErrorAction SilentlyContinue
            $State.PublicationMask[$name] = @{Exists=($null -ne $value);Value=$(if($value){$value.Value}else{$null})}
        }
        Set-Variable $name -Value $false -Scope Global
    }
}

function Test-SmartM365InterruptedRun {
    $state=Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    return ($null -ne $state -and ($state.ExitCode -eq 130 -or ($state.InCleanup -and -not $state.ErrorRecord -and
        (-not $state.ReachedTerminal -or ($state.Finalizing -and $state.RequireReturn -and -not $state.ReturnedNormally -and -not $state.ExplicitExit)))))
}

function Invoke-SmartM365RunCleanup {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State,[Parameter(Mandatory)][string]$Name,[Parameter(Mandatory)][scriptblock]$Action)
    try { & $Action | Out-Null }
    catch {
        $message = "Cleanup failed ($Name): $($_.Exception.Message)"
        $State.CleanupWarnings.Add($message)
        if ($State.Initialized -and (Get-Command WriteLog -ErrorAction SilentlyContinue)) {
            try { WriteLog -Message $message -Level WARNING } catch {}
        }
        # Cleanup reporting must never prevent the other resources being released.
        try { [Console]::Error.WriteLine($message) } catch {}
    }
}

function Set-SmartM365InventoryConnectionOwnership {
    param([ValidateSet('Graph','Exchange','Teams')][string]$Connection,[bool]$Owned)
    $state=Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if($state){$state.($Connection+'Owned')=$Owned}
}

function Close-SmartM365InventoryConnections {
    param([Parameter(Mandatory)]$State)
    foreach($entry in @(@('Graph','Disconnect-MgGraph'),@('Exchange','Disconnect-ExchangeOnline'),@('Teams','Disconnect-MicrosoftTeams'))){
        $kind=$entry[0];$command=$entry[1]
        if($State.($kind+'Owned') -and (Get-Command $command -ErrorAction SilentlyContinue)){
            Invoke-SmartM365RunCleanup $State $kind {
                if($kind -eq 'Exchange'){ & $command -Confirm:$false -ErrorAction Stop | Out-Null }
                else{ & $command -ErrorAction Stop | Out-Null }
                $State.($kind+'Owned')=$false
            }
        }
    }
}

function Start-SmartM365RunTranscript {
    [CmdletBinding()]
    param([string]$Path,[string]$LiteralPath,[switch]$Append,[switch]$Force,[switch]$NoClobber,
        [switch]$IncludeInvocationHeader,[switch]$UseMinimalHeader,[string]$OutputDirectory)
    $result = Microsoft.PowerShell.Host\Start-Transcript @PSBoundParameters
    $state = Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($state) { $state.TranscriptOwned=$true }
    return $result
}

function Stop-SmartM365RunTranscript {
    [CmdletBinding()]
    param()
    $state = Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($state -and -not $state.TranscriptOwned) { return }
    $result = Microsoft.PowerShell.Host\Stop-Transcript @PSBoundParameters
    if ($state) { $state.TranscriptOwned=$false }
    return $result
}

function Get-SmartM365InventoryCompletionStatus {
    param([string]$Status,[AllowNull()]$ErrorRecord)
    $state = Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if (-not $state) { return $Status }
    if ($Status -eq 'Cancelled' -or $state.ExitCode -eq 130) { $state.ExitCode=130; Disable-SmartM365InterruptedPublication $state; return 'Cancelled' }
    if (-not $state.InCleanup) {
        $state.ReachedTerminal=$true
        if ($Status -eq 'Failed') { $state.ExitCode=1 }
        return $Status
    }
    if ($state.ErrorRecord -or $ErrorRecord -or $state.ExitCode -notin @(0,3,130)) { if($state.ExitCode -in @(0,3)){ $state.ExitCode=1 }; return 'Failed' }
    if (-not $state.ReachedTerminal -or ($state.Finalizing -and $state.RequireReturn -and -not $state.ReturnedNormally -and -not $state.ExplicitExit)) {
        $errorCount = Get-Variable SmartM365ErrorCount -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($errorCount -gt 0 -or $Status -eq 'Failed') { $state.ExitCode=1; return 'Failed' }
        $state.ExitCode=130
        Disable-SmartM365InterruptedPublication $state
        return 'Cancelled'
    }
    if ($state.Finalizing -and $state.CleanupWarnings.Count -gt 0 -and $Status -in @('Auto','Success')) { return 'CompletedWithWarnings' }
    return $Status
}

function Complete-SmartM365InventoryRun {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State,[scriptblock]$Cleanup)
    if ($State.Finalized) { return }
    $State.Finalized=$true; $State.InCleanup=$true; $State.Finalizing=$true
    try {
        if (Test-SmartM365InterruptedRun) { Disable-SmartM365InterruptedPublication $State }
        if ($Cleanup) { Invoke-SmartM365RunCleanup $State 'entry resources' $Cleanup }
        Close-SmartM365InventoryConnections $State
        $completion = $State.PendingCompletion
        if ($State.TranscriptOwned -and -not ($completion -and $completion.CloseTranscriptBeforeUpload)) {
            Invoke-SmartM365RunCleanup $State 'transcript' { Stop-SmartM365RunTranscript -ErrorAction Stop }
        }
        $status = Get-SmartM365InventoryCompletionStatus -Status 'Auto' -ErrorRecord $State.ErrorRecord
        $completionCommand=Get-Command Complete-SmartM365ExecutionContext,Complete-CoreSmartM365ExecutionContext -ErrorAction SilentlyContinue | Select-Object -First 1
        # Core banner first: Get-Command keeps the order of the names, and a dot-sourced TenantContext banner
        # would otherwise be chosen even when the prefixed Core one is loaded.
        $bannerCommand=Get-Command Write-CoreSmartM365CompletionBanner,Write-SmartM365CompletionBanner -ErrorAction SilentlyContinue | Select-Object -First 1
        $written = Get-Variable SmartM365ExecutionSummaryWritten -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($State.PendingBanner -and -not $completion) {
            $banner = @{}; foreach($key in $State.PendingBanner.Keys){ $banner[$key]=$State.PendingBanner[$key] }
            $banner.Status = Get-SmartM365InventoryCompletionStatus -Status $banner.Status -ErrorRecord $State.ErrorRecord
            if($State.CleanupWarnings.Count){$banner.WarningCount=[Math]::Max([int]$banner.WarningCount,$State.CleanupWarnings.Count)}
            if($bannerCommand){Invoke-SmartM365RunCleanup $State 'terminal banner' { & $bannerCommand @banner }}
        }
        # Bespoke entry initialization still owns an explicitly deferred completion.
        elseif (($State.Initialized -or $completion) -and -not $written -and $completionCommand) {
            Invoke-SmartM365RunCleanup $State 'terminal summary' {
                $options = @{Status=$status; ErrorRecord=$State.ErrorRecord; FailureStage=$(if($status -eq 'Cancelled'){'Interrupted'}else{'Inventory'})}
                if($completion){ foreach($key in $completion.Keys){$options[$key]=$completion[$key]} }
                if($State.ErrorRecord){$options.ErrorRecord=$State.ErrorRecord}
                $options.Status=Get-SmartM365InventoryCompletionStatus -Status $options.Status -ErrorRecord $options.ErrorRecord
                & $completionCommand @options
            }
        }
        $finalStatus = Get-Variable SmartM365ExecutionStatus -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($State.ErrorRecord -or (($State.Initialized -or $completion) -and $finalStatus -eq 'Failed')) {
            if ($State.ExitCode -in @(0,3)) { $State.ExitCode=1 }
        }
        if ($State.TranscriptOwned) { Invoke-SmartM365RunCleanup $State 'remaining transcript' { Stop-SmartM365RunTranscript -ErrorAction Stop } }
        Close-SmartM365InventoryConnections $State
        if ($State.ExitCode) { [Environment]::ExitCode=$State.ExitCode }
    }
    finally {
        foreach ($name in $State.PublicationMask.Keys) {
            $previous=$State.PublicationMask[$name]
            if ($previous.Exists) { Set-Variable $name -Value $previous.Value -Scope Global }
            else { Remove-Variable $name -Scope Global -ErrorAction SilentlyContinue }
        }
        if ($State.Previous) { Set-Variable SmartM365InventoryRun -Value $State.Previous.Value -Scope Global }
        else { Remove-Variable SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue }
    }
}
