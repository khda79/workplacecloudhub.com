#Requires -Version 5.1

<#
.SYNOPSIS
Installs or removes the SmartM365 Exchange Health Checker scheduled task on Exchange servers.

.DESCRIPTION
Run once from the collector server (the one that has the repository) for the Exchange servers that
the infrastructure collector cannot run Health Checker on locally. PowerShell remoting (WinRM) is
not used: files are deployed through the administrative share (SMB) and the task is registered
through CIM over DCOM.

For each server the installer:
- checks that the service account is a member of the local Administrators group, directly or
  through domain groups (members read over CIM/DCOM, the account's groups read from Active Directory
  with a read-only LDAP query in the context of the account that runs the installer). A server where
  it is proven not to be is refused; when it cannot be determined, a warning is written;
- refuses the server when the parent folder C:\Program Files\SmartM365 is owned or writable by
  anyone else than SYSTEM, Administrators or TrustedInstaller (the task runs elevated);
- creates <root>\Bin and <root>\Output under C:\Program Files\SmartM365\HealthChecker, where standard
  users cannot create files, owned by Administrators with explicit permissions: SYSTEM and
  Administrators full control, the service account read on Bin and modify on Output, inheritance removed;
- copies HealthChecker.ps1 (checked first: expected SHA256 and Microsoft signature) and the local
  runner, only when the file on the server differs;
- registers the task \WCH\SmartM365 Exchange Health Checker: Windows PowerShell 5.1 runs the local
  runner daily at -DailyAt, with highest privileges, ignoring a new start while a run is in
  progress, stopped after one hour.

The task runs under a dedicated domain service account. Its password is requested once through
Get-Credential, passed to Register-ScheduledTask, and never written to a file, a log or the output.
When the password expires or is changed, run the installer again: it updates the task on every
server (re-registration replaces the stored credential).

Running the installer again is safe: folders, permissions, files and the task are brought to the
same state.

.PARAMETER ComputerName
Exchange servers to install on or remove from. Required. Several names can also be given as one string
separated by commas or semicolons (powershell -File passes 'A,B' as a single value).

.PARAMETER ServiceAccount
Dedicated domain service account (DOMAIN\user or user@domain) used to run the task. SYSTEM and
LocalSystem are refused. Required when installing; ignored when uninstalling.

.PARAMETER DailyAt
Daily start time of the task, HH:mm, server local time. Default 01:00.

.PARAMETER StartNow
Starts the task on every server after registration and waits for the run result, to validate the
service account (rights and Exchange Management Shell loading). Use it on one server first.

.PARAMETER Uninstall
Removes the task and the Bin and Output folders instead of installing. No credential is requested.

.PARAMETER Interactive
Forces the guided interactive workflow. This is used by the CMD launcher.

.EXAMPLE
./Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1 -ComputerName 'EXSRV01' -ServiceAccount 'CONTOSO\svc-exchange-hc' -StartNow

.EXAMPLE
./Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1 -ComputerName 'EXSRV01','EXSRV02' -ServiceAccount 'CONTOSO\svc-exchange-hc'

.EXAMPLE
./Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1 -ComputerName 'EXSRV01' -Uninstall

.VERSION
1.0.4

.REQUIREMENTS
    Windows PowerShell 5.1, elevated. Started from PowerShell 7, the script restarts itself in Windows
    PowerShell 5.1 (the ScheduledTasks module would otherwise run through Windows compatibility, where
    CIM sessions cannot be passed); it requests UAC elevation when needed.
    Run by an account that is local administrator of every target server (C$ share and CIM/DCOM).
    Network from this server to the targets: RPC 135 and dynamic RPC ports (DCOM), SMB 445.
    Active Directory read access (LDAP) from this server, for the local administrator check.
    Service account: local administrator of every target, member of the Exchange role group
    View-Only Organization Management, and allowed to log on as a batch job (Task Scheduler grants
    the right at registration unless a Group Policy overrides it).
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string[]]$ComputerName = @(),
    [string]$ServiceAccount = '',
    [ValidatePattern('^([01][0-9]|2[0-3]):[0-5][0-9]$')]
    [string]$DailyAt = '01:00',
    [switch]$StartNow,
    [switch]$Uninstall,
    [switch]$Interactive
)

# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
function Repair-HealthCheckerModulePath {
    # A Windows PowerShell 5.1 process started from PowerShell 7 (the orchestrator, a pwsh console) inherits its
    # PSModulePath, and the PowerShell 7 builds of built-in modules such as Microsoft.PowerShell.Security
    # cannot load in 5.1 (Get-Acl, Get-AuthenticodeSignature). Windows PowerShell modules come first again.
    if ($PSVersionTable.PSEdition -ne 'Desktop') { return }
    $kept = New-Object System.Collections.Generic.List[string]
    $kept.Add((Join-Path -Path $PSHOME -ChildPath 'Modules'))
    foreach ($entry in @([string]$env:PSModulePath -split [IO.Path]::PathSeparator)) {
        $trimmed = $entry.Trim().TrimEnd('\')
        if (-not $trimmed -or $trimmed -match '(?i)(^|\\)PowerShell\\(7[^\\]*\\)?Modules$') { continue }
        if (-not ($kept -contains $trimmed)) { $kept.Add($trimmed) }
    }
    $env:PSModulePath = $kept -join [IO.Path]::PathSeparator
}
Repair-HealthCheckerModulePath
$ErrorActionPreference = 'Stop'
# Must match the collector constants (checked by the offline test).
$taskPath = '\WCH\'
$taskName = 'SmartM365 Exchange Health Checker'
$remoteRootPath = 'C:\Program Files\SmartM365\HealthChecker'
$healthCheckerExpectedSha256 = 'BF8758E96E8760226E3094D45A84909AFFA9F91167DE11F739D8E3D9343229DC'
$runnerFileName = 'SmartM365-Exchange-OnPrem-HealthChecker-LocalRunner.ps1'
$taskTimeLimit = New-TimeSpan -Hours 1
$startNowWaitMinutes = 40

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$healthCheckerSourcePath = Join-Path $scriptDirectory 'ThirdParty\CSS-Exchange\HealthChecker\HealthChecker.ps1'
$runnerSourcePath = Join-Path $scriptDirectory $runnerFileName
$smartM365Root = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptDirectory)))
$tenantContextPath = Join-Path $smartM365Root 'Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) {
    throw "SmartM365 tenant context not found: $tenantContextPath"
}
. $tenantContextPath
$script:InstallerStartedAt = Get-Date
$global:SmartM365ScriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$global:SmartM365ExecutionStartTime = $script:InstallerStartedAt
$installerLogFolder = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath 'SmartM365\Logs\ExchangeHealthChecker'
if (-not (Test-Path -LiteralPath $installerLogFolder)) { New-Item -ItemType Directory -Path $installerLogFolder -Force | Out-Null }
$global:LogTextFile = Join-Path -Path $installerLogFolder -ChildPath ('{0}_{1}.log' -f $global:SmartM365ScriptName, $script:InstallerStartedAt.ToString('yyyyMMdd_HHmmss'))
$script:InstallerCompletionStatus = 'Auto'
# Read by the completion banner: every failure is counted, so a failed server never ends as a success.
$global:SmartM365ErrorCount = 0
$global:SmartM365WarningCount = 0

Write-SmartM365StartupBanner

function Write-InstallerMessage {
    param([Parameter(Mandatory = $true)][string]$Message)
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Microsoft.PowerShell.Utility\Write-Host $line
    Add-Content -LiteralPath $global:LogTextFile -Value $line -Encoding UTF8
}

function Write-InstallerError {
    param([Parameter(Mandatory = $true)][string]$Message)
    $global:SmartM365ErrorCount++
    Write-InstallerMessage ('ERROR: {0}' -f $Message)
}

function Write-InstallerWarning {
    param([Parameter(Mandatory = $true)][string]$Message)
    $global:SmartM365WarningCount++
    Write-InstallerMessage ('WARNING: {0}' -f $Message)
}

function ConvertFrom-HealthCheckerSidBytes {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)
    return (New-Object System.Security.Principal.SecurityIdentifier($Bytes, 0)).Value
}

function Get-HealthCheckerAccountTokenSid {
    param([Parameter(Mandatory = $true)][string]$AccountSid)

    # Read-only LDAP, in the context of the account that runs the installer (no credential is asked):
    # the account SID and every security group it belongs to, nested domain groups included (tokenGroups).
    $entry = New-Object System.DirectoryServices.DirectoryEntry(('LDAP://<SID={0}>' -f $AccountSid))
    try {
        $entry.RefreshCache([string[]]@('tokenGroups'))
        $sids = New-Object System.Collections.Generic.List[string]
        $sids.Add($AccountSid)
        foreach ($bytes in @($entry.Properties['tokenGroups'])) { $sids.Add((ConvertFrom-HealthCheckerSidBytes -Bytes ([byte[]]$bytes))) }
        return $sids.ToArray()
    }
    finally {
        $entry.Dispose()
    }
}

function Get-HealthCheckerLocalAdministratorMember {
    param([Parameter(Mandatory = $true)]$Session)

    $group = Get-CimInstance -CimSession $Session -ClassName Win32_Group -Filter "LocalAccount = TRUE AND SID = 'S-1-5-32-544'" -ErrorAction Stop
    if ($null -eq $group) { throw 'The local Administrators group was not found.' }
    return @(Get-CimAssociatedInstance -CimSession $Session -InputObject $group -Association Win32_GroupUser -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{ Sid = [string]$_.SID; IsLocalGroup = ($_.CimClass.CimClassName -eq 'Win32_Group' -and [bool]$_.LocalAccount) }
        })
}

function Test-HealthCheckerAdministratorMembership {
    param(
        [AllowNull()][string[]]$AccountSids,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Members
    )

    if ($null -eq $AccountSids -or $AccountSids.Count -eq 0) { return 'Unknown' }
    foreach ($member in $Members) {
        if ($member.Sid -and $AccountSids -contains [string]$member.Sid) { return 'Member' }
    }
    # A local group inside Administrators can hold the account indirectly; it is not resolved here.
    if (@($Members | Where-Object { $_.IsLocalGroup }).Count -gt 0) { return 'Unknown' }
    return 'NotMember'
}

function ConvertTo-HealthCheckerComputerList {
    param([AllowNull()][AllowEmptyCollection()][string[]]$Value)

    # powershell -File passes 'A,B' as a single string: commas and semicolons separate names in every form.
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($Value)) {
        foreach ($part in ([string]$item -split '[,;]')) {
            $name = $part.Trim()
            if ($name -and -not ($names -contains $name)) { $names.Add($name) }
        }
    }
    return $names.ToArray()
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-AccountSid {
    param([Parameter(Mandatory = $true)][string]$AccountName)
    try {
        $account = [Security.Principal.NTAccount]::new($AccountName)
        return $account.Translate([Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        throw "Service account '$AccountName' could not be resolved. Verify the account name and domain connectivity. $($_.Exception.Message)"
    }
}

function ConvertTo-QuotedArgument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

function Read-InstallerChoice {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][string[]]$AllowedValues
    )

    while ($true) {
        $value = Read-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Prompt)
        if ($value -in $AllowedValues) {
            return $value
        }
        Write-InstallerMessage ('Invalid choice. Allowed values: {0}' -f ($AllowedValues -join ', '))
    }
}

function Read-InstallerYesNo {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    while ($true) {
        $value = Read-Host ('[{0}] {1} [Y/N]' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Prompt)
        switch ($value.Trim().ToLowerInvariant()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
            default { Write-InstallerMessage 'Invalid choice. Enter Y or N.' }
        }
    }
}

function Wait-InteractiveClose {
    if ($script:InteractiveMode) {
        [void](Read-Host ('[{0}] Press Enter to close' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    }
}

function Complete-InstallerExecution {
    param(
        [ValidateSet('Auto', 'Success', 'Failed', 'CompletedWithWarnings')]
        [string]$Status = $script:InstallerCompletionStatus
    )
    Write-SmartM365CompletionBanner -Status $Status -ScriptName $global:SmartM365ScriptName -StartedAt $script:InstallerStartedAt -LogPath $global:LogTextFile
}

function Test-HealthCheckerScriptIntegrity {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedSha256
    )

    $result = [pscustomobject]@{ IsValid = $false; Reason = $null; Sha256 = $null; SignatureStatus = $null; Signer = $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Reason = "Health Checker script not found: $Path"
        return $result
    }

    $result.Sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($result.Sha256 -ne $ExpectedSha256.ToUpperInvariant()) {
        $result.Reason = "SHA256 mismatch: expected $($ExpectedSha256.ToUpperInvariant()), found $($result.Sha256)."
        return $result
    }

    $signature = Get-AuthenticodeSignature -FilePath $Path
    $result.SignatureStatus = [string]$signature.Status
    if ($signature.SignerCertificate) { $result.Signer = [string]$signature.SignerCertificate.Subject }
    if ([string]$signature.Status -ne 'Valid' -or -not $signature.SignerCertificate) {
        $result.Reason = "Authenticode signature is not valid (status $($signature.Status))."
        return $result
    }
    $subject = [string]$signature.SignerCertificate.Subject
    $issuer = [string]$signature.SignerCertificate.Issuer
    if ($subject -notmatch '(^|,\s*)CN=Microsoft Corporation(,|$)' -or $subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)' -or $issuer -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
        $result.Reason = "Authenticode signer is not Microsoft Corporation: $subject"
        return $result
    }

    $result.IsValid = $true
    return $result
}

function ConvertTo-AdministrativeSharePath {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][string]$LocalPath
    )

    if ($LocalPath -notmatch '^(?<Drive>[A-Za-z]):\\(?<Rest>.+)$') {
        throw "Remote root must be a local absolute path with a drive letter: $LocalPath"
    }
    return ('\\{0}\{1}$\{2}' -f $Computer, $Matches['Drive'].ToUpperInvariant(), $Matches['Rest'])
}

function Set-HealthCheckerFolderAclRule {
    param(
        [Parameter(Mandatory = $true)][System.Security.AccessControl.DirectorySecurity]$Acl,
        [Parameter(Mandatory = $true)][string]$ServiceAccountSid,
        [Parameter(Mandatory = $true)][ValidateSet('ReadAndExecute', 'Modify')][string]$ServiceAccountRight
    )

    # The task runs elevated: Administrators own the folder and only explicit entries remain.
    $acl = $Acl
    $acl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
    $acl.SetAccessRuleProtection($true, $false)
    foreach ($existingRule in @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))) {
        [void]$acl.RemoveAccessRuleSpecific($existingRule)
    }
    $inheritance = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    $propagation = [System.Security.AccessControl.PropagationFlags]::None
    $allow = [System.Security.AccessControl.AccessControlType]::Allow
    foreach ($sid in @('S-1-5-18', 'S-1-5-32-544')) {
        $identity = New-Object System.Security.Principal.SecurityIdentifier($sid)
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($identity, 'FullControl', $inheritance, $propagation, $allow)))
    }
    $serviceIdentity = New-Object System.Security.Principal.SecurityIdentifier($ServiceAccountSid)
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($serviceIdentity, $ServiceAccountRight, $inheritance, $propagation, $allow)))
    return $acl
}

function Get-HealthCheckerUnsafeAclEntry {
    param([Parameter(Mandatory = $true)][System.Security.AccessControl.FileSystemSecurity]$Acl)

    # SYSTEM, Administrators, TrustedInstaller; CREATOR OWNER only grants rights to whoever creates a child.
    $trustedSids = @('S-1-5-18', 'S-1-5-32-544', 'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')
    $writeRights = [System.Security.AccessControl.FileSystemRights]'WriteData, AppendData, WriteExtendedAttributes, DeleteSubdirectoriesAndFiles, WriteAttributes, Delete, ChangePermissions, TakeOwnership'
    $unsafe = New-Object System.Collections.Generic.List[string]
    $owner = [string]$Acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value
    if ($owner -notin $trustedSids) { $unsafe.Add("owner $owner") }
    foreach ($rule in @($Acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))) {
        $sid = [string]$rule.IdentityReference.Value
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow -or $sid -in $trustedSids -or $sid -eq 'S-1-3-0') { continue }
        # Generic rights (GENERIC_ALL 0x10000000, GENERIC_WRITE 0x40000000) are not mapped by .NET: test them too.
        $mask = [int64][int]$rule.FileSystemRights -band 4294967295
        if (($mask -band [int64]$writeRights) -ne 0 -or ($mask -band 0x50000000) -ne 0) { $unsafe.Add(('{0} {1}' -f $sid, $rule.FileSystemRights)) }
    }
    return $unsafe.ToArray()
}

function Assert-HealthCheckerParentFolder {
    param([Parameter(Mandatory = $true)][string]$Path)

    # A parent that someone else owns or can write lets them swap the folder the elevated task runs from.
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        # Created here, it keeps the Program Files permissions; only its owner is set to Administrators.
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        $parentAcl = Get-Acl -LiteralPath $Path
        $parentAcl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
        Set-Acl -LiteralPath $Path -AclObject $parentAcl
    }
    $unsafe = @(Get-HealthCheckerUnsafeAclEntry -Acl (Get-Acl -LiteralPath $Path))
    if ($unsafe.Count -gt 0) {
        throw ("The parent folder is not restricted to administrators ({0}): {1}. Fix its owner and permissions first." -f ($unsafe -join '; '), $Path)
    }
}

function Set-HealthCheckerFolder {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ServiceAccountSid,
        [Parameter(Mandatory = $true)][ValidateSet('ReadAndExecute', 'Modify')][string]$ServiceAccountRight
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    $acl = Set-HealthCheckerFolderAclRule -Acl (Get-Acl -LiteralPath $Path) -ServiceAccountSid $ServiceAccountSid -ServiceAccountRight $ServiceAccountRight
    Set-Acl -LiteralPath $Path -AclObject $acl
}

function Copy-HealthCheckerFileIfChanged {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
    if ((Test-Path -LiteralPath $DestinationPath -PathType Leaf) -and (Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash -eq $sourceHash) {
        return 'Unchanged'
    }
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    if ((Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash -ne $sourceHash) {
        throw "Copied file does not match its source: $DestinationPath"
    }
    return 'Copied'
}

function New-HealthCheckerDcomSessionOption {
    # DCOM, not WinRM: PowerShell remoting is disabled on the Exchange servers. The DCOM parameter set
    # (-Impersonation/-PacketPrivacy/-PacketIntegrity) returns DComSessionOptions; it cannot be combined with -Protocol.
    return New-CimSessionOption -Impersonation Impersonate -PacketPrivacy -PacketIntegrity
}

function New-HealthCheckerCimSession {
    param([Parameter(Mandatory = $true)][string]$Computer)

    return New-CimSession -ComputerName $Computer -SessionOption (New-HealthCheckerDcomSessionOption) -ErrorAction Stop
}

function Wait-HealthCheckerFirstRun {
    param(
        [Parameter(Mandatory = $true)][string]$OutputSharePath,
        [AllowNull()][string]$PreviousRunToken,
        [Parameter(Mandatory = $true)][int]$TimeoutMinutes
    )

    $statusPath = Join-Path $OutputSharePath 'status.json'
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 15
        if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { continue }
        try { $status = Get-Content -LiteralPath $statusPath -Raw | ConvertFrom-Json } catch { continue }
        if ([string]$status.RunToken -eq [string]$PreviousRunToken -or [string]$status.State -eq 'Running') { continue }
        return $status
    }
    return $null
}

$script:InteractiveMode = $Interactive -or $PSBoundParameters.Count -eq 0

trap {
    Write-InstallerError $_.Exception.Message
    $script:InstallerCompletionStatus = 'Failed'
    Wait-InteractiveClose
    Complete-InstallerExecution -Status 'Failed'
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
}

$isAdministrator = Test-IsAdministrator
$isWindowsPowerShell = $PSVersionTable.PSEdition -eq 'Desktop'
if (-not $isAdministrator -or -not $isWindowsPowerShell) {
    # Windows PowerShell 5.1 only: under PowerShell 7 the ScheduledTasks module is loaded through Windows
    # compatibility (implicit remoting) and receives CIM sessions deserialized, so registration fails.
    $windowsPowerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $windowsPowerShellPath -PathType Leaf)) {
        throw 'Windows PowerShell 5.1 was not found; the installer requires it.'
    }
    foreach ($unsafeValue in @($ServiceAccount) + @($ComputerName)) {
        if (-not [string]::IsNullOrWhiteSpace($unsafeValue) -and $unsafeValue -match '[\r\n"]') {
            throw 'ServiceAccount and ComputerName must not contain quotes or line breaks.'
        }
    }

    $elevationArguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (ConvertTo-QuotedArgument -Value $MyInvocation.MyCommand.Path))
    if ($script:InteractiveMode) {
        $elevationArguments += '-Interactive'
    }
    else {
        if ($ComputerName.Count -gt 0) {
            $elevationArguments += @('-ComputerName', ((@($ComputerName) | ForEach-Object { ConvertTo-QuotedArgument -Value $_ }) -join ','))
        }
        if (-not [string]::IsNullOrWhiteSpace($ServiceAccount)) {
            $elevationArguments += @('-ServiceAccount', (ConvertTo-QuotedArgument -Value $ServiceAccount))
        }
        $elevationArguments += @('-DailyAt', $DailyAt)
        if ($StartNow) { $elevationArguments += '-StartNow' }
        if ($Uninstall) { $elevationArguments += '-Uninstall' }
        if ($WhatIfPreference) { $elevationArguments += '-WhatIf' }
    }

    $elevationWorkingDirectory = $scriptDirectory
    if ($scriptDirectory.StartsWith('\\', [StringComparison]::Ordinal)) {
        $elevationWorkingDirectory = $env:SystemRoot
    }

    $restartParameters = @{
        FilePath         = $windowsPowerShellPath
        ArgumentList     = ($elevationArguments -join ' ')
        WorkingDirectory = $elevationWorkingDirectory
        Wait             = $true
        PassThru         = $true
    }
    if ($isAdministrator) {
        Write-InstallerMessage 'The installer runs in Windows PowerShell 5.1. Restarting it there.'
        $restartParameters.NoNewWindow = $true
    }
    else {
        Write-InstallerMessage 'Administrator rights and Windows PowerShell 5.1 are required. Requesting elevation through UAC.'
        $restartParameters.Verb = 'RunAs'
    }
    $elevationProcess = Start-Process @restartParameters
    $script:InstallerCompletionStatus = if ($elevationProcess.ExitCode -eq 0) { 'Auto' } else { 'Failed' }
    Complete-InstallerExecution
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=($elevationProcess.ExitCode); exit $smartM365RunGuard.ExitCode
}

if ($script:InteractiveMode) {
    Write-InstallerMessage 'SmartM365 Exchange Health Checker scheduled-task installer'
    Write-InstallerMessage '1 - Install or update the scheduled task'
    Write-InstallerMessage '2 - Uninstall the scheduled task'
    $actionChoice = Read-InstallerChoice -Prompt 'Select an action [1-2]' -AllowedValues @('1', '2')
    $Uninstall = $actionChoice -eq '2'

    while ($true) {
        $candidateComputers = Read-Host ('[{0}] Exchange servers, separated by commas' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        $ComputerName = @(ConvertTo-HealthCheckerComputerList -Value $candidateComputers)
        if ($ComputerName.Count -gt 0) { break }
        Write-InstallerMessage 'At least one server is required.'
    }

    if ($Uninstall) {
        if (-not (Read-InstallerYesNo -Prompt ("Remove '{0}' from {1} server(s)?" -f $taskName, $ComputerName.Count))) {
            Write-InstallerMessage 'Uninstall cancelled.'
            Wait-InteractiveClose
            Complete-InstallerExecution -Status 'CompletedWithWarnings'
            $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
        }
        $StartNow = $false
        $ServiceAccount = ''
    }
    else {
        while ($true) {
            $candidateAccount = Read-Host ('[{0}] Dedicated service account (DOMAIN\user or user@domain)' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            if ([string]::IsNullOrWhiteSpace($candidateAccount)) {
                Write-InstallerMessage 'The service account is required.'
                continue
            }
            if ($candidateAccount -notmatch '^(?:[A-Za-z0-9._-]+\\[A-Za-z0-9._+-]+|[A-Za-z0-9._+-]+@[A-Za-z0-9.-]+)$') {
                Write-InstallerMessage 'Invalid account format. Use DOMAIN\user or user@domain without spaces.'
                continue
            }
            $ServiceAccount = $candidateAccount
            break
        }
        $StartNow = Read-InstallerYesNo -Prompt 'Start the task now and wait for its result (recommended on one server first)?'
    }
}

$ComputerName = @(ConvertTo-HealthCheckerComputerList -Value $ComputerName)
if ($ComputerName.Count -eq 0) {
    throw 'ComputerName is required.'
}
foreach ($computer in $ComputerName) {
    if ($computer -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]*$') {
        throw "Invalid server name: $computer"
    }
}

if ($Uninstall) {
    if ($StartNow) {
        throw 'StartNow cannot be combined with Uninstall.'
    }

    $failedServers = 0
    foreach ($computer in $ComputerName) {
        $session = $null
        try {
            $session = New-HealthCheckerCimSession -Computer $computer
            $existingTask = Get-ScheduledTask -CimSession $session -TaskPath $taskPath -TaskName $taskName -ErrorAction SilentlyContinue
            if ($null -eq $existingTask) {
                Write-InstallerMessage "${computer}: scheduled task does not exist: $taskPath$taskName"
            }
            elseif ($PSCmdlet.ShouldProcess("$computer $taskPath$taskName", 'Unregister scheduled task')) {
                Unregister-ScheduledTask -CimSession $session -TaskPath $taskPath -TaskName $taskName -Confirm:$false
                Write-InstallerMessage "${computer}: scheduled task removed: $taskPath$taskName"
            }
            $rootShare = ConvertTo-AdministrativeSharePath -Computer $computer -LocalPath $remoteRootPath
            foreach ($folder in @((Join-Path $rootShare 'Bin'), (Join-Path $rootShare 'Output'))) {
                if ((Test-Path -LiteralPath $folder -PathType Container) -and $PSCmdlet.ShouldProcess($folder, 'Remove Health Checker folder')) {
                    Remove-Item -LiteralPath $folder -Recurse -Force
                    Write-InstallerMessage "${computer}: folder removed: $folder"
                }
            }
        }
        catch {
            $failedServers++
            Write-InstallerError ("{0}: uninstall failed: {1}" -f $computer, $_.Exception.Message)
        }
        finally {
            if ($null -ne $session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
        }
    }

    if ($failedServers -gt 0) { $script:InstallerCompletionStatus = 'Failed' }
    Wait-InteractiveClose
    Complete-InstallerExecution
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
}

if ([string]::IsNullOrWhiteSpace($ServiceAccount)) {
    throw 'ServiceAccount is required when installing the scheduled task.'
}
$normalizedServiceAccount = $ServiceAccount.Trim().ToLowerInvariant()
$forbiddenServiceAccounts = @('system', 'localsystem', 'nt authority\system', 's-1-5-18')
if ($normalizedServiceAccount -in $forbiddenServiceAccounts) {
    throw 'SYSTEM and LocalSystem are refused. Use a dedicated domain service account.'
}
$serviceAccountSid = Resolve-AccountSid -AccountName $ServiceAccount
if ($serviceAccountSid -eq 'S-1-5-18') {
    throw 'SYSTEM and LocalSystem are refused. Use a dedicated domain service account.'
}

$integrity = Test-HealthCheckerScriptIntegrity -Path $healthCheckerSourcePath -ExpectedSha256 $healthCheckerExpectedSha256
if (-not $integrity.IsValid) {
    throw "The vendored Health Checker failed its integrity check and is not deployed: $($integrity.Reason)"
}
if (-not (Test-Path -LiteralPath $runnerSourcePath -PathType Leaf)) {
    throw "Local runner not found: $runnerSourcePath"
}

$actionArguments = @(
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    (ConvertTo-QuotedArgument -Value (Join-Path (Join-Path $remoteRootPath 'Bin') $runnerFileName))
) -join ' '
$description = 'Runs the Microsoft CSS-Exchange Health Checker for this server only, through the SmartM365 local runner. Started daily and on demand by the SmartM365 Exchange infrastructure collector.'

# Read once, before any server: the service account and all its groups, for the local administrator check.
$accountTokenSids = $null
try { $accountTokenSids = @(Get-HealthCheckerAccountTokenSid -AccountSid $serviceAccountSid) }
catch { Write-InstallerWarning ("The groups of $ServiceAccount could not be read from Active Directory, so the local administrator check cannot conclude: {0}" -f $_.Exception.Message) }

$credential = $null
$passwordPointer = [IntPtr]::Zero
$plainTextPassword = $null
$failedServers = 0

try {
    if (-not $PSCmdlet.ShouldProcess(($ComputerName -join ', '), "Deploy files and register $taskPath$taskName as $ServiceAccount")) {
        Wait-InteractiveClose
        Complete-InstallerExecution -Status 'CompletedWithWarnings'
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    $credentialParameters = @{
        UserName = $ServiceAccount
        Message = "Enter the password for the SmartM365 Exchange Health Checker service account ($ServiceAccount)."
    }
    $credential = Get-Credential @credentialParameters
    if ($null -eq $credential) {
        throw 'Credential entry was cancelled.'
    }
    $credentialSid = Resolve-AccountSid -AccountName $credential.UserName
    if ($credentialSid -ne $serviceAccountSid) {
        throw "The entered credential belongs to '$($credential.UserName)', not '$ServiceAccount'."
    }
    $passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
    $plainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)

    foreach ($computer in $ComputerName) {
        $session = $null
        try {
            $session = New-HealthCheckerCimSession -Computer $computer
            # Without local administrator rights the task never starts the runner: refuse such a server up front.
            $membership = 'Unknown'
            $membershipDetail = ''
            try { $membership = Test-HealthCheckerAdministratorMembership -AccountSids $accountTokenSids -Members @(Get-HealthCheckerLocalAdministratorMember -Session $session) }
            catch { $membershipDetail = " ($($_.Exception.Message))" }
            if ($membership -eq 'NotMember') {
                throw ("the service account {0} is not a local administrator on {1}: add it, directly or through a domain group, then run the installer again." -f $ServiceAccount, $computer)
            }
            if ($membership -eq 'Unknown') {
                Write-InstallerWarning ("{0}: local administrator membership of {1} could not be confirmed{2}. Use -StartNow to validate the first run." -f $computer, $ServiceAccount, $membershipDetail)
            }

            $rootShare = ConvertTo-AdministrativeSharePath -Computer $computer -LocalPath $remoteRootPath
            $binShare = Join-Path $rootShare 'Bin'
            $outputShare = Join-Path $rootShare 'Output'
            Assert-HealthCheckerParentFolder -Path (Split-Path -Path $rootShare -Parent)
            Set-HealthCheckerFolder -Path $rootShare -ServiceAccountSid $serviceAccountSid -ServiceAccountRight 'ReadAndExecute'
            Set-HealthCheckerFolder -Path $binShare -ServiceAccountSid $serviceAccountSid -ServiceAccountRight 'ReadAndExecute'
            Set-HealthCheckerFolder -Path $outputShare -ServiceAccountSid $serviceAccountSid -ServiceAccountRight 'Modify'
            $healthCheckerCopy = Copy-HealthCheckerFileIfChanged -SourcePath $healthCheckerSourcePath -DestinationPath (Join-Path $binShare 'HealthChecker.ps1')
            $runnerCopy = Copy-HealthCheckerFileIfChanged -SourcePath $runnerSourcePath -DestinationPath (Join-Path $binShare $runnerFileName)
            Write-InstallerMessage ("{0}: folders and permissions set; HealthChecker.ps1 {1}; runner {2}." -f $computer, $healthCheckerCopy.ToLowerInvariant(), $runnerCopy.ToLowerInvariant())

            $action = New-ScheduledTaskAction -CimSession $session -Execute '%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument $actionArguments -WorkingDirectory (Join-Path $remoteRootPath 'Bin')
            $trigger = New-ScheduledTaskTrigger -CimSession $session -Daily -At ([datetime]::ParseExact($DailyAt, 'HH:mm', [Globalization.CultureInfo]::InvariantCulture))
            $settings = New-ScheduledTaskSettingsSet -CimSession $session -MultipleInstances IgnoreNew -StartWhenAvailable -ExecutionTimeLimit $taskTimeLimit
            $previousStatusPath = Join-Path $outputShare 'status.json'
            $previousRunToken = $null
            if (Test-Path -LiteralPath $previousStatusPath -PathType Leaf) {
                try { $previousRunToken = [string](Get-Content -LiteralPath $previousStatusPath -Raw | ConvertFrom-Json).RunToken } catch { $previousRunToken = $null }
            }

            $registrationParameters = @{
                CimSession = $session
                TaskPath = $taskPath
                TaskName = $taskName
                Action = $action
                Trigger = $trigger
                Settings = $settings
                Description = $description
                User = $ServiceAccount
                Password = $plainTextPassword
                RunLevel = 'Highest'
                Force = $true
            }
            Register-ScheduledTask @registrationParameters | Out-Null
            $registered = Get-ScheduledTask -CimSession $session -TaskPath $taskPath -TaskName $taskName
            Write-InstallerMessage ("{0}: scheduled task registered: {1}{2}, run account {3}, daily at {4}." -f $computer, $taskPath, $taskName, $registered.Principal.UserId, $DailyAt)

            if ($StartNow) {
                Start-ScheduledTask -CimSession $session -TaskPath $taskPath -TaskName $taskName
                Write-InstallerMessage ("{0}: task started; waiting up to {1} minutes for its result." -f $computer, $startNowWaitMinutes)
                $firstRun = Wait-HealthCheckerFirstRun -OutputSharePath $outputShare -PreviousRunToken $previousRunToken -TimeoutMinutes $startNowWaitMinutes
                if ($null -eq $firstRun) {
                    $info = Get-ScheduledTaskInfo -CimSession $session -TaskPath $taskPath -TaskName $taskName
                    throw ("no run result within {0} minutes (last task result 0x{1:X8}). A logon failure means a wrong password or a missing 'Log on as a batch job' right." -f $startNowWaitMinutes, [uint32]$info.LastTaskResult)
                }
                Write-InstallerMessage ("{0}: run result {1} in {2}s. {3}" -f $computer, $firstRun.State, $firstRun.DurationSeconds, $firstRun.Message)
                if ([string]$firstRun.State -ne 'Completed') {
                    $failedServers++
                    Write-InstallerError ("{0}: the task is installed but its first run ended with {1}." -f $computer, $firstRun.State)
                }
            }
        }
        catch {
            $failedServers++
            Write-InstallerError ("{0}: installation failed: {1}" -f $computer, $_.Exception.Message)
        }
        finally {
            if ($null -ne $session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
        }
    }

    if ($failedServers -gt 0) { $script:InstallerCompletionStatus = 'Failed' }
    $smartM365RunGuard.ReachedTerminal=$true
}
finally {
    $smartM365RunGuard.InCleanup=$true

    $plainTextPassword = $null
    $credential = $null
    if ($passwordPointer -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
    }
}

Wait-InteractiveClose

Complete-InstallerExecution
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}
