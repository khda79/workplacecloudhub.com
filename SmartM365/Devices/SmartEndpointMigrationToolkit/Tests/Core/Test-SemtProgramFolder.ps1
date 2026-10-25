<#
.SYNOPSIS
    Offline tests of the Program Files location of every script that runs as SYSTEM.
.DESCRIPTION
    Scripts started as SYSTEM (Intune runner, retry-after-reboot tasks, endpoint started by
    PsExec) are installed under %ProgramFiles%, never under ProgramData.
    1. The region "SEMT trusted program folder v1" is identical in both endpoint scripts, the
       Intune installer, the Intune runner and the Core module.
    2. Trust rules: real checks on Program Files, on a folder owned by the test account, on a
       folder whose permissions cannot be read and on a reparse point; rule-level checks on
       synthetic security descriptors (owner, write entries, inherit-only CREATOR OWNER, generic
       rights, deny entries).
    3. Folder chain: created components get the protected ACL, existing trusted components are
       left unchanged, an existing untrusted component is never repaired and stops the
       operation, a folder that cannot be protected is removed again (real, non-elevated).
    4. Retry after reboot (Windows 11 and Hybrid Join endpoint functions): runner, state and the
       endpoint copy are written under the Program Files folder, the task points there, an
       untrusted folder stops before the task is registered, cleanup removes the copy.
    3 bis. ProgramData data chain: only the SEMT components are changed; shared parents (WCH,
       WCH\SmartM365) are verified read-only and siblings (other product, legacy cache reused in
       place, explicit sibling) keep owner, ACL and content; existing content without verified
       provenance is refused, never adopted by a permission replacement; reparse points and hard
       links are refused before any change, the reproduced case included.
    5. PsExec payload (Core): the endpoint is copied under Program Files on the target and the
       PsExec path points there; an untrusted program folder stops the copy.
    Temporary folders only; schtasks, administrative shares and ACL changes that need elevation
    are replaced by stubs at the boundary. Prints one PASS/FAIL line per assertion and
    "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$packageRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$w11Endpoint = Join-Path $packageRoot 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$hjEndpoint = Join-Path $packageRoot 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
$installer = Join-Path $packageRoot 'Campaigns\Windows11Upgrade\IntuneWin32\Source\SmartM365-EndpointMigration-Install.ps1'
$runner = Join-Path $packageRoot 'Campaigns\Windows11Upgrade\IntuneWin32\Source\SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1'
$coreRemote = Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\Functions\Semt.Remote.ps1'
$regionName = 'SEMT trusted program folder v1'

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

function Get-RegionText {
    param([string]$Path, [string]$Region)
    $match = [regex]::Match([IO.File]::ReadAllText($Path), ('(?s)#region {0}.*?#endregion {0}' -f [regex]::Escape($Region)))
    if (-not $match.Success) { throw "Region '$Region' not found in $Path" }
    return ($match.Value -replace "`r`n", "`n")
}

function Get-FunctionText {
    param([string]$Path, [string[]]$Name)
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    return (@($Name | ForEach-Object {
        $functionName = $_
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        if (-not $definition) { throw "Function $functionName not found in $Path" }
        $definition.Extent.Text
    }) -join "`n`n")
}

function Get-ThrowMessage {
    param([scriptblock]$Action)
    try { & $Action; return '' } catch { return $_.Exception.Message }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-ProgramFolder-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    #region 1. Shared region
    $reference = Get-RegionText -Path $w11Endpoint -Region $regionName
    foreach ($copy in @($hjEndpoint, $installer, $runner, $coreRemote)) {
        Test-Step ("trusted program folder region identical in {0}" -f (Split-Path -Path $copy -Leaf)) { (Get-RegionText -Path $copy -Region $regionName) -ceq $reference }
    }
    #endregion

    . ([scriptblock]::Create($reference))

    #region 2. Trust rules
    Test-Step 'real: Program Files itself is trusted (owner TrustedInstaller, no write entry for others)' { (Test-SemtTrustedProgramPath -Path $env:ProgramFiles).Trusted }
    $own = Join-Path $tempRoot 'own'; New-Item -ItemType Directory -Path $own | Out-Null
    Test-Step 'real: folder owned by the test account is Untrusted (proven, not Unreadable)' { $r = Test-SemtTrustedProgramPath -Path $own; -not $r.Trusted -and $r.State -eq 'Untrusted' }
    Test-Step 'real: missing folder is Unreadable, never Trusted' { $r = Test-SemtTrustedProgramPath -Path (Join-Path $tempRoot 'missing'); -not $r.Trusted -and $r.State -eq 'Unreadable' }
    $junction = Join-Path $own 'link'
    $null = & cmd.exe /d /c ('mklink /J "{0}" "{1}"' -f $junction,$tempRoot) 2>&1
    Test-Step 'real: reparse point inside the folder is Untrusted' { $r = Test-SemtTrustedProgramPath -Path $junction; -not $r.Trusted -and $r.Detail -match 'Reparse point' }
    $null = & cmd.exe /d /c ('rmdir "{0}"' -f $junction) 2>&1

    # Rule level: Get-Acl is replaced by a function returning a security descriptor built from
    # SDDL, so owners and entries that need elevation can be checked without changing anything.
    $script:FakeSddl = @{}
    $ruleCase = Join-Path $tempRoot 'rules'; New-Item -ItemType Directory -Path $ruleCase | Out-Null
    function Get-Acl {
        param([string]$LiteralPath, $ErrorAction)
        if (-not $script:FakeSddl.ContainsKey($LiteralPath)) { throw 'Access denied (synthetic).' }
        $security = New-Object System.Security.AccessControl.DirectorySecurity
        $security.SetSecurityDescriptorSddlForm($script:FakeSddl[$LiteralPath])
        return $security
    }
    $ruleCases = @(
        @{ Name = 'owner Administrators, SYSTEM/Administrators full, Users read'; Sddl = 'O:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;0x1200a9;;;BU)'; Trusted = $true },
        @{ Name = 'owner TrustedInstaller with inherit-only CREATOR OWNER full control'; Sddl = 'O:S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464D:PAI(A;;FA;;;S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464)(A;OICIIO;GA;;;CO)(A;OICI;0x1200a9;;;BU)'; Trusted = $true },
        @{ Name = 'deny entry for Users is not a write grant'; Sddl = 'O:SYD:PAI(D;OICI;FA;;;BU)(A;OICI;FA;;;SY)'; Trusted = $true },
        @{ Name = 'owner is a standard account'; Sddl = 'O:S-1-5-21-1-2-3-1001D:PAI(A;OICI;FA;;;SY)'; Trusted = $false },
        @{ Name = 'Users can create files'; Sddl = 'O:BAD:PAI(A;OICI;FA;;;BA)(A;OICI;0x1200af;;;BU)'; Trusted = $false },
        @{ Name = 'Authenticated Users can change permissions'; Sddl = 'O:BAD:PAI(A;OICI;FA;;;BA)(A;;0x40000;;;AU)'; Trusted = $false },
        @{ Name = 'Users can delete'; Sddl = 'O:BAD:PAI(A;OICI;FA;;;BA)(A;;0x10000;;;BU)'; Trusted = $false },
        @{ Name = 'inherit-only generic write for Users'; Sddl = 'O:BAD:PAI(A;OICI;FA;;;BA)(A;OICIIO;GW;;;BU)'; Trusted = $false },
        @{ Name = 'CREATOR OWNER full control that also applies to this folder'; Sddl = 'O:BAD:PAI(A;OICI;FA;;;BA)(A;OICI;FA;;;CO)'; Trusted = $false }
    )
    foreach ($ruleCaseItem in $ruleCases) {
        $script:FakeSddl = @{ (Get-Item -LiteralPath $ruleCase).FullName = $ruleCaseItem.Sddl }
        $r = Test-SemtTrustedProgramPath -Path $ruleCase
        Write-TestResult -Condition ($r.Trusted -eq $ruleCaseItem.Trusted -and $r.State -eq $(if ($ruleCaseItem.Trusted) { 'Trusted' } else { 'Untrusted' })) -Message ("rule: {0} -> {1}" -f $ruleCaseItem.Name,$(if ($ruleCaseItem.Trusted) { 'Trusted' } else { 'Untrusted' })) -Detail ('{0} {1}' -f $r.State,$r.Detail)
    }
    $child = Join-Path $ruleCase 'child.ps1'; Set-Content -LiteralPath $child -Value '#' -Encoding ASCII
    $script:FakeSddl = @{ (Get-Item -LiteralPath $ruleCase).FullName = 'O:BAD:PAI(A;OICI;FA;;;BA)'; (Get-Item -LiteralPath $child).FullName = 'O:BAD:(A;;FA;;;BA)(A;;0x2;;;BU)' }
    Test-Step 'rule: -Recurse finds a writable file below a protected folder' { $r = Test-SemtTrustedProgramPath -Path $ruleCase -Recurse; -not $r.Trusted -and $r.Detail -match 'child\.ps1' }
    Test-Step 'rule: without -Recurse only the folder itself is checked' { (Test-SemtTrustedProgramPath -Path $ruleCase).Trusted }
    $script:FakeSddl = @{ (Get-Item -LiteralPath $ruleCase).FullName = 'O:BAD:PAI(A;OICI;FA;;;BA)' }
    Test-Step 'rule: permissions that cannot be read are Unreadable, never Trusted' { $r = Test-SemtTrustedProgramPath -Path $ruleCase -Recurse; -not $r.Trusted -and $r.State -eq 'Unreadable' }
    $sharedCases = @(
        @{ Name = 'ProgramData-like: Users can create items, inherit-only CREATOR OWNER'; Sddl = 'O:BAD:(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICIIO;GA;;;CO)(A;OICI;0x1200a9;;;BU)(A;CI;0x116;;;BU)'; Trusted = $true },
        @{ Name = 'Users full control on new children only (inherit-only)'; Sddl = 'O:SYD:(A;OICI;FA;;;SY)(A;OICIIO;FA;;;BU)'; Trusted = $true },
        @{ Name = 'Users can delete children'; Sddl = 'O:BAD:(A;OICI;FA;;;BA)(A;;0x40;;;BU)'; Trusted = $false },
        @{ Name = 'Users can change permissions'; Sddl = 'O:BAD:(A;OICI;FA;;;BA)(A;;0x40000;;;BU)'; Trusted = $false },
        @{ Name = 'Users can modify (delete included)'; Sddl = 'O:BAD:(A;OICI;FA;;;BA)(A;OICI;0x1301bf;;;BU)'; Trusted = $false },
        @{ Name = 'owner is a standard account'; Sddl = 'O:S-1-5-21-1-2-3-1001D:(A;OICI;FA;;;SY)'; Trusted = $false }
    )
    foreach ($sharedCase in $sharedCases) {
        $script:FakeSddl = @{ (Get-Item -LiteralPath $ruleCase).FullName = $sharedCase.Sddl }
        $r = Test-SemtTrustedSharedParent -Path $ruleCase
        Write-TestResult -Condition ($r.Trusted -eq $sharedCase.Trusted) -Message ("shared parent rule (read-only): {0} -> {1}" -f $sharedCase.Name,$(if ($sharedCase.Trusted) { 'Trusted' } else { 'Untrusted' })) -Detail ('{0} {1}' -f $r.State,$r.Detail)
    }
    $script:FakeSddl = @{}
    Test-Step 'shared parent rule: permissions that cannot be read are Unreadable, never Trusted' { $r = Test-SemtTrustedSharedParent -Path $ruleCase; -not $r.Trusted -and $r.State -eq 'Unreadable' }
    Remove-Item -Path Function:\Get-Acl
    #endregion

    #region 3. Folder chain
    $base = Join-Path $tempRoot 'PF'; New-Item -ItemType Directory -Path $base | Out-Null
    $target = Join-Path $base 'WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\Intune'
    Test-Step 'real: base folder owned by the test account: refused, nothing created' { (Get-ThrowMessage { Initialize-SemtTrustedProgramFolder -Path $target -BasePath $base }) -like 'SEMT_PROGRAM_FOLDER_UNTRUSTED*' -and -not (Test-Path -LiteralPath (Join-Path $base 'WCH')) }
    Test-Step 'path outside the Program Files base: refused' { (Get-ThrowMessage { Initialize-SemtTrustedProgramFolder -Path (Join-Path $tempRoot 'elsewhere') -BasePath $base }) -like '*is not below*' }

    $script:TrustedPaths = New-Object System.Collections.Generic.List[string]
    $script:AclCalls = New-Object System.Collections.Generic.List[string]
    $script:UntrustedPath = ''
    function Test-SemtTrustedProgramPath {
        param([string]$Path, [switch]$Recurse)
        $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
        if (-not (Test-Path -LiteralPath $full)) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = 'missing' } }
        if ($script:UntrustedPath -and $full -ieq $script:UntrustedPath) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = 'synthetic' } }
        $script:TrustedPaths.Add($full)
        return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
    }
    function Set-SemtProgramFolderAcl { param([string]$Path) $script:AclCalls.Add([IO.Path]::GetFullPath($Path).TrimEnd('\')) }
    $existing = Join-Path $base 'WCH'; New-Item -ItemType Directory -Path $existing | Out-Null
    $result = Initialize-SemtTrustedProgramFolder -Path $target -BasePath $base
    Test-Step 'chain: missing components are created and get the protected ACL' { $result -eq $target -and (Test-Path -LiteralPath $target) -and $script:AclCalls.Count -eq 4 -and $script:AclCalls -contains $target }
    Test-Step 'chain: an existing trusted component is verified and left unchanged' { $script:AclCalls -notcontains $existing -and $script:TrustedPaths -contains $existing }
    $script:AclCalls.Clear(); $script:UntrustedPath = Join-Path $base 'WCH\SmartM365'
    Test-Step 'chain: an existing untrusted component stops the operation and is never repaired' { (Get-ThrowMessage { Initialize-SemtTrustedProgramFolder -Path $target -BasePath $base }) -like 'SEMT_PROGRAM_FOLDER_UNTRUSTED*SmartM365*' -and $script:AclCalls.Count -eq 0 }
    $script:UntrustedPath = ''
    Remove-Item -Path Function:\Set-SemtProgramFolderAcl
    . ([scriptblock]::Create((Get-FunctionText -Path $w11Endpoint -Name 'Set-SemtProgramFolderAcl')))
    $isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $fresh = Join-Path $base 'Fresh\Folder'
    $message = Get-ThrowMessage { Initialize-SemtTrustedProgramFolder -Path $fresh -BasePath $base }
    if ($isElevated) {
        Test-Step 'real ACL change (elevated): created folder owned by Administrators with the protected DACL' { $message -eq '' -and (Get-Acl -LiteralPath $fresh).Sddl -match '^O:BA' }
    }
    else {
        Write-TestResult -Condition ($message -like 'SEMT_PROGRAM_FOLDER_UNTRUSTED: permissions of*' -and -not (Test-Path -LiteralPath (Join-Path $base 'Fresh'))) -Message 'real ACL change (not elevated): owner cannot be set, folder removed again, operation stopped' -Detail $message
    }
    $file = Join-Path $target 'payload.ps1'; $source = Join-Path $tempRoot 'payload.ps1'; Set-Content -LiteralPath $source -Value '# payload' -Encoding ASCII
    $installed = Install-SemtTrustedProgramFile -Source $source -Destination $file
    Test-Step 'program file: copied through a staging name and verified by SHA256' { $installed.Sha256 -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -and @(Get-ChildItem -LiteralPath $target -Force -Filter '*.tmp').Count -eq 0 }
    $script:UntrustedPath = $file
    Test-Step 'program file: an installed file that is not trusted stops the operation' { (Get-ThrowMessage { Install-SemtTrustedProgramFile -Source $source -Destination $file }) -like 'SEMT_PROGRAM_FOLDER_UNTRUSTED*' }
    $script:UntrustedPath = ''
    #endregion

    #region 3 bis. ProgramData data folder chain
    Remove-Item -Path Function:\Test-SemtTrustedProgramPath -ErrorAction SilentlyContinue
    . ([scriptblock]::Create($reference))
    $dataBase = Join-Path $tempRoot 'PD'; New-Item -ItemType Directory -Path $dataBase | Out-Null
    $productRoot = Join-Path $dataBase 'WCH\SmartM365\SmartEndpointMigrationToolkit'
    $dataRoot = Join-Path $productRoot 'Windows11Upgrade'
    Test-Step 'data chain: data root outside the ProgramData base is refused' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path (Join-Path $tempRoot 'elsewhere') -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*is not below*' }
    Test-Step 'data chain: data root that is not a campaign folder below SmartEndpointMigrationToolkit is refused, nothing created' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path (Join-Path $dataBase 'WCH\SmartM365\OtherProduct\Data') -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*not a campaign folder*' -and -not (Test-Path -LiteralPath (Join-Path $dataBase 'WCH')) }
    $linkedBase = Join-Path $tempRoot 'PD-link'
    $null = & cmd.exe /d /c ('mklink /J "{0}" "{1}"' -f $linkedBase,$dataBase) 2>&1
    Test-Step 'data chain: a reparse point as base or parent is refused before any change' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path (Join-Path $linkedBase 'WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade') -BasePath $linkedBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*reparse point*' -and -not (Test-Path -LiteralPath (Join-Path $dataBase 'WCH')) }
    $null = & cmd.exe /d /c ('rmdir "{0}"' -f $linkedBase) 2>&1
    $isElevatedData = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isElevatedData) {
        # Real checks, nothing simulated: a shared folder created by the test account is not
        # owned by SYSTEM or Administrators, so it is refused and removed again.
        $message = Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }
        Write-TestResult -Condition ($message -like 'SEMT_DATA_FOLDER_UNTRUSTED: shared folder*created here is not protected*' -and -not (Test-Path -LiteralPath (Join-Path $dataBase 'WCH'))) -Message 'data chain, real (not elevated): shared folder that cannot be proven protected: blocking failure, created folder removed' -Detail $message
    }

    # Pre-scan before a DACL replacement (real hard link and junction).
    $scan = Join-Path $tempRoot 'scan'; New-Item -ItemType Directory -Path (Join-Path $scan 'sub') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $scan 'sub\a.log') -Value 'x' -Encoding ASCII
    Test-Step 'pre-scan: folder without links passes' { (Get-ThrowMessage { Assert-SemtNoLinkedItem -Path $scan }) -eq '' }
    $null = & cmd.exe /d /c ('mklink /H "{0}" "{1}"' -f (Join-Path $scan 'sub\b.log'),(Join-Path $scan 'sub\a.log')) 2>&1
    Test-Step 'pre-scan, real hard link: refused (its permissions would change in another place)' { (Get-ThrowMessage { Assert-SemtNoLinkedItem -Path $scan }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*hard link*' }
    Remove-Item -LiteralPath (Join-Path $scan 'sub\b.log') -Force
    $null = & cmd.exe /d /c ('mklink /J "{0}" "{1}"' -f (Join-Path $scan 'sub\j'),$tempRoot) 2>&1
    Test-Step 'pre-scan, real junction: refused and never followed' { (Get-ThrowMessage { Assert-SemtNoLinkedItem -Path $scan }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*reparse point*' }
    $null = & cmd.exe /d /c ('rmdir "{0}"' -f (Join-Path $scan 'sub\j')) 2>&1

    # Chain behaviour. Simulated privileged operation: Set-SemtProgramFolderAcl really replaces the
    # DACL (protected, inheritance propagated by Windows) but keeps the test account as owner,
    # which needs no elevation. Trust is synthetic: temporary items belong to the test account.
    $currentSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $script:AclCalls = New-Object System.Collections.Generic.List[string]
    $script:Canonical = @{}
    $script:UntrustedItems = @{}
    $script:UntrustedShared = @{}
    $script:AfterAcl = $null
    function Set-SemtProgramFolderAcl {
        param([string]$Path)
        $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
        $script:AclCalls.Add($full)
        # DACL only: protected, every entry replaced (SACL and owner untouched).
        $directory = New-Object System.IO.DirectoryInfo($full)
        $security = if ($PSVersionTable.PSEdition -eq 'Core') { [System.IO.FileSystemAclExtensions]::GetAccessControl($directory, [System.Security.AccessControl.AccessControlSections]::Access) } else { $directory.GetAccessControl([System.Security.AccessControl.AccessControlSections]::Access) }
        $security.SetAccessRuleProtection($true, $false)
        foreach ($rule in @($security.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))) { [void]$security.RemoveAccessRuleSpecific($rule) }
        foreach ($sid in @('S-1-5-18', 'S-1-5-32-544', $currentSid)) { $security.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(([Security.Principal.SecurityIdentifier]$sid), 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow'))) }
        $security.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(([Security.Principal.SecurityIdentifier]'S-1-5-32-545'), 'ReadAndExecute, Synchronize', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
        if ($PSVersionTable.PSEdition -eq 'Core') { [System.IO.FileSystemAclExtensions]::SetAccessControl($directory, $security) } else { $directory.SetAccessControl($security) }
        $script:Canonical[$full] = $true; $script:UntrustedItems.Remove($full)
        if ($script:AfterAcl) { & $script:AfterAcl $full }
    }
    function Test-SemtCanonicalFolderAcl { param([string]$Path) return [bool]$script:Canonical[[IO.Path]::GetFullPath($Path).TrimEnd('\')] }
    function Test-SemtTrustedSharedParent {
        param([string]$Path)
        $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
        if ($script:UntrustedShared.ContainsKey($full)) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('synthetic {0}' -f $full) } }
        return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
    }
    function Test-SemtTrustedProgramPath {
        param([string]$Path, [switch]$Recurse)
        $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
        $items = @(Get-Item -LiteralPath $full -Force)
        if ($Recurse -and $items[0].PSIsContainer) { $items += @(Get-ChildItem -LiteralPath $full -Recurse -Force) }
        foreach ($item in $items) {
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Reparse point {0}' -f $item.FullName) } }
            if ($script:UntrustedItems.ContainsKey($item.FullName)) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('synthetic {0}' -f $item.FullName) } }
        }
        return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
    }
    function Get-SiblingState {
        # Owner, DACL (SDDL) and content of the sibling fixtures: proves that they were not changed.
        return (@($script:SiblingRoots | ForEach-Object { @(Get-Item -LiteralPath $_ -Force) + @(Get-ChildItem -LiteralPath $_ -Recurse -Force) } | Sort-Object FullName | ForEach-Object {
            $hash = if ($_.PSIsContainer) { 'DIR' } else { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
            '{0}|{1}|{2}' -f $_.FullName,$hash,(Get-Acl -LiteralPath $_.FullName).Sddl
        }) -join "`n")
    }
    function Reset-ChainCase { $script:AclCalls.Clear(); $script:UntrustedItems = @{}; $script:UntrustedShared = @{}; $script:AfterAcl = $null }

    # Sibling fixtures under the shared parents: another product, a legacy cache reused in place
    # (ReuseInPlace) and a folder with an explicit entry. All but the last inherit their ACL.
    $sharedWch = Join-Path $dataBase 'WCH'; $sharedSmart = Join-Path $sharedWch 'SmartM365'
    $otherProduct = Join-Path $sharedWch 'OtherProduct'; $legacyCache = Join-Path $sharedSmart 'Windows11UpgradeToolkit\SetupMedia\Win11-fr-FR'; $explicitSibling = Join-Path $sharedSmart 'IntuneHybridJoinToolkit'
    foreach ($folder in @((Join-Path $otherProduct 'Cache'), (Join-Path $legacyCache 'sources'), $explicitSibling)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    Set-Content -LiteralPath (Join-Path $otherProduct 'Cache\data.bin') -Value 'other product' -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $legacyCache 'setup.exe') -Value 'legacy setup' -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $legacyCache 'sources\install.wim') -Value 'legacy image' -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $explicitSibling 'state.json') -Value '{}' -Encoding ASCII
    $explicitAcl = Get-Acl -LiteralPath $explicitSibling
    $explicitAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(([Security.Principal.SecurityIdentifier]'S-1-5-32-545'), 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
    Set-Acl -LiteralPath $explicitSibling -AclObject $explicitAcl
    $script:SiblingRoots = @($otherProduct, (Join-Path $sharedSmart 'Windows11UpgradeToolkit'), $explicitSibling)
    $siblingsBefore = Get-SiblingState
    Test-Step 'fixtures: sibling folders inherit their ACL from the shared parents (inheritance active)' { -not (Get-Acl -LiteralPath $otherProduct).AreAccessRulesProtected -and -not (Get-Acl -LiteralPath $legacyCache).AreAccessRulesProtected -and @((Get-Acl -LiteralPath (Join-Path $legacyCache 'setup.exe')).Access | Where-Object IsInherited).Count -gt 0 }

    Reset-ChainCase
    $result = Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase
    Test-Step 'data chain: only the SEMT components are created and protected (product folder, then data root)' { $result -eq $dataRoot -and $script:AclCalls.Count -eq 2 -and $script:AclCalls[0] -eq $productRoot -and $script:AclCalls[1] -eq $dataRoot }
    Test-Step 'data chain: shared parents WCH and WCH\SmartM365 are never changed' { $script:AclCalls -notcontains $sharedWch -and $script:AclCalls -notcontains $sharedSmart -and -not (Get-Acl -LiteralPath $sharedWch).AreAccessRulesProtected -and -not (Get-Acl -LiteralPath $sharedSmart).AreAccessRulesProtected }
    Test-Step 'siblings: owner, ACL and content of the other product, the legacy cache and the explicit sibling unchanged after the creation' { (Get-SiblingState) -ceq $siblingsBefore }
    $script:AclCalls.Clear()
    [void](Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase)
    Test-Step 'data chain: a chain that already has exactly the protected ACL is verified, not rewritten' { $script:AclCalls.Count -eq 0 }

    # Trusted but not canonical (for example an older version): the replacement propagates below
    # the product folder only; siblings stay unchanged; the other campaign stays as it is.
    $otherCampaign = Join-Path $productRoot 'HybridJoin\Logs'; New-Item -ItemType Directory -Path $otherCampaign -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $otherCampaign 'run.log') -Value 'trusted log' -Encoding ASCII
    New-Item -ItemType Directory -Path (Join-Path $dataRoot 'State') -Force | Out-Null; Set-Content -LiteralPath (Join-Path $dataRoot 'State\state.json') -Value '{"trusted":true}' -Encoding ASCII
    Reset-ChainCase; $script:Canonical = @{}
    [void](Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase)
    Test-Step 'data chain: trusted, non-canonical SEMT components are replaced, shared parents still untouched' { $script:AclCalls.Count -eq 2 -and $script:AclCalls -notcontains $sharedWch -and $script:AclCalls -notcontains $sharedSmart }
    Test-Step 'siblings: unchanged after a real DACL replacement with inheritance propagated below the product folder' { (Get-SiblingState) -ceq $siblingsBefore }
    Test-Step 'trusted existing content is kept (content unchanged)' { (Get-Content -LiteralPath (Join-Path $dataRoot 'State\state.json') -Raw).Trim() -eq '{"trusted":true}' }

    # Shared parent not protected: blocking failure, nothing changed anywhere.
    Reset-ChainCase; $script:Canonical = @{}; $script:UntrustedShared = @{ $sharedSmart = $true }
    $failure = Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }
    Test-Step 'shared parent not protected: blocking failure, it is never changed, zero ACL change' { $failure -like 'SEMT_DATA_FOLDER_UNTRUSTED: shared folder*SmartM365*never changed; nothing was changed.' -and $script:AclCalls.Count -eq 0 -and (Get-SiblingState) -ceq $siblingsBefore }
    Test-Step 'verification only: an unprotected shared parent makes the data folder untrusted' { $r = Test-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase; -not $r.Trusted -and $r.Detail -match 'SmartM365' }

    # Provenance: a permission replacement never makes existing content trusted.
    $planted = Join-Path $dataRoot 'State\planted.json'; Set-Content -LiteralPath $planted -Value '{"planted":true}' -Encoding ASCII
    $plantedAcl = (Get-Acl -LiteralPath $planted).Sddl
    Reset-ChainCase; $script:Canonical = @{ $productRoot = $true; $dataRoot = $true }; $script:UntrustedItems = @{ $planted = $true }
    $failure = Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }
    Test-Step 'provenance: untrusted state file below a protected chain is refused, never reset or adopted (zero ACL change)' { $failure -like 'SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance*planted.json*nothing was changed.' -and $script:AclCalls.Count -eq 0 -and (Get-Acl -LiteralPath $planted).Sddl -eq $plantedAcl -and (Get-Content -LiteralPath $planted -Raw).Trim() -eq '{"planted":true}' }
    $script:Canonical = @{}
    $failure = Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }
    Test-Step 'provenance: same file when a DACL replacement is planned: still refused before the replacement (zero ACL change)' { $failure -like 'SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance*planted.json*' -and $script:AclCalls.Count -eq 0 -and (Get-Acl -LiteralPath $planted).Sddl -eq $plantedAcl }
    Remove-Item -LiteralPath $planted -Force

    $plantedMedia = Join-Path $dataRoot 'SetupMedia\Win11-fr-FR'; New-Item -ItemType Directory -Path (Join-Path $plantedMedia 'sources') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $plantedMedia 'setup.exe') -Value 'planted' -Encoding ASCII; Set-Content -LiteralPath (Join-Path $plantedMedia 'SmartM365-SetupMediaManifest.sha256.csv') -Value 'RelativePath,Length,SHA256' -Encoding ASCII
    Reset-ChainCase; $script:Canonical = @{}; $script:UntrustedItems = @{ $plantedMedia = $true }
    Test-Step 'provenance: setup media in a folder another account could modify is refused, even with its own manifest (zero ACL change)' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance*Win11-fr-FR*' -and $script:AclCalls.Count -eq 0 }
    Remove-Item -LiteralPath (Join-Path $dataRoot 'SetupMedia') -Recurse -Force

    Reset-ChainCase; $script:Canonical = @{}; $script:UntrustedItems = @{ $dataRoot = $true }
    Test-Step 'provenance: unprotected SEMT data root with content is refused, nothing adopted (zero ACL change)' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED:*was not protected*not empty*nothing was changed.' -and $script:AclCalls.Count -eq 0 }

    # Other SEMT campaign inside the propagation scope.
    Reset-ChainCase; $script:Canonical = @{}; $script:UntrustedItems = @{ (Join-Path $otherCampaign 'run.log') = $true }
    Test-Step 'propagation scope: untrusted content of another SEMT campaign is refused before the product folder is replaced (zero ACL change)' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance*run.log*' -and $script:AclCalls.Count -eq 0 }
    $script:Canonical = @{ $productRoot = $true; $dataRoot = $true }
    Test-Step 'outside the scope: other campaign content is not changed or required when nothing propagates to it' { (Get-ThrowMessage { [void](Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase) }) -eq '' -and $script:AclCalls.Count -eq 0 -and (Get-Content -LiteralPath (Join-Path $otherCampaign 'run.log') -Raw).Trim() -eq 'trusted log' }

    # Reproduced case: parents canonical, untrusted file listed before a junction.
    $script:OwnSddl = (Get-Acl -LiteralPath (Join-Path $tempRoot 'own')).Sddl
    $reproFile = Join-Path $dataRoot 'A-untrusted.json'; Set-Content -LiteralPath $reproFile -Value '{}' -Encoding ASCII
    $reproAcl = (Get-Acl -LiteralPath $reproFile).Sddl
    $null = & cmd.exe /d /c ('mklink /J "{0}" "{1}"' -f (Join-Path $dataRoot 'Z-junction'),(Join-Path $tempRoot 'own')) 2>&1
    Reset-ChainCase; $script:Canonical = @{ $productRoot = $true; $dataRoot = $true }; $script:UntrustedItems = @{ $reproFile = $true }
    $failure = Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }
    Test-Step 'reproduced case: parents canonical, untrusted file before a junction: refused with zero ACL change (file and target unchanged)' { $failure -like 'SEMT_DATA_FOLDER_UNTRUSTED*nothing was changed.' -and $script:AclCalls.Count -eq 0 -and (Get-Acl -LiteralPath $reproFile).Sddl -eq $reproAcl -and (Get-Acl -LiteralPath (Join-Path $tempRoot 'own')).Sddl -eq $script:OwnSddl }
    $script:UntrustedItems = @{}
    Test-Step 'reproduced case: junction alone below a canonical chain: refused, zero ACL change' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*Z-junction*' -and $script:AclCalls.Count -eq 0 }
    $null = & cmd.exe /d /c ('rmdir "{0}"' -f (Join-Path $dataRoot 'Z-junction')) 2>&1
    Remove-Item -LiteralPath $reproFile -Force

    # Links in the propagation scope when a replacement is planned.
    $null = & cmd.exe /d /c ('mklink /H "{0}" "{1}"' -f (Join-Path $dataRoot 'State\linked.json'),(Join-Path $dataRoot 'State\state.json')) 2>&1
    Reset-ChainCase; $script:Canonical = @{}
    Test-Step 'links: a real hard link below the data root stops the operation before any permission change' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*hard link*nothing was changed.' -and $script:AclCalls.Count -eq 0 }
    Remove-Item -LiteralPath (Join-Path $dataRoot 'State\linked.json') -Force
    $null = & cmd.exe /d /c ('mklink /H "{0}" "{1}"' -f (Join-Path $otherCampaign 'linked.log'),(Join-Path $otherCampaign 'run.log')) 2>&1
    Reset-ChainCase; $script:Canonical = @{}
    Test-Step 'links: a real hard link in another SEMT campaign stops the replacement of the product folder (zero ACL change)' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*linked.log*hard link*' -and $script:AclCalls.Count -eq 0 }
    Remove-Item -LiteralPath (Join-Path $otherCampaign 'linked.log') -Force
    $null = & cmd.exe /d /c ('mklink /J "{0}" "{1}"' -f (Join-Path $otherCampaign 'j'),(Join-Path $tempRoot 'own')) 2>&1
    Reset-ChainCase; $script:Canonical = @{}
    Test-Step 'links: a real junction in another SEMT campaign stops the replacement of the product folder (zero ACL change)' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED*' -and $script:AclCalls.Count -eq 0 }
    $null = & cmd.exe /d /c ('rmdir "{0}"' -f (Join-Path $otherCampaign 'j')) 2>&1

    # Change between verification and use: content written into a folder before it is protected.
    $emptyCampaign = Join-Path $productRoot 'EmptyCampaign'; New-Item -ItemType Directory -Path $emptyCampaign | Out-Null
    Reset-ChainCase; $script:Canonical = @{ $productRoot = $true }; $script:UntrustedItems = @{ $emptyCampaign = $true }
    [void](Initialize-SemtTrustedDataFolder -Path $emptyCampaign -BasePath $dataBase)
    Test-Step 'empty unprotected SEMT folder: protected and verified (nothing to adopt)' { $script:AclCalls.Count -eq 1 -and $script:AclCalls[0] -eq $emptyCampaign }
    Reset-ChainCase; $script:Canonical = @{ $productRoot = $true }; $script:UntrustedItems = @{ $emptyCampaign = $true }
    $script:AfterAcl = { param($Path) Set-Content -LiteralPath (Join-Path $Path 'late.json') -Value '{}' -Encoding ASCII }
    Test-Step 'change after the read-only pass: content written before protection is detected and refused' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $emptyCampaign -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED: content appeared in*EmptyCampaign*' }
    Remove-Item -LiteralPath $emptyCampaign -Recurse -Force
    $newCampaign = Join-Path $productRoot 'NewCampaign'
    Reset-ChainCase; $script:Canonical = @{ $productRoot = $true }
    $script:AfterAcl = { param($Path) Set-Content -LiteralPath (Join-Path $Path 'late.json') -Value '{}' -Encoding ASCII }
    Test-Step 'change after the read-only pass: content in a folder created here is detected and refused' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $newCampaign -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED: content appeared in*NewCampaign*' }
    Remove-Item -LiteralPath $newCampaign -Recurse -Force -ErrorAction SilentlyContinue

    Reset-ChainCase; $script:Canonical = @{}
    function Set-SemtProgramFolderAcl { param([string]$Path) throw 'synthetic: owner cannot be set' }
    Test-Step 'data chain: owner and DACL that cannot be replaced: blocking failure' { (Get-ThrowMessage { Initialize-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase }) -like 'SEMT_DATA_FOLDER_UNTRUSTED: owner and permissions of*could not be replaced*' }
    Test-Step 'siblings: unchanged after every case above' { (Get-SiblingState) -ceq $siblingsBefore }
    $script:UntrustedItems = @{ (Join-Path $dataRoot 'State\state.json') = $true }
    Test-Step 'verification only: an untrusted item makes the data folder untrusted' { -not (Test-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase).Trusted }
    $script:UntrustedItems = @{ $productRoot = $true }
    Test-Step 'verification only: an untrusted SEMT component makes the data folder untrusted' { $r = Test-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase; -not $r.Trusted -and $r.Detail -match 'SmartEndpointMigrationToolkit' }
    $script:UntrustedItems = @{}
    Test-Step 'verification only: trusted chain and content: trusted' { (Test-SemtTrustedDataFolder -Path $dataRoot -BasePath $dataBase).Trusted }
    foreach ($name in @('Set-SemtProgramFolderAcl','Test-SemtCanonicalFolderAcl','Test-SemtTrustedProgramPath','Test-SemtTrustedSharedParent')) { Remove-Item -Path ('Function:\' + $name) -ErrorAction SilentlyContinue }
    . ([scriptblock]::Create($reference))

    # Real endpoints: a data root that is not below ProgramData stops them before any write.
    $ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    foreach ($endpoint in @($w11Endpoint, $hjEndpoint)) {
        $badRoot = Join-Path $tempRoot ('endpoint-data-' + [IO.Path]::GetFileNameWithoutExtension($endpoint))
        $ErrorActionPreference = 'Continue'
        $output = @(& $ps51 -NoProfile -ExecutionPolicy Bypass -File $endpoint -DataRoot $badRoot -CorrelationId 'qualif-data-boundary' 2>&1 | ForEach-Object { [string]$_ })
        $exit = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        Write-TestResult -Condition ($exit -eq 1 -and ($output -join "`n") -match 'FINAL_STATUS=DATA_FOLDER_UNTRUSTED; EXIT_CODE=1; NEXT_ACTION=VERIFY_DATA_FOLDER_PERMISSIONS; CORRELATION_ID=qualif-data-boundary' -and -not (Test-Path -LiteralPath $badRoot)) -Message ("real endpoint {0}: data root outside ProgramData: DATA_FOLDER_UNTRUSTED, exit 1, nothing written" -f (Split-Path -Path $endpoint -Leaf)) -Detail ("exit={0}; {1}" -f $exit,(($output | Select-Object -Last 3) -join ' | '))
    }
    #endregion

    #region 4. Retry after reboot (endpoint functions)
    function Invoke-RetryCase {
        param([string]$Campaign, [string]$EndpointPath, [string]$ProgramFolderState = 'Trusted')
        $case = Join-Path $tempRoot ('retry-' + $Campaign + '-' + $ProgramFolderState)
        $script:StateDir = Join-Path $case 'Data\State'; $StateDir = $script:StateDir
        $script:DataRoot = Join-Path $case 'Data'; $DataRoot = $script:DataRoot
        New-Item -ItemType Directory -Path $script:StateDir -Force | Out-Null
        $script:RetryAfterRebootProgramDir = Join-Path $case ('ProgramFiles\WCH\SmartM365\SmartEndpointMigrationToolkit\{0}\RetryAfterReboot' -f $Campaign)
        $script:RetryAfterRebootStatePath = Join-Path $script:RetryAfterRebootProgramDir 'RetryAfterReboot.json'
        $script:RetryAfterRebootRunnerPath = Join-Path $script:RetryAfterRebootProgramDir 'RetryAfterRebootRunner.ps1'
        $script:RetryAfterRebootScriptPath = Join-Path $script:RetryAfterRebootProgramDir (Split-Path -Path $EndpointPath -Leaf)
        $script:RetryAfterRebootTaskName = ('SmartM365-SEMT-{0}-RetryAfterReboot' -f $Campaign)
        $script:EndpointScriptPath = Join-Path $case ('Data\{0}' -f (Split-Path -Path $EndpointPath -Leaf))
        Copy-Item -LiteralPath $EndpointPath -Destination $script:EndpointScriptPath
        $script:Schtasks = New-Object System.Collections.Generic.List[string]
        $script:ProgramFolderState = $ProgramFolderState
        $script:RetryAfterRebootMaxAttempts = 3; $RetryAfterRebootMaxAttempts = 3
        $script:RetryAfterRebootDelaySeconds = 0; $RetryAfterRebootDelaySeconds = 0
        $script:ComputerName = 'SEMT-QUALIF-PC'; $ComputerName = 'SEMT-QUALIF-PC'
        $names = @('Register-RetryAfterRebootTask','Write-RetryAfterRebootRunner','Write-RetryAfterRebootState','Read-RetryAfterRebootState','Unregister-RetryAfterRebootTask','Start-RetryAfterRebootTaskRun','Set-RetryAfterRebootStateProperty')
        . ([scriptblock]::Create((Get-FunctionText -Path $EndpointPath -Name $names)))
        function Initialize-SemtTrustedProgramFolder { param([string]$Path, [string]$BasePath) New-Item -ItemType Directory -Path $Path -Force | Out-Null; return $Path }
        function Test-SemtTrustedProgramPath {
            param([string]$Path, [switch]$Recurse)
            if ($script:ProgramFolderState -ne 'Trusted') { return [pscustomobject]@{ Trusted = $false; State = $script:ProgramFolderState; Detail = 'synthetic' } }
            return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
        }
        function New-RetryAfterRebootArgumentList { return @('-DataRoot', $script:DataRoot) }
        function Protect-Windows11ToolkitPathAcl { param([string]$Path, [switch]$Directory, [switch]$RequireSuccess) }
        function Protect-RetryAfterRebootPathAcl { param([string]$Path, [switch]$Directory) }
        function New-SmartDirectory { param([string]$Path) New-Item -ItemType Directory -Path $Path -Force | Out-Null }
        function Write-SmartLog { param([string]$Message, [string]$Level = 'INFO') }
        function Write-RunLog { param([string]$Message) }
        function schtasks.exe { $script:Schtasks.Add(($args -join ' ')); $global:LASTEXITCODE = 0 }
        $outcome = [pscustomobject]@{ Error = ''; Case = $case; TaskRunError = ''; Unregistered = $false }
        try { Register-RetryAfterRebootTask -Reason 'qualification' } catch { $outcome.Error = $_.Exception.Message }
        if (-not $outcome.Error) {
            $outcome | Add-Member -NotePropertyName State -NotePropertyValue (Get-Content -LiteralPath $script:RetryAfterRebootStatePath -Raw | ConvertFrom-Json)
            $outcome | Add-Member -NotePropertyName Runner -NotePropertyValue (Get-Content -LiteralPath $script:RetryAfterRebootRunnerPath -Raw)
            $script:ProgramFolderState = 'Untrusted'
            try { [void](Start-RetryAfterRebootTaskRun) } catch { $outcome.TaskRunError = $_.Exception.Message }
            $script:ProgramFolderState = 'Trusted'
            Unregister-RetryAfterRebootTask -Reason 'qualification'
            $outcome.Unregistered = -not (Test-Path -LiteralPath $script:RetryAfterRebootScriptPath) -and -not (Test-Path -LiteralPath $script:RetryAfterRebootRunnerPath) -and -not (Test-Path -LiteralPath $script:RetryAfterRebootStatePath)
        }
        return $outcome
    }

    foreach ($retry in @(@{ Campaign = 'Windows11Upgrade'; Path = $w11Endpoint }, @{ Campaign = 'HybridJoin'; Path = $hjEndpoint })) {
        $o = Invoke-RetryCase -Campaign $retry.Campaign -EndpointPath $retry.Path
        $programDir = Join-Path $o.Case ('ProgramFiles\WCH\SmartM365\SmartEndpointMigrationToolkit\{0}\RetryAfterReboot' -f $retry.Campaign)
        $copy = Join-Path $programDir (Split-Path -Path $retry.Path -Leaf)
        Test-Step ("{0} retry: state points to the endpoint copy under Program Files (same SHA256)" -f $retry.Campaign) { $o.Error -eq '' -and [string]$o.State.ScriptPath -eq $copy -and -not $o.Unregistered -eq $false }
        Test-Step ("{0} retry: the SYSTEM task starts the runner under Program Files" -f $retry.Campaign) { @($script:Schtasks | Where-Object { $_ -like '/Create*' }).Count -eq 1 -and @($script:Schtasks | Where-Object { $_ -like ('*{0}*' -f (Join-Path $programDir 'RetryAfterRebootRunner.ps1')) }).Count -eq 1 }
        Test-Step ("{0} retry: runner reads its state from Program Files; no script or state under ProgramData" -f $retry.Campaign) { $o.Runner.Contains((Join-Path $programDir 'RetryAfterReboot.json').Replace("'", "''")) -and @(Get-ChildItem -LiteralPath (Join-Path $o.Case 'Data\State') -Force).Count -eq 0 }
        Test-Step ("{0} retry: a task run refuses an untrusted program folder" -f $retry.Campaign) { $o.TaskRunError -like 'SEMT_PROGRAM_FOLDER_UNTRUSTED*' }
        Test-Step ("{0} retry: cleanup removes runner, state and endpoint copy" -f $retry.Campaign) { $o.Unregistered }
        $o = Invoke-RetryCase -Campaign $retry.Campaign -EndpointPath $retry.Path -ProgramFolderState 'Untrusted'
        Test-Step ("{0} retry: untrusted program folder: refused before the task is registered" -f $retry.Campaign) { $o.Error -like 'SEMT_PROGRAM_FOLDER_UNTRUSTED*' -and @($script:Schtasks | Where-Object { $_ -like '/Create*' }).Count -eq 0 }
    }
    Remove-Item -Path Function:\Test-SemtTrustedProgramPath -ErrorAction SilentlyContinue
    #endregion

    #region 5. PsExec payload (Core)
    $core = Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force -PassThru -ErrorAction Stop
    $targetRoot = Join-Path $tempRoot 'target'
    New-Item -ItemType Directory -Path (Join-Path $targetRoot 'C\Program Files') -Force | Out-Null
    $log = Join-Path $tempRoot 'worker.log'
    $copyCase = {
        param([string]$TrustState)
        & $core {
            param($TargetRoot, $TrustState)
            # Administrative shares map to a temporary target tree; trust and ACL boundaries are stubbed.
            function script:ConvertTo-SemtAdminSharePath { param([string]$ComputerName, [string]$LocalPath) return (Join-Path $TargetRoot ($LocalPath.Substring(0, 1) + $LocalPath.Substring(2))) }
            function script:Protect-SemtRemotePathAcl { param([string]$Path, [switch]$Directory) return 'Protected' }
            function script:Set-SemtProgramFolderAcl { param([string]$Path) }
            $script:QualifTrustState = $TrustState
            function script:Test-SemtTrustedProgramPath { param([string]$Path, [switch]$Recurse) [pscustomobject]@{ Trusted = ($script:QualifTrustState -eq 'Trusted'); State = $script:QualifTrustState; Detail = 'synthetic' } }
        } $targetRoot $TrustState
    }
    $endpointSha = (Get-FileHash -LiteralPath $w11Endpoint -Algorithm SHA256).Hash
    $dataRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade'
    $programRoot = & $core { Get-SemtTargetProgramRoot -Campaign 'Windows11Upgrade' }
    & $copyCase 'Trusted'
    $copied = Copy-SemtRemotePayload -ConnectionTarget 'SEMT-QUALIF-PC' -LocalScriptPath $w11Endpoint -ExpectedSha256 $endpointSha -RemoteDataRoot $dataRoot -RemoteProductRoot (Split-Path -Path $dataRoot -Parent) -RemoteProgramRoot $programRoot -LogPath $log
    $programCopy = Join-Path $targetRoot ('C' + $programRoot.Substring(2) + '\' + (Split-Path -Path $w11Endpoint -Leaf))
    Test-Step 'PsExec: target program root is C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>' { $programRoot -eq 'C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade' }
    Test-Step 'PsExec: endpoint copied under Program Files on the target with the expected SHA256' { (Test-Path -LiteralPath $programCopy) -and $copied.RemoteSha256 -eq $endpointSha }
    Test-Step 'PsExec: the script path given to PsExec is under Program Files' { [string]$copied.RemoteScriptPath -eq (Join-Path $programRoot (Split-Path -Path $w11Endpoint -Leaf)) }
    Test-Step 'PsExec: no endpoint script under the ProgramData data root of the target' { @(Get-ChildItem -LiteralPath (Join-Path $targetRoot 'C\ProgramData') -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq '.ps1' }).Count -eq 0 }
    Remove-Item -LiteralPath $programCopy -Force
    & $copyCase 'Untrusted'
    Test-Step 'PsExec: untrusted program folder on the target: copy refused, nothing copied' { (Get-ThrowMessage { Copy-SemtRemotePayload -ConnectionTarget 'SEMT-QUALIF-PC' -LocalScriptPath $w11Endpoint -ExpectedSha256 $endpointSha -RemoteDataRoot $dataRoot -RemoteProductRoot (Split-Path -Path $dataRoot -Parent) -RemoteProgramRoot $programRoot -LogPath $log }) -like 'SEMT_PROGRAM_FOLDER_UNTRUSTED*' -and -not (Test-Path -LiteralPath $programCopy) }
    Remove-Module -Name $core.Name -Force
    $workerText = [IO.File]::ReadAllText((Join-Path $packageRoot 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'))
    $engineText = [IO.File]::ReadAllText((Join-Path $packageRoot 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'))
    Test-Step 'PsExec: the launcher contract carries RemoteProgramRoot and the worker refuses a contract without it' { $engineText.Contains('RemoteProgramRoot = $script:TargetProgramRoot') -and $workerText.Contains("SEMT_CONTRACT_INVALID: Endpoint.RemoteProgramRoot is missing") -and $workerText.Contains('-RemoteScriptPath ([string]$copied.RemoteScriptPath)') }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    try { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction Stop } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDpgMRywYcCM9ia
# 7KxAN3DwtMzKaXJqdwzoOXwAqkPP1qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIANSQT4AjsQefThTA8hPNTghJSamQ8DMXPMC+tzP4JIWMA0GCSqG
# SIb3DQEBAQUABIIBgKnAjkOMObtoJklE0eG7VWfjHg0I+r7AW1VIogRIGgmWI34f
# NG0HTNo6Q1vLcOAYnymTXM1T+D4nIIJM6pskbQvd8IE5LmCzHfBq7akSUeora9iN
# x/YLmO7qj3kCBab2iQkLNHlRSqE03NvX0pFs0bTpUtP0fQS4O5QJzVXmZurfmT8P
# l51qlGxzr6EVz79mbWZ5w8bcS/pwClFsGhBSX+J9Jr5NSAZYF1m+8BxNb250otGA
# 7fN86RJto4FAf7fDeufEcjX1vouccdTMnLFLf05AOS7uy4/Kr3QWvpUGfrfs53B1
# seusMwN4ubfLMzMrSqnPWgHnNKx9tBFLdQkdMK90A4m/SFLRa5FD7cH5Ky2yk1FO
# vYQGiqbks7S+gjSrTVzMkwNkaYFCo0xhcBzqC8oCDWtVFhchv+iSZjmSbR9kVciR
# 8iFoSH3UicwAbyBj9Oz1FGYtJAIMrhaKFkP9TPKoHqPuITEm7vxACxkx3JAjz3xq
# tpj2/C7GatwoEX0mhqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NThaMC8GCSqGSIb3DQEJBDEiBCB2wKu+vDCwaRh0DkJ7U+g+k5cyzEKcSliwQaez
# L80nqzANBgkqhkiG9w0BAQEFAASCAgAF1cXqTa3167ydQdw/kNC/Z/gVjIUxK0+3
# pzO471yY+MiMws0Q/vHe5ztCc9E/i/i2oZNJwlIsczVl5cTR+6QXI9VLTKGhnMdy
# jcYx1yn6/ycIpmgIp3rzH8wG63wtTDHlOs6RIA21qV6GlauHWXAiuiAZkfEVP9m9
# S+xCxhqx+Nhgt1Pq3g6PL1qGrX7qsfE57JJtZe4pVguOsazhOK5kzqu8hFvPI2eG
# aDPMPv+/CHS4vjCblGlB2Sc8g7Lx+F6tgxrsg5tXSRRiBxN8TxMwhm9YyAW9SmRA
# uw8XFP3QbYC8H3DHoy5to7NDSthFXWeoBdMwQyuLNBPMv8QlS7mHlNR0/g1wyFaj
# qnVm5cVPFiF2QN6A5jvz5hKAh2MUOr9tS1onx9dZpDRDuvwLr1KcNoizHVLgAnRq
# uR7w/VhQoCy4LpNPGuCgP26Hv16Nech4CXi4prYevSjLifOXJ/Iz6zHtxKXtaLdW
# L8PaCh8/5NY+SgOEkMDIwrctfwg3t8hUwXTNUwJTvArHAEl7o1b3ArAZM3pdnN5t
# bCw7Cyp3zkZVwl1CLjEJe5kPpNQ1ugR1annd1IZqugx6Lcm+r1rRFFz+SDV4Kg/3
# 20xNdwWV0MJBCRghQhGaSmhs2AqKY2d3zgu+HIPc0KPR7zjHxfpLyGeNyiD4FfqC
# 5uYxTWQHXQ==
# SIG # End signature block
