<#
.SYNOPSIS
    Offline tests of the Windows 11 setup media resolution for PsExec and Intune.
.DESCRIPTION
    Synthetic media only (tiny files with a real SHA256 manifest), temporary folders only.
    1. The shared resolution region is identical in the endpoint script and the Intune helper.
    2. Resolution decisions: new cache, legacy Intune-owned, legacy adoptable, absent, corrupt,
       incompatible (language and media build), in use (live lock, live process), indeterminate,
       untrusted (insufficient protection, proven with real ACLs) separated from indeterminate
       (permissions or folder tree that cannot be read, real ACLs) and from absent.
    3. Endpoint (Resolve-SetupUpgradeExecutable with stubbed Windows Setup checks): selected path
       used, no copy when media is reusable, legacy media never modified, unknown path rejected,
       duplication blocked, adoption.
    4. Installer, runner and detection: the real scripts run in a child Windows PowerShell from a
       temporary copy where a shim replaces only the operational boundaries (HKLM registry, task
       scheduler, OS query, ACL hardening, endpoint process start, legacy registry ownership).
       Robocopy and the file system are real.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
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
$campaignRoot = Join-Path $packageRoot 'Campaigns\Windows11Upgrade'
$endpointPath = Join-Path $campaignRoot 'Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$intuneSource = Join-Path $campaignRoot 'IntuneWin32\Source'
$helperPath = Join-Path $intuneSource 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

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

function Get-FileSha {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function New-SyntheticMedia {
    # setup.exe, sources\install.wim, sources\lang.ini and a real SHA256 manifest.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$Language = 'fr-FR', [string]$Build = 'A')
    New-Item -ItemType Directory -Path (Join-Path $Path 'sources') -Force | Out-Null
    [IO.File]::WriteAllBytes((Join-Path $Path 'setup.exe'), [Text.Encoding]::ASCII.GetBytes(('SYNTHETIC-SETUP-{0}-' -f $Build) + ('x' * 4096)))
    [IO.File]::WriteAllBytes((Join-Path $Path 'sources\install.wim'), [Text.Encoding]::ASCII.GetBytes(('SYNTHETIC-WIM-{0}-' -f $Build) + ('y' * 8192)))
    Set-Content -LiteralPath (Join-Path $Path 'sources\lang.ini') -Value ("[Available UI Languages]`r`n{0} = 3`r`n" -f $Language) -Encoding ASCII
    $rows = foreach ($file in @(Get-ChildItem -LiteralPath $Path -Recurse -File | Where-Object { $_.Name -ne 'SmartM365-SetupMediaManifest.sha256.csv' })) {
        [pscustomobject]@{ RelativePath = $file.FullName.Substring($Path.Length + 1); Length = $file.Length; SHA256 = (Get-FileSha $file.FullName) }
    }
    $rows | Export-Csv -LiteralPath (Join-Path $Path 'SmartM365-SetupMediaManifest.sha256.csv') -NoTypeInformation -Encoding ASCII
}

if (-not ('SemtTestAcl' -as [type]) -or -not [SemtTestAcl].GetMethod('SetDacl')) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SemtTestAcl {
    [StructLayout(LayoutKind.Sequential)] struct UnicodeString { public ushort Length; public ushort MaximumLength; public IntPtr Buffer; }
    [StructLayout(LayoutKind.Sequential)] struct ObjectAttributes { public int Length; public IntPtr RootDirectory; public IntPtr ObjectName; public uint Attributes; public IntPtr SecurityDescriptor; public IntPtr SecurityQualityOfService; }
    [StructLayout(LayoutKind.Sequential)] struct IoStatusBlock { public IntPtr Status; public IntPtr Information; }
    [DllImport("ntdll.dll")] static extern int NtOpenFile(out IntPtr handle, uint access, ref ObjectAttributes attributes, out IoStatusBlock status, uint share, uint options);
    [DllImport("ntdll.dll")] static extern int NtClose(IntPtr handle);
    [DllImport("ntdll.dll")] static extern int NtSetSecurityObject(IntPtr handle, uint info, byte[] descriptor);
    // Opens with WRITE_DAC only (no SYNCHRONIZE, no READ_CONTROL), then writes the DACL of a
    // self-relative security descriptor directly (no read of the current descriptor).
    public static void SetDacl(string path, byte[] descriptor) {
        string ntPath = @"\??\" + path;
        IntPtr buffer = Marshal.StringToHGlobalUni(ntPath);
        IntPtr name = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(UnicodeString)));
        try {
            UnicodeString us = new UnicodeString(); us.Buffer = buffer; us.Length = (ushort)(ntPath.Length * 2); us.MaximumLength = (ushort)(ntPath.Length * 2 + 2);
            Marshal.StructureToPtr(us, name, false);
            ObjectAttributes oa = new ObjectAttributes(); oa.Length = Marshal.SizeOf(typeof(ObjectAttributes)); oa.ObjectName = name; oa.Attributes = 0x40;
            IntPtr handle; IoStatusBlock iosb;
            int status = NtOpenFile(out handle, 0x00040000, ref oa, out iosb, 7, 0);
            if (status != 0) throw new InvalidOperationException(string.Format("NtOpenFile(WRITE_DAC) failed: 0x{0:X8} {1}", status, path));
            try {
                int rc = NtSetSecurityObject(handle, 0x4u, descriptor);
                if (rc != 0) throw new InvalidOperationException(string.Format("NtSetSecurityObject failed: 0x{0:X8} {1}", rc, path));
            } finally { NtClose(handle); }
        } finally { Marshal.FreeHGlobal(name); Marshal.FreeHGlobal(buffer); }
    }
}
'@
}
function Set-TestSddl {
    # Writes only a protected DACL through a WRITE_DAC handle (works when READ_CONTROL and
    # SYNCHRONIZE are not granted), on both PowerShell engines.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Sddl)
    $raw = New-Object System.Security.AccessControl.RawSecurityDescriptor($Sddl)
    $bytes = New-Object byte[] $raw.BinaryLength
    $raw.GetBinaryForm($bytes, 0)
    [SemtTestAcl]::SetDacl([System.IO.Path]::GetFullPath($Path), $bytes)
}

function Get-TreeState {
    # Content hashes and DACLs of every item: proves that a folder was not modified.
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return '<absent>' }
    $items = @(Get-Item -LiteralPath $Path -Force) + @(Get-ChildItem -LiteralPath $Path -Recurse -Force)
    return (@($items | Sort-Object FullName | ForEach-Object {
        $hash = if ($_.PSIsContainer) { 'DIR' } else { Get-FileSha $_.FullName }
        '{0}|{1}|{2}|{3}' -f $_.FullName.Substring($Path.Length),$hash,$_.LastWriteTimeUtc.Ticks,(Get-Acl -LiteralPath $_.FullName).Sddl
    }) -join "`n")
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-MediaTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    #region 1. Shared region
    Test-Step 'setup media resolution region is identical in the endpoint and the Intune helper' { (Get-RegionText -Path $endpointPath -Region 'SEMT setup media resolution v1') -ceq (Get-RegionText -Path $helperPath -Region 'SEMT setup media resolution v1') }
    Test-Step 'resolution region writes nothing (no Move/Remove/Set-Content/icacls/registry write)' { (Get-RegionText -Path $endpointPath -Region 'SEMT setup media resolution v1') -notmatch '(Move-Item|Remove-Item|Set-Content|Add-Content|Out-File|WriteAll|Directory\]::Move|icacls|Set-Acl|SetValue|CreateSubKey|DeleteSubKey|New-Item)' }
    #endregion

    #region 2. Resolution decisions (Intune validator with real SHA256 checks)
    . $helperPath
    # Real trust check: temporary folders are owned and writable by the test account, system
    # folders are not. The synthetic cases below then stub the check as "trusted".
    $untrustedProbe = Join-Path $tempRoot 'trust-probe'; New-Item -ItemType Directory -Path $untrustedProbe -Force | Out-Null
    Test-Step 'trust check (real): folder owned and writable by a standard account is Untrusted (proven)' { $t = Test-SemtMediaWriteRestricted -Path $untrustedProbe; -not $t.Trusted -and $t.State -eq 'Untrusted' }
    Test-Step 'trust check (real): system folder writable only by SYSTEM/Administrators/TrustedInstaller is trusted' { $t = Test-SemtMediaWriteRestricted -Path (Join-Path $env:SystemRoot 'System32\drivers\etc'); $t.Trusted -and $t.State -eq 'Trusted' }
    $script:FakeOwners = @{}
    $script:OwnerState = 'Read'
    $owner = { param($Path) [pscustomobject]@{ State = $script:OwnerState; Owners = @($script:FakeOwners[(Split-Path $Path -Leaf)] | Where-Object { $_ }); Detail = 'synthetic' } }
    function New-Case { param([string]$Name) $root = Join-Path $tempRoot ('case-' + $Name); New-Item -ItemType Directory -Path $root -Force | Out-Null; return [pscustomobject]@{ Data = (Join-Path $root 'PD\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade'); Legacy = (Join-Path $root 'Legacy'); NewCache = (Join-Path $root 'PD\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\SetupMedia\Win11-fr-FR'); LegacyCache = (Join-Path $root 'Legacy\SetupMedia\Win11-fr-FR'); Program = (Join-Path $root 'ProgramFiles\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade') } }
    function Resolve-Case { param($Case, [string]$Language = 'fr-FR', [string]$Reference = '') Resolve-SemtIntuneSetupMedia -DataRoot $Case.Data -SetupCacheFolder 'Win11-fr-FR' -Language $Language -ReferenceManifestPath $Reference -LegacyRoot $Case.Legacy -OwnerProvider $owner }

    # Real permissions, no stub: legacy media in a temporary folder is owned and writable by the
    # test account (insufficient protection); a media folder without READ_CONTROL for anyone makes
    # its permissions unreadable, and one without list access makes its tree unreadable. OWNER
    # RIGHTS keeps only WRITE_DAC, so the test restores the DACL afterwards.
    $case = New-Case 'real-untrusted'; New-SyntheticMedia -Path $case.LegacyCache
    $before = Get-TreeState $case.LegacyCache
    $r = Resolve-Case $case
    Test-Step 'real ACL: present legacy media with insufficient protection: Decision=Untrusted, never Absent' { $r.Decision -eq 'Untrusted' -and $r.LegacyState -eq 'Untrusted' -and $r.Path -eq '' -and $r.LegacyPath -eq $case.LegacyCache }
    Test-Step 'real ACL: untrusted legacy media requires a duplication authorization for any second copy' { $r.DuplicationRequired -and $r.UntrustedPath -eq $case.LegacyCache }
    Test-Step 'real ACL: untrusted legacy media is not changed by the resolution (content, dates, DACL)' { (Get-TreeState $case.LegacyCache) -ceq $before }

    $case = New-Case 'real-unreadable'; New-SyntheticMedia -Path $case.LegacyCache
    $lockedFolder = $case.LegacyCache
    Set-TestSddl -Path $lockedFolder -Sddl 'D:P(A;;0x1100A1;;;WD)(A;;0x40000;;;OW)'
    try {
        Test-Step 'trust check (real): permissions that cannot be read: State=Unreadable (not Untrusted)' { $t = Test-SemtMediaWriteRestricted -Path $case.LegacyCache; -not $t.Trusted -and $t.State -eq 'Unreadable' }
        $r = Resolve-Case $case
        Test-Step 'real ACL: unreadable permissions of legacy media: Decision=Indeterminate (blocks)' { $r.Decision -eq 'Indeterminate' -and $r.LegacyState -eq 'Indeterminate' -and $r.LegacyDetail -match 'could not be verified' }
    }
    finally { Set-TestSddl -Path $lockedFolder -Sddl 'D:PAI(A;OICI;FA;;;WD)' }

    $case = New-Case 'real-unlisted'; New-SyntheticMedia -Path $case.LegacyCache
    $lockedFolder = Join-Path $case.LegacyCache 'sources\locked'; New-Item -ItemType Directory -Path $lockedFolder -Force | Out-Null
    Set-TestSddl -Path $lockedFolder -Sddl 'D:P(A;;0x10080;;;WD)(A;;0x40000;;;OW)'
    try {
        $r = Resolve-Case $case
        Test-Step 'real ACL: legacy media folder tree that cannot be listed: Decision=Indeterminate' { $r.Decision -eq 'Indeterminate' -and $r.LegacyDetail -match 'tree could not be read' }
    }
    finally { Set-TestSddl -Path $lockedFolder -Sddl 'D:PAI(A;OICI;FA;;;WD)' }

    $script:TrustAll = $true
    $script:TrustState = ''
    $script:TrustThrows = $false
    function Test-SemtMediaWriteRestricted {
        param([string]$Path)
        if ($script:TrustThrows) { throw 'synthetic trust failure' }
        if ($script:TrustState -eq 'Null') { return $null }
        if ($script:TrustState) { return [pscustomobject]@{ Trusted = ($script:TrustState -eq 'Trusted'); State = $script:TrustState; Detail = 'synthetic' } }
        [pscustomobject]@{ Trusted = $script:TrustAll; Detail = 'synthetic' }
    }

    $case = New-Case 'new'; New-SyntheticMedia -Path $case.NewCache
    $r = Resolve-Case $case
    Test-Step 'valid new cache: NewCacheValid, path of this product, not external' { $r.Decision -eq 'NewCacheValid' -and $r.Path -eq $case.NewCache -and -not $r.External }

    $case = New-Case 'owned'; New-SyntheticMedia -Path $case.LegacyCache; $script:FakeOwners['Win11-fr-FR'] = 'SmartM365-W11-Win11-fr-FR:Installed'
    $before = Get-TreeState $case.LegacyCache
    $r = Resolve-Case $case
    Test-Step 'legacy media owned by its Intune package: LegacyIntuneOwned, used in place, external' { $r.Decision -eq 'LegacyIntuneOwned' -and $r.Path -eq $case.LegacyCache -and $r.External -and $r.Owner -like '*Installed*' }
    Test-Step 'resolution leaves the legacy media unchanged (content, dates, ACL)' { (Get-TreeState $case.LegacyCache) -ceq $before }
    $script:FakeOwners = @{}

    $case = New-Case 'adopt'; New-SyntheticMedia -Path $case.LegacyCache
    $r = Resolve-Case $case
    Test-Step 'idle unmanaged legacy media: LegacyAdoptable' { $r.Decision -eq 'LegacyAdoptable' -and $r.Path -eq $case.LegacyCache -and -not $r.External }

    $case = New-Case 'absent'
    Test-Step 'no media: Absent' { (Resolve-Case $case).Decision -eq 'Absent' }

    $case = New-Case 'corrupt'; New-SyntheticMedia -Path $case.NewCache; Add-Content -LiteralPath (Join-Path $case.NewCache 'sources\install.wim') -Value 'tampered'
    $r = Resolve-Case $case
    Test-Step 'tampered new cache (SHA256 mismatch): Corrupt' { $r.Decision -eq 'Corrupt' -and $r.NewCacheState -eq 'Corrupt' -and $r.NewCacheDetail -match 'SHA256 mismatch|Length mismatch' }

    $case = New-Case 'lang'; New-SyntheticMedia -Path $case.NewCache -Language 'de-DE'
    Test-Step 'other language: Incompatible' { (Resolve-Case $case -Language 'fr-FR').Decision -eq 'Incompatible' }

    $case = New-Case 'build'; New-SyntheticMedia -Path $case.NewCache -Build 'A'; $reference = Join-Path $tempRoot 'reference-B'; New-SyntheticMedia -Path $reference -Build 'B'
    Test-Step 'other media build than the packaged media: Incompatible' { (Resolve-Case $case -Reference (Join-Path $reference 'SmartM365-SetupMediaManifest.sha256.csv')).Decision -eq 'Incompatible' }

    # Manifest required: a missing manifest or a missing required reference refuses the media.
    $case = New-Case 'no-manifest'; New-SyntheticMedia -Path $case.NewCache; Remove-Item -LiteralPath (Join-Path $case.NewCache 'SmartM365-SetupMediaManifest.sha256.csv') -Force
    Test-Step 'new cache without its SHA256 manifest: Corrupt, never NewCacheValid' { $r = Resolve-Case $case; $r.Decision -eq 'Corrupt' -and $r.NewCacheDetail -match 'manifest not found' }
    $case = New-Case 'no-reference'; New-SyntheticMedia -Path $case.NewCache
    Test-Step 'required package reference manifest missing: refused (Corrupt), the comparison is never skipped' { $r = Resolve-Case $case -Reference (Join-Path $tempRoot 'missing-reference\SmartM365-SetupMediaManifest.sha256.csv'); $r.Decision -eq 'Corrupt' -and $r.NewCacheDetail -match 'Reference setup media manifest' }
    $case = New-Case 'self-manifest'; New-SyntheticMedia -Path $case.NewCache -Build 'A'
    $referenceA = Join-Path $tempRoot 'reference-A'; New-SyntheticMedia -Path $referenceA -Build 'A'
    [IO.File]::WriteAllBytes((Join-Path $case.NewCache 'setup.exe'), [Text.Encoding]::ASCII.GetBytes('SYNTHETIC-SETUP-REPLACED-' + ('z' * 4096)))
    $rows = foreach ($file in @(Get-ChildItem -LiteralPath $case.NewCache -Recurse -File | Where-Object { $_.Name -ne 'SmartM365-SetupMediaManifest.sha256.csv' })) { [pscustomobject]@{ RelativePath = $file.FullName.Substring($case.NewCache.Length + 1); Length = $file.Length; SHA256 = (Get-FileSha $file.FullName) } }
    $rows | Export-Csv -LiteralPath (Join-Path $case.NewCache 'SmartM365-SetupMediaManifest.sha256.csv') -NoTypeInformation -Encoding ASCII
    Test-Step 'replaced file with a manifest regenerated next to it: valid on its own manifest only, refused against the package reference' { $alone = Resolve-Case $case; $withReference = Resolve-Case $case -Reference (Join-Path $referenceA 'SmartM365-SetupMediaManifest.sha256.csv'); $alone.Decision -eq 'NewCacheValid' -and $withReference.Decision -eq 'Incompatible' }

    # Real endpoint check (Test-SetupMedia -ValidateManifest): the manifest result is checked.
    & {
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($endpointPath, [ref]$tokens, [ref]$errors)
        foreach ($name in @('Test-SetupMedia','Test-SetupMediaIntegrityManifest','Get-SetupMediaIntegrityManifestPath')) {
            $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
            . ([scriptblock]::Create($definition.Extent.Text))
        }
        function Test-SetupExecutableSignature { param([string]$SetupExe) }
        function Test-SetupInstallImageReadable { param([string]$MediaRoot) return $true }
        function Write-SmartLog { param([string]$Message, [string]$Level = 'INFO') }
        function Get-SmartFileSha256 { param([string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
        $script:SetupMediaManifestFileName = 'SmartM365-SetupMediaManifest.sha256.csv'
        $media = Join-Path $tempRoot 'endpoint-manifest\Win11-fr-FR'; New-SyntheticMedia -Path $media
        [IO.File]::WriteAllBytes((Join-Path $media 'setup.exe'), [Text.Encoding]::ASCII.GetBytes('SYNTHETIC-SETUP-' + ('x' * 70000)))
        $rows = foreach ($file in @(Get-ChildItem -LiteralPath $media -Recurse -File | Where-Object { $_.Name -ne 'SmartM365-SetupMediaManifest.sha256.csv' })) { [pscustomobject]@{ RelativePath = $file.FullName.Substring($media.Length + 1); Length = $file.Length; SHA256 = (Get-FileSha $file.FullName) } }
        $rows | Export-Csv -LiteralPath (Join-Path $media 'SmartM365-SetupMediaManifest.sha256.csv') -NoTypeInformation -Encoding ASCII
        Test-Step 'endpoint Test-SetupMedia (real): media with a valid manifest is accepted' { (Test-SetupMedia -MediaPath $media -ValidateManifest) -like '*setup.exe' }
        Remove-Item -LiteralPath (Join-Path $media 'SmartM365-SetupMediaManifest.sha256.csv') -Force
        Test-Step 'endpoint Test-SetupMedia (real): missing manifest is refused (the $false result is no longer ignored)' { $m = ''; try { [void](Test-SetupMedia -MediaPath $media -ValidateManifest) } catch { $m = $_.Exception.Message }; $m -like 'Setup media integrity manifest validation failed.*manifest not found*' }
    }

    $case = New-Case 'notadoptable'; New-SyntheticMedia -Path $case.LegacyCache; New-Item -ItemType Directory -Path $case.NewCache -Force | Out-Null
    $r = Resolve-Case $case
    Test-Step 'valid legacy media next to a broken new cache: not adopted, LegacyValidRemains' { $r.Decision -eq 'Corrupt' -and $r.LegacyState -eq 'NotAdoptable' -and $r.LegacyValidRemains }

    $case = New-Case 'lock'; New-SyntheticMedia -Path $case.LegacyCache
    $lockDir = Join-Path $case.Legacy 'Locks\SetupCache-Win11-fr-FR.lock'; New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value ('{ "PID": ' + $PID + ' }') -Encoding ASCII
    Test-Step 'legacy media lock held by a live process: InUse' { (Resolve-SemtIntuneSetupMedia -DataRoot $case.Data -SetupCacheFolder 'Win11-fr-FR' -Language 'fr-FR' -LegacyRoot $case.Legacy -OwnerProvider $owner).Decision -eq 'InUse' }
    Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value '{ broken' -Encoding ASCII
    Test-Step 'unreadable legacy media lock: Indeterminate' { (Resolve-SemtIntuneSetupMedia -DataRoot $case.Data -SetupCacheFolder 'Win11-fr-FR' -Language 'fr-FR' -LegacyRoot $case.Legacy -OwnerProvider $owner).Decision -eq 'Indeterminate' }
    Remove-Item -LiteralPath $lockDir -Recurse -Force

    $case = New-Case 'process'; New-SyntheticMedia -Path $case.LegacyCache
    $holder = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\cmd.exe') -ArgumentList @('/d','/c','ping -n 60 127.0.0.1 >nul & rem',('"{0}\setup.exe"' -f $case.LegacyCache)) -WindowStyle Hidden -PassThru
    try {
        Start-Sleep -Seconds 1
        Test-Step 'a running process uses the legacy media: InUse (real process scan)' { (Resolve-SemtIntuneSetupMedia -DataRoot $case.Data -SetupCacheFolder 'Win11-fr-FR' -Language 'fr-FR' -LegacyRoot $case.Legacy -OwnerProvider $owner).Decision -eq 'InUse' }
    }
    finally { Stop-Process -Id $holder.Id -Force -ErrorAction SilentlyContinue }

    $case = New-Case 'ownerunreadable'; New-SyntheticMedia -Path $case.LegacyCache; $script:OwnerState = 'Unreadable'
    Test-Step 'legacy ownership unreadable: Indeterminate' { (Resolve-Case $case).Decision -eq 'Indeterminate' }
    $script:OwnerState = 'Read'

    $case = New-Case 'untrusted'; New-SyntheticMedia -Path $case.LegacyCache; $script:FakeOwners['Win11-fr-FR'] = 'SmartM365-W11-Win11-fr-FR:Installed'; $script:TrustAll = $false
    $r = Resolve-Case $case
    Test-Step 'legacy media writable by a non-administrator: never used, explicit Untrusted decision, duplication required' { $r.Decision -eq 'Untrusted' -and $r.LegacyState -eq 'Untrusted' -and $r.DuplicationRequired -and $r.Path -eq '' -and -not $r.External }
    $script:TrustAll = $true; $script:FakeOwners = @{}

    $case = New-Case 'trust-unreadable'; New-SyntheticMedia -Path $case.LegacyCache
    $script:TrustState = 'Unreadable'
    Test-Step 'trust provider reports unreadable permissions: Indeterminate, not Absent or Untrusted' { $r = Resolve-Case $case; $r.Decision -eq 'Indeterminate' -and -not $r.DuplicationRequired }
    $script:TrustState = 'Null'
    Test-Step 'trust provider returns nothing: Indeterminate' { (Resolve-Case $case).Decision -eq 'Indeterminate' }
    $script:TrustState = 'Unknown'
    Test-Step 'trust provider returns an unknown state: Indeterminate' { (Resolve-Case $case).Decision -eq 'Indeterminate' }
    $script:TrustState = ''; $script:TrustThrows = $true
    Test-Step 'trust provider fails with an error: Indeterminate' { (Resolve-Case $case).Decision -eq 'Indeterminate' }
    $script:TrustThrows = $false
    $script:TrustState = 'Untrusted'; New-SyntheticMedia -Path $case.NewCache -Language 'de-DE'
    Test-Step 'untrusted legacy media next to an incompatible new cache: Untrusted reported first' { $r = Resolve-Case $case; $r.Decision -eq 'Untrusted' -and $r.NewCacheState -eq 'Incompatible' -and $r.DuplicationRequired }
    $script:TrustState = ''
    Test-Step 'no legacy media at all: Absent without duplication requirement' { $r = Resolve-Case (New-Case 'absent2'); $r.Decision -eq 'Absent' -and -not $r.DuplicationRequired -and $r.LegacyState -eq 'Absent' }

    $case = New-Case 'reparse'; New-SyntheticMedia -Path $case.LegacyCache; $script:FakeOwners['Win11-fr-FR'] = 'SmartM365-W11-Win11-fr-FR:Installed'
    $junctionTarget = Join-Path $tempRoot 'junction-target'; New-Item -ItemType Directory -Path $junctionTarget -Force | Out-Null
    $null = & cmd.exe /d /c ('mklink /J "{0}" "{1}"' -f (Join-Path $case.LegacyCache 'sources\link'),$junctionTarget) 2>&1
    $r = Resolve-Case $case
    Test-Step 'legacy media containing a reparse point: never used, even when owned by Intune (Untrusted)' { $r.Decision -eq 'Untrusted' -and $r.LegacyState -eq 'Untrusted' -and $r.DuplicationRequired }
    $null = & cmd.exe /d /c ('rmdir "{0}"' -f (Join-Path $case.LegacyCache 'sources\link')) 2>&1
    $script:FakeOwners = @{}

    $case = New-Case 'traversal'; New-Item -ItemType Directory -Path $case.Legacy -Force | Out-Null
    $evil = Join-Path $tempRoot 'case-traversal\Evil\Win11-fr-FR'; New-SyntheticMedia -Path $evil
    $sibling = $case.Legacy + 'Evil\SetupMedia\Win11-fr-FR'; New-SyntheticMedia -Path $sibling
    foreach ($recorded in @((Join-Path $case.Legacy '..\Evil\Win11-fr-FR'), $sibling)) {
        Set-Content -LiteralPath (Join-Path $case.Legacy 'LastRun.json') -Value ((@{ SetupCachePath = $recorded } | ConvertTo-Json)) -Encoding UTF8
        $candidates = @(Get-SemtIntuneLegacySetupMediaCandidates -SetupCacheFolder 'Win11-fr-FR' -LegacyRoot $case.Legacy)
        $r = Resolve-Case $case
        Write-TestResult -Condition ($candidates.Count -eq 1 -and $candidates[0] -eq $case.LegacyCache -and $r.Decision -eq 'Absent') -Message ("LastRun.json path outside <legacy>\SetupMedia is ignored: {0}" -f $(if ($recorded -match '\.\.') { 'parent traversal' } else { 'sibling prefix' })) -Detail (($candidates -join ';') + ' ' + $r.Decision)
    }
    Set-Content -LiteralPath (Join-Path $case.Legacy 'LastRun.json') -Value ((@{ SetupCachePath = (Join-Path $case.Legacy 'SetupMedia\Win11-fr-FR') } | ConvertTo-Json)) -Encoding UTF8
    Test-Step 'LastRun.json path inside <legacy>\SetupMedia is accepted' { @(Get-SemtIntuneLegacySetupMediaCandidates -SetupCacheFolder 'Win11-fr-FR' -LegacyRoot $case.Legacy) -contains $case.LegacyCache }
    #endregion

    #region 3. Endpoint selection (Windows Setup checks stubbed on synthetic media)
    $tokens = $null; $errors = $null
    $endpointAst = [System.Management.Automation.Language.Parser]::ParseFile($endpointPath, [ref]$tokens, [ref]$errors)
    foreach ($region in @('SEMT shared target guards v1','SEMT setup media resolution v1','SEMT legacy setup media reuse')) { . ([scriptblock]::Create((Get-RegionText -Path $endpointPath -Region $region))) }
    foreach ($name in @('Resolve-SetupUpgradeExecutable','Resolve-SetupCachePath','Convert-ToSafePathSegment')) {
        $definition = $endpointAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    function Test-SemtMediaWriteRestricted {
        param([string]$Path)
        if ($script:TrustState) { return [pscustomobject]@{ Trusted = ($script:TrustState -eq 'Trusted'); State = $script:TrustState; Detail = 'synthetic' } }
        [pscustomobject]@{ Trusted = $script:TrustAll; Detail = 'synthetic' }
    }
    $script:Log = New-Object System.Collections.Generic.List[string]
    $script:Copies = New-Object System.Collections.Generic.List[string]
    $script:Clears = New-Object System.Collections.Generic.List[string]
    function Write-SmartLog { param([string]$Message, [string]$Level = 'INFO') $script:Log.Add(('{0} {1}' -f $Level,$Message)) }
    function New-SmartDirectory { param([string]$Path) New-Item -ItemType Directory -Path $Path -Force | Out-Null }
    function Resolve-SetupLanguageRequirement { param([string]$RequestedLanguage) return 'fr-FR' }
    function Get-EffectiveSetupSourceCandidates { return @($script:SourceCandidates) }
    function Resolve-PreferredSetupSourcePath { param($SourcePaths, $ExpectedLanguage) return @($SourcePaths)[0] }
    function Get-SetupMediaFingerprint { param($MediaRoot, $ExpectedLanguage) return (Get-FileSha (Join-Path $MediaRoot 'SmartM365-SetupMediaManifest.sha256.csv')) }
    function Test-SetupFingerprintMatch { param($Left, $Right) return ($Left -eq $Right) }
    function Test-SetupMedia {
        param([string]$MediaPath, [string]$ExpectedLanguage, [switch]$ValidateManifest)
        $check = Test-SemtIntuneSetupMediaCandidate -Path $MediaPath -ExpectedLanguage $ExpectedLanguage
        if (-not $check.Valid) { if ($check.Category -eq 'Incompatible') { throw ('Setup media language mismatch. {0}' -f $check.Detail) } else { throw $check.Detail } }
        return (Join-Path $MediaPath 'setup.exe')
    }
    function Test-SetupCacheReady { param([string]$CachePath, [string]$ExpectedLanguage) return (Test-SetupMedia -MediaPath $CachePath -ExpectedLanguage $ExpectedLanguage) }
    function Copy-SetupMediaToLocalCache { param($SourcePath, $CachePath, $ExpectedLanguage) $script:Copies.Add($CachePath); return (Join-Path $CachePath 'setup.exe') }
    function Clear-SetupCachePath { param($CachePath, $Reason) $script:Clears.Add($CachePath) }
    function Get-SemtLegacySetupMediaOwners { param([string]$Leaf, [string]$RegistryKeyPath) [pscustomobject]@{ State = 'Read'; Owners = @($script:FakeOwners[$Leaf] | Where-Object { $_ }); Detail = '' } }

    function Set-EndpointCase {
        param($Case, [string[]]$Sources = @(), [switch]$Duplication, [string]$MediaPath = '')
        $script:AllowSetupUpgrade = $true; $script:DirectSetupUpgrade = $false; $script:SetupExecutionMode = 'LocalCache'
        $script:SetupLanguage = 'fr-FR'; $script:SetupMediaId = 'Win11'; $script:SetupCacheRoot = (Join-Path $Case.Data 'SetupMedia')
        $script:DataRoot = $Case.Data; $script:StateDir = (Join-Path $Case.Data 'State'); $script:ReuseLegacySetupMedia = $true
        $script:AllowSetupMediaDuplication = [bool]$Duplication; $script:SkipSetupMediaPreCopy = $false; $script:SetupMediaPath = $MediaPath
        $script:SourceCandidates = $Sources; $script:SemtLegacyWindows11Root = $Case.Legacy; $script:SetupCacheAction = ''
        $script:Copies.Clear(); $script:Clears.Clear()
    }

    $case = New-Case 'ep-new'; New-SyntheticMedia -Path $case.NewCache; Set-EndpointCase $case -MediaPath $case.NewCache
    $exe = Resolve-SetupUpgradeExecutable
    Test-Step 'endpoint: valid new cache is used and nothing is copied' { $exe -eq (Join-Path $case.NewCache 'setup.exe') -and $script:Copies.Count -eq 0 -and $script:Clears.Count -eq 0 }

    $case = New-Case 'ep-owned'; New-SyntheticMedia -Path $case.LegacyCache; $script:FakeOwners['Win11-fr-FR'] = 'SmartM365-W11-Win11-fr-FR:Installed'
    $before = Get-TreeState $case.LegacyCache
    Set-EndpointCase $case -MediaPath $case.LegacyCache
    $exe = Resolve-SetupUpgradeExecutable
    Test-Step 'endpoint: Intune-owned legacy media path selected by the runner is used in place' { $exe -eq (Join-Path $case.LegacyCache 'setup.exe') -and $script:SetupCacheAction -eq 'LegacyReusedInPlace' }
    Test-Step 'endpoint: no copy and no new cache created for Intune-owned legacy media' { $script:Copies.Count -eq 0 -and -not (Test-Path -LiteralPath $case.NewCache) }
    Test-Step 'endpoint: Intune-owned legacy media not moved, cleaned, re-permissioned or marked' { (Get-TreeState $case.LegacyCache) -ceq $before -and -not (Test-Path -LiteralPath (Join-Path $case.LegacyCache 'SmartM365-SetupMedia.json')) }
    $script:FakeOwners = @{}

    $case = New-Case 'ep-adopt'; New-SyntheticMedia -Path $case.LegacyCache; Set-EndpointCase $case -MediaPath $case.LegacyCache
    $exe = Resolve-SetupUpgradeExecutable
    Test-Step 'endpoint: adoptable legacy media is moved to the new cache and used, without copy' { $exe -eq (Join-Path $case.NewCache 'setup.exe') -and $script:SetupCacheAction -eq 'LegacyMovedAndValidated' -and -not (Test-Path -LiteralPath $case.LegacyCache) -and $script:Copies.Count -eq 0 }
    Test-Step 'endpoint: adoption evidence recorded with outcome Moved' { $file = @(Get-ChildItem -LiteralPath (Join-Path $case.Data 'State\MediaAdoption') -Recurse -Filter 'adoption.json')[0]; (Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json).Outcome -eq 'Moved' }

    $case = New-Case 'ep-absent'; $source = Join-Path $tempRoot 'source-A'; if (-not (Test-Path $source)) { New-SyntheticMedia -Path $source }; Set-EndpointCase $case -Sources @($source)
    [void](Resolve-SetupUpgradeExecutable)
    Test-Step 'endpoint: no reusable media and a source: one copy into the new cache' { $script:Copies.Count -eq 1 -and $script:Copies[0] -eq $case.NewCache }

    $case = New-Case 'ep-noSource'; Set-EndpointCase $case
    Test-Step 'endpoint: no reusable media and no source: explicit error, no copy' { try { [void](Resolve-SetupUpgradeExecutable); $false } catch { $_.Exception.Message -like 'Local setup cache is not ready and no SetupSourcePath*' -and $script:Copies.Count -eq 0 } }

    $case = New-Case 'ep-dup'; New-SyntheticMedia -Path $case.LegacyCache; New-Item -ItemType Directory -Path $case.NewCache -Force | Out-Null; Set-EndpointCase $case -Sources @($source)
    $before = Get-TreeState $case.LegacyCache
    Test-Step 'endpoint: valid legacy media not adoptable: duplication blocked before any change' { try { [void](Resolve-SetupUpgradeExecutable); $false } catch { $_.Exception.Message -like 'Setup media duplication blocked.*' -and $script:Copies.Count -eq 0 -and $script:Clears.Count -eq 0 -and (Get-TreeState $case.LegacyCache) -ceq $before } }
    Set-EndpointCase $case -Sources @($source) -Duplication
    [void](Resolve-SetupUpgradeExecutable)
    Test-Step 'endpoint: explicit duplication authorization allows the copy and logs it' { $script:Copies.Count -eq 1 -and @($script:Log | Where-Object { $_ -match 'Second setup media copy allowed' }).Count -ge 1 }

    # Untrusted and unreadable legacy media (PsExec path)
    $case = New-Case 'ep-untrusted'; New-SyntheticMedia -Path $case.LegacyCache; $script:TrustState = 'Untrusted'; Set-EndpointCase $case -Sources @($source)
    $before = Get-TreeState $case.LegacyCache
    Test-Step 'endpoint: untrusted legacy media, no authorization: explicit untrusted error, no copy, no cache clear' { try { [void](Resolve-SetupUpgradeExecutable); $false } catch { $_.Exception.Message -like 'Setup media legacy untrusted.*' -and $script:Copies.Count -eq 0 -and $script:Clears.Count -eq 0 } }
    Test-Step 'endpoint: untrusted legacy media not used, not moved and its permissions not changed' { (Get-TreeState $case.LegacyCache) -ceq $before -and -not (Test-Path -LiteralPath $case.NewCache) -and $script:SemtLegacyMediaDecision.Decision -eq 'Untrusted' }
    Set-EndpointCase $case -Sources @($source) -Duplication; $script:Log.Clear()
    [void](Resolve-SetupUpgradeExecutable)
    Test-Step 'endpoint: untrusted legacy media with explicit authorization: one copy into the new cache, logged' { $script:Copies.Count -eq 1 -and $script:Copies[0] -eq $case.NewCache -and @($script:Log | Where-Object { $_ -match 'Second setup media copy allowed.*Decision=Untrusted' }).Count -ge 1 }
    Test-Step 'endpoint: authorized copy leaves the untrusted legacy media unchanged' { (Get-TreeState $case.LegacyCache) -ceq $before }
    Set-EndpointCase $case -Sources @($source); $script:SkipSetupMediaPreCopy = $true
    Test-Step 'endpoint: untrusted legacy media in existing-media-only mode keeps its explicit status' { try { [void](Resolve-SetupUpgradeExecutable); $false } catch { $_.Exception.Message -like 'Setup media legacy untrusted.*' -and $script:Copies.Count -eq 0 } }
    $script:TrustState = 'Unreadable'; Set-EndpointCase $case -Sources @($source) -Duplication
    Test-Step 'endpoint: unreadable legacy permissions: indeterminate error, no copy even with authorization' { try { [void](Resolve-SetupUpgradeExecutable); $false } catch { $_.Exception.Message -like 'Setup media state is indeterminate.*' -and $script:Copies.Count -eq 0 -and $script:Clears.Count -eq 0 -and (Get-TreeState $case.LegacyCache) -ceq $before } }
    $script:TrustState = ''

    $case = New-Case 'ep-inuse'; New-SyntheticMedia -Path $case.LegacyCache
    $lockDir = Join-Path $case.Legacy 'Locks\SetupCache-Win11-fr-FR.lock'; New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value ('{ "PID": ' + $PID + ' }') -Encoding ASCII
    Set-EndpointCase $case -Sources @($source)
    $before = Get-TreeState $case.LegacyCache
    Test-Step 'endpoint: legacy media in use: explicit SETUP_MEDIA_IN_USE error, nothing copied or moved' { try { [void](Resolve-SetupUpgradeExecutable); $false } catch { $_.Exception.Message -like 'Setup media is in use.*' -and $script:Copies.Count -eq 0 -and (Get-TreeState $case.LegacyCache) -ceq $before } }

    $case = New-Case 'ep-reject'; New-SyntheticMedia -Path $case.NewCache; Set-EndpointCase $case -MediaPath (Join-Path $tempRoot 'elsewhere\Win11-fr-FR')
    Test-Step 'endpoint: an unknown -SetupMediaPath is rejected' { try { [void](Resolve-SetupUpgradeExecutable); $false } catch { $_.Exception.Message -like 'Setup media path not recognized.*' } }
    Test-Step 'endpoint: status mapping exists for the new media errors' { $text = [IO.File]::ReadAllText($endpointPath); $text.Contains("'SETUP_MEDIA_IN_USE'") -and $text.Contains("'SETUP_MEDIA_STATE_INDETERMINATE'") -and $text.Contains("'SETUP_MEDIA_PATH_REJECTED'") -and $text.Contains("'SETUP_MEDIA_LEGACY_UNTRUSTED'") }
    Test-Step 'campaign: SETUP_MEDIA_LEGACY_UNTRUSTED and SETUP_MEDIA_STATE_INDETERMINATE are operator actions' {
        Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force -ErrorAction Stop
        $campaignModule = Import-Module (Join-Path $campaignRoot 'SmartEndpointMigration.Campaign.Windows11Upgrade.psd1') -MinimumVersion 1.0.0 -Force -PassThru -ErrorAction Stop
        $categories = & $campaignModule { @('SETUP_MEDIA_LEGACY_UNTRUSTED','SETUP_MEDIA_STATE_INDETERMINATE','SETUP_MEDIA_IN_USE') | ForEach-Object { Get-Windows11FailureCategory -Result ([pscustomobject]@{ Status = $_; Detail = '' }) } }
        Remove-Module -Name $campaignModule.Name -Force
        $categories[0] -eq 'OperatorAction' -and $categories[1] -eq 'OperatorAction' -and $categories[2] -eq 'ExecutionTransient'
    }
    #endregion

    #region 4. Installer, runner and detection (real scripts, shimmed boundaries)
    $shimPath = Join-Path $tempRoot 'shim.ps1'
    @'
# Test shim: replaces only operational boundaries. Records every call in $env:SEMT_SHIM_RECORD.
function Write-ShimRecord { param([string]$Kind, $Data) Add-Content -LiteralPath $env:SEMT_SHIM_RECORD -Value ((@{ Kind = $Kind; Data = $Data } | ConvertTo-Json -Compress -Depth 5)) -Encoding UTF8 }
function Set-Registry64String { param([string]$SubKey, [string]$Name, [string]$Value) Write-ShimRecord 'Registry' @{ SubKey = $SubKey; Name = $Name; Value = $Value } }
function Remove-Registry64SubKeyTree { param([string]$SubKey) Write-ShimRecord 'RegistryRemove' @{ SubKey = $SubKey } }
function Protect-ToolkitPathAcl { param([string]$Path, [switch]$Directory) Write-ShimRecord 'Acl' @{ Path = $Path } }
function Get-OsFamily { return 'Windows10' }
function Get-CimInstance { param([string]$ClassName, $ErrorAction) if ($ClassName -eq 'Win32_OperatingSystem') { return [pscustomobject]@{ BuildNumber = '19045'; Caption = 'Microsoft Windows 10 Enterprise' } }; return @() }
function New-ScheduledTaskAction { param($Execute, $Argument) Write-ShimRecord 'TaskAction' @{ Execute = $Execute; Argument = $Argument }; return 'action' }
function New-ScheduledTaskTrigger { param([switch]$Once, $At, $RepetitionInterval, $RepetitionDuration) return 'trigger' }
function New-ScheduledTaskPrincipal { param($UserId, $LogonType, $RunLevel) return 'principal' }
function New-ScheduledTaskSettingsSet { param([switch]$StartWhenAvailable, [switch]$AllowStartIfOnBatteries, [switch]$DontStopIfGoingOnBatteries, $MultipleInstances) return 'settings' }
function Register-ScheduledTask { param($TaskName, $Action, $Trigger, $Principal, $Settings, [switch]$Force) Write-ShimRecord 'TaskRegister' @{ TaskName = $TaskName } }
function Start-ScheduledTask { param($TaskName) Write-ShimRecord 'TaskStart' @{ TaskName = $TaskName } }
function Unregister-ScheduledTask { param($TaskName, $Confirm, $ErrorAction) Write-ShimRecord 'TaskUnregister' @{ TaskName = $TaskName } }
function Start-Process {
    param($FilePath, $ArgumentList, [switch]$Wait, [switch]$PassThru, [switch]$NoNewWindow, $RedirectStandardOutput, $RedirectStandardError)
    Write-ShimRecord 'EndpointStart' @{ Arguments = @($ArgumentList) }
    return [pscustomobject]@{ ExitCode = [int]$env:SEMT_SHIM_ENDPOINT_EXIT }
}
# Program Files boundary: temporary folders belong to the test account, so the real check reports
# them as Untrusted and the real ACL change needs elevation. SEMT_SHIM_PROGRAM_TRUST selects a
# synthetic state (default Trusted), or 'Real' keeps the real check and the real ACL change.
if ($env:SEMT_SHIM_PROGRAM_TRUST -ne 'Real') {
    function Test-SemtTrustedProgramPath {
        param([string]$Path, [switch]$Recurse)
        # The synthetic state applies to the Program Files tree of the case only.
        $state = if ($env:SEMT_SHIM_PROGRAM_TRUST -and ([string]$Path -like '*\ProgramFiles\*' -or [string]$Path -like '*\ProgramFiles')) { $env:SEMT_SHIM_PROGRAM_TRUST } else { 'Trusted' }
        [pscustomobject]@{ Trusted = ($state -eq 'Trusted'); State = $state; Detail = 'shim' }
    }
    function Set-SemtProgramFolderAcl { param([string]$Path) Write-ShimRecord 'ProgramAcl' @{ Path = $Path } }
    function Test-SemtCanonicalFolderAcl { param([string]$Path) return $true }
    function Test-SemtTrustedSharedParent { param([string]$Path) [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = 'shim' } }
}
# Data folder boundary: SEMT_SHIM_DATA_TRUST=Untrusted makes the ProgramData chain untrusted.
if ($env:SEMT_SHIM_DATA_TRUST -eq 'Untrusted') {
    function Initialize-SemtTrustedDataFolder { param([string]$Path, [string]$BasePath) throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} (Untrusted) synthetic' -f $Path) }
    function Test-SemtTrustedDataFolder { param([string]$Path, [string]$BasePath) [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = 'synthetic' } }
}
if (Get-Command -Name Resolve-SemtIntuneSetupMedia -ErrorAction SilentlyContinue) {
    $script:SemtIntuneLegacyWindows11Root = $env:SEMT_SHIM_LEGACY_ROOT
    # Temporary test folders are writable by the test account, so the real check reports them as
    # Untrusted. SEMT_SHIM_TRUST_STATE selects a synthetic state, or 'Real' keeps the real check.
    if ($env:SEMT_SHIM_TRUST_STATE -ne 'Real') {
        function Test-SemtMediaWriteRestricted {
            param([string]$Path)
            $state = if ($env:SEMT_SHIM_TRUST_STATE) { $env:SEMT_SHIM_TRUST_STATE } else { 'Trusted' }
            [pscustomobject]@{ Trusted = ($state -eq 'Trusted'); State = $state; Detail = 'shim' }
        }
    }
    function Get-SemtLegacySetupMediaOwners {
        param([string]$Leaf, [string]$RegistryKeyPath)
        $owners = @()
        if ($env:SEMT_SHIM_LEGACY_OWNER) { $owners = @($env:SEMT_SHIM_LEGACY_OWNER) }
        return [pscustomobject]@{ State = 'Read'; Owners = $owners; Detail = 'shim' }
    }
}
'@ | Set-Content -LiteralPath $shimPath -Encoding ASCII

    function New-ShimmedScript {
        param([string]$Source, [string]$Destination, [string[]]$Anchors)
        $text = [IO.File]::ReadAllText($Source)
        foreach ($anchor in $Anchors) {
            $index = $text.IndexOf($anchor)
            if ($index -lt 0) { throw "Shim anchor not found in ${Source}: $anchor" }
            if ($anchor.StartsWith('. ')) { $index = $text.IndexOf("`n", $index) + 1; $text = $text.Insert($index, (". '{0}'`r`n" -f $shimPath)) }
            else { $text = $text.Insert($index, (". '{0}'`r`n" -f $shimPath)) }
        }
        [IO.File]::WriteAllText($Destination, $text, [Text.UTF8Encoding]::new($false))
    }

    function New-IntunePackage {
        param([string]$Name, [switch]$CacheOnly, [string]$Build = 'A', [switch]$AllowDuplication)
        $pkg = Join-Path $tempRoot ('pkg-' + $Name)
        New-Item -ItemType Directory -Path $pkg -Force | Out-Null
        New-ShimmedScript -Source (Join-Path $intuneSource 'SmartM365-EndpointMigration-Install.ps1') -Destination (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -Anchors @('# SEMT data folder boundary', '. $integrityHelperPackagePath')
        Copy-Item -LiteralPath $helperPath -Destination $pkg
        Copy-Item -LiteralPath (Join-Path $intuneSource 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1') -Destination $pkg
        Copy-Item -LiteralPath $endpointPath -Destination $pkg
        $mode = if ($CacheOnly) { 'WithCacheOnly' } else { 'WithMedia' }
        [pscustomobject]@{ PackageId = ('SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR' + $(if ($CacheOnly) { '-WithCacheOnly' } else { '' })); PackageVersion = '1.0.0'; MediaId = 'Win11'; Language = 'fr-FR'; SetupCacheFolder = 'Win11-fr-FR'; PackageMode = $mode; RequiresExistingSetupCache = [bool]$CacheOnly; AllowSetupMediaDuplication = [bool]$AllowDuplication } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $pkg 'PackageManifest.json') -Encoding ASCII
        if (-not $CacheOnly) { New-SyntheticMedia -Path (Join-Path $pkg 'SetupMedia\Win11-fr-FR') -Build $Build }
        return $pkg
    }

    function Invoke-Shimmed {
        param([string]$Script, [string]$DataRoot, [string]$LegacyRoot, [string]$Owner = '', [int]$EndpointExit = 0, [string]$TrustState = '', [string]$ProgramTrust = '', [string]$DataTrust = '', [string[]]$ExtraArguments = @())
        $record = Join-Path $tempRoot ('record-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.jsonl')
        $env:SEMT_SHIM_RECORD = $record; $env:SEMT_SHIM_LEGACY_ROOT = $LegacyRoot; $env:SEMT_SHIM_LEGACY_OWNER = $Owner; $env:SEMT_SHIM_ENDPOINT_EXIT = [string]$EndpointExit; $env:SEMT_SHIM_TRUST_STATE = $TrustState; $env:SEMT_SHIM_PROGRAM_TRUST = $ProgramTrust; $env:SEMT_SHIM_DATA_TRUST = $DataTrust
        $arguments = @('-DataRoot', $DataRoot)
        # The case folder plays the role of the system drive: <case>\PD is ProgramData.
        $caseRoot = $DataRoot.Substring(0, $DataRoot.IndexOf('\PD\WCH\'))
        New-Item -ItemType Directory -Path (Join-Path $caseRoot 'PD') -Force | Out-Null
        if ((Split-Path -Path $Script -Leaf) -eq 'SmartM365-EndpointMigration-Install.ps1') {
            # The installer writes its scripts to a temporary Program Files tree of the case.
            New-Item -ItemType Directory -Path (Join-Path $caseRoot 'ProgramFiles') -Force | Out-Null
            $arguments += @('-ProgramRoot', (Join-Path $caseRoot 'ProgramFiles\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade'), '-ProgramFilesPath', (Join-Path $caseRoot 'ProgramFiles'))
        }
        $arguments += @('-ProgramDataPath', (Join-Path $caseRoot 'PD'))
        $arguments += $ExtraArguments
        try {
            $ErrorActionPreference = 'Continue'
            $output = @(& $powershell -NoProfile -ExecutionPolicy Bypass -File $Script @arguments 2>&1 | ForEach-Object { [string]$_ })
            $exit = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = 'Stop'; foreach ($name in 'SEMT_SHIM_RECORD','SEMT_SHIM_LEGACY_ROOT','SEMT_SHIM_LEGACY_OWNER','SEMT_SHIM_ENDPOINT_EXIT','SEMT_SHIM_TRUST_STATE','SEMT_SHIM_PROGRAM_TRUST','SEMT_SHIM_DATA_TRUST') { Remove-Item -Path ('Env:' + $name) -ErrorAction SilentlyContinue } }
        $calls = @(); if (Test-Path -LiteralPath $record) { $calls = @(Get-Content -LiteralPath $record | ForEach-Object { $_ | ConvertFrom-Json }) }
        $registry = @{}; foreach ($call in @($calls | Where-Object Kind -eq 'Registry')) { $registry[[string]$call.Data.Name] = [string]$call.Data.Value }
        return [pscustomobject]@{ Exit = $exit; Output = ($output -join "`n"); Calls = $calls; Registry = $registry; Log = $(if (Test-Path -LiteralPath (Join-Path $DataRoot 'Logs\Intune')) { (Get-ChildItem -LiteralPath (Join-Path $DataRoot 'Logs\Intune') -Filter 'SEMT-Intune-*.log' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n" } else { '' }) }
    }

    # Installer
    $case = New-Case 'inst-new'; $pkg = New-IntunePackage 'inst-new'; New-SyntheticMedia -Path $case.NewCache -Build 'A'
    $before = Get-TreeState $case.NewCache
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer WithMedia + valid new cache: success, no recopy (cache unchanged)' { $run.Exit -eq 0 -and (Get-TreeState $case.NewCache) -ceq $before -and $run.Log -notmatch 'Copying packaged setup media' }
    Test-Step 'installer records the selected media in the package state' { $run.Registry['SetupMediaPath'] -eq $case.NewCache -and $run.Registry['SetupMediaSource'] -eq 'NewCache' -and $run.Registry['InstallState'] -eq 'Installed' }

    $case = New-Case 'inst-owned'; $pkg = New-IntunePackage 'inst-owned'; New-SyntheticMedia -Path $case.LegacyCache -Build 'A'
    $before = Get-TreeState $case.LegacyCache
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -Owner 'SmartM365-W11-Win11-fr-FR:Installed'
    Test-Step 'installer WithMedia + legacy Intune-owned media: success without copy' { $run.Exit -eq 0 -and -not (Test-Path -LiteralPath $case.NewCache) -and $run.Log -notmatch 'Copying packaged setup media' }
    Test-Step 'installer: legacy Intune-owned media unchanged, no ACL hardening on it' { (Get-TreeState $case.LegacyCache) -ceq $before -and @($run.Calls | Where-Object { $_.Kind -eq 'Acl' -and [string]$_.Data.Path -like ($case.Legacy + '*') }).Count -eq 0 }
    Test-Step 'installer records the external legacy path and source' { $run.Registry['SetupMediaPath'] -eq $case.LegacyCache -and $run.Registry['SetupMediaSource'] -eq 'LegacyExternal' -and $run.Registry['SetupMediaExternal'] -eq 'True' }

    $case = New-Case 'inst-copy'; $pkg = New-IntunePackage 'inst-copy'
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer WithMedia + no media: packaged media copied (real robocopy) and valid' { $run.Exit -eq 0 -and (Get-FileSha (Join-Path $case.NewCache 'sources\install.wim')) -eq (Get-FileSha (Join-Path $pkg 'SetupMedia\Win11-fr-FR\sources\install.wim')) -and $run.Registry['SetupMediaSource'] -eq 'PackageCopy' }

    $case = New-Case 'inst-incompatible'; $pkg = New-IntunePackage 'inst-incompatible' -Build 'B'; New-SyntheticMedia -Path $case.NewCache -Build 'A'
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer WithMedia + cache of another media build: replaced by the packaged media' { $run.Exit -eq 0 -and (Get-FileSha (Join-Path $case.NewCache 'setup.exe')) -eq (Get-FileSha (Join-Path $pkg 'SetupMedia\Win11-fr-FR\setup.exe')) }

    $case = New-Case 'inst-inuse'; $pkg = New-IntunePackage 'inst-inuse'; New-SyntheticMedia -Path $case.LegacyCache
    $lockDir = Join-Path $case.Legacy 'Locks\SetupCache-Win11-fr-FR.lock'; New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value ('{ "PID": ' + $PID + ' }') -Encoding ASCII
    $before = Get-TreeState $case.LegacyCache
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer: legacy operation active on the media: exit 1618, nothing copied or changed' { $run.Exit -eq 1618 -and -not (Test-Path -LiteralPath $case.NewCache) -and (Get-TreeState $case.LegacyCache) -ceq $before -and -not $run.Registry.ContainsKey('InstallState') }

    $case = New-Case 'inst-dup'; $pkg = New-IntunePackage 'inst-dup'; New-SyntheticMedia -Path $case.LegacyCache; New-Item -ItemType Directory -Path $case.NewCache -Force | Out-Null
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer: valid legacy media that cannot be used: duplication blocked (exit 1, no copy)' { $run.Exit -eq 1 -and $run.Log -match 'SEMT_SETUP_MEDIA_DUPLICATION_BLOCKED' -and @(Get-ChildItem -LiteralPath $case.NewCache -Force).Count -eq 0 }

    # Untrusted and unreadable legacy media (Intune path)
    $case = New-Case 'inst-untrusted'; $pkg = New-IntunePackage 'inst-untrusted'; New-SyntheticMedia -Path $case.LegacyCache
    $before = Get-TreeState $case.LegacyCache
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -TrustState 'Real'
    Test-Step 'installer WithMedia, real ACL check, legacy media writable by the test account: explicit untrusted failure (exit 1)' { $run.Exit -eq 1 -and $run.Log -match 'SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED' -and $run.Log -match 'Decision=Untrusted' }
    Test-Step 'installer: untrusted legacy media: no copy by default, nothing installed, no ACL change' { -not (Test-Path -LiteralPath $case.NewCache) -and -not $run.Registry.ContainsKey('InstallState') -and @($run.Calls | Where-Object { $_.Kind -eq 'Acl' -and [string]$_.Data.Path -like ($case.Legacy + '*') }).Count -eq 0 -and (Get-TreeState $case.LegacyCache) -ceq $before }
    $pkg = New-IntunePackage 'inst-untrusted-dup' -AllowDuplication
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -TrustState 'Real'
    Test-Step 'installer WithMedia + untrusted legacy media + AllowSetupMediaDuplication: packaged copy made and logged' { $run.Exit -eq 0 -and (Test-Path -LiteralPath (Join-Path $case.NewCache 'setup.exe')) -and $run.Registry['SetupMediaSource'] -eq 'PackageCopy' -and $run.Log -match 'SEMT_SETUP_MEDIA_DUPLICATION_ALLOWED' }
    Test-Step 'installer: authorized copy leaves the untrusted legacy media unchanged' { (Get-TreeState $case.LegacyCache) -ceq $before -and @($run.Calls | Where-Object { $_.Kind -eq 'Acl' -and [string]$_.Data.Path -like ($case.Legacy + '*') }).Count -eq 0 }

    $case = New-Case 'inst-unreadable'; $pkg = New-IntunePackage 'inst-unreadable' -AllowDuplication; New-SyntheticMedia -Path $case.LegacyCache
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -TrustState 'Unreadable'
    Test-Step 'installer: unreadable legacy permissions: indeterminate failure, no copy even with AllowSetupMediaDuplication' { $run.Exit -eq 1 -and $run.Log -match 'SEMT_SETUP_MEDIA_STATE_INDETERMINATE' -and -not (Test-Path -LiteralPath $case.NewCache) -and -not $run.Registry.ContainsKey('InstallState') }

    $case = New-Case 'inst-real-unreadable'; $pkg = New-IntunePackage 'inst-real-unreadable' -AllowDuplication; New-SyntheticMedia -Path $case.LegacyCache
    $lockedFolder = $case.LegacyCache
    Set-TestSddl -Path $lockedFolder -Sddl 'D:P(A;;0x1100A1;;;WD)(A;;0x40000;;;OW)'
    try {
        $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -TrustState 'Real'
        Test-Step 'installer, real ACL check, legacy item without readable permissions: indeterminate failure, no copy' { $run.Exit -eq 1 -and $run.Log -match 'SEMT_SETUP_MEDIA_STATE_INDETERMINATE' -and -not (Test-Path -LiteralPath $case.NewCache) }
    }
    finally { Set-TestSddl -Path $lockedFolder -Sddl 'D:PAI(A;OICI;FA;;;WD)' }

    $case = New-Case 'inst-co-untrusted'; $pkg = New-IntunePackage 'inst-co-untrusted' -CacheOnly; New-SyntheticMedia -Path $case.LegacyCache
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -TrustState 'Untrusted'
    Test-Step 'installer WithCacheOnly + untrusted legacy media: explicit untrusted failure (not reported as absent)' { $run.Exit -eq 1 -and $run.Log -match 'SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED' -and $run.Log -notmatch 'SEMT_SETUP_MEDIA_NOT_REUSABLE' -and -not $run.Registry.ContainsKey('InstallState') }

    $case = New-Case 'inst-co-absent'; $pkg = New-IntunePackage 'inst-co-absent' -CacheOnly
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer WithCacheOnly + no reusable media: explicit failure, nothing installed' { $run.Exit -eq 1 -and $run.Log -match 'SEMT_SETUP_MEDIA_NOT_REUSABLE' -and -not $run.Registry.ContainsKey('InstallState') }

    $case = New-Case 'inst-co-adopt'; $pkg = New-IntunePackage 'inst-co-adopt' -CacheOnly; New-SyntheticMedia -Path $case.LegacyCache
    $before = Get-TreeState $case.LegacyCache
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer WithCacheOnly + adoptable legacy media: success, media left for the endpoint to adopt' { $run.Exit -eq 0 -and $run.Registry['SetupMediaSource'] -eq 'LegacyAdoptable' -and $run.Registry['SetupMediaPath'] -eq $case.LegacyCache -and (Get-TreeState $case.LegacyCache) -ceq $before }

    # Runner
    function New-RunnerData {
        param($Case, [string]$Name)
        $programIntune = Join-Path $Case.Program 'Intune'
        New-Item -ItemType Directory -Path $programIntune -Force | Out-Null
        New-Item -ItemType Directory -Path $Case.Data -Force | Out-Null
        New-ShimmedScript -Source (Join-Path $intuneSource 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1') -Destination (Join-Path $programIntune 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1') -Anchors @('trap {', '. $integrityHelperPath')
        Copy-Item -LiteralPath $helperPath -Destination $programIntune
        Copy-Item -LiteralPath $endpointPath -Destination $Case.Program
        [pscustomobject]@{ PackageId = 'SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR'; PackageVersion = '1.0.0'; MediaId = 'Win11'; Language = 'fr-FR'; SetupCacheFolder = 'Win11-fr-FR'; PackageMode = 'WithMedia' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $programIntune 'PackageManifest.json') -Encoding ASCII
        return (Join-Path $programIntune 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1')
    }
    function Get-EndpointArguments { param($Run) $start = @($Run.Calls | Where-Object Kind -eq 'EndpointStart'); if ($start.Count -eq 0) { return @() }; return @($start[0].Data.Arguments) }
    function Get-ArgumentValue { param([string[]]$Arguments, [string]$Name) $i = [Array]::IndexOf($Arguments, $Name); if ($i -lt 0 -or $i + 1 -ge $Arguments.Count) { return $null }; return $Arguments[$i + 1] }

    $case = New-Case 'run-owned'; New-SyntheticMedia -Path $case.LegacyCache; $runner = New-RunnerData $case
    $before = Get-TreeState $case.LegacyCache
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy -Owner 'SmartM365-W11-Win11-fr-FR:Installed'
    $arguments = Get-EndpointArguments $run
    Test-Step 'runner: legacy Intune-owned media path is passed to the endpoint (-SetupMediaPath)' { $run.Exit -eq 0 -and (Get-ArgumentValue $arguments '-SetupMediaPath') -eq $case.LegacyCache }
    Test-Step 'runner: no RepairRequired for valid external media' { $run.Registry['InstallState'] -ne 'RepairRequired' -and @($run.Calls | Where-Object { $_.Kind -eq 'Registry' -and $_.Data.Value -eq 'RepairRequired' }).Count -eq 0 }
    Test-Step 'runner: legacy media unchanged' { (Get-TreeState $case.LegacyCache) -ceq $before }

    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy -Owner 'SmartM365-W11-Win11-fr-FR:Installed' -EndpointExit 1
    Test-Step 'runner: endpoint failure with valid external media does not trigger a repair loop' { @($run.Calls | Where-Object { $_.Kind -eq 'Registry' -and $_.Data.Value -eq 'RepairRequired' }).Count -eq 0 -and @($run.Calls | Where-Object Kind -eq 'TaskUnregister').Count -eq 0 }

    $case = New-Case 'run-adopt'; New-SyntheticMedia -Path $case.LegacyCache; $runner = New-RunnerData $case
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'runner: adoptable legacy media: endpoint started with the legacy path, nothing moved by the runner' { (Get-ArgumentValue (Get-EndpointArguments $run) '-SetupMediaPath') -eq $case.LegacyCache -and (Test-Path -LiteralPath (Join-Path $case.LegacyCache 'setup.exe')) -and -not (Test-Path -LiteralPath $case.NewCache) }

    $case = New-Case 'run-absent'; $runner = New-RunnerData $case
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'runner: no usable media: RepairRequired with the media decision, endpoint not started' { $run.Registry['InstallState'] -eq 'RepairRequired' -and $run.Registry['RepairReason'] -like 'SEMT_MEDIA_ABSENT*' -and @($run.Calls | Where-Object Kind -eq 'EndpointStart').Count -eq 0 }

    $case = New-Case 'run-untrusted'; New-SyntheticMedia -Path $case.LegacyCache; $runner = New-RunnerData $case
    $before = Get-TreeState $case.LegacyCache
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy -Owner 'SmartM365-W11-Win11-fr-FR:Installed' -TrustState 'Real'
    Test-Step 'runner, real ACL check: untrusted legacy media (even Intune-owned) is not used: SEMT_MEDIA_UNTRUSTED, endpoint not started' { $run.Registry['RepairReason'] -like 'SEMT_MEDIA_UNTRUSTED*' -and @($run.Calls | Where-Object Kind -eq 'EndpointStart').Count -eq 0 }
    Test-Step 'runner: untrusted legacy media unchanged and no copy' { (Get-TreeState $case.LegacyCache) -ceq $before -and -not (Test-Path -LiteralPath $case.NewCache) }

    $case = New-Case 'run-unreadable'; New-SyntheticMedia -Path $case.LegacyCache; $runner = New-RunnerData $case
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy -TrustState 'Unreadable'
    Test-Step 'runner: unreadable legacy permissions: cycle skipped, no repair, endpoint not started' { $run.Exit -eq 0 -and @($run.Calls | Where-Object Kind -eq 'EndpointStart').Count -eq 0 -and -not $run.Registry.ContainsKey('InstallState') -and $run.Log -match 'Decision=Indeterminate' }

    $case = New-Case 'run-inuse'; New-SyntheticMedia -Path $case.LegacyCache; $runner = New-RunnerData $case
    $lockDir = Join-Path $case.Legacy 'Locks\SetupCache-Win11-fr-FR.lock'; New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $lockDir 'lock.json') -Value ('{ "PID": ' + $PID + ' }') -Encoding ASCII
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'runner: media in use: cycle skipped (exit 0), no repair, endpoint not started' { $run.Exit -eq 0 -and @($run.Calls | Where-Object Kind -eq 'EndpointStart').Count -eq 0 -and -not $run.Registry.ContainsKey('InstallState') }

    $case = New-Case 'run-new'; New-SyntheticMedia -Path $case.NewCache; $runner = New-RunnerData $case
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'runner: valid new cache: endpoint started with the new cache path' { (Get-ArgumentValue (Get-EndpointArguments $run) '-SetupMediaPath') -eq $case.NewCache -and $run.Registry['SetupMediaSource'] -eq 'NewCache' }

    # Program Files location of the scripts run as SYSTEM (installer, runner)
    $case = New-Case 'pf-install'; $pkg = New-IntunePackage 'pf-install'; New-SyntheticMedia -Path $case.NewCache
    New-Item -ItemType Directory -Path (Join-Path $case.Data 'Intune') -Force | Out-Null
    $legacyRunnerCopy = Join-Path $case.Data 'Intune\SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1'
    Set-Content -LiteralPath $legacyRunnerCopy -Value '# script copy of an earlier qualification build' -Encoding ASCII
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    $installed = @(
        @{ Source = (Join-Path $pkg 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'); Target = (Join-Path $case.Program 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1') },
        @{ Source = (Join-Path $pkg 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1'); Target = (Join-Path $case.Program 'Intune\SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1') },
        @{ Source = (Join-Path $pkg 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'); Target = (Join-Path $case.Program 'Intune\SmartM365-EndpointMigration-SetupMediaIntegrity.ps1') },
        @{ Source = (Join-Path $pkg 'PackageManifest.json'); Target = (Join-Path $case.Program 'Intune\PackageManifest.json') })
    Test-Step 'installer: endpoint, runner, helper and manifest installed under Program Files with the package hashes' { $run.Exit -eq 0 -and @($installed | Where-Object { -not (Test-Path -LiteralPath $_.Target) -or (Get-FileSha $_.Target) -ne (Get-FileSha $_.Source) }).Count -eq 0 }
    $dataScripts = @(Get-ChildItem -LiteralPath $case.Data -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq '.ps1' } | ForEach-Object { $_.FullName })
    Write-TestResult -Condition (-not (Test-Path -LiteralPath $legacyRunnerCopy) -and $dataScripts.Count -eq 0) -Message 'installer: no script is written under ProgramData; the old runner copy there is removed' -Detail ($dataScripts -join '; ')
    Test-Step 'installer: the SYSTEM task starts the runner from Program Files' { $action = @($run.Calls | Where-Object Kind -eq 'TaskAction'); $action.Count -eq 1 -and [string]$action[0].Data.Argument -like ('*-File "{0}"*' -f (Join-Path $case.Program 'Intune\SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1')) -and [string]$action[0].Data.Argument -notlike ('*-File "{0}*' -f $case.Data) }
    Test-Step 'installer: each Program Files folder it creates gets the protected ACL' { $acl = @($run.Calls | Where-Object Kind -eq 'ProgramAcl' | ForEach-Object { [string]$_.Data.Path } | Where-Object { $_ -like '*\ProgramFiles\*' }); $acl -contains (Join-Path $case.Program 'Intune') -and $acl -contains $case.Program -and $acl.Count -eq 5 }

    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -ExtraArguments @('-Uninstall')
    Test-Step 'installer -Uninstall: task removed and installed program files and folders removed' { $run.Exit -eq 0 -and -not (Test-Path -LiteralPath $case.Program) -and @($run.Calls | Where-Object Kind -eq 'TaskUnregister').Count -eq 1 }

    $case = New-Case 'pf-untrusted'; $pkg = New-IntunePackage 'pf-untrusted'
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -ProgramTrust 'Untrusted'
    Test-Step 'installer: untrusted Program Files folder: exit 1, SEMT_PROGRAM_FOLDER_UNTRUSTED, no task, no media copy, not installed' { $run.Exit -eq 1 -and $run.Log -match 'SEMT_PROGRAM_FOLDER_UNTRUSTED' -and @($run.Calls | Where-Object { $_.Kind -in @('TaskRegister','TaskAction') }).Count -eq 0 -and -not (Test-Path -LiteralPath $case.NewCache) -and -not $run.Registry.ContainsKey('InstallState') }

    $case = New-Case 'pf-real'; $pkg = New-IntunePackage 'pf-real'
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -ProgramTrust 'Real'
    Test-Step 'installer, real checks (not elevated): owner Administrators cannot be set: refused, nothing installed, written or registered' { $run.Exit -eq 1 -and $run.Output -match 'SEMT_DATA_FOLDER_UNTRUSTED' -and -not (Test-Path -LiteralPath (Join-Path $case.Program 'Intune')) -and -not (Test-Path -LiteralPath $case.Data) -and @($run.Calls | Where-Object Kind -eq 'TaskRegister').Count -eq 0 }

    $case = New-Case 'pd-untrusted'; $pkg = New-IntunePackage 'pd-untrusted'
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy -DataTrust 'Untrusted'
    Test-Step 'installer: ProgramData chain that cannot be made trusted: exit 1, SEMT_DATA_FOLDER_UNTRUSTED, nothing written, installed or registered' { $run.Exit -eq 1 -and $run.Output -match 'SEMT_DATA_FOLDER_UNTRUSTED' -and -not (Test-Path -LiteralPath $case.Data) -and -not (Test-Path -LiteralPath (Join-Path $case.Program 'Intune')) -and @($run.Calls | Where-Object { $_.Kind -in @('TaskRegister','TaskAction','Registry') }).Count -eq 0 }
    $case = New-Case 'pd-chain'; $pkg = New-IntunePackage 'pd-chain'
    $run = Invoke-Shimmed -Script (Join-Path $pkg 'SmartM365-EndpointMigration-Install.ps1') -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'installer: only the SEMT components (product folder, then data root) get owner Administrators and the protected DACL, before anything is written; shared parents WCH and SmartM365 never' { $acl = @($run.Calls | Where-Object Kind -eq 'ProgramAcl' | ForEach-Object { [string]$_.Data.Path } | Where-Object { $_ -like '*\PD\*' }); $shared = Split-Path -Path (Split-Path -Path $case.Data -Parent) -Parent; $run.Exit -eq 0 -and $acl.Count -eq 2 -and $acl[0] -eq (Split-Path -Path $case.Data -Parent) -and $acl[1] -eq $case.Data -and $acl -notcontains $shared -and $acl -notcontains (Split-Path -Path $shared -Parent) }

    $case = New-Case 'pf-runner'; New-SyntheticMedia -Path $case.NewCache; $runner = New-RunnerData $case
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy
    Test-Step 'runner: trusted Program Files folder: endpoint started from Program Files' { $start = @($run.Calls | Where-Object Kind -eq 'EndpointStart'); $start.Count -eq 1 -and @($start[0].Data.Arguments) -contains (Join-Path $case.Program 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1') }
    foreach ($programTrust in @('Untrusted','Unreadable','Real')) {
        $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy -ProgramTrust $programTrust
        Write-TestResult -Condition ($run.Exit -eq 1 -and $run.Output -match 'SEMT_PROGRAM_FOLDER_UNTRUSTED' -and @($run.Calls | Where-Object Kind -eq 'EndpointStart').Count -eq 0 -and @($run.Calls | Where-Object Kind -eq 'Registry').Count -eq 0) -Message ("runner: program folder {0}: exit 1, nothing loaded, started or written" -f $(if ($programTrust -eq 'Real') { 'owned by a standard account (real check)' } else { $programTrust })) -Detail ("exit={0}" -f $run.Exit)
    }

    $case = New-Case 'pd-runner'; New-SyntheticMedia -Path $case.NewCache; $runner = New-RunnerData $case
    $logsBefore = Test-Path -LiteralPath (Join-Path $case.Data 'Logs')
    $run = Invoke-Shimmed -Script $runner -DataRoot $case.Data -LegacyRoot $case.Legacy -DataTrust 'Untrusted'
    Test-Step 'runner: ProgramData chain not trusted: endpoint not started, nothing written under the data root' { $run.Exit -eq 1 -and $run.Output -match 'SEMT_DATA_FOLDER_UNTRUSTED' -and @($run.Calls | Where-Object Kind -eq 'EndpointStart').Count -eq 0 -and (Test-Path -LiteralPath (Join-Path $case.Data 'Logs')) -eq $logsBefore }
    Test-Step 'runner: ProgramData chain not trusted: package marked RepairRequired (registry) and task removed, so the installer hardens it again' { $run.Registry['InstallState'] -eq 'RepairRequired' -and $run.Registry['RepairReason'] -like 'SEMT_DATA_FOLDER_UNTRUSTED*' -and @($run.Calls | Where-Object Kind -eq 'TaskUnregister').Count -eq 1 }

    # Detection (generated from the template)
    $detectTemplate = [IO.File]::ReadAllText((Join-Path $intuneSource 'SmartM365-EndpointMigration-Detect-Template.ps1')).Replace('__PACKAGE_ID__','SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR').Replace('__PACKAGE_VERSION__','1.0.0')
    foreach ($detectCase in @(@{ Name = 'external'; State = 'Installed'; Source = 'LegacyExternal'; Expected = 0 }, @{ Name = 'new'; State = 'Installed'; Source = 'NewCache'; Expected = 0 }, @{ Name = 'repair'; State = 'RepairRequired'; Source = 'NewCache'; Expected = 1 })) {
        $detectShim = @"
function Get-CimInstance { param([string]`$ClassName, `$ErrorAction) [pscustomobject]@{ BuildNumber = '19045'; Caption = 'Microsoft Windows 10 Enterprise' } }
function Get-Registry64PackageState { param([string]`$SubKey) [pscustomobject]@{ PackageId = 'SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR'; PackageVersion = '1.0.0'; InstallState = '$($detectCase.State)'; SetupMediaSource = '$($detectCase.Source)'; SetupMediaPath = 'C:\Legacy\Win11-fr-FR' } }
"@
        $index = $detectTemplate.IndexOf("try {`r`n    `$os = Get-CimInstance")
        if ($index -lt 0) { $index = $detectTemplate.IndexOf("try {`n    `$os = Get-CimInstance") }
        $detectFile = Join-Path $tempRoot ('detect-{0}.ps1' -f $detectCase.Name)
        [IO.File]::WriteAllText($detectFile, $detectTemplate.Insert($index, $detectShim + "`r`n"), [Text.UTF8Encoding]::new($false))
        $ErrorActionPreference = 'Continue'
        $detectOutput = @(& $powershell -NoProfile -ExecutionPolicy Bypass -File $detectFile 2>&1 | ForEach-Object { [string]$_ })
        $detectExit = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        Write-TestResult -Condition ($detectExit -eq $detectCase.Expected) -Message ("detection: InstallState={0}, SetupMediaSource={1} -> exit {2}" -f $detectCase.State,$detectCase.Source,$detectCase.Expected) -Detail (($detectOutput -join ' ') + " exit=$detectExit")
    }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    try { [IO.Directory]::Delete('\\?\' + $tempRoot, $true) } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7h52vLWK89dT4
# aNVZtJJje/E9Cu+YLd4m8W1Iml3OVaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP69+IEhUXRtIgV+3lSwN15+V9H+I/ukeUI+KJjQtFLgMA0GCSqG
# SIb3DQEBAQUABIIBgAB++Z44aAZyzfYKPVarJuqNlTnNnW4fN+5y+IGhSThGbRs9
# v0EMmgysdxcob2u9JjrFbyp7ST9tuxw2ANeSCtFDrR2rvd+NrI+6BLgpm1PKOkdN
# GzJT4t0Ib/lzmVx3Kmo+aGzSnrYS5RmoxYj+pdHdC2dyMHdv1Pg9Ph0ifwJZM9WX
# OrxHiqtc++KigcrEvOOw6ZgwzEIONcAiWPOryWJ+NTA+hcT0XtMLQ6IX5LVcp7Nw
# 4gq+cSzMoLNhejuC9nXsfo39rO8Z4XZdWqvgCnMa/3WP5xqp8kBKjklSkzor2+gc
# UjUwIKZoh0VVILgKUqQsFh7l1SGWCKZ/YQDRDmot0sSSP5ZX4QcgRMg9KxEYzRGE
# 1QBfLXSFHxcKQYRhbNCiG0B5RdXuLQ76j0rQfRvtv7AjN+2evJJ31myY/F9L4gOV
# 4vSeeSh92WyJbY2/u6xYPR34FM83YN9TZKHvwTMt86N906iWIiXfG5EsraUUqx+6
# n4E8WUWavNVpdRoftKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE3
# MDBaMC8GCSqGSIb3DQEJBDEiBCDkTE7B8wWn6z6mNacDjTgo09Dd3uvF74xkxjMn
# bUEFqTANBgkqhkiG9w0BAQEFAASCAgBYZhRk8uW36zU/K2/pwGPQy3EKXb7lPYbQ
# jLH+IhB0u/sjSvGFassnCxiBQl6Eg6M2iGNKAVHsGM3aXt0KBPQMvVKnehzY4Gum
# GXWnkuCFsYZ/80EDn35GQgFjp4EsCrU2nsPmCGHNuLJiHgvx6jcTpdNe92dXuh2m
# +ELJNnpX1NbMAvvQ13X/GG4gtxTDtxG4WngbPakexH5ke3WWDDfJb/b5DTZM+aVp
# e4z2HPAVXj/5wH1tsH9AwjNcuWfSXeGgUvevT0YWuj7Ar+szzAf70uUA2+WgCt4O
# gP2DkduGQPEKbEXEgLjT5/btHJcMck2yCsB4P/h6Dy4oXeBWIm9Z8DGA+L51SDhw
# 1J5UdujQ4g6BNimnSw19xOg+T9yw/cCFjce/5H626NnEsVhH8edhPkFQS2c3ve0N
# eZ5GzmFp/5no/7OkaUoFuRlTw4AYLeJ370GjJJaUGdAiuWEbGj3XqxvwMe+i/xfZ
# XBZEFz6jn9XPhWB2+AmuqWPA3LAhiHS16r8jinUo8dcODzVn9MBLK4Ed2bFYQjMG
# AHS9ZLuSef0vHrTUyK1D6PWm1hJQ/bf/uzce7Q95fE61ax56P1/H8VfEAQ6nxvd0
# 3ohtxdm+EyLdg+LVizuWjOxL4JrUtRzzK2+Ge9vmBUzE+SICn2Q8ZyMdt7CNagUs
# RzKNU4Zjhg==
# SIG # End signature block
