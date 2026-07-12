#=============================================================================================================================
#
# Script Name:     HardwareReadiness.ps1
# Description:     Verifies the hardware compliance. Return code 0 for success. 
#                  In case of failure, returns non zero error code along with error message.

# This script is not supported under any Microsoft standard support program or service and is distributed under the MIT license

# Copyright (C) 2021 Microsoft Corporation

# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation
# files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy,
# modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software
# is furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
# WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
# COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
# ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

#=============================================================================================================================

$exitCode = 0

[int]$MinOSDiskSizeGB = 64
[int]$MinMemoryGB = 4
[Uint32]$MinClockSpeedMHz = 1000
[Uint32]$MinLogicalCores = 2
[Uint16]$RequiredAddressWidth = 64

$PASS_STRING = "PASS"
$FAIL_STRING = "FAIL"
$FAILED_TO_RUN_STRING = "FAILED TO RUN"
$UNDETERMINED_CAPS_STRING = "UNDETERMINED"
$UNDETERMINED_STRING = "Undetermined"
$CAPABLE_STRING = "Capable"
$NOT_CAPABLE_STRING = "Not capable"
$CAPABLE_CAPS_STRING = "CAPABLE"
$NOT_CAPABLE_CAPS_STRING = "NOT CAPABLE"
$STORAGE_STRING = "Storage"
$OS_DISK_SIZE_STRING = "OSDiskSize"
$MEMORY_STRING = "Memory"
$SYSTEM_MEMORY_STRING = "System_Memory"
$GB_UNIT_STRING = "GB"
$TPM_STRING = "TPM"
$TPM_VERSION_STRING = "TPMVersion"
$PROCESSOR_STRING = "Processor"
$SECUREBOOT_STRING = "SecureBoot"
$I7_7820HQ_CPU_STRING = "i7-7820hq CPU"

# 0=name of check, 1=attribute checked, 2=value, 3=PASS/FAIL/UNDETERMINED
$logFormat = '{0}: {1}={2}. {3}; '

# 0=name of check, 1=attribute checked, 2=value, 3=unit of the value, 4=PASS/FAIL/UNDETERMINED
$logFormatWithUnit = '{0}: {1}={2}{3}. {4}; '

# 0=name of check.
$logFormatReturnReason = '{0}, '

# 0=exception.
$logFormatException = '{0}; '

# 0=name of check, 1= attribute checked and its value, 2=PASS/FAIL/UNDETERMINED
$logFormatWithBlob = '{0}: {1}. {2}; '

# return returnCode is -1 when an exception is thrown. 1 if the value does not meet requirements. 0 if successful. -2 default, script didn't run.
$outObject = @{ returnCode = -2; returnResult = $FAILED_TO_RUN_STRING; returnReason = ""; logging = "" }

# NOT CAPABLE(1) state takes precedence over UNDETERMINED(-1) state
function Private:UpdateReturnCode {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(-2, 1)]
        [int] $ReturnCode
    )

    Switch ($ReturnCode) {

        0 {
            if ($outObject.returnCode -eq -2) {
                $outObject.returnCode = $ReturnCode
            }
        }
        1 {
            $outObject.returnCode = $ReturnCode
        }
        -1 {
            if ($outObject.returnCode -ne 1) {
                $outObject.returnCode = $ReturnCode
            }
        }
    }
}

$Source = @"
using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;

    public class CpuFamilyResult
    {
        public bool IsValid { get; set; }
        public string Message { get; set; }
    }

    public class CpuFamily
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct SYSTEM_INFO
        {
            public ushort ProcessorArchitecture;
            ushort Reserved;
            public uint PageSize;
            public IntPtr MinimumApplicationAddress;
            public IntPtr MaximumApplicationAddress;
            public IntPtr ActiveProcessorMask;
            public uint NumberOfProcessors;
            public uint ProcessorType;
            public uint AllocationGranularity;
            public ushort ProcessorLevel;
            public ushort ProcessorRevision;
        }

        [DllImport("kernel32.dll")]
        internal static extern void GetNativeSystemInfo(ref SYSTEM_INFO lpSystemInfo);

        public enum ProcessorFeature : uint
        {
            ARM_SUPPORTED_INSTRUCTIONS = 34
        }

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsProcessorFeaturePresent(ProcessorFeature processorFeature);

        private const ushort PROCESSOR_ARCHITECTURE_X86 = 0;
        private const ushort PROCESSOR_ARCHITECTURE_ARM64 = 12;
        private const ushort PROCESSOR_ARCHITECTURE_X64 = 9;

        private const string INTEL_MANUFACTURER = "GenuineIntel";
        private const string AMD_MANUFACTURER = "AuthenticAMD";
        private const string QUALCOMM_MANUFACTURER = "Qualcomm Technologies Inc";

        public static CpuFamilyResult Validate(string manufacturer, ushort processorArchitecture)
        {
            CpuFamilyResult cpuFamilyResult = new CpuFamilyResult();

            if (string.IsNullOrWhiteSpace(manufacturer))
            {
                cpuFamilyResult.IsValid = false;
                cpuFamilyResult.Message = "Manufacturer is null or empty";
                return cpuFamilyResult;
            }

            string registryPath = "HKEY_LOCAL_MACHINE\\Hardware\\Description\\System\\CentralProcessor\\0";
            SYSTEM_INFO sysInfo = new SYSTEM_INFO();
            GetNativeSystemInfo(ref sysInfo);

            switch (processorArchitecture)
            {
                case PROCESSOR_ARCHITECTURE_ARM64:

                    if (manufacturer.Equals(QUALCOMM_MANUFACTURER, StringComparison.OrdinalIgnoreCase))
                    {
                        bool isArmv81Supported = IsProcessorFeaturePresent(ProcessorFeature.ARM_SUPPORTED_INSTRUCTIONS);

                        if (!isArmv81Supported)
                        {
                            string registryName = "CP 4030";
                            long registryValue = (long)Registry.GetValue(registryPath, registryName, -1);
                            long atomicResult = (registryValue >> 20) & 0xF;

                            if (atomicResult >= 2)
                            {
                                isArmv81Supported = true;
                            }
                        }

                        cpuFamilyResult.IsValid = isArmv81Supported;
                        cpuFamilyResult.Message = isArmv81Supported ? "" : "Processor does not implement ARM v8.1 atomic instruction";
                    }
                    else
                    {
                        cpuFamilyResult.IsValid = false;
                        cpuFamilyResult.Message = "The processor isn't currently supported for Windows 11";
                    }

                    break;

                case PROCESSOR_ARCHITECTURE_X64:
                case PROCESSOR_ARCHITECTURE_X86:

                    int cpuFamily = sysInfo.ProcessorLevel;
                    int cpuModel = (sysInfo.ProcessorRevision >> 8) & 0xFF;
                    int cpuStepping = sysInfo.ProcessorRevision & 0xFF;

                    if (manufacturer.Equals(INTEL_MANUFACTURER, StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            cpuFamilyResult.IsValid = true;
                            cpuFamilyResult.Message = "";

                            if (cpuFamily >= 6 && cpuModel <= 95 && !(cpuFamily == 6 && cpuModel == 85))
                            {
                                cpuFamilyResult.IsValid = false;
                                cpuFamilyResult.Message = "";
                            }
                            else if (cpuFamily == 6 && (cpuModel == 142 || cpuModel == 158) && cpuStepping == 9)
                            {
                                string registryName = "Platform Specific Field 1";
                                int registryValue = (int)Registry.GetValue(registryPath, registryName, -1);

                                if ((cpuModel == 142 && registryValue != 16) || (cpuModel == 158 && registryValue != 8))
                                {
                                    cpuFamilyResult.IsValid = false;
                                }
                                cpuFamilyResult.Message = "PlatformId " + registryValue;
                            }
                        }
                        catch (Exception ex)
                        {
                            cpuFamilyResult.IsValid = false;
                            cpuFamilyResult.Message = "Exception:" + ex.GetType().Name;
                        }
                    }
                    else if (manufacturer.Equals(AMD_MANUFACTURER, StringComparison.OrdinalIgnoreCase))
                    {
                        cpuFamilyResult.IsValid = true;
                        cpuFamilyResult.Message = "";

                        if (cpuFamily < 23 || (cpuFamily == 23 && (cpuModel == 1 || cpuModel == 17)))
                        {
                            cpuFamilyResult.IsValid = false;
                        }
                    }
                    else
                    {
                        cpuFamilyResult.IsValid = false;
                        cpuFamilyResult.Message = "Unsupported Manufacturer: " + manufacturer + ", Architecture: " + processorArchitecture + ", CPUFamily: " + sysInfo.ProcessorLevel + ", ProcessorRevision: " + sysInfo.ProcessorRevision;
                    }

                    break;

                default:
                    cpuFamilyResult.IsValid = false;
                    cpuFamilyResult.Message = "Unsupported CPU category. Manufacturer: " + manufacturer + ", Architecture: " + processorArchitecture + ", CPUFamily: " + sysInfo.ProcessorLevel + ", ProcessorRevision: " + sysInfo.ProcessorRevision;
                    break;
            }
            return cpuFamilyResult;
        }
    }
"@

# Storage
try {
    $osDrive = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -Property SystemDrive
    $osDriveSize = Get-WmiObject -Class Win32_LogicalDisk -filter "DeviceID='$($osDrive.SystemDrive)'" | Select-Object @{Name = "SizeGB"; Expression = { $_.Size / 1GB -as [int] } }  

    if ($null -eq $osDriveSize) {
        UpdateReturnCode -ReturnCode 1
        $outObject.returnReason += $logFormatReturnReason -f $STORAGE_STRING
        $outObject.logging += $logFormatWithBlob -f $STORAGE_STRING, "Storage is null", $FAIL_STRING
        $exitCode = 1
    }
    elseif ($osDriveSize.SizeGB -lt $MinOSDiskSizeGB) {
        UpdateReturnCode -ReturnCode 1
        $outObject.returnReason += $logFormatReturnReason -f $STORAGE_STRING
        $outObject.logging += $logFormatWithUnit -f $STORAGE_STRING, $OS_DISK_SIZE_STRING, ($osDriveSize.SizeGB), $GB_UNIT_STRING, $FAIL_STRING
        $exitCode = 1
    }
    else {
        $outObject.logging += $logFormatWithUnit -f $STORAGE_STRING, $OS_DISK_SIZE_STRING, ($osDriveSize.SizeGB), $GB_UNIT_STRING, $PASS_STRING
        UpdateReturnCode -ReturnCode 0
    }
}
catch {
    UpdateReturnCode -ReturnCode -1
    $outObject.logging += $logFormat -f $STORAGE_STRING, $OS_DISK_SIZE_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
    $outObject.logging += $logFormatException -f "$($_.Exception.GetType().Name) $($_.Exception.Message)"
    $exitCode = 1
}

# Memory (bytes)
try {
    $memory = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | Select-Object @{Name = "SizeGB"; Expression = { $_.Sum / 1GB -as [int] } }

    if ($null -eq $memory) {
        UpdateReturnCode -ReturnCode 1
        $outObject.returnReason += $logFormatReturnReason -f $MEMORY_STRING
        $outObject.logging += $logFormatWithBlob -f $MEMORY_STRING, "Memory is null", $FAIL_STRING
        $exitCode = 1
    }
    elseif ($memory.SizeGB -lt $MinMemoryGB) {
        UpdateReturnCode -ReturnCode 1
        $outObject.returnReason += $logFormatReturnReason -f $MEMORY_STRING
        $outObject.logging += $logFormatWithUnit -f $MEMORY_STRING, $SYSTEM_MEMORY_STRING, ($memory.SizeGB), $GB_UNIT_STRING, $FAIL_STRING
        $exitCode = 1
    }
    else {
        $outObject.logging += $logFormatWithUnit -f $MEMORY_STRING, $SYSTEM_MEMORY_STRING, ($memory.SizeGB), $GB_UNIT_STRING, $PASS_STRING
        UpdateReturnCode -ReturnCode 0
    }
}
catch {
    UpdateReturnCode -ReturnCode -1
    $outObject.logging += $logFormat -f $MEMORY_STRING, $SYSTEM_MEMORY_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
    $outObject.logging += $logFormatException -f "$($_.Exception.GetType().Name) $($_.Exception.Message)"
    $exitCode = 1
}

# TPM
try {
    $tpm = Get-Tpm

    if ($null -eq $tpm) {
        UpdateReturnCode -ReturnCode 1
        $outObject.returnReason += $logFormatReturnReason -f $TPM_STRING
        $outObject.logging += $logFormatWithBlob -f $TPM_STRING, "TPM is null", $FAIL_STRING
        $exitCode = 1
    }
    elseif ($tpm.TpmPresent) {
        $tpmVersion = Get-WmiObject -Class Win32_Tpm -Namespace root\CIMV2\Security\MicrosoftTpm | Select-Object -Property SpecVersion

        if ($null -eq $tpmVersion.SpecVersion) {
            UpdateReturnCode -ReturnCode 1
            $outObject.returnReason += $logFormatReturnReason -f $TPM_STRING
            $outObject.logging += $logFormat -f $TPM_STRING, $TPM_VERSION_STRING, "null", $FAIL_STRING
            $exitCode = 1
        }

        $majorVersion = $tpmVersion.SpecVersion.Split(",")[0] -as [int]
        if ($majorVersion -lt 2) {
            UpdateReturnCode -ReturnCode 1
            $outObject.returnReason += $logFormatReturnReason -f $TPM_STRING
            $outObject.logging += $logFormat -f $TPM_STRING, $TPM_VERSION_STRING, ($tpmVersion.SpecVersion), $FAIL_STRING
            $exitCode = 1
        }
        else {
            $outObject.logging += $logFormat -f $TPM_STRING, $TPM_VERSION_STRING, ($tpmVersion.SpecVersion), $PASS_STRING
            UpdateReturnCode -ReturnCode 0
        }
    }
    else {
        if ($tpm.GetType().Name -eq "String") {
            UpdateReturnCode -ReturnCode -1
            $outObject.logging += $logFormat -f $TPM_STRING, $TPM_VERSION_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
            $outObject.logging += $logFormatException -f $tpm
        }
        else {
            UpdateReturnCode -ReturnCode  1
            $outObject.returnReason += $logFormatReturnReason -f $TPM_STRING
            $outObject.logging += $logFormat -f $TPM_STRING, $TPM_VERSION_STRING, ($tpm.TpmPresent), $FAIL_STRING
        }
        $exitCode = 1
    }
}
catch {
    UpdateReturnCode -ReturnCode -1
    $outObject.logging += $logFormat -f $TPM_STRING, $TPM_VERSION_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
    $outObject.logging += $logFormatException -f "$($_.Exception.GetType().Name) $($_.Exception.Message)"
    $exitCode = 1
}

# CPU Details
$cpuDetails;
try {
    $cpuDetails = @(Get-WmiObject -Class Win32_Processor)[0]

    if ($null -eq $cpuDetails) {
        UpdateReturnCode -ReturnCode 1
        $exitCode = 1
        $outObject.returnReason += $logFormatReturnReason -f $PROCESSOR_STRING
        $outObject.logging += $logFormatWithBlob -f $PROCESSOR_STRING, "CpuDetails is null", $FAIL_STRING
    }
    else {
        $processorCheckFailed = $false

        # AddressWidth
        if ($null -eq $cpuDetails.AddressWidth -or $cpuDetails.AddressWidth -ne $RequiredAddressWidth) {
            UpdateReturnCode -ReturnCode 1
            $processorCheckFailed = $true
            $exitCode = 1
        }

        # ClockSpeed is in MHz
        if ($null -eq $cpuDetails.MaxClockSpeed -or $cpuDetails.MaxClockSpeed -le $MinClockSpeedMHz) {
            UpdateReturnCode -ReturnCode 1;
            $processorCheckFailed = $true
            $exitCode = 1
        }

        # Number of Logical Cores
        if ($null -eq $cpuDetails.NumberOfLogicalProcessors -or $cpuDetails.NumberOfLogicalProcessors -lt $MinLogicalCores) {
            UpdateReturnCode -ReturnCode 1
            $processorCheckFailed = $true
            $exitCode = 1
        }

        # CPU Family
        Add-Type -TypeDefinition $Source
        $cpuFamilyResult = [CpuFamily]::Validate([String]$cpuDetails.Manufacturer, [uint16]$cpuDetails.Architecture)

        $cpuDetailsLog = "{AddressWidth=$($cpuDetails.AddressWidth); MaxClockSpeed=$($cpuDetails.MaxClockSpeed); NumberOfLogicalCores=$($cpuDetails.NumberOfLogicalProcessors); Manufacturer=$($cpuDetails.Manufacturer); Caption=$($cpuDetails.Caption); $($cpuFamilyResult.Message)}"

        if (!$cpuFamilyResult.IsValid) {
            UpdateReturnCode -ReturnCode 1
            $processorCheckFailed = $true
            $exitCode = 1
        }

        if ($processorCheckFailed) {
            $outObject.returnReason += $logFormatReturnReason -f $PROCESSOR_STRING
            $outObject.logging += $logFormatWithBlob -f $PROCESSOR_STRING, ($cpuDetailsLog), $FAIL_STRING
        }
        else {
            $outObject.logging += $logFormatWithBlob -f $PROCESSOR_STRING, ($cpuDetailsLog), $PASS_STRING
            UpdateReturnCode -ReturnCode 0
        }
    }
}
catch {
    UpdateReturnCode -ReturnCode -1
    $outObject.logging += $logFormat -f $PROCESSOR_STRING, $PROCESSOR_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
    $outObject.logging += $logFormatException -f "$($_.Exception.GetType().Name) $($_.Exception.Message)"
    $exitCode = 1
}

# SecureBooot
try {
    $isSecureBootEnabled = Confirm-SecureBootUEFI
    $outObject.logging += $logFormatWithBlob -f $SECUREBOOT_STRING, $CAPABLE_STRING, $PASS_STRING
    UpdateReturnCode -ReturnCode 0
}
catch [System.PlatformNotSupportedException] {
    # PlatformNotSupportedException "Cmdlet not supported on this platform." - SecureBoot is not supported or is non-UEFI computer.
    UpdateReturnCode -ReturnCode 1
    $outObject.returnReason += $logFormatReturnReason -f $SECUREBOOT_STRING
    $outObject.logging += $logFormatWithBlob -f $SECUREBOOT_STRING, $NOT_CAPABLE_STRING, $FAIL_STRING
    $exitCode = 1
}
catch [System.UnauthorizedAccessException] {
    UpdateReturnCode -ReturnCode -1
    $outObject.logging += $logFormatWithBlob -f $SECUREBOOT_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
    $outObject.logging += $logFormatException -f "$($_.Exception.GetType().Name) $($_.Exception.Message)"
    $exitCode = 1
}
catch {
    UpdateReturnCode -ReturnCode -1
    $outObject.logging += $logFormatWithBlob -f $SECUREBOOT_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
    $outObject.logging += $logFormatException -f "$($_.Exception.GetType().Name) $($_.Exception.Message)"
    $exitCode = 1
}

# i7-7820hq CPU
try {
    $supportedDevices = @('surface studio 2', 'precision 5520')
    $systemInfo = @(Get-WmiObject -Class Win32_ComputerSystem)[0]

    if ($null -ne $cpuDetails) {
        if ($cpuDetails.Name -match 'i7-7820hq cpu @ 2.90ghz'){
            $modelOrSKUCheckLog = $systemInfo.Model.Trim()
            if ($supportedDevices -contains $modelOrSKUCheckLog){
                $outObject.logging += $logFormatWithBlob -f $I7_7820HQ_CPU_STRING, $modelOrSKUCheckLog, $PASS_STRING
                $outObject.returnCode = 0
                $exitCode = 0
            }
        }
    }
}
catch {
    if ($outObject.returnCode -ne 0){
        UpdateReturnCode -ReturnCode -1
        $outObject.logging += $logFormatWithBlob -f $I7_7820HQ_CPU_STRING, $UNDETERMINED_STRING, $UNDETERMINED_CAPS_STRING
        $outObject.logging += $logFormatException -f "$($_.Exception.GetType().Name) $($_.Exception.Message)"
        $exitCode = 1
    }
}

Switch ($outObject.returnCode) {

    0 { $outObject.returnResult = $CAPABLE_CAPS_STRING }
    1 { $outObject.returnResult = $NOT_CAPABLE_CAPS_STRING }
    -1 { $outObject.returnResult = $UNDETERMINED_CAPS_STRING }
    -2 { $outObject.returnResult = $FAILED_TO_RUN_STRING }
}

$outObject | ConvertTo-Json -Compress
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBiHhm0Lk+/A8qC
# V96ruzLc51r2jcv0Tmi4ylIArUm6AaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAOK1cnlY1uVXtWlrxL/Ws/VGIoBgKL0ztA9fj8dDfLhzANBgkqhkiG9w0B
# AQEFAASCAYC+I9TDdMv1BMRht4uDQ5d2vVAVjkF7X1m70TQAIR2gJPaWqLgzk9Mc
# BB18qUya70GJd8AGeflex2dS6M+n4n2VZ9FWyCdUAXmB3H2pcu7eY/HyQjAqzjoT
# puvNZkRm1qM+XfnF5fkT0D5HEMrLBY5LRk/bWmrkQZkLUNn2I8FrLPQwVbI2/GTt
# MN9bxDxtv8/7vMNu6uvPvbtkxOw7AfWNTrhnrMJyQwfSzSQbzV99/o/QbU+TNLIB
# wqiipMN52RsMcmdhKC/xWT/FdaK28VYMVVPWdvcafdEDHOIVuy0wri/z4B01+IMb
# YXAsOjpa54Oh+wemnYi6LidCwqgHrzyDUWQhCVyUu7FL3IQWLjYJXroEv2CaYXHU
# GEU5Ztv9xj0VkKyJnSURZmBr/fPikO3BL3Xld5TgVd2EszV6y99FQevfyd6Ymcps
# abPcOwVpXFv2P133nn4NTBJL52whott937A+wqR9fMbCbfDeIGtX5u/jNgQAlTtN
# NcR5+2QWJ3s=
# SIG # End signature block
