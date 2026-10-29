<#
.SYNOPSIS
Compiled CSV validation and serialization for large datasets. Dot-sourced Core helper.
.VERSION
1.0.0
.NOTES
PowerShell 7 only; the type is compiled on first use, for datasets of SmartM365CsvFastMinimumRows rows
or more. Each helper reproduces one PowerShell path exactly (Assert-SmartM365CsvDataCompleteness loops,
Select-Object -Property | Export-Csv with UTF8 or utf8BOM). Input outside that exact behaviour returns
$null or $false and the caller runs its original PowerShell path, so results never depend on the path.
#>
[CmdletBinding()]
param()
$script:SmartM365CsvFastMinimumRows = 5000
$script:SmartM365CsvFastState = $null

function Initialize-SmartM365CsvFast {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    if ($null -ne $script:SmartM365CsvFastState) { return $script:SmartM365CsvFastState }
    $script:SmartM365CsvFastState = $false
    if ($PSVersionTable.PSEdition -ne 'Core') { return $false }
    try {
        if (-not ('SmartM365.Core.CsvFast' -as [type])) {
            Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;

namespace SmartM365.Core
{
    public sealed class CsvRowScan
    {
        public bool Supported;
        public string MissingIdentityColumn;
        public bool DuplicateKey;
        public int MissingRows;
        public bool ThresholdReached;
        public string[] Examples = new string[0];
    }

    public static class CsvFast
    {
        // $row.PSObject.Properties[$name]: PowerShell returns $null for an empty name, C# throws.
        private static PSPropertyInfo GetProperty(PSObject item, string name)
        {
            return string.IsNullOrEmpty(name) ? null : item.Properties[name];
        }

        // Same text as a PowerShell [string] cast.
        private static string AsText(object value)
        {
            if (value == null) { return string.Empty; }
            string text = value as string;
            if (text != null) { return text; }
            object converted = LanguagePrimitives.ConvertTo(value, typeof(string), CultureInfo.InvariantCulture);
            return converted == null ? string.Empty : (string)converted;
        }

        // Unique-key and critical-field loops of Assert-SmartM365CsvDataCompleteness, in the same order.
        public static CsvRowScan Scan(object[] rows, string[] uniqueFields, string[] criticalFields, int failMinRows, double failPercent)
        {
            CsvRowScan result = new CsvRowScan();
            foreach (object row in rows) { if (row == null) { return result; } }
            result.Supported = true;
            if (uniqueFields.Length > 0)
            {
                HashSet<string> keys = new HashSet<string>(StringComparer.Ordinal);
                StringBuilder key = new StringBuilder();
                foreach (object row in rows)
                {
                    PSObject item = PSObject.AsPSObject(row);
                    key.Length = 0;
                    foreach (string field in uniqueFields)
                    {
                        PSPropertyInfo property = GetProperty(item, field);
                        if (property == null) { result.MissingIdentityColumn = field; return result; }
                        string value = AsText(property.Value).Trim().ToLowerInvariant();
                        // Length-prefixed parts: as injective as the JSON array key of the PowerShell path.
                        key.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value).Append('|');
                    }
                    if (!keys.Add(key.ToString())) { result.DuplicateKey = true; return result; }
                }
            }
            List<string> examples = new List<string>();
            int rowNumber = 1;
            foreach (object row in rows)
            {
                PSObject item = PSObject.AsPSObject(row);
                bool rowMissing = false;
                foreach (string field in criticalFields)
                {
                    PSPropertyInfo property = GetProperty(item, field);
                    object value = property == null ? null : property.Value;
                    if (value == null || string.IsNullOrWhiteSpace(AsText(value)))
                    {
                        rowMissing = true;
                        if (examples.Count < 10) { examples.Add(string.Format(CultureInfo.CurrentCulture, "row {0}, field {1}", rowNumber, field)); }
                    }
                }
                if (rowMissing)
                {
                    result.MissingRows++;
                    double percent = Math.Round(100.0 * result.MissingRows / rows.Length, 4);
                    bool rowThreshold = failMinRows > 0 && result.MissingRows >= failMinRows;
                    bool percentThreshold = failPercent > 0 && percent >= failPercent;
                    if (rowThreshold || percentThreshold) { result.ThresholdReached = true; break; }
                }
                rowNumber++;
            }
            result.Examples = examples.ToArray();
            return result;
        }

        private static void AppendQuoted(StringBuilder line, string value)
        {
            // Export-Csv (UseQuotes Always): a null value is an empty unquoted field.
            if (value == null) { return; }
            line.Append('"');
            if (value.IndexOf('"') >= 0) { line.Append(value.Replace("\"", "\"\"")); } else { line.Append(value); }
            line.Append('"');
        }

        // Select-Object -Property <columns> | Export-Csv -NoTypeInformation for PSObject rows holding note
        // properties. Returns false, with no file left, when a row is outside that exact behaviour.
        public static bool WriteQuoted(object[] rows, string[] columns, string path, char delimiter, bool byteOrderMark)
        {
            if (rows.Length == 0 || columns.Length == 0) { return false; }
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string column in columns)
            {
                if (string.IsNullOrEmpty(column) || column.IndexOfAny(new char[] { '*', '?', '[', ']' }) >= 0 || !seen.Add(column)) { return false; }
            }
            PSObject first = rows[0] as PSObject;
            if (first == null) { return false; }
            // Select-Object names each property after the matching property of the row; Export-Csv takes the
            // header from the first row.
            string[] header = new string[columns.Length];
            for (int index = 0; index < columns.Length; index++)
            {
                PSPropertyInfo property = first.Properties[columns[index]];
                header[index] = property == null ? columns[index] : property.Name;
            }
            bool written = false;
            try
            {
                using (StreamWriter writer = new StreamWriter(path, false, new UTF8Encoding(byteOrderMark), 1 << 20))
                {
                    StringBuilder line = new StringBuilder(1024);
                    for (int index = 0; index < header.Length; index++)
                    {
                        if (index > 0) { line.Append(delimiter); }
                        AppendQuoted(line, header[index]);
                    }
                    writer.WriteLine(line.ToString());
                    foreach (object row in rows)
                    {
                        PSObject item = row as PSObject;
                        if (item == null || !(item.BaseObject is PSCustomObject)) { return false; }
                        line.Length = 0;
                        for (int index = 0; index < columns.Length; index++)
                        {
                            if (index > 0) { line.Append(delimiter); }
                            PSPropertyInfo property = item.Properties[columns[index]];
                            if (property == null) { continue; }
                            if (!(property is PSNoteProperty)) { return false; }
                            object value = property.Value;
                            AppendQuoted(line, value == null ? null : value.ToString());
                        }
                        writer.WriteLine(line.ToString());
                    }
                }
                written = true;
                return true;
            }
            finally
            {
                if (!written && File.Exists(path)) { File.Delete(path); }
            }
        }
    }
}
'@
        }
        $script:SmartM365CsvFastState = $true
    }
    catch {
        WriteLog -Message ("Compiled CSV helpers unavailable; PowerShell CSV path used: {0}" -f $_.Exception.Message) -Level 'WARNING'
    }
    return $script:SmartM365CsvFastState
}

# Unique-key and critical-field scan; $null means the caller runs its PowerShell loops.
function Invoke-SmartM365CsvRowScan {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [AllowEmptyCollection()][string[]]$UniqueFields = @(),
        [AllowEmptyCollection()][string[]]$CriticalFields = @(),
        [int]$FailMinRows,
        [double]$FailPercent
    )
    if ($null -eq $Rows -or $Rows.Count -lt $script:SmartM365CsvFastMinimumRows -or -not (Initialize-SmartM365CsvFast)) { return $null }
    if ($null -eq $UniqueFields) { $UniqueFields = @() }
    if ($null -eq $CriticalFields) { $CriticalFields = @() }
    $scan = [SmartM365.Core.CsvFast]::Scan($Rows, $UniqueFields, $CriticalFields, $FailMinRows, $FailPercent)
    if (-not $scan.Supported) { return $null }
    return $scan
}

# Writes the rows as Export-Csv would; $false means nothing was written and the caller runs Export-Csv.
function Write-SmartM365CsvRowsFast {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$Path,
        [string]$Encoding = 'UTF8',
        [string]$Delimiter = ','
    )
    if ($null -eq $Rows -or $Rows.Count -lt $script:SmartM365CsvFastMinimumRows -or $Delimiter.Length -ne 1) { return $false }
    if ($Encoding -notin @('UTF8','utf8BOM') -or -not (Initialize-SmartM365CsvFast)) { return $false }
    return [SmartM365.Core.CsvFast]::WriteQuoted($Rows, $Columns, $Path, $Delimiter[0], ($Encoding -eq 'utf8BOM'))
}
