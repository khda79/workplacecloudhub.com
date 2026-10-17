// Fast strict CSV reading for prepared evidence. Only an unambiguous RFC 4180
// subset is accepted; any other construct returns null so the caller falls back
// to its reference parser (TextFieldParser or Import-Csv) with unchanged results.
// Values and identities are never included in diagnostics.
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Management.Automation;

namespace Workplace.Prepared {
    public sealed class TenantScan {
        public long Rows { get; set; }
        public bool HasTenantKey { get; set; }
    }

    sealed class StrictCsvUnsupported : Exception { }

    sealed class StrictCsvTokenizer : IDisposable {
        public const int Quoted = 1;
        public const int QuotedLineBreak = 2;
        public const int EdgeWhiteSpace = 4;
        const int WhiteSpaceOnly = 8;
        // Raw characters per record, kept below the 10,000,000-character TextFieldParser buffer.
        const int MaxRecordChars = 8000000;
        readonly StreamReader reader;
        readonly char delimiter;
        readonly char[] buffer = new char[1 << 16];
        readonly StringBuilder field = new StringBuilder();
        int length, position;

        public StrictCsvTokenizer(string path, char delimiter) {
            // Same decoding as the reference parsers: UTF-8 with byte-order-mark detection.
            reader = new StreamReader(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16), new UTF8Encoding(false), true, 1 << 16);
            this.delimiter = delimiter;
        }

        public void Dispose() { reader.Dispose(); }

        bool Fill() {
            if (position < length) return true;
            length = reader.Read(buffer, 0, buffer.Length);
            position = 0;
            if (length > 0) return true;
            length = 0;
            return false;
        }

        int Peek() { return Fill() ? buffer[position] : -1; }

        int Next() { return Fill() ? buffer[position++] : -1; }

        // Reads one record; values of columns not selected by keep are null. Returns false
        // at end of file and throws StrictCsvUnsupported outside the strict subset.
        public bool ReadRecord(List<string> values, List<int> flags, Func<int, bool> keep) {
            values.Clear(); flags.Clear();
            if (Peek() < 0) return false;
            int recordChars = 0;
            while (true) {
                bool wanted = keep == null || keep(values.Count);
                field.Clear();
                int flag = 0, c, fieldChars = 0;
                if (Peek() == '"') {
                    position++; flag |= Quoted; recordChars++;
                    while (true) {
                        if (!Fill()) throw new StrictCsvUnsupported();
                        int start = position;
                        while (position < length) {
                            char ch = buffer[position];
                            if (ch == '"') break;
                            if (ch == '\r' || ch == '\n') flag |= QuotedLineBreak;
                            position++;
                        }
                        if (wanted) field.Append(buffer, start, position - start);
                        recordChars += position - start;
                        if (recordChars > MaxRecordChars) throw new StrictCsvUnsupported();
                        if (position == length) continue;
                        position++; recordChars++;
                        if (Peek() != '"') break;
                        position++; recordChars++;
                        if (wanted) field.Append('"');
                    }
                    c = Next();
                } else {
                    bool allWhite = true, firstWhite = false, lastWhite = false;
                    while (Fill()) {
                        int start = position;
                        while (position < length) {
                            char ch = buffer[position];
                            if (ch == delimiter || ch == '\r' || ch == '\n') break;
                            if (ch == '"') throw new StrictCsvUnsupported();
                            bool white = Char.IsWhiteSpace(ch);
                            if (fieldChars == 0) firstWhite = white;
                            lastWhite = white; allWhite &= white;
                            fieldChars++; position++;
                        }
                        if (wanted) field.Append(buffer, start, position - start);
                        recordChars += position - start;
                        if (recordChars > MaxRecordChars) throw new StrictCsvUnsupported();
                        if (position < length) break;
                    }
                    if (fieldChars > 0 && (firstWhite || lastWhite)) flag |= EdgeWhiteSpace;
                    if (allWhite) flag |= WhiteSpaceOnly;
                    c = Next();
                }
                values.Add(wanted ? field.ToString() : null); flags.Add(flag);
                if (++recordChars > MaxRecordChars) throw new StrictCsvUnsupported();
                if (c == delimiter) continue;
                if (c < 0 || c == '\n') break;
                if (c == '\r' && Peek() == '\n') { position++; break; }
                // Lone carriage return or text after a closing quote.
                throw new StrictCsvUnsupported();
            }
            // Reference parsers skip blank lines; leave them to the reference path.
            if (values.Count == 1 && (flags[0] & WhiteSpaceOnly) != 0) throw new StrictCsvUnsupported();
            return true;
        }
    }

    public static class StrictCsv {
        const int ValueAnomaly = StrictCsvTokenizer.QuotedLineBreak | StrictCsvTokenizer.EdgeWhiteSpace;

        static bool HeaderSupported(List<string> header, List<int> flags) {
            if (header.Count == 0) return false;
            var ordinal = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var culture = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            var invariant = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
            for (int i = 0; i < header.Count; i++) {
                if ((flags[i] & ValueAnomaly) != 0 || String.IsNullOrWhiteSpace(header[i])) return false;
                if (!ordinal.Add(header[i]) || !culture.Add(header[i]) || !invariant.Add(header[i])) return false;
            }
            return true;
        }

        // Same checks as the TextFieldParser tenant validator. Null means "use the reference parser",
        // including for every failure, so error messages and record numbers stay identical.
        public static TenantScan CheckTenant(string path, char delimiter, string tenant, bool allowMissingTenantKey) {
            try {
                using (var tokenizer = new StrictCsvTokenizer(path, delimiter)) {
                    var values = new List<string>(); var flags = new List<int>();
                    if (!tokenizer.ReadRecord(values, flags, null) || !HeaderSupported(values, flags)) return null;
                    int columns = values.Count;
                    int index = values.IndexOf("TenantKey");
                    if (index < 0 && !allowMissingTenantKey) return null;
                    Func<int, bool> keep = column => column == index;
                    long rows = 0;
                    while (tokenizer.ReadRecord(values, flags, keep)) {
                        rows++;
                        if (values.Count != columns) return null;
                        if (index >= 0 && ((flags[index] & ValueAnomaly) != 0 ||
                            String.IsNullOrWhiteSpace(values[index]) || !String.Equals(values[index], tenant, StringComparison.Ordinal))) return null;
                    }
                    return new TenantScan { Rows = rows, HasTenantKey = index >= 0 };
                }
            } catch (Exception) {
                return null;
            }
        }

        // Selected columns as Import-Csv | Select-Object would return them. Null means
        // "use Import-Csv", including for missing columns and every failure.
        public static PSObject[] ReadColumns(string path, char delimiter, string[] columns) {
            try {
                using (var tokenizer = new StrictCsvTokenizer(path, delimiter)) {
                    var values = new List<string>(); var flags = new List<int>();
                    if (!tokenizer.ReadRecord(values, flags, null) || !HeaderSupported(values, flags) || values[0].StartsWith("#", StringComparison.Ordinal)) return null;
                    int count = values.Count;
                    var indexes = new int[columns.Length]; var names = new string[columns.Length];
                    var selected = new bool[count];
                    for (int c = 0; c < columns.Length; c++) {
                        string column = columns[c];
                        indexes[c] = values.FindIndex(name => String.Equals(name, column, StringComparison.OrdinalIgnoreCase));
                        if (indexes[c] < 0) return null;
                        names[c] = values[indexes[c]];
                        selected[indexes[c]] = true;
                    }
                    Func<int, bool> keep = column => column < count && selected[column];
                    var rows = new List<PSObject>();
                    while (tokenizer.ReadRecord(values, flags, keep)) {
                        if (values.Count != count) return null;
                        var row = new PSObject();
                        for (int c = 0; c < indexes.Length; c++) {
                            if ((flags[indexes[c]] & ValueAnomaly) != 0) return null;
                            row.Properties.Add(new PSNoteProperty(names[c], values[indexes[c]]));
                        }
                        rows.Add(row);
                    }
                    return rows.ToArray();
                }
            } catch (Exception) {
                return null;
            }
        }
    }
}
