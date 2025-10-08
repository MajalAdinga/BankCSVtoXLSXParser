using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using BankCSVtoXLSXParser.Utils;

namespace BankCSVtoXLSXParser.Parsers
{
    /// <summary>
    /// Parser for Standard Bank statement exports.
    /// Outputs columns:
    /// Ext. Tran. ID | Ext. Ref. Nbr. | Tran. Date (yyyy-MM-dd) | Tran. Desc | Receipt | Disbursement
    /// </summary>
    public sealed class StandardBankParser : IBankParser
    {
        public string Name => "Standard Bank";
        public string ShortName => "StandardBank";

        // IMPORTANT: Must remain yyyy-MM-dd to satisfy the common schema contract (IBankParser).
        private const string DateOutputFormat = "yyyy-MM-dd";
        private static readonly Regex YyyyMmDd = new Regex(@"^20\d{6}$", RegexOptions.Compiled);

        public bool IsMatch(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                int headerHits = 0, yyyymmddHits = 0, signedPaddedHits = 0, quotedPatternHits = 0, lines = 0;
                var signedPadded = new Regex(@"^[\+\-]0+\d+(\.\d+)?$");

                foreach (var line in File.ReadLines(filePath))
                {
                    if (lines++ >= 50) break;
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    string probe = line.Trim().Trim('"');

                    if (probe.IndexOf("ACC-NO", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        probe.IndexOf("ACCOUNT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        probe.IndexOf("BRANCH", StringComparison.OrdinalIgnoreCase) >= 0)
                        headerHits++;

                    char delim = CSVHelper.DetectDelimiter(probe);
                    var tokens = (probe.Contains(delim) ? CSVHelper.SplitCSVLine(probe, delim) : probe.Split(' '))
                                .Select(x => x.Trim().Trim('"'));

                    foreach (var t in tokens)
                    {
                        if (YyyyMmDd.IsMatch(t)) yyyymmddHits++;
                        if (signedPadded.IsMatch(t)) signedPaddedHits++;
                    }

                    if (line.Length > 20 && line[0] == '"' && line.Count(c => c == '"') >= 10)
                        quotedPatternHits++;
                }

                if (headerHits >= 1 && (yyyymmddHits >= 2 || signedPaddedHits >= 2)) return true;
                if (quotedPatternHits >= 3 && signedPaddedHits >= 2) return true;
                return false;
            }
            catch
            {
                return false;
            }
        }

        public DataTable Parse(string filePath)
        {
            DataTable dt = NewSchema();
            int sequentialId = 1;

            using (var sr = new StreamReader(filePath))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    string upper = line.ToUpperInvariant();
                    if (upper.Contains("ACCOUNT") || upper.StartsWith("\"ALL\"") || upper.Contains("BRANCH") ||
                        upper.Contains("ACC-NO") || upper.Contains("OPEN BALANCE") || upper.Contains("CLOSE BALANCE"))
                        continue;

                    string[] fields;
                    if (IsQuotedStandardBankLine(line)) fields = ParseQuotedStandardBankLine(line);
                    else
                    {
                        char detected = CSVHelper.DetectDelimiter(line);
                        if (line.IndexOf(detected) >= 0) fields = CSVHelper.SplitCSVLine(line, detected);
                        else if (Regex.IsMatch(line, @"\s{2,}"))
                        {
                            fields = Regex.Split(line.Trim(), @"\s{2,}")
                                          .Select(v => v.Trim(' ', '\t', '"'))
                                          .Where(v => v.Length > 0)
                                          .ToArray();
                        }
                        else fields = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    }

                    bool looksQuotedACB = fields.Length >= 5 && IsStdBankDate(fields[1]) && IsSignedPaddedAmount(fields[3]);
                    bool looksHeaderStyle = fields.Length >= 4 &&
                                            (fields[0].Contains("/") || fields[0].Contains("-") || IsStdBankDate(fields[0]) || YyyyMmDd.IsMatch(fields[0]));
                    if (!looksQuotedACB && !looksHeaderStyle) continue;

                    var row = dt.NewRow();

                    if (looksQuotedACB)
                    {
                        row["Ext. Tran. ID"] = SafeId(fields[0], sequentialId);
                        row["Tran. Date"] = ConvertDate(fields[1].Trim());
                        string desc = CleanDescription(fields[4]);
                        row["Tran. Desc"] = desc;
                        row["Ext. Ref. Nbr."] = fields.Length > 2 ? fields[2].Trim() : "";

                        string amountRaw = fields[3].Trim();
                        decimal amount;
                        if (TryParseAmount(amountRaw, out amount)) AssignAmount(row, amount);
                        else AssignAmount(row, 0m);
                    }
                    else
                    {
                        row["Ext. Tran. ID"] = sequentialId.ToString();
                        row["Tran. Date"] = ConvertDate(fields[0].Trim());
                        string description = fields.Length > 3 ? CleanDescription(fields[3]) : "";
                        row["Tran. Desc"] = description;
                        row["Ext. Ref. Nbr."] = fields.Length > 1 ? fields[1].Trim() : "";

                        decimal amount;
                        if (fields.Length > 2 && TryParseAmount(fields[2], out amount)) AssignAmount(row, amount);
                        else AssignAmount(row, 0m);
                    }

                    sequentialId++;
                    dt.Rows.Add(row);
                }
            }

            return dt;
        }

        private static DataTable NewSchema()
        {
            var dt = new DataTable();
            dt.Columns.Add("Ext. Tran. ID", typeof(string));
            dt.Columns.Add("Ext. Ref. Nbr.", typeof(string));
            dt.Columns.Add("Tran. Date", typeof(string));
            dt.Columns.Add("Tran. Desc", typeof(string));
            dt.Columns.Add("Receipt", typeof(string));
            dt.Columns.Add("Disbursement", typeof(string));
            return dt;
        }

        private bool IsQuotedStandardBankLine(string line) =>
            line.Length > 20 && line[0] == '"' && line.Count(c => c == '"') >= 10;

        private string[] ParseQuotedStandardBankLine(string line)
        {
            var parts = new System.Collections.Generic.List<string>();
            var cur = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"') { inQuotes = !inQuotes; continue; }
                if (c == ',' && !inQuotes)
                {
                    parts.Add(cur.ToString());
                    cur.Length = 0;
                }
                else cur.Append(c);
            }
            if (cur.Length > 0) parts.Add(cur.ToString());
            return parts.ToArray();
        }

        private bool IsStdBankDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            string v = value.Trim();
            if (v.Length == 9 && v.StartsWith("0")) v = v.Substring(1);
            if (v.Length != 8 || !v.All(char.IsDigit)) return false;
            return v.StartsWith("20");
        }

        private bool IsSignedPaddedAmount(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            value = value.Trim();
            if (!(value.StartsWith("+") || value.StartsWith("-"))) return false;
            value = value.Substring(1);
            return value.Replace(".", "").All(char.IsDigit);
        }

        private bool TryParseAmount(string raw, out decimal amount)
        {
            amount = 0m;
            if (string.IsNullOrWhiteSpace(raw)) return false;
            string s = raw.Trim();

            bool negative = false;
            if (s.StartsWith("+")) s = s.Substring(1);
            else if (s.StartsWith("-")) { negative = true; s = s.Substring(1); }

            if (s.IndexOf(',') >= 0 && s.IndexOf('.') >= 0) s = s.Replace(",", "");
            else if (s.Count(c => c == ',') > 1) s = s.Replace(",", "");

            if (s.Length > 1 && s[0] == '0')
            {
                int i = 0;
                while (i < s.Length - 1 && s[i] == '0' && s[i + 1] != '.') i++;
                s = s.Substring(i);
            }

            decimal parsed;
            if (!decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out parsed))
            {
                if (!decimal.TryParse(s, out parsed)) return false;
            }

            amount = negative ? -parsed : parsed;
            return true;
        }

        private void AssignAmount(DataRow row, decimal amount)
        {
            if (amount >= 0)
            {
                row["Receipt"] = FormatAmount(amount);
                row["Disbursement"] = "0";
            }
            else
            {
                row["Receipt"] = "0";
                row["Disbursement"] = FormatAmount(Math.Abs(amount));
            }
        }

        private string FormatAmount(decimal value)
        {
            if (decimal.Truncate(value) == value) return value.ToString("0", CultureInfo.InvariantCulture);
            return value.ToString("0.00", CultureInfo.InvariantCulture);
        }

        private string CleanDescription(string desc)
        {
            if (string.IsNullOrEmpty(desc)) return "";
            desc = desc.Replace("\t", " ").Replace("\r", " ").Replace("\n", " ");
            while (desc.Contains("  ")) desc = desc.Replace("  ", " ");
            return desc.Trim();
        }

        private string ConvertDate(string d)
        {
            if (string.IsNullOrWhiteSpace(d)) return "";
            string v = d.Trim();
            if (v.Length == 9 && v.StartsWith("0")) v = v.Substring(1);

            if (v.Length == 8 && v.All(char.IsDigit))
            {
                int y = int.Parse(v.Substring(0, 4));
                int m = int.Parse(v.Substring(4, 2));
                int day = int.Parse(v.Substring(6, 2));
                return new DateTime(y, m, day).ToString(DateOutputFormat);
            }

            DateTime parsed;
            string[] fmts = { "dd/MM/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "yyyy-MM-dd" };
            if (DateTime.TryParseExact(v, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
                return parsed.ToString(DateOutputFormat);
            if (DateTime.TryParse(v, out parsed))
                return parsed.ToString(DateOutputFormat);
            return v;
        }

        private string SafeId(string sourceField, int fallbackSeq)
        {
            if (string.IsNullOrWhiteSpace(sourceField)) return fallbackSeq.ToString();
            string trimmed = sourceField.Trim().Trim('"');
            if (trimmed.Length == 0) return fallbackSeq.ToString();
            return trimmed;
        }
    }
}