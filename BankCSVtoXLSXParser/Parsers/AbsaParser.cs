using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using BankCSVtoXLSXParser.Utils;

namespace BankCSVtoXLSXParser.Parsers
{
    /// <summary>
    /// ABSA statement parser.
    /// </summary>
    /// <remarks>
    /// Parsing strategy:
    /// - Supports CSV and TXT inputs. Delimiter is auto-detected (comma, semicolon, tab, pipe),
    ///   with a fallback to multi-space separation and then single-space tokens.
    /// - Detects typical ABSA transaction layout via heuristics:
    ///   • Presence of a YYYYMMDD date token (e.g., 20230501) near the beginning of a row.
    ///   • Indicators like DT/CT and some known ABSA keywords (e.g., CASHFOCUS, SETTLEMENT).
    /// - Amount detection:
    ///   • Prefers pattern DT|CT, then D|C, then <amount>.
    ///   • Falls back to the first parsable signed numeric token that is not a date.
    /// - Description:
    ///   • Built from tokens after the amount up to the next date-like token, skipping short uppercase codes.
    /// - Reference:
    ///   • Extracted from description using common patterns, or any long numeric token when absent.
    /// Output schema:
    /// - "Ext. Tran. ID", "Ext. Ref. Nbr.", "Tran. Date" (yyyy-MM-dd), "Tran. Desc", "Receipt", "Disbursement"
    /// </remarks>
    public sealed class AbsaParser : IBankParser
    {
        /// <inheritdoc />
        public string Name => "ABSA";

        /// <inheritdoc />
        public string ShortName => "ABSA";

        /// <summary>
        /// Regex used to detect YYYYMMDD dates beginning with 20 (e.g., 20230501).
        /// </summary>
        private static readonly Regex YyyyMmDd = new Regex(@"^20\d{6}$", RegexOptions.Compiled);

        /// <summary>
        /// Quickly checks if the file likely represents an ABSA layout.
        /// </summary>
        /// <param name="filePath">Absolute path to the input file.</param>
        /// <returns>True when heuristics indicate ABSA-like content; otherwise false.</returns>
        /// <remarks>
        /// Heuristics used:
        /// - A YYYYMMDD token in one of the first few columns.
        /// - Presence of DT/CT tokens.
        /// - Presence of common ABSA keywords (e.g., CASHFOCUS, SETTLEMENT).
        /// Only the first ~40 lines are inspected for performance.
        /// </remarks>
        public bool IsMatch(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                int hits = 0;
                int lines = 0;

                foreach (var line in File.ReadLines(filePath).Take(40))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    lines++;

                    // Tokenize robustly: detect delimiter or use multi-space fallback
                    var parts = Tokenize(line);

                    // Third token often is yyyymmdd
                    if (parts.Length >= 3 && YyyyMmDd.IsMatch(parts[2])) hits++;

                    // Presence of DT/CT and common ABSA keywords
                    if (parts.Any(p => p.Equals("DT", StringComparison.OrdinalIgnoreCase) ||
                                       p.Equals("CT", StringComparison.OrdinalIgnoreCase)))
                        hits++;

                    if (parts.Any(p => p.IndexOf("CASHFOCUS", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                       p.IndexOf("SETTLEMENT", StringComparison.OrdinalIgnoreCase) >= 0))
                        hits++;
                }

                return lines > 0 && hits >= 3;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Parses the supplied ABSA CSV/TXT file into the common transaction DataTable schema.
        /// </summary>
        /// <param name="filePath">Absolute path to the ABSA CSV/TXT file.</param>
        /// <returns>DataTable populated with normalized transaction rows.</returns>
        /// <exception cref="IOException">Thrown when the file cannot be read.</exception>
        public DataTable Parse(string filePath)
        {
            DataTable dt = NewSchema();
            int id = 1;

            foreach (var raw in File.ReadLines(filePath))
            {
                if (string.IsNullOrWhiteSpace(raw)) continue;

                var tokens = Tokenize(raw);
                if (tokens.Length < 4) continue;

                // Skip balance and header-like rows
                string joinedUpper = string.Join(",", tokens).ToUpperInvariant();
                if (joinedUpper.Contains("BALANCE B/FORWARD") || joinedUpper.Contains("BALANCE B\\FORWARD") ||
                    joinedUpper.Contains("ACCOUNT") || joinedUpper.StartsWith("ALL,"))
                    continue;

                // Find date token (first YYYYMMDD or dd/MM/yyyy etc.)
                string date = FindDate(tokens);
                if (string.IsNullOrEmpty(date)) continue;

                // Locate DT/CT + D/C + amount
                int idxType = Array.FindIndex(tokens, t => string.Equals(t, "DT", StringComparison.OrdinalIgnoreCase) ||
                                                           string.Equals(t, "CT", StringComparison.OrdinalIgnoreCase));
                bool? creditFlag = null;
                if (idxType >= 0)
                    creditFlag = string.Equals(tokens[idxType], "CT", StringComparison.OrdinalIgnoreCase);

                decimal? amount = null;
                int idxAmount = -1;

                // Preferred pattern: DT|CT, D|C, amount
                if (idxType >= 0)
                {
                    int idxDC = idxType + 1;
                    if (idxDC < tokens.Length &&
                        (string.Equals(tokens[idxDC], "D", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(tokens[idxDC], "C", StringComparison.OrdinalIgnoreCase)))
                    {
                        creditFlag = string.Equals(tokens[idxDC], "C", StringComparison.OrdinalIgnoreCase);
                        int idxAmtCandidate = idxDC + 1;
                        if (idxAmtCandidate < tokens.Length && TryParseAmount(tokens[idxAmtCandidate], out var a1))
                        {
                            amount = Math.Abs(a1);
                            idxAmount = idxAmtCandidate;
                        }
                    }
                }

                // Fallback: First signed numeric that is not a date
                if (!amount.HasValue)
                {
                    for (int i = 0; i < tokens.Length; i++)
                    {
                        if (YyyyMmDd.IsMatch(tokens[i])) continue;
                        if (TryParseAmount(tokens[i], out var a2))
                        {
                            amount = a2;
                            idxAmount = i;
                            if (!creditFlag.HasValue)
                                creditFlag = a2 >= 0m;
                            break;
                        }
                    }
                }

                // If no amount could be resolved, skip the line
                if (!amount.HasValue) continue;

                // Description: tokens after detected amount until next date-like token
                string description = BuildDescription(tokens, Math.Max(idxType, idxAmount));

                // Reference: try description; fallback to any long numeric token
                string reference = ReferenceExtractor.ExtractReference(description);
                if (string.IsNullOrEmpty(reference))
                {
                    var refTok = tokens.FirstOrDefault(t =>
                        !YyyyMmDd.IsMatch(t) && t.All(char.IsDigit) && t.Length >= 7);
                    if (!string.IsNullOrEmpty(refTok)) reference = refTok;
                }

                var row = dt.NewRow();
                row["Ext. Tran. ID"] = id.ToString();
                id++;
                row["Ext. Ref. Nbr."] = reference ?? "";
                row["Tran. Date"] = date;
                row["Tran. Desc"] = description;

                // Map to Receipt / Disbursement
                decimal signed = amount.Value;
                if (creditFlag.HasValue && !creditFlag.Value) signed = -Math.Abs(signed);

                if (signed >= 0)
                {
                    row["Receipt"] = signed.ToString("0.00");
                    row["Disbursement"] = "0.00";
                }
                else
                {
                    row["Receipt"] = "0.00";
                    row["Disbursement"] = Math.Abs(signed).ToString("0.00");
                }

                dt.Rows.Add(row);
            }

            return dt;
        }

        /// <summary>
        /// Splits a raw input line into tokens. Supports:
        /// - Auto-detected delimiters (comma, semicolon, tab, pipe)
        /// - Multi-space separated tokens
        /// - Single-space fallback
        /// </summary>
        /// <param name="line">Raw line to split.</param>
        /// <returns>Token array (never null).</returns>
        private static string[] Tokenize(string line)
        {
            if (string.IsNullOrEmpty(line)) return Array.Empty<string>();

            char delim = CSVHelper.DetectDelimiter(line);
            if (line.IndexOf(delim) >= 0)
            {
                return CSVHelper.SplitCSVLine(line, delim)
                                .Select(t => (t ?? "").Trim(' ', '\t', '"'))
                                .ToArray();
            }

            if (Regex.IsMatch(line, @"\s{2,}"))
            {
                return Regex.Split(line.Trim(), @"\s{2,}")
                            .Select(t => t.Trim(' ', '\t', '"'))
                            .Where(t => t.Length > 0)
                            .ToArray();
            }

            return line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        /// <summary>
        /// Creates a new DataTable with the standard output schema.
        /// </summary>
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

        /// <summary>
        /// Attempts to find and normalize a date token from the supplied token list.
        /// </summary>
        /// <param name="tokens">Tokenized row.</param>
        /// <returns>A yyyy-MM-dd date string when found; otherwise an empty string.</returns>
        private static string FindDate(string[] tokens)
        {
            var ymd = tokens.FirstOrDefault(t => YyyyMmDd.IsMatch(t));
            if (!string.IsNullOrEmpty(ymd))
            {
                int y = int.Parse(ymd.Substring(0, 4));
                int m = int.Parse(ymd.Substring(4, 2));
                int d = int.Parse(ymd.Substring(6, 2));
                try { return new DateTime(y, m, d).ToString("yyyy-MM-dd"); } catch { }
            }

            foreach (var t in tokens)
            {
                DateTime parsed;
                if (DateTime.TryParseExact(t.Trim(),
                    new[] { "dd/MM/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "yyyy-MM-dd" },
                    CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out parsed))
                    return parsed.ToString("yyyy-MM-dd");

                if (DateTime.TryParse(t, out parsed))
                    return parsed.ToString("yyyy-MM-dd");
            }

            return "";
        }

        /// <summary>
        /// Builds a human-readable description from tokens after the detected amount.
        /// Skips short all-letter codes and stops at the next date-like token.
        /// </summary>
        /// <param name="tokens">Tokenized input row.</param>
        /// <param name="startIndex">Index of the amount or type token from which to start building the description.</param>
        /// <returns>Normalized description string.</returns>
        private static string BuildDescription(string[] tokens, int startIndex)
        {
            if (startIndex < 0) startIndex = 0;
            var parts = tokens
                .Skip(startIndex + 1)
                .TakeWhile(t => !YyyyMmDd.IsMatch(t))
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Select(t => t.Trim());

            // Remove short uppercase codes that look like column codes (e.g., DID, CFA?, CMT, ACC)
            parts = parts.Where(t =>
            {
                var up = t.ToUpperInvariant();
                return !(up.Length <= 5 && up.All(c => char.IsLetter(c)));
            });

            var description = string.Join(" ", parts);
            description = Regex.Replace(description, @"\s{2,}", " ").Trim(' ', '.', ',');
            return description;
        }

        /// <summary>
        /// Tries to parse a token as a monetary amount.
        /// </summary>
        /// <param name="raw">Raw token value (may contain currency, grouping, signs).</param>
        /// <param name="value">Parsed decimal value when successful.</param>
        /// <returns>True if parsing succeeded; otherwise false.</returns>
        /// <remarks>
        /// Handling:
        /// - Removes currency symbols (R, r, $).
        /// - Recognizes parentheses and trailing minus for negatives.
        /// - Handles leading sign (+/-).
        /// - Normalizes grouping/decimal separators (commas/dots).
        /// - Ignores 8-digit YYYYMMDD-like tokens.
        /// </remarks>
        private static bool TryParseAmount(string raw, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            string s = raw.Trim();
            s = s.Replace("R", "").Replace("r", "").Replace("$", "").Trim();

            bool neg = false;
            if (s.StartsWith("(") && s.EndsWith(")"))
            {
                neg = true;
                s = s.Substring(1, s.Length - 2);
            }

            if (s.EndsWith("-"))
            {
                neg = true;
                s = s.Substring(0, s.Length - 1);
            }

            if (s.StartsWith("+")) s = s.Substring(1);
            else if (s.StartsWith("-"))
            {
                neg = true;
                s = s.Substring(1);
            }

            // Ignore dates like 20230501 (8 digits, all numeric)
            if (s.Length == 8 && s.All(char.IsDigit))
                return false;

            // Thousands separators
            if (s.IndexOf(',') >= 0 && s.IndexOf('.') >= 0)
                s = s.Replace(",", "");
            else if (s.Count(c => c == ',') > 1)
                s = s.Replace(",", "");
            else if (s.Count(c => c == ',') == 1 && !s.Contains("."))
                s = s.Replace(",", ".");

            decimal parsed;
            if (!decimal.TryParse(s, System.Globalization.NumberStyles.Number, CultureInfo.InvariantCulture, out parsed))
            {
                if (!decimal.TryParse(s, out parsed))
                    return false;
            }

            value = neg ? -Math.Abs(parsed) : parsed;
            return true;
        }
    }
}