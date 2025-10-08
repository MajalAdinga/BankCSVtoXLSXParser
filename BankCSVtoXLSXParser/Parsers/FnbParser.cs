using System;
using System.Collections.Generic;
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
    /// Parser for First National Bank (FNB) statement exports.
    /// Produces a normalized DataTable with the standard schema used by the application.
    /// </summary>
    /// <remarks>
    /// - Supports CSV and TXT inputs (comma/semicolon/tab/pipe or multi-space separated).
    /// - Attempts header detection; if no header is present, infers column layout heuristically.
    /// - Maps amounts to Receipt/Disbursement and extracts a best-effort reference number.
    /// </remarks>
    public sealed class FnbParser : IBankParser
    {
        /// <summary>
        /// Gets the human-friendly bank name.
        /// </summary>
        public string Name => "FNB";

        /// <summary>
        /// Gets the short name token used for file naming, worksheet naming, and color selection.
        /// </summary>
        public string ShortName => "FNB";

        /// <summary>
        /// Lightweight heuristic to decide if a file likely represents an FNB export.
        /// </summary>
        /// <param name="filePath">Absolute path to the CSV/TXT file.</param>
        /// <returns>
        /// True if common FNB header/data cues are detected within the first ~40 lines; otherwise false.
        /// </returns>
        /// <remarks>
        /// Heuristics:
        /// - Presence of header-like keywords (DATE, DESC/NARRATION/DETAIL, AMOUNT/DEBIT/CREDIT/DR/CR, REFERENCE/REF).
        /// - Lines that begin with a date pattern (dd/MM, dd-MM, yyyy-MM-dd, yyyy/MM/dd, or 8-digit yyyymmdd starting with '20').
        /// </remarks>
        public bool IsMatch(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return false;

                int headerScore = 0;
                int dateLikeLines = 0;
                int linesRead = 0;

                foreach (var line in File.ReadLines(filePath))
                {
                    if (linesRead >= 40) break;
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    linesRead++;

                    string probe = line.Trim();

                    if (Regex.IsMatch(probe, @"DATE", RegexOptions.IgnoreCase)) headerScore++;
                    if (Regex.IsMatch(probe, @"DESC|NARRATION|DETAIL", RegexOptions.IgnoreCase)) headerScore++;
                    if (Regex.IsMatch(probe, @"AMOUNT|DEBIT|CREDIT|DR|CR", RegexOptions.IgnoreCase)) headerScore++;
                    if (Regex.IsMatch(probe, @"REFERENCE|REF", RegexOptions.IgnoreCase)) headerScore++;

                    if (Regex.IsMatch(probe, @"^\s*(\d{2}[/-]\d{2}[/-]\d{2,4}|\d{4}[/-]\d{2}[/-]\d{2}|20\d{6})"))
                        dateLikeLines++;
                }

                return headerScore >= 2 || dateLikeLines >= 5;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Parses an FNB statement file and returns a normalized DataTable.
        /// </summary>
        /// <param name="filePath">Absolute path to the CSV/TXT file.</param>
        /// <returns>
        /// DataTable with columns: Ext. Tran. ID, Ext. Ref. Nbr., Tran. Date (yyyy-MM-dd), Tran. Desc, Receipt, Disbursement.
        /// </returns>
        /// <exception cref="Exception">Thrown if no valid transaction rows are found.</exception>
        public DataTable Parse(string filePath)
        {
            DataTable dt = NewSchema();

            int transactionIdCounter = 1;
            int dataRowsAdded = 0;

            int dateIdx = -1, descIdx = -1, amountIdx = -1, creditIdx = -1, debitIdx = -1, referenceIdx = -1;
            bool headerDetected = false, structureInferred = false;

            using (var sr = new StreamReader(filePath))
            {
                string rawLine;
                while ((rawLine = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(rawLine))
                        continue;

                    // Strip BOM if present on first line
                    if (!headerDetected && !structureInferred && rawLine.Length > 0 && rawLine[0] == '\uFEFF')
                        rawLine = rawLine.TrimStart('\uFEFF');

                    string[] values;
                    char delimiter = CSVHelper.DetectDelimiter(rawLine);
                    bool hasStandardDelimiter = rawLine.IndexOf(delimiter) >= 0;

                    if (hasStandardDelimiter)
                    {
                        values = CSVHelper.SplitCSVLine(rawLine, delimiter)
                            .Select(v => v.Trim())
                            .ToArray();
                    }
                    else if (Regex.IsMatch(rawLine, @"\s{2,}"))
                    {
                        values = Regex.Split(rawLine.Trim(), @"\s{2,}")
                                      .Select(v => v.Trim(' ', '\t', '"'))
                                      .Where(v => v.Length > 0)
                                      .ToArray();
                    }
                    else
                    {
                        values = new[] { rawLine.Trim() };
                    }

                    if (values.Length == 0) continue;

                    // One-time header detection or structure inference
                    if (!headerDetected && !structureInferred)
                    {
                        if (IsLikelyHeader(values))
                        {
                            DetectColumns(values, out dateIdx, out descIdx, out amountIdx, out creditIdx, out debitIdx, out referenceIdx);
                            headerDetected = true;
                            continue;
                        }

                        if (IsDateValue(values[0]))
                        {
                            TryInferLayout(values, ref dateIdx, ref descIdx, ref amountIdx, ref creditIdx, ref debitIdx, ref referenceIdx);
                            structureInferred = true;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    // Skip subsequent header-like lines
                    if (headerDetected && IsLikelyHeader(values))
                        continue;

                    // Validate date column
                    if (dateIdx < 0 || dateIdx >= values.Length || !IsDateValue(values[dateIdx]))
                    {
                        string joined = string.Join(" ", values).ToUpperInvariant();
                        if (joined.Contains("OPEN") || joined.Contains("CLOSE") || joined.Contains("BALANCE"))
                            continue;
                        continue;
                    }

                    // Build row
                    var row = dt.NewRow();
                    row["Ext. Tran. ID"] = transactionIdCounter.ToString();
                    transactionIdCounter++;

                    row["Tran. Date"] = ConvertDate(values[dateIdx]);

                    if (descIdx >= 0 && descIdx < values.Length)
                        row["Tran. Desc"] = values[descIdx];
                    else
                        row["Tran. Desc"] = BuildDescription(values, dateIdx, amountIdx, creditIdx, debitIdx, referenceIdx);

                    string refCandidate = referenceIdx >= 0 && referenceIdx < values.Length ? values[referenceIdx] : "";
                    if (string.IsNullOrEmpty(refCandidate))
                        refCandidate = ExtractReference(row["Tran. Desc"].ToString());
                    row["Ext. Ref. Nbr."] = refCandidate;

                    bool amountAssigned = false;

                    // Single Amount column
                    if (amountIdx >= 0 && amountIdx < values.Length)
                    {
                        AssignAmountFromSingle(values[amountIdx], row);
                        amountAssigned = true;
                    }
                    else
                    {
                        // Separate Credit/Debit columns
                        decimal credit = 0m, debit = 0m;
                        bool creditOk = creditIdx >= 0 && creditIdx < values.Length && TryParseDecimal(values[creditIdx], out credit);
                        bool debitOk = debitIdx >= 0 && debitIdx < values.Length && TryParseDecimal(values[debitIdx], out debit);

                        if (creditOk || debitOk)
                        {
                            row["Receipt"] = creditOk && credit > 0 ? credit.ToString("0.00") : "0.00";
                            row["Disbursement"] = debitOk && debit > 0 ? debit.ToString("0.00") : "0.00";
                            amountAssigned = true;
                        }
                    }

                    // Last resort inference
                    if (!amountAssigned)
                    {
                        decimal inferred;
                        if (TryFindSingleAmount(values, out inferred))
                        {
                            if (inferred >= 0)
                            {
                                row["Receipt"] = inferred.ToString("0.00");
                                row["Disbursement"] = "0.00";
                            }
                            else
                            {
                                row["Receipt"] = "0.00";
                                row["Disbursement"] = Math.Abs(inferred).ToString("0.00");
                            }
                        }
                        else
                        {
                            row["Receipt"] = "0.00";
                            row["Disbursement"] = "0.00";
                        }
                    }

                    dt.Rows.Add(row);
                    dataRowsAdded++;
                }
            }

            if (dataRowsAdded == 0)
                throw new Exception("No valid transaction data found in FNB file.");

            return dt;
        }

        /// <summary>
        /// Creates a new DataTable initialized with the standard output schema.
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
        /// Determines if the provided token array resembles a header row.
        /// </summary>
        /// <param name="values">Tokenized row to inspect.</param>
        /// <returns>True if a minimal header score threshold is reached; otherwise false.</returns>
        private bool IsLikelyHeader(string[] values)
        {
            if (values == null || values.Length < 2) return false;
            int score = 0;
            foreach (var t0 in values)
            {
                var t = t0.Trim().ToUpperInvariant();
                if (t.Contains("DATE")) score++;
                else if (t.Contains("DESC") || t.Contains("DESCRIPTION") || t.Contains("NARRATION")) score++;
                else if (t.Contains("AMOUNT")) score++;
                else if (t.Contains("DEBIT") || t == "DR") score++;
                else if (t.Contains("CREDIT") || t == "CR") score++;
                else if (t.Contains("BALANCE")) score++;
                else if (t.Contains("REFERENCE") || t == "REF") score++;
            }
            return score >= 2;
        }

        /// <summary>
        /// Maps header tokens to known column indices.
        /// </summary>
        /// <param name="headerValues">The header row tokens.</param>
        /// <param name="dateIdx">Out: index of the Date column; -1 if not found.</param>
        /// <param name="descIdx">Out: index of the Description column; -1 if not found.</param>
        /// <param name="amountIdx">Out: index of the Amount column; -1 if not found.</param>
        /// <param name="creditIdx">Out: index of the Credit column; -1 if not found.</param>
        /// <param name="debitIdx">Out: index of the Debit column; -1 if not found.</param>
        /// <param name="referenceIdx">Out: index of the Reference column; -1 if not found.</param>
        private void DetectColumns(string[] headerValues, out int dateIdx, out int descIdx, out int amountIdx, out int creditIdx, out int debitIdx, out int referenceIdx)
        {
            dateIdx = descIdx = amountIdx = creditIdx = debitIdx = referenceIdx = -1;

            for (int i = 0; i < headerValues.Length; i++)
            {
                string h = headerValues[i].Trim().ToUpperInvariant();

                if (dateIdx == -1 && h.Contains("DATE")) dateIdx = i;
                else if (descIdx == -1 && (h.Contains("DESC") || h.Contains("DESCRIPTION") || h.Contains("NARRATION") || h.Contains("DETAIL"))) descIdx = i;
                else if (referenceIdx == -1 && (h.Contains("REFERENCE") || h == "REF" || h.Contains("REF."))) referenceIdx = i;
                else if (amountIdx == -1 && h.Contains("AMOUNT")) amountIdx = i;
                else if (creditIdx == -1 && (h.Contains("CREDIT") || h == "CR" || h.Contains("CR. RECEIPTS"))) creditIdx = i;
                else if (debitIdx == -1 && (h.Contains("DEBIT") || h == "DR" || h.Contains("DR. PAYMENTS"))) debitIdx = i;
            }

            // Prefer explicit Credit/Debit columns over a generic Amount column
            if (creditIdx >= 0 && debitIdx >= 0) amountIdx = -1;
        }

        /// <summary>
        /// Attempts to infer a basic column layout from a data row when no header is present.
        /// </summary>
        /// <param name="values">Tokenized data row.</param>
        /// <param name="dateIdx">Ref: inferred date index.</param>
        /// <param name="descIdx">Ref: inferred description index.</param>
        /// <param name="amountIdx">Ref: inferred amount index (if single-amount layout).</param>
        /// <param name="creditIdx">Ref: inferred credit index (for split credit/debit layouts).</param>
        /// <param name="debitIdx">Ref: inferred debit index (for split credit/debit layouts).</param>
        /// <param name="referenceIdx">Ref: inferred reference index (if any).</param>
        private void TryInferLayout(string[] values, ref int dateIdx, ref int descIdx, ref int amountIdx, ref int creditIdx, ref int debitIdx, ref int referenceIdx)
        {
            if (dateIdx < 0) dateIdx = 0;

            var numericIndexes = new List<int>();
            for (int i = 1; i < values.Length; i++)
            {
                decimal tmp;
                if (TryParseDecimal(values[i], out tmp))
                    numericIndexes.Add(i);
            }

            if (numericIndexes.Count == 1) amountIdx = numericIndexes[0];
            else if (numericIndexes.Count >= 2)
            {
                debitIdx = numericIndexes[0];
                creditIdx = numericIndexes[1];
            }

            if (descIdx < 0)
            {
                for (int i = 1; i < values.Length; i++)
                {
                    if (i == amountIdx || i == creditIdx || i == debitIdx) continue;
                    if (IsDateValue(values[i])) continue;

                    decimal tmp;
                    if (TryParseDecimal(values[i], out tmp)) continue;

                    descIdx = i;
                    break;
                }
            }
        }

        /// <summary>
        /// Tries to parse a token into a decimal using invariant culture (with fallback to current culture).
        /// </summary>
        /// <param name="raw">Raw token text.</param>
        /// <param name="value">Out: parsed decimal value when successful.</param>
        /// <returns>True if parsing succeeded; otherwise false.</returns>
        private bool TryParseDecimal(string raw, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            string a = CleanAmountString(raw);

            decimal parsed;
            if (decimal.TryParse(a, NumberStyles.Number, CultureInfo.InvariantCulture, out parsed))
            {
                value = parsed;
                return true;
            }
            if (decimal.TryParse(a, out parsed))
            {
                value = parsed; return true;
            }
            return false;
        }

        /// <summary>
        /// Normalizes a raw amount string by removing currency codes, handling DR/CR and signs,
        /// and unifying thousand/decimal separators.
        /// </summary>
        /// <param name="s">Raw amount string.</param>
        /// <returns>Cleaned numeric string suitable for decimal parsing.</returns>
        private string CleanAmountString(string s)
        {
            string a = s.Trim().Replace("R", "").Replace("r", "").Replace("$", "").Trim();

            bool neg = false;
            if (a.StartsWith("(") && a.EndsWith(")")) { neg = true; a = a.Substring(1, a.Length - 2); }
            if (a.EndsWith("-")) { neg = true; a = a.Substring(0, a.Length - 1); }

            if (a.EndsWith("CR", StringComparison.OrdinalIgnoreCase)) a = a.Substring(0, a.Length - 2).Trim();
            else if (a.EndsWith("DR", StringComparison.OrdinalIgnoreCase)) { a = a.Substring(0, a.Length - 2).Trim(); neg = true; }

            if (a.StartsWith("+")) a = a.Substring(1);
            else if (a.StartsWith("-")) { neg = true; a = a.Substring(1); }

            if (a.IndexOf(',') >= 0 && a.IndexOf('.') >= 0) a = a.Replace(",", "");
            else if (a.Count(c => c == ',') > 1) a = a.Replace(",", "");
            else if (a.Count(c => c == ',') == 1 && !a.Contains(".")) a = a.Replace(",", ".");

            a = a.Replace(" ", "");
            if (neg) a = "-" + a;
            return a;
        }

        /// <summary>
        /// Assigns a single parsed amount to either Receipt or Disbursement on the given row.
        /// Handles negative signs and DR/CR semantics.
        /// </summary>
        /// <param name="amountStr">Raw amount token.</param>
        /// <param name="row">Target DataRow to populate.</param>
        private void AssignAmountFromSingle(string amountStr, DataRow row)
        {
            if (string.IsNullOrWhiteSpace(amountStr))
            {
                row["Receipt"] = "0.00"; row["Disbursement"] = "0.00"; return;
            }

            bool isNegative = false;
            string cleaned = amountStr.Trim().Replace("R", "").Replace("r", "").Replace("$", "").Trim();

            if (cleaned.StartsWith("(") && cleaned.EndsWith(")")) { isNegative = true; cleaned = cleaned.Substring(1, cleaned.Length - 2); }
            if (cleaned.EndsWith("-")) { isNegative = true; cleaned = cleaned.Substring(0, cleaned.Length - 1); }
            if (cleaned.StartsWith("-")) { isNegative = true; cleaned = cleaned.Substring(1); }
            else if (cleaned.StartsWith("+")) { cleaned = cleaned.Substring(1); }

            if (cleaned.EndsWith("CR", StringComparison.OrdinalIgnoreCase)) { cleaned = cleaned.Substring(0, cleaned.Length - 2).Trim(); }
            else if (cleaned.EndsWith("DR", StringComparison.OrdinalIgnoreCase)) { cleaned = cleaned.Substring(0, cleaned.Length - 2).Trim(); isNegative = true; }

            if (cleaned.Contains(",") && cleaned.Contains(".")) cleaned = cleaned.Replace(",", "");
            else if (cleaned.Count(c => c == ',') > 1) cleaned = cleaned.Replace(",", "");
            else if (cleaned.Count(c => c == ',') == 1 && !cleaned.Contains(".")) cleaned = cleaned.Replace(",", ".");

            cleaned = cleaned.Replace(" ", "");

            decimal amount;
            if (!decimal.TryParse(cleaned, NumberStyles.Number, CultureInfo.InvariantCulture, out amount))
            {
                if (!decimal.TryParse(cleaned, out amount))
                {
                    row["Receipt"] = "0.00"; row["Disbursement"] = "0.00"; return;
                }
            }

            if (isNegative) amount = -Math.Abs(amount);

            if (amount >= 0) { row["Receipt"] = amount.ToString("0.00"); row["Disbursement"] = "0.00"; }
            else { row["Receipt"] = "0.00"; row["Disbursement"] = Math.Abs(amount).ToString("0.00"); }
        }

        /// <summary>
        /// Builds a description string by concatenating non-date, non-numeric tokens,
        /// excluding known indices (date/amount/credit/debit/reference).
        /// </summary>
        /// <param name="values">Tokenized row.</param>
        /// <param name="dateIdx">Index of the date column.</param>
        /// <param name="amountIdx">Index of the amount column (or -1 if not present).</param>
        /// <param name="creditIdx">Index of the credit column (or -1 if not present).</param>
        /// <param name="debitIdx">Index of the debit column (or -1 if not present).</param>
        /// <param name="referenceIdx">Index of the reference column (or -1 if not present).</param>
        /// <returns>A trimmed description string.</returns>
        private string BuildDescription(string[] values, int dateIdx, int amountIdx, int creditIdx, int debitIdx, int referenceIdx)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < values.Length; i++)
            {
                if (i == dateIdx || i == amountIdx || i == creditIdx || i == debitIdx || i == referenceIdx) continue;
                if (string.IsNullOrWhiteSpace(values[i])) continue;
                decimal d; if (TryParseDecimal(values[i], out d)) continue;
                if (IsDateValue(values[i])) continue;
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(values[i].Trim());
            }
            return sb.ToString().Trim();
        }

        /// <summary>
        /// Attempts to extract a transaction reference number from the provided text using common patterns.
        /// </summary>
        /// <param name="text">Source text (typically the transaction description).</param>
        /// <returns>A numeric reference string when detected; otherwise an empty string.</returns>
        private string ExtractReference(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            var startNumMatch = Regex.Match(text, @"^(\d{4,})");
            if (startNumMatch.Success) return startNumMatch.Groups[1].Value;

            var refMatch = Regex.Match(text, @"(?:REF|Ref|REFERENCE)[\s:#]*(\d+)", RegexOptions.IgnoreCase);
            if (refMatch.Success) return refMatch.Groups[1].Value;

            var trnMatch = Regex.Match(text, @"(?:TRN|TRANS|TRANSACTION)[\s:#]*(\d+)", RegexOptions.IgnoreCase);
            if (trnMatch.Success) return trnMatch.Groups[1].Value;

            var pmtMatch = Regex.Match(text, @"(?:PMT|PAYMENT)[\s:#]*(\d+)", RegexOptions.IgnoreCase);
            if (pmtMatch.Success) return pmtMatch.Groups[1].Value;

            var numMatch = Regex.Match(text, @"\b(\d{5,})\b");
            if (numMatch.Success) return numMatch.Groups[1].Value;
            return "";
        }

        /// <summary>
        /// Attempts to locate exactly one numeric token in the row and returns it as the inferred amount.
        /// </summary>
        /// <param name="values">Tokenized row values.</param>
        /// <param name="amount">Out: the inferred amount when exactly one numeric candidate exists.</param>
        /// <returns>True if a single candidate was found; otherwise false.</returns>
        private bool TryFindSingleAmount(string[] values, out decimal amount)
        {
            amount = 0m;
            var candidates = new System.Collections.Generic.List<decimal>();
            foreach (var v in values)
            {
                decimal tmp;
                if (TryParseDecimal(v, out tmp)) candidates.Add(tmp);
            }
            if (candidates.Count == 1) { amount = candidates[0]; return true; }
            return false;
        }

        /// <summary>
        /// Determines whether the provided value looks like a date.
        /// </summary>
        /// <param name="value">Token value to test.</param>
        /// <returns>True if the token resembles a date; otherwise false.</returns>
        private bool IsDateValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            string v = value.Trim();
            if (v.Contains("/") || v.Contains("-"))
            {
                DateTime dt; if (DateTime.TryParse(v, out dt)) return true;
            }
            if (v.Length == 8 && v.All(char.IsDigit) && v.StartsWith("20")) return true;
            DateTime parsed; return DateTime.TryParse(v, out parsed);
        }

        /// <summary>
        /// Converts a date token to yyyy-MM-dd if possible; otherwise returns a trimmed original value.
        /// </summary>
        /// <param name="dateStr">Source date string.</param>
        /// <returns>Normalized yyyy-MM-dd string or the original trimmed value when parsing fails.</returns>
        private string ConvertDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return "";
            DateTime date;
            string[] formats = { "yyyy/MM/dd", "yyyy-MM-dd", "dd/MM/yyyy", "dd-MM-yyyy", "dd MMM yyyy", "dd MMMM yyyy", "yyyy/MM/dd HH:mm", "dd/MM/yyyy HH:mm", "yyyyMMdd" };
            if (DateTime.TryParseExact(dateStr.Trim(), formats, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out date))
                return date.ToString("yyyy-MM-dd");
            if (DateTime.TryParse(dateStr, out date))
                return date.ToString("yyyy-MM-dd");
            return dateStr.Trim();
        }
    }
}