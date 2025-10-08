using System;
using System.Data;
using System.Globalization;
using System.IO;
using BankCSVtoXLSXParser.Utils;
                
namespace BankCSVtoXLSXParser.Parsers
{
    /// <summary>
    /// Generic/fallback parser for delimited bank statement files.
    /// </summary>
    /// <remarks>
    /// This parser makes minimal assumptions about the input structure and is used as a last resort
    /// when no bank-specific parser matches. It attempts to:
    /// - Infer the delimiter from the first line (comma, semicolon, tab, pipe).
    /// - Map columns by index into the app's standard schema:
    ///   [1] Ext. Ref. Nbr., [2] Tran. Date, [3] Tran. Desc, [4] Amount.
    /// - Normalize dates to yyyy-MM-dd when possible.
    /// - Allocate positive amounts to Receipt and negative amounts to Disbursement.
    /// </remarks>
    public sealed class GenericCsvParser : IBankParser
    {
        /// <summary>
        /// Gets the human-friendly name for this parser.
        /// </summary>
        public string Name => "Generic CSV";

        /// <summary>
        /// Gets the short token for this parser used for file naming, sheet naming, and color selection.
        /// </summary>
        public string ShortName => "Bank";

        /// <summary>
        /// Indicates whether this parser supports the specified file.
        /// </summary>
        /// <param name="filePath">Absolute path to the input file.</param>
        /// <returns>
        /// Always returns true so the factory can fall back to this parser if no bank-specific parser matches.
        /// </returns>
        /// <remarks>
        /// Ensure this parser is evaluated last by the factory to avoid masking bank-specific matches.
        /// </remarks>
        public bool IsMatch(string filePath)
        {
            // Always a safe fallback
            return true;
        }

        /// <summary>
        /// Parses a delimited text file into the standard transaction DataTable schema.
        /// </summary>
        /// <param name="filePath">Absolute path to the CSV/TXT file to parse.</param>
        /// <returns>
        /// A <see cref="DataTable"/> with columns:
        /// "Ext. Tran. ID", "Ext. Ref. Nbr.", "Tran. Date" (yyyy-MM-dd when parseable),
        /// "Tran. Desc", "Receipt", "Disbursement".
        /// </returns>
        /// <remarks>
        /// - The delimiter is inferred from the first line via <see cref="CSVHelper.DetectDelimiter(string)"/>
        /// - Column mapping is index-based: reference (1), date (2), description (3), amount (4).
        /// - Amount parsing tries invariant culture first, then current culture; positive values map to Receipt,
        ///   negative values map to Disbursement.
        /// - Blank/whitespace lines are skipped.
        /// </remarks>
        /// <exception cref="IOException">Thrown if the file cannot be read.</exception>
        public DataTable Parse(string filePath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Ext. Tran. ID", typeof(string));
            dt.Columns.Add("Ext. Ref. Nbr.", typeof(string));
            dt.Columns.Add("Tran. Date", typeof(string));
            dt.Columns.Add("Tran. Desc", typeof(string));
            dt.Columns.Add("Receipt", typeof(string));
            dt.Columns.Add("Disbursement", typeof(string));

            using (var sr = new StreamReader(filePath))
            {
                string firstLine = sr.ReadLine();
                if (string.IsNullOrEmpty(firstLine)) return dt;

                char delimiter = CSVHelper.DetectDelimiter(firstLine);

                int id = 1;
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    string[] values = CSVHelper.SplitCSVLine(line, delimiter);
                    var dr = dt.NewRow();

                    dr["Ext. Tran. ID"] = id.ToString();
                    id++;

                    if (values.Length > 1) dr["Ext. Ref. Nbr."] = values[1].Trim();
                    if (values.Length > 2)
                    {
                        DateTime d;
                        dr["Tran. Date"] = DateTime.TryParse(values[2].Trim(), out d) ? d.ToString("yyyy-MM-dd") : values[2].Trim();
                    }
                    if (values.Length > 3) dr["Tran. Desc"] = values[3].Trim();
                    if (values.Length > 4)
                    {
                        decimal amount;
                        if (decimal.TryParse(values[4].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out amount) ||
                            decimal.TryParse(values[4].Trim(), out amount))
                        {
                            if (amount >= 0) { dr["Receipt"] = amount.ToString("0.00"); dr["Disbursement"] = "0.00"; }
                            else { dr["Receipt"] = "0.00"; dr["Disbursement"] = Math.Abs(amount).ToString("0.00"); }
                        }
                        else
                        {
                            dr["Receipt"] = "0.00"; dr["Disbursement"] = "0.00";
                        }
                    }

                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }
    }
}