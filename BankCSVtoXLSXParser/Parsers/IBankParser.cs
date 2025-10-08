using System.Data;

namespace BankCSVtoXLSXParser.Parsers
{
    /// <summary>
    /// Defines the contract for a bank statement parser.
    /// Implementations should detect whether a given file matches their bank format
    /// and convert it to the common DataTable schema used by the application.
    /// </summary>
    public interface IBankParser
    {
        /// <summary>
        /// Human-friendly bank name (e.g., "ABSA", "FNB", "Standard Bank").
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Short token used for sheet naming, file naming, and color selection (e.g., "ABSA", "FNB", "StandardBank").
        /// </summary>
        string ShortName { get; }

        /// <summary>
        /// Lightweight sniff to determine whether the supplied file is likely in this bank's format.
        /// Implementations should avoid reading the entire file and rely on small samples/heuristics only.
        /// </summary>
        /// <param name="filePath">Absolute path to the CSV/TXT to inspect.</param>
        /// <returns>True if the file likely matches this parser's format; otherwise false.</returns>
        bool IsMatch(string filePath);

        /// <summary>
        /// Parses the specified bank statement file and converts it to the common schema:
        /// Columns:
        ///  - Ext. Tran. ID (string)
        ///  - Ext. Ref. Nbr. (string)
        ///  - Tran. Date (string, yyyy-MM-dd)
        ///  - Tran. Desc (string)
        ///  - Receipt (string, "0.00" format)
        ///  - Disbursement (string, "0.00" format)
        /// </summary>
        /// <param name="filePath">Absolute path to the CSV/TXT file to parse.</param>
        /// <returns>A populated DataTable in the common schema. May be empty if no rows were parsed.</returns>
        DataTable Parse(string filePath);
    }
}