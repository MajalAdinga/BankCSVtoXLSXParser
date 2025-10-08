using System;
using System.Collections.Generic;
using System.IO;

namespace BankCSVtoXLSXParser.Parsers
{
    /// <summary>
    /// Factory responsible for selecting the appropriate bank statement parser for a given file.
    /// </summary>
    /// <remarks>
    /// Resolution order:
    /// <list type="number">
    ///   <item>
    ///     <description>
    ///       If the user explicitly selected a bank name in the UI, the factory honors that selection
    ///       and returns the corresponding parser (ABSA, FNB, or Standard Bank).
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       Otherwise, the factory iterates through all registered parsers and calls
    ///       <see cref="IBankParser.IsMatch(string)"/> to heuristically detect a match using a small sample of the file.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       If no parser matches, the factory returns the <see cref="GenericCsvParser"/> as a safe fallback.
    ///     </description>
    ///   </item>
    /// </list>
    /// Thread-safety: resolution is stateless and the static parser list is read-only, so this type is safe to use across threads.
    /// </remarks>
    public static class BankParserFactory
    {
        /// <summary>
        /// Registered parsers used for auto-detection (in priority order).
        /// </summary>
        private static readonly List<IBankParser> Parsers = new List<IBankParser>
        {
            new AbsaParser(),
            new FnbParser(),
            new StandardBankParser(),
            new GenericCsvParser()
        };

        /// <summary>
        /// Resolves the most suitable <see cref="IBankParser"/> for the provided file.
        /// </summary>
        /// <param name="bankName">
        /// User-selected bank name from the UI (may be null/empty). If provided, the selection is honored first.
        /// Matching is case-insensitive and looks for keywords:
        /// "ABSA", "FNB" / "FIRST NATIONAL", or "STANDARD".
        /// </param>
        /// <param name="filePath">Absolute path to the CSV/TXT file to parse.</param>
        /// <returns>
        /// A concrete <see cref="IBankParser"/> implementation. Returns <see cref="GenericCsvParser"/> if no bank-specific parser matches.
        /// </returns>
        /// <exception cref="FileNotFoundException">Thrown if the file path does not exist.</exception>
        /// <remarks>
        /// - When <paramref name="bankName"/> is provided, this method returns a new instance of the corresponding parser
        ///   without attempting auto-detection.
        /// - During auto-detection, each registered parser’s <see cref="IBankParser.IsMatch(string)"/> is called in order until one returns true.
        /// - Any exceptions thrown by <see cref="IBankParser.IsMatch(string)"/> are swallowed to keep resolution robust.
        /// </remarks>
        public static IBankParser Resolve(string bankName, string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("Input file not found.", filePath);

            // If user has explicitly selected a bank, honor it first
            if (!string.IsNullOrWhiteSpace(bankName))
            {
                string b = bankName.ToUpperInvariant();
                if (b.Contains("ABSA")) return new AbsaParser();
                if (b.Contains("FNB") || b.Contains("FIRST NATIONAL")) return new FnbParser();
                if (b.Contains("STANDARD")) return new StandardBankParser();
            }

            // Otherwise try auto-detect
            foreach (var p in Parsers)
            {
                try
                {
                    if (p.IsMatch(filePath))
                        return p;
                }
                catch
                {
                    // Ignore sniff errors and continue with the next parser
                }
            }

            // Fallback
            return new GenericCsvParser();
        }
    }
}