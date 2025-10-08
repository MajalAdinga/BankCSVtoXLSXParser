        // File: Utils/CSVHelper.cs
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BankCSVtoXLSXParser.Utils
{
    /// <summary>
    /// CSV helper utilities for delimiter detection and robust line splitting.
    /// </summary>
    /// <remarks>
    /// This helper focuses on single-line tokenization and does not perform multi-line CSV parsing.
    /// It supports RFC4180-style quoting for fields (double quotes) when splitting.
    /// </remarks>
    public static class CSVHelper
    {
        /// <summary>
        /// Detects the most likely field delimiter for a given line of text.
        /// </summary>
        /// <param name="line">A single line of text to inspect.</param>
        /// <returns>
        /// The best-guess delimiter character among comma (<c>,</c>), semicolon (<c>;</c>), tab (<c>\t</c>),
        /// and pipe (<c>|</c>). Defaults to comma when no candidate has a positive count.
        /// </returns>
        /// <remarks>
        /// Strategy:
        /// - Count occurrences of candidate delimiters in the line.
        /// - Return the delimiter with the highest count (ties are resolved by priority: tab, semicolon, pipe, comma).
        /// - If all counts are zero, returns comma.
        /// </remarks>
        public static char DetectDelimiter(string line)
        {
            int commaCount = line.Count(c => c == ',');
            int semicolonCount = line.Count(c => c == ';');
            int tabCount = line.Count(c => c == '\t');
            int pipeCount = line.Count(c => c == '|');

            int maxCount = System.Math.Max(System.Math.Max(commaCount, semicolonCount),
                                           System.Math.Max(tabCount, pipeCount));

            if (tabCount == maxCount && tabCount > 0) return '\t';
            if (semicolonCount == maxCount && semicolonCount > 0) return ';';
            if (pipeCount == maxCount && pipeCount > 0) return '|';
            return ',';
        }

        /// <summary>
        /// Splits a single CSV line into fields using the specified delimiter, honoring RFC4180-style quoting.
        /// </summary>
        /// <param name="line">The raw line to split.</param>
        /// <param name="delimiter">The delimiter to use (e.g., ',', ';', '\t', '|').</param>
        /// <returns>An array of field values as strings (never null; may be empty).</returns>
        /// <remarks>
        /// Quoting rules:
        /// - Double quotes toggle quoted mode; delimiters inside quoted sections are treated as literal characters.
        /// - Escaped quotes inside a quoted field are represented as two consecutive quotes (<c>""</c>)
        ///   and are converted to a single double quote in the output.
        /// - Trailing/leading quotes are not automatically trimmed outside quoted sections.
        /// Performance:
        /// - Uses a reusable <see cref="StringBuilder"/> and a single pass over the input line.
        /// </remarks>
        public static string[] SplitCSVLine(string line, char delimiter)
        {
            var result = new List<string>();
            var currentField = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == delimiter && !inQuotes)
                {
                    result.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }

            result.Add(currentField.ToString());
            return result.ToArray();
        }

        /// <summary>
        /// Parses a comma-separated line that may contain quoted fields, honoring RFC4180-style quoting.
        /// </summary>
        /// <param name="line">The raw comma-separated line to parse.</param>
        /// <returns>An array of field values as strings.</returns>
        /// <remarks>
        /// - This is a convenience overload specialized for comma delimiters.
        /// - Double quotes toggle quoted mode; commas inside quoted sections are treated as literal characters.
        /// - This method does not unescape doubled quotes (<c>""</c>) to a single quote; use <see cref="SplitCSVLine"/> if unescaping is required.
        /// </remarks>
        public static string[] ParseQuotedLine(string line)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            if (current.Length > 0 || inQuotes)
            {
                result.Add(current.ToString());
            }

            return result.ToArray();
        }
    }
}
