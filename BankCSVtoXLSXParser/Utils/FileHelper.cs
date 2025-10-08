// File: Utils/FileHelper.cs
using System.IO;
using System.Linq;

namespace BankCSVtoXLSXParser.Utils
{
    /// <summary>
    /// File-related utility helpers used across the application.
    /// </summary>
    /// <remarks>
    /// This helper provides lightweight checks for input files such as quoted CSV detection,
    /// extension retrieval, and basic validation of allowed file types for parsing.
    /// All methods are safe for use in UI flows; any I/O exceptions are swallowed
    /// where noted to keep the UX responsive and resilient.
    /// </remarks>
    public static class FileHelper
    {
        /// <summary>
        /// Heuristically determines whether the specified file appears to be a
        /// quoted CSV layout (e.g., "field1","field2",...).
        /// </summary>
        /// <param name="filePath">Absolute path to the file to probe.</param>
        /// <returns>
        /// True if the first few lines contain evidence of a quoted CSV format; otherwise false.
        /// </returns>
        /// <remarks>
        /// Implementation details:
        /// <list type="bullet">
        ///   <item>Reads up to the first 5 lines to minimize I/O.</item>
        ///   <item>Checks for common quoted-field patterns such as <c>"",""</c> separators.</item>
        ///   <item>
        ///     Looks for known tokens often present in bank exports (e.g., <c>"0</c>, <c>"ACB</c>, <c>"INTPAY</c>, <c>"PAY</c>).
        ///   </item>
        ///   <item>Any exceptions (e.g., file not found, access denied) result in a return value of false.</item>
        /// </list>
        /// Note: This is only a heuristic and should not be treated as a definitive CSV validator.
        /// </remarks>
        public static bool IsQuotedFormat(string filePath)
        {
            try
            {
                string[] sampleLines = File.ReadLines(filePath).Take(5).ToArray();

                foreach (string line in sampleLines)
                {
                    if (line.Contains("\"") && line.Contains("\",\"") &&
                        (line.Contains("\"0") || line.Contains("\"ACB") ||
                         line.Contains("\"INTPAY") || line.Contains("\"PAY")))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                // Any I/O or access exceptions are treated as "not quoted"
                return false;
            }
        }

        /// <summary>
        /// Returns the lowercase file extension for the provided path (including the leading dot).
        /// </summary>
        /// <param name="filePath">Absolute or relative file path.</param>
        /// <returns>
        /// The file extension in lowercase (e.g., <c>".csv"</c>, <c>".txt"</c>), or an empty string if none exists.
        /// </returns>
        public static string GetFileExtension(string filePath)
        {
            return Path.GetExtension(filePath).ToLower();
        }

        /// <summary>
        /// Indicates whether the file has a supported extension for bank parsing.
        /// </summary>
        /// <param name="filePath">Absolute or relative file path.</param>
        /// <returns>
        /// True if the file extension is <c>.csv</c> or <c>.txt</c>; otherwise false.
        /// </returns>
        /// <remarks>
        /// This check is purely extension-based and does not validate the file's content.
        /// </remarks>
        public static bool IsValidBankFile(string filePath)
        {
            string ext = GetFileExtension(filePath);
            return ext == ".csv" || ext == ".txt";
        }
    }
}