// File: Utils/DateHelper.cs
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace BankCSVtoXLSXParser.Utils
{
    /// <summary>
    /// Provides helper methods for recognizing and converting date values to a standard format.
    /// </summary>
    /// <remarks>
    /// - Standard output format is yyyy-MM-dd.
    /// - Includes special handling for:
    ///   • Standard Bank compact format with leading zero (e.g., 020241101 -> 20241101 -> 2024-11-01).
    ///   • Raw YYYYMMDD tokens (e.g., 20230501 -> 2023-05-01).
    /// - Falls back to parsing common localized formats via <see cref="DateTime.TryParse"/> when needed.
    /// </remarks>
    public static class DateHelper
    {
        /// <summary>
        /// Converts an input date string to the standard yyyy-MM-dd format when possible.
        /// </summary>
        /// <param name="dateStr">The input date string to convert (e.g., 20230501, 2023/05/01, 01-05-2023).</param>
        /// <returns>
        /// The normalized date in yyyy-MM-dd format if parsing succeeds; otherwise, the original input trimmed.
        /// Returns an empty string when the input is null or whitespace.
        /// </returns>
        /// <remarks>
        /// Processing order:
        /// 1) Strips a leading zero for Standard Bank compact format (length 9, starting with '0').
        /// 2) Parses strict YYYYMMDD tokens.
        /// 3) Attempts a set of common exact formats (e.g., yyyy/MM/dd, dd-MM-yyyy, dd MMM yyyy).
        /// 4) Falls back to <see cref="DateTime.TryParse(string, out DateTime)"/>.
        /// </remarks>
        public static string ConvertToStandardFormat(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr))
                return "";

            // Handle Standard Bank format: 020241101 (0YYYYMMDD)
            if (dateStr.Length == 9 && dateStr.StartsWith("0"))
            {
                dateStr = dateStr.Substring(1);
            }

            // Handle YYYYMMDD format
            if (dateStr.Length == 8 && !dateStr.Contains("/") && !dateStr.Contains("-"))
            {
                string year = dateStr.Substring(0, 4);
                string month = dateStr.Substring(4, 2);
                string day = dateStr.Substring(6, 2);

                try
                {
                    DateTime dt = new DateTime(int.Parse(year), int.Parse(month), int.Parse(day));
                    return dt.ToString("yyyy-MM-dd");
                }
                catch
                {
                    return dateStr;
                }
            }

            // Try common date formats
            string[] formats = {
                "yyyy/MM/dd", "yyyy-MM-dd", "dd/MM/yyyy", "dd-MM-yyyy",
                "dd MMM yyyy", "dd MMMM yyyy", "yyyy/MM/dd HH:mm", "dd/MM/yyyy HH:mm"
            };

            DateTime date;
            if (DateTime.TryParseExact(dateStr, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date.ToString("yyyy-MM-dd");
            }

            if (DateTime.TryParse(dateStr, out date))
            {
                return date.ToString("yyyy-MM-dd");
            }

            return dateStr;
        }

        /// <summary>
        /// Heuristically determines whether a string likely represents a date.
        /// </summary>
        /// <param name="value">The input string to test.</param>
        /// <returns>
        /// True if the value appears to be a date (contains separators parsable by <see cref="DateTime.TryParse"/>,
        /// or starts with "20" and has at least 4 characters); otherwise false.
        /// </returns>
        /// <remarks>
        /// This is a permissive check intended for parsing pipelines. It does not guarantee that
        /// <see cref="ConvertToStandardFormat(string)"/> will succeed, only that the token looks date-like.
        /// </remarks>
        public static bool IsDateValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            if (value.Contains("/") || value.Contains("-"))
            {
                DateTime testDate;
                return DateTime.TryParse(value, out testDate);
            }

            if (value.Length >= 4 && value.StartsWith("20"))
                return true;

            return false;
        }
    }
}