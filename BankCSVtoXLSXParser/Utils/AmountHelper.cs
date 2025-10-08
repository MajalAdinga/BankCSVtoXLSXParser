// File: Utils/AmountHelper.cs
using System;

namespace BankCSVtoXLSXParser.Utils
{
    /// <summary>
    /// Utility functions for cleaning and mapping monetary values to the app's
    /// standard output format (Receipt/Disbursement).
    /// </summary>
    /// <remarks>
    /// All helpers are culture-agnostic and expect normalized tokens. Currency symbols
    /// (e.g., R, $) and thousands separators should be removed or will be handled where noted.
    /// </remarks>
    public static class AmountHelper
    {
        /// <summary>
        /// Removes leading zeros from a numeric string while preserving the decimal part, if present.
        /// </summary>
        /// <param name="value">A numeric string (e.g., "000123.45", "0007", "0.50").</param>
        /// <returns>
        /// The input without leading zeros. If the integer part becomes empty, it returns "0".
        /// For example:
        /// - "000123.45" => "123.45"
        /// - "0000" => "0"
        /// - "000.50" => "0.50"
        /// </returns>
        public static string RemoveLeadingZeros(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            if (value.Contains("."))
            {
                string[] parts = value.Split('.');
                string intPart = parts[0].TrimStart('0');
                if (string.IsNullOrEmpty(intPart))
                    intPart = "0";
                return intPart + "." + parts[1];
            }

            string result = value.TrimStart('0');
            return string.IsNullOrEmpty(result) ? "0" : result;
        }

        /// <summary>
        /// Cleans a raw amount string and maps it to Receipt and Disbursement strings
        /// in the app's standard "0.00" format.
        /// </summary>
        /// <param name="amountStr">
        /// Raw amount value which may include currency symbols (e.g., "R", "$"),
        /// whitespace, thousands separators, signs (+/-), or parentheses to indicate negatives.
        /// Examples: "R 1,234.56", "-250.00", "(100.00)", "+75", "00045.00".
        /// </param>
        /// <returns>
        /// A tuple where:
        /// - <c>receipt</c> is the formatted positive amount when the input is non-negative, otherwise "0.00".
        /// - <c>disbursement</c> is the formatted positive amount when the input is negative, otherwise "0.00".
        /// </returns>
        /// <remarks>
        /// Processing steps:
        /// 1) Trims spaces and removes common currency symbols (R, r, $).
        /// 2) Removes thousands separators when both comma and decimal point are present.
        /// 3) Interprets negative values via a leading '-' or surrounding parentheses.
        /// 4) Removes a leading '+' when present.
        /// 5) Removes leading zeros from the integer part.
        /// 6) Parses to <see cref="decimal"/> and formats to "0.00".
        /// 
        /// If parsing fails at any step, both Receipt and Disbursement return "0.00".
        /// </remarks>
        public static (string receipt, string disbursement) ProcessAmount(string amountStr)
        {
            if (string.IsNullOrWhiteSpace(amountStr))
                return ("0.00", "0.00");

            // Clean amount
            amountStr = amountStr.Trim();
            amountStr = amountStr.Replace("R", "").Replace("r", "").Replace("$", "").Replace(" ", "");

            // Handle thousands separators
            if (amountStr.Contains(",") && amountStr.Contains("."))
            {
                amountStr = amountStr.Replace(",", "");
            }

            // Check for negative
            bool isNegative = false;

            if (amountStr.StartsWith("-"))
            {
                isNegative = true;
                amountStr = amountStr.Substring(1);
            }
            else if (amountStr.StartsWith("(") && amountStr.EndsWith(")"))
            {
                isNegative = true;
                amountStr = amountStr.Replace("(", "").Replace(")", "");
            }
            else if (amountStr.StartsWith("+"))
            {
                amountStr = amountStr.Substring(1);
            }

            // Remove leading zeros
            amountStr = RemoveLeadingZeros(amountStr);

            decimal amount;
            if (decimal.TryParse(amountStr, out amount))
            {
                amount = Math.Abs(amount);
                string formattedAmount = amount.ToString("0.00");

                if (isNegative)
                    return ("0.00", formattedAmount);
                else
                    return (formattedAmount, "0.00");
            }

            return ("0.00", "0.00");
        }
    }
}