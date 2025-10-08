// File: Utils/ReferenceExtractor.cs
using System.Text.RegularExpressions;

namespace BankCSVtoXLSXParser.Utils
{
    public static class ReferenceExtractor
    {
        public static string ExtractReference(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";

            // Pattern 1: REF: followed by number
            var refMatch = Regex.Match(text, @"(?:REF|Ref)[:]\s*(\d+)", RegexOptions.IgnoreCase);
            if (refMatch.Success)
                return refMatch.Groups[1].Value;

            // Pattern 2: Numbers at the end
            var endNumberMatch = Regex.Match(text, @"\b(\d{5,})\s*$");
            if (endNumberMatch.Success)
                return endNumberMatch.Groups[1].Value;

            // Pattern 3: Customer numbers
            var custMatch = Regex.Match(text, @"(?:Cust No|Customer)\s+(\d+)", RegexOptions.IgnoreCase);
            if (custMatch.Success)
                return custMatch.Groups[1].Value;

            // Pattern 4: Reference in parentheses
            var parenMatch = Regex.Match(text, @"\)(\d{5,})");
            if (parenMatch.Success)
                return parenMatch.Groups[1].Value;

            // Pattern 5: Any 5-6 digit number
            var numberMatch = Regex.Match(text, @"\b(\d{5,6})\b");
            if (numberMatch.Success)
                return numberMatch.Groups[1].Value;

            return "";
        }
    }
}