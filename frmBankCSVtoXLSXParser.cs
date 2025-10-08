private DataTable ParseCPOASBFormat(string filePath)
{
    DataTable dt = new DataTable();

    // Acumatica-compatible columns
    dt.Columns.Add("Ext. Tran. ID", typeof(string));
    dt.Columns.Add("Ext. Ref. Nbr.", typeof(string));
    dt.Columns.Add("Tran. Date", typeof(string));
    dt.Columns.Add("Tran. Desc", typeof(string));
    dt.Columns.Add("Receipt", typeof(string));
    dt.Columns.Add("Disbursement", typeof(string));

    using (StreamReader sr = new StreamReader(filePath))
    {
        string line;
        bool foundDataSection = false;
        int transactionId = 1;

        while ((line = sr.ReadLine()) != null)
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Skip header section until we find transaction data
            if (!foundDataSection)
            {
                if (line.Contains("\",\"") && !line.Contains("BRANCH"))
                {
                    foundDataSection = true;
                }
                else
                {
                    continue;
                }
            }

            string[] fields = ParseQuotedLine(line);

            // Validate minimum required fields
            if (fields.Length < 5)
                continue;

            // Skip OPEN/CLOSE entries
            string tranType = fields[2].Trim().ToUpper();
            if (tranType == "OPEN" || tranType == "CLOSE")
                continue;

            DataRow row = dt.NewRow();

            // Map fields according to Acumatica format
            row["Ext. Tran. ID"] = RemoveLeadingZeros(fields[0].Trim());
            
            // Convert and validate date
            string dateStr = RemoveLeadingZeros(fields[1].Trim());
            string convertedDate = ConvertDateTime(dateStr);
            if (string.IsNullOrEmpty(convertedDate))
                continue;
            
            row["Tran. Date"] = convertedDate;
            row["Ext. Ref. Nbr."] = fields[2].Trim();
            row["Tran. Desc"] = fields[4].Trim();

            // Handle amount - split into Receipt/Disbursement
            string amountStr = fields[3].Trim().Replace(",", "");
            if (amountStr.StartsWith("+"))
            {
                row["Receipt"] = RemoveLeadingZeros(amountStr.Substring(1));
                row["Disbursement"] = "0.00";
            }
            else if (amountStr.StartsWith("-"))
            {
                row["Receipt"] = "0.00";
                row["Disbursement"] = RemoveLeadingZeros(amountStr.Substring(1));
            }
            else
            {
                // Default to receipt if no sign
                row["Receipt"] = RemoveLeadingZeros(amountStr);
                row["Disbursement"] = "0.00";
            }

            dt.Rows.Add(row);
            transactionId++;
        }
    }

    return dt;
}

// Add these methods to your #region Helper Methods section

private string RemoveLeadingZeros(string input)
{
    try
    {
        // Remove any leading zeros from the entire string
        dateStr = dateStr.TrimStart('0');

        // Handle different date formats
        if (dateStr.Length == 8 || dateStr.Length == 9)
        {
            // Format: YYYYMMDD or 0YYYYMMDD
            if (dateStr.Length == 9)
                dateStr = dateStr.Substring(1);

            string year = dateStr.Substring(0, 4);
            string month = dateStr.Substring(4, 2);
            string day = dateStr.Substring(6, 2);

            DateTime dt = DateTime.ParseExact($"{year}{month}{day}", "yyyyMMdd", null);
            return dt.ToString("yyyy-MM-dd");
        }
        else if (dateStr.Contains("/") || dateStr.Contains("-"))
        {
            string[] parts = dateStr.Split(new[] { '/', '-' });
            if (parts.Length == 3)
            {
                // Assume DD/MM/YYYY or DD-MM-YYYY format
                string day = parts[0].PadLeft(2, '0');
                string month = parts[1].PadLeft(2, '0');
                string year = parts[2].Length == 2 ? "20" + parts[2] : parts[2];

                DateTime dt = DateTime.ParseExact($"{year}{month}{day}", "yyyyMMdd", null);
                return dt.ToString("yyyy-MM-dd");
            }
        }

        // If no specific format matches, try general parsing
        DateTime generalDate = DateTime.Parse(dateStr);
        return generalDate.ToString("yyyy-MM-dd");
    }
    catch
    {
        return "";
    }
}

private string CleanNumericValue(string value)
{
    if (string.IsNullOrEmpty(value))
        return "0.00";

    // Remove currency symbols and spaces
    value = value.Replace("R", "")
                 .Replace("$", "")
                 .Replace(" ", "")
                 .Replace(",", "");

    // Handle parentheses for negative numbers
    if (value.StartsWith("(") && value.EndsWith(")"))
    {
        value = "-" + value.Trim('(', ')');
    }

    // Parse and format to ensure proper decimal places
    if (decimal.TryParse(value, out decimal numericValue))
    {
        return numericValue.ToString("0.00");
    }

    return "0.00";
}