using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using BankCSVtoXLSXParser.Parsers; // ADDED

namespace BankCSVtoXLSXParser
{
    public partial class frmBankCSVtoXLSXParser : Form
    {
        /// <summary>
        /// Tracks whether the last click on the textbox was a double-click to differentiate
        /// between single and double click behaviors.
        /// </summary>
        private bool isDoubleClick = false;

        /// <summary>
        /// Stores the last recorded error (location + message) for user friendly display after exceptions.
        /// </summary>
        private string lastError = "";

        private const string SimpleLogFileName = "error_log.txt";
        private const string VerboseLogFileName = "error_log_verbose.txt";
        private bool verboseLogging = false;

        /// <summary>
        /// Initializes the form, its components, and performs initial layout and control setup.
        /// Any exception encountered is logged and surfaced to the user.
        /// </summary>
        public frmBankCSVtoXLSXParser()
        {
            try
            {
                InitializeComponent();
                SetupForm();
            }
            catch (Exception ex)
            {
                LogError("Constructor", ex);
                MessageBox.Show($"Error initializing application: {ex.Message}", "Startup Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Performs one-time UI initialization: window title, control behaviors, anchoring, and sizing.
        /// </summary>
        private void SetupForm()
        {
            try
            {
                Version version = Assembly.GetExecutingAssembly().GetName().Version;
                this.Text = $"Bank CSV to XLSX Parser - V{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
                cmbBank.DropDownStyle = ComboBoxStyle.DropDownList;

                txtFilePath.ReadOnly = false;
                txtFilePath.BackColor = Color.White;
                txtFilePath.Cursor = Cursors.IBeam;

                txtFilePath.Click += TxtFilePath_Click;
                txtFilePath.DoubleClick += TxtFilePath_DoubleClick;

                this.MaximumSize = new Size(2000, this.Height);
                this.MinimumSize = new Size(369, this.Height);

                txtFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                cmbBank.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                btnLoad.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                btnConvert.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

                btnLoad.Location = new Point(12, this.ClientSize.Height - btnLoad.Height - 12);
                btnConvert.Location = new Point(this.ClientSize.Width - btnConvert.Width - 12,
                    this.ClientSize.Height - btnConvert.Height - 12);
            }
            catch (Exception ex)
            {
                LogError("SetupForm", ex);
            }
        }

        #region Event Handlers

        /// <summary>
        /// Handles form load; sets initial enabled/disabled state of controls based on bank selection.
        /// </summary>
        private void frmBankCSVtoXLSXParser_Load(object sender, EventArgs e)
        {
            try
            {
                if (cmbBank.SelectedIndex == -1 && cmbBank.Items.Count > 0)
                {
                    int targetIndex = -1;

                    // Prefer ABSA
                    for (int i = 0; i < cmbBank.Items.Count; i++)
                    {
                        var text = cmbBank.Items[i]?.ToString();
                        if (!string.IsNullOrEmpty(text) && text.IndexOf("ABSA", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            targetIndex = i;
                            break;
                        }
                    }

                    // If no ABSA, fall back to FNB, else first item
                    if (targetIndex == -1)
                    {
                        for (int i = 0; i < cmbBank.Items.Count; i++)
                        {
                            var text = cmbBank.Items[i]?.ToString();
                            if (!string.IsNullOrEmpty(text) &&
                                (text.IndexOf("FNB", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 text.IndexOf("First National", StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                targetIndex = i;
                                break;
                            }
                        }
                    }

                    cmbBank.SelectedIndex = targetIndex >= 0 ? targetIndex : 0;
                }

                if (cmbBank.SelectedItem != null && cmbBank.SelectedItem.ToString() != "")
                {
                    txtFilePath.Enabled = true;
                    btnLoad.Enabled = true;
                    btnConvert.Enabled = !string.IsNullOrEmpty(txtFilePath.Text);
                }
                else
                {
                    btnConvert.Enabled = false;
                    txtFilePath.Enabled = false;
                    btnLoad.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                LogError("Form_Load", ex);
            }
        }

        /// <summary>
        /// Enables file path selection and convert button when a valid bank selection is made.
        /// </summary>
        private void cmbBank_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbBank.SelectedItem != null && cmbBank.SelectedItem.ToString() != "")
                {
                    btnConvert.Enabled = !string.IsNullOrEmpty(txtFilePath.Text);
                    txtFilePath.Enabled = true;
                    btnLoad.Enabled = true;
                }
                else
                {
                    btnConvert.Enabled = false;
                    txtFilePath.Enabled = false;
                    btnLoad.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                LogError("Bank_SelectedIndexChanged", ex);
            }
        }

        /// <summary>
        /// Enables the convert button when a file path becomes available and a bank is selected.
        /// </summary>
        private void txtFilePath_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbBank.SelectedItem != null && !string.IsNullOrEmpty(txtFilePath.Text))
                {
                    btnConvert.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                LogError("FilePath_TextChanged", ex);
            }
        }

        /// <summary>
        /// Single click flag reset for distinguishing single from double-click.
        /// </summary>
        private void TxtFilePath_Click(object sender, EventArgs e)
        {
            isDoubleClick = false;
        }

        /// <summary>
        /// Opens file selection dialog on double-click when enabled.
        /// </summary>
        private void TxtFilePath_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                isDoubleClick = true;
                if (txtFilePath.Enabled)
                {
                    LoadFile();
                }
            }
            catch (Exception ex)
            {
                LogError("FilePath_DoubleClick", ex);
            }
        }

        /// <summary>
        /// Handles explicit file load button click.
        /// </summary>
        private void btnLoad_Click(object sender, EventArgs e)
        {
            try
            {
                LoadFile();
            }
            catch (Exception ex)
            {
                LogError("Load_Click", ex);
            }
        }

        /// <summary>
        /// Opens a file open dialog allowing the user to pick a CSV/TXT file
        /// and updates the file path textbox. Enables conversion if bank selected.
        /// </summary>
        private void LoadFile()
        {
            try
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt";
                openFileDialog.Title = "Select a CSV or TXT file";
                openFileDialog.FileName = "";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;

                    if (cmbBank.SelectedItem != null)
                    {
                        btnConvert.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("LoadFile", ex);
                MessageBox.Show("Failed to load file.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtFilePath.Text))
                {
                    MessageBox.Show("Please select a file first.", "Validation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (cmbBank.SelectedItem == null)
                {
                    MessageBox.Show("Please select a bank first.", "Validation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!File.Exists(txtFilePath.Text))
                {
                    MessageBox.Show("The selected file does not exist.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var parser = BankParserFactory.Resolve(cmbBank.SelectedItem?.ToString() ?? "", txtFilePath.Text);
                string bankShortName = parser.ShortName; // ensures colors/file name match detected bank

                SaveFileDialog saveDialog = new SaveFileDialog();
                string originalName = Path.GetFileNameWithoutExtension(txtFilePath.Text);
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.Title = "Save Excel file as";
                saveDialog.FileName = $"{originalName}_{bankShortName}_converted.xlsx";
                saveDialog.InitialDirectory = Path.GetDirectoryName(txtFilePath.Text);

                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                string outputPath = saveDialog.FileName;

                this.Cursor = Cursors.WaitCursor;
                btnConvert.Enabled = false;

                try
                {
                    DataTable dt = parser.Parse(txtFilePath.Text); // use the same parser
                    if (dt == null || dt.Rows.Count == 0)
                    {
                        MessageBox.Show("No data found in the file.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    await System.Threading.Tasks.Task.Run(() =>
                    {
                        WriteExcelWithBankColors(dt, outputPath, bankShortName); // ABSA colors applied here
                    });

                    // Success then CLOSE prompt
                    MessageBox.Show("The file has been converted successfully.", "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    using (var dlg = new ClosePromptForm("Close the app now?"))
                    {
                        var result = dlg.ShowDialog(this);
                        if (result == DialogResult.Yes)
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                try { this.Close(); } catch { }
                                try { Application.Exit(); } catch { }
                            }));
                        }
                        // If No: leave the app open so the user can continue converting files.
                    }
                }
                catch (Exception ex)
                {
                    LogError("Convert_Processing", ex);
                    throw;
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                    btnConvert.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                Application.UseWaitCursor = false;
                this.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                btnConvert.Enabled = true;

                LogError("Convert_Click", ex);
                MessageBox.Show("Conversion failed.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helper Methods

        private string GetBankShortName(string bankName)
        {
            if (bankName.IndexOf("ABSA", StringComparison.OrdinalIgnoreCase) >= 0)
                return "ABSA";
            if (bankName.IndexOf("FNB", StringComparison.OrdinalIgnoreCase) >= 0 ||
                bankName.IndexOf("First National", StringComparison.OrdinalIgnoreCase) >= 0)
                return "FNB";
            if (bankName.IndexOf("Standard", StringComparison.OrdinalIgnoreCase) >= 0)
                return "StandardBank";
            return "Bank";
        }

        private DataTable ParseBankFile(string filePath, string bankName)
        {
            try
            {
                var parser = BankParserFactory.Resolve(bankName, filePath);
                return parser.Parse(filePath);
            }
            catch (Exception ex)
            {
                LogError("ParseBankFile", ex);
                throw;
            }
        }

        private void WriteExcelWithBankColors(DataTable dt, string outputPath, string bankName)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;

            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                xlWorkBook = xlApp.Workbooks.Add();
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[1];
                xlWorkSheet.Name = $"{bankName} Transactions";

                Color headerColor;
                Color alternateRowColor;

                if (string.Equals(bankName, "FNB", StringComparison.OrdinalIgnoreCase))
                {
                    headerColor = Color.FromArgb(0, 168, 168);
                    alternateRowColor = Color.FromArgb(230, 250, 250);
                }
                else if (string.Equals(bankName, "StandardBank", StringComparison.OrdinalIgnoreCase))
                {
                    headerColor = Color.FromArgb(0, 83, 159);
                    alternateRowColor = Color.FromArgb(230, 240, 255);
                }
                else if (string.Equals(bankName, "ABSA", StringComparison.OrdinalIgnoreCase)) // ADDED ABSA
                {
                    // ABSA brand red and a soft rose alternate row
                    headerColor = Color.FromArgb(186, 12, 47);     // #BA0C2F
                    alternateRowColor = Color.FromArgb(255, 240, 244);
                }
                else
                {
                    headerColor = Color.FromArgb(68, 114, 196);
                    alternateRowColor = Color.FromArgb(242, 242, 242);
                }

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Excel.Range headerCell = xlWorkSheet.Cells[1, i + 1];
                    headerCell.Value = dt.Columns[i].ColumnName;
                    headerCell.Font.Bold = true;
                    headerCell.Font.Color = ColorTranslator.ToOle(Color.White);
                    headerCell.Interior.Color = ColorTranslator.ToOle(headerColor);
                    headerCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }

                xlWorkSheet.Rows[1].RowHeight = 30;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        xlWorkSheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    }

                    if (i % 2 == 1)
                    {
                        Excel.Range rowRange = xlWorkSheet.Range[
                            xlWorkSheet.Cells[i + 2, 1],
                            xlWorkSheet.Cells[i + 2, dt.Columns.Count]
                        ];
                        rowRange.Interior.Color = ColorTranslator.ToOle(alternateRowColor);
                    }
                }

                xlWorkSheet.Columns.AutoFit();

                Excel.Range dataRange = xlWorkSheet.UsedRange;
                dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                dataRange.Borders.Color = ColorTranslator.ToOle(Color.LightGray);

                try
                {
                    if (xlWorkSheet.AutoFilterMode)
                        xlWorkSheet.AutoFilterMode = false;

                    Excel.Range headerRow = xlWorkSheet.Range[
                        xlWorkSheet.Cells[1, 1],
                        xlWorkSheet.Cells[1, dt.Columns.Count]
                    ];
                    headerRow.AutoFilter();
                }
                catch { }

                try
                {
                    xlWorkSheet.Application.ActiveWindow.SplitRow = 1;
                    xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
                }
                catch { }

                xlWorkBook.SaveAs(outputPath);
                xlWorkBook.Close(false);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                LogError("WriteExcel", ex);
                throw;
            }
            finally
            {
                try
                {
                    if (xlWorkSheet != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
                    if (xlWorkBook != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                    if (xlApp != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    LogError("Excel_Cleanup", ex);
                }
            }
        }

        #endregion

        #region Error Logging

        private void LogError(string location, Exception ex)
        {
            lastError = $"{location}: {ex.Message}";
            System.Diagnostics.Debug.WriteLine($"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {lastError}");
            try
            {
                string logFile = Path.Combine(Application.StartupPath, "error_log.txt");
                File.AppendAllText(logFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {lastError}\n{ex.StackTrace}\n\n");
            }
            catch { }
        }

        #endregion

        // --- Added missing FNB helper methods to resolve CS0103 references ---

        private bool IsDateValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;
            string v = value.Trim();
            if (v.Contains("/") || v.Contains("-"))
            {
                DateTime dt;
                if (DateTime.TryParse(v, out dt))
                    return true;
            }
            if (v.Length == 8 && v.All(char.IsDigit) && v.StartsWith("20"))
                return true;
            DateTime parsed;
            return DateTime.TryParse(v, out parsed);
        }

        private string ConvertFNBDate(string dateStr)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dateStr))
                    return "";
                DateTime date;
                string[] formats = {
                    "yyyy/MM/dd","yyyy-MM-dd","dd/MM/yyyy","dd-MM-yyyy","dd MMM yyyy",
                    "dd MMMM yyyy","yyyy/MM/dd HH:mm","dd/MM/yyyy HH:mm","yyyyMMdd"
                };
                if (DateTime.TryParseExact(dateStr.Trim(), formats,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out date))
                    return date.ToString("yyyy-MM-dd");
                if (DateTime.TryParse(dateStr, out date))
                    return date.ToString("yyyy-MM-dd");
                return dateStr.Trim();
            }
            catch (Exception ex)
            {
                LogError("ConvertFNBDate", ex);
                return dateStr;
            }
        }

        // ADDED: local FNB decimal parser used by description/amount inference
        private bool TryParseFNBDecimal(string raw, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            string cleaned = CleanFNBAmountString(raw);

            decimal parsed;
            if (decimal.TryParse(cleaned, System.Globalization.NumberStyles.Number,
                System.Globalization.CultureInfo.InvariantCulture, out parsed))
            {
                value = parsed;
                return true;
            }
            if (decimal.TryParse(cleaned, out parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }

        // ADDED: minimal amount string cleaner (currency, signs, CR/DR, separators)
        private string CleanFNBAmountString(string amountStr)
        {
            if (string.IsNullOrWhiteSpace(amountStr))
                return "";

            string a = amountStr.Trim();

            a = a.Replace("R", "").Replace("r", "").Replace("$", "").Trim();

            bool neg = false;

            if (a.StartsWith("(") && a.EndsWith(")"))
            {
                neg = true;
                a = a.Substring(1, a.Length - 2);
            }

            if (a.EndsWith("-"))
            {
                neg = true;
                a = a.Substring(0, a.Length - 1);
            }

            if (a.EndsWith("CR", StringComparison.OrdinalIgnoreCase))
            {
                a = a.Substring(0, a.Length - 2).Trim();
            }
            else if (a.EndsWith("DR", StringComparison.OrdinalIgnoreCase))
            {
                a = a.Substring(0, a.Length - 2).Trim();
                neg = true;
            }

            if (a.StartsWith("+"))
                a = a.Substring(1);
            else if (a.StartsWith("-"))
            {
                neg = true;
                a = a.Substring(1);
            }

            if (a.IndexOf(',') >= 0 && a.IndexOf('.') >= 0)
                a = a.Replace(",", "");
            else if (a.Count(c => c == ',') > 1)
                a = a.Replace(",", "");
            else if (a.Count(c => c == ',') == 1 && !a.Contains("."))
                a = a.Replace(",", ".");

            a = a.Replace(" ", "");

            if (neg)
                a = "-" + a;

            return a;
        }

        private string BuildFNBDescription(string[] values, int dateIdx, int amountIdx,
            int creditIdx, int debitIdx, int referenceIdx)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < values.Length; i++)
            {
                if (i == dateIdx || i == amountIdx || i == creditIdx || i == debitIdx || i == referenceIdx)
                    continue;
                if (string.IsNullOrWhiteSpace(values[i]))
                    continue;
                decimal d;
                if (TryParseFNBDecimal(values[i], out d))
                    continue;
                if (IsDateValue(values[i]))
                    continue;
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(values[i].Trim());
            }
            return sb.ToString().Trim();
        }

        private string ExtractFNBReference(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";
            var startNumMatch = System.Text.RegularExpressions.Regex.Match(text, @"^(\d{4,})");
            if (startNumMatch.Success) return startNumMatch.Groups[1].Value;

            var refMatch = System.Text.RegularExpressions.Regex.Match(text,
                @"(?:REF|Ref|REFERENCE)[\s:#]*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (refMatch.Success) return refMatch.Groups[1].Value;

            var trnMatch = System.Text.RegularExpressions.Regex.Match(text,
                @"(?:TRN|TRANS|TRANSACTION)[\s:#]*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (trnMatch.Success) return trnMatch.Groups[1].Value;

            var pmtMatch = System.Text.RegularExpressions.Regex.Match(text,
                @"(?:PMT|PAYMENT)[\s:#]*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (pmtMatch.Success) return pmtMatch.Groups[1].Value;

            var numMatch = System.Text.RegularExpressions.Regex.Match(text, @"\b(\d{5,})\b");
            if (numMatch.Success) return numMatch.Groups[1].Value;

            return "";
        }

        private bool TryFindSingleAmount(string[] values, out decimal amount)
        {
            amount = 0m;
            var candidates = new List<decimal>();
            foreach (var v in values)
            {
                decimal tmp;
                if (TryParseFNBDecimal(v, out tmp))
                    candidates.Add(tmp);
            }
            if (candidates.Count == 1)
            {
                amount = candidates[0];
                return true;
            }
            return false;
        }
    }
}