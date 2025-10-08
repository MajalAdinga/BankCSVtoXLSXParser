using System;
using System.Drawing;
using System.Windows.Forms;

namespace BankCSVtoXLSXParser
{
    public sealed class ClosePromptForm : Form
    {
        private readonly string _message;

        public ClosePromptForm(string message = "CLOSE")
        {
            _message = string.IsNullOrEmpty(message) ? "CLOSE" : message;

            AutoScaleMode = AutoScaleMode.Font;
            Text = "Close";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            TopMost = true;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                ColumnCount = 1,
                RowCount = 3 // message | spacer | buttons
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 8F));   // small spacer
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));  // button row

            var lbl = new Label
            {
                Text = _message,
                AutoSize = true,
                Dock = DockStyle.None,
                Anchor = AnchorStyles.None,
                TextAlign = ContentAlignment.MiddleCenter,
                Padding = new Padding(2)
            };

            var spacer = new Panel { Dock = DockStyle.Fill };

            var btnRow = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0, 2, 0, 0)
            };
            btnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            btnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            btnRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Match original app button size (75 x 24)
            var yes = new Button
            {
                Text = "Yes",
                DialogResult = DialogResult.Yes,
                AutoSize = false,
                Size = new Size(75, 24),
                MinimumSize = new Size(75, 24),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Margin = new Padding(6),
                UseVisualStyleBackColor = true
            };
            var no = new Button
            {
                Text = "No",
                DialogResult = DialogResult.No,
                AutoSize = false,
                Size = new Size(75, 24),
                MinimumSize = new Size(75, 24),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Margin = new Padding(6),
                UseVisualStyleBackColor = true
            };

            btnRow.Controls.Add(yes, 0, 0);
            btnRow.Controls.Add(no, 1, 0);

            layout.Controls.Add(lbl, 0, 0);
            layout.Controls.Add(spacer, 0, 1);
            layout.Controls.Add(btnRow, 0, 2);
            Controls.Add(layout);

            AcceptButton = yes;
            CancelButton = no;

            MinimumSize = new Size(280, 150); // compact minimum
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            // Copy size (compact) and icon from owner (fallback to first open form)
            Form src = Owner as Form;
            if (src == null && Application.OpenForms.Count > 0)
                src = Application.OpenForms[0];

            if (src != null)
            {
                const double scale = 0.25;
                var cs = src.ClientSize;
                int w = (int)Math.Round(cs.Width * scale);
                int h = (int)Math.Round(cs.Height * scale);
                Size = new Size(Math.Max(MinimumSize.Width, w), Math.Max(MinimumSize.Height, h));

                try { Font = src.Font; } catch { }
                try { Icon = src.Icon; } catch { }
            }

            PerformLayout();
            Invalidate();
        }
    }
}