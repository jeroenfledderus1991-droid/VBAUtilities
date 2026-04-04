using System;
using System.Drawing;
using System.Windows.Forms;

namespace VBEAddIn
{
    /// <summary>
    /// Compacte "Wat is er nieuw?" notificatie die éénmalig verschijnt na een update.
    /// </summary>
    internal class WhatsNewForm : Form
    {
        private RichTextBox rtbEntries;
        private Button btnClose;
        private Button btnShowAll;
        private Label lblTitle;

        internal WhatsNewForm(ChangelogEntry entry)
        {
            InitializeForm(entry);
        }

        private void InitializeForm(ChangelogEntry entry)
        {
            this.Text = "VBE Code Tools — Wat is er nieuw?";
            this.Width = 460;
            this.Height = 300;
            this.StartPosition = FormStartPosition.Manual;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Positioneer rechts onderin het scherm
            Rectangle workArea = Screen.PrimaryScreen.WorkingArea;
            this.Left = workArea.Right - this.Width - 16;
            this.Top = workArea.Bottom - this.Height - 16;

            // Titel
            lblTitle = new Label();
            lblTitle.Text = "Nieuw in versie " + entry.Version + "  (" + entry.Date + ")";
            lblTitle.Left = 12;
            lblTitle.Top = 12;
            lblTitle.Width = this.Width - 28;
            lblTitle.Height = 20;
            lblTitle.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            lblTitle.ForeColor = Color.FromArgb(0, 80, 160);
            this.Controls.Add(lblTitle);

            // Entries
            rtbEntries = new RichTextBox();
            rtbEntries.Left = 12;
            rtbEntries.Top = 38;
            rtbEntries.Width = this.Width - 28;
            rtbEntries.Height = 190;
            rtbEntries.ReadOnly = true;
            rtbEntries.BorderStyle = BorderStyle.None;
            rtbEntries.BackColor = this.BackColor;
            rtbEntries.Font = new Font("Segoe UI", 9.5f);
            this.Controls.Add(rtbEntries);

            foreach (string line in entry.Lines)
            {
                AppendLine(line);
            }

            // Knoppen
            btnShowAll = new Button();
            btnShowAll.Text = "Alle versies bekijken";
            btnShowAll.Left = 12;
            btnShowAll.Top = 238;
            btnShowAll.Width = 150;
            btnShowAll.Height = 26;
            btnShowAll.Click += (s, e) => { new ChangelogForm().ShowDialog(); };
            this.Controls.Add(btnShowAll);

            btnClose = new Button();
            btnClose.Text = "Sluiten";
            btnClose.Left = this.Width - 96;
            btnClose.Top = 238;
            btnClose.Width = 72;
            btnClose.Height = 26;
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }

        private void AppendLine(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                rtbEntries.AppendText(Environment.NewLine);
                return;
            }

            string prefix = line.Length >= 1 ? line.Substring(0, 1) : "";
            Color color;
            switch (prefix)
            {
                case "+": color = Color.FromArgb(0, 120, 0);   break;
                case "*": color = Color.FromArgb(0, 80, 160);  break;
                case "-": color = Color.FromArgb(180, 0, 0);   break;
                default:  color = SystemColors.ControlText;    break;
            }

            int start = rtbEntries.TextLength;
            rtbEntries.AppendText(line + Environment.NewLine);
            rtbEntries.Select(start, line.Length);
            rtbEntries.SelectionColor = color;
            rtbEntries.SelectionLength = 0;
        }
    }
}
