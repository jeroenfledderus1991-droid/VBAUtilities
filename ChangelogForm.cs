using System;
using System.Drawing;
using System.Windows.Forms;

namespace VBEAddIn
{
    internal class ChangelogForm : Form
    {
        private ListBox lstVersions;
        private RichTextBox rtbDetails;
        private Label lblVersion;
        private Button btnClose;

        internal ChangelogForm()
        {
            InitializeForm();
            PopulateVersionList();
        }

        private void InitializeForm()
        {
            this.Text = "VBE Code Tools — Versiegeschiedenis";
            this.ClientSize = new Size(624, 388);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int margin = 12;
            int buttonHeight = 26;
            int buttonWidth = 80;
            int closeTop = this.ClientSize.Height - margin - buttonHeight;

            // Left panel: version list
            lstVersions = new ListBox();
            lstVersions.Left = margin;
            lstVersions.Top = margin;
            lstVersions.Width = 130;
            lstVersions.Height = closeTop - margin - 8;
            lstVersions.Font = new Font("Segoe UI", 9.5f);
            lstVersions.SelectedIndexChanged += OnVersionSelected;
            this.Controls.Add(lstVersions);

            // Version label above detail area
            lblVersion = new Label();
            lblVersion.Left = 156;
            lblVersion.Top = 12;
            lblVersion.Width = 460;
            lblVersion.Height = 20;
            lblVersion.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            this.Controls.Add(lblVersion);

            // Right panel: detail text
            rtbDetails = new RichTextBox();
            rtbDetails.Left = 156;
            rtbDetails.Top = 36;
            rtbDetails.Width = 460;
            rtbDetails.Height = closeTop - rtbDetails.Top - 8;
            rtbDetails.ReadOnly = true;
            rtbDetails.BorderStyle = BorderStyle.Fixed3D;
            rtbDetails.Font = new Font("Segoe UI", 9.5f);
            rtbDetails.BackColor = SystemColors.Window;
            this.Controls.Add(rtbDetails);

            // Close button
            btnClose = new Button();
            btnClose.Text = "Sluiten";
            btnClose.Width = buttonWidth;
            btnClose.Height = buttonHeight;
            btnClose.Left = this.ClientSize.Width - margin - btnClose.Width;
            btnClose.Top = closeTop;
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }

        private void PopulateVersionList()
        {
            foreach (ChangelogEntry entry in ChangelogData.Entries)
            {
                lstVersions.Items.Add("v" + entry.Version);
            }

            if (lstVersions.Items.Count > 0)
            {
                lstVersions.SelectedIndex = 0;
            }
        }

        private void OnVersionSelected(object sender, EventArgs e)
        {
            int idx = lstVersions.SelectedIndex;
            if (idx < 0 || idx >= ChangelogData.Entries.Length) return;

            ChangelogEntry entry = ChangelogData.Entries[idx];
            lblVersion.Text = "Versie " + entry.Version + "  —  " + entry.Date;

            rtbDetails.Clear();
            foreach (string line in entry.Lines)
            {
                AppendLine(line);
            }
        }

        private void AppendLine(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                rtbDetails.AppendText(Environment.NewLine);
                return;
            }

            string prefix = line.Length >= 1 ? line.Substring(0, 1) : "";
            Color color;
            switch (prefix)
            {
                case "+": color = Color.FromArgb(0, 120, 0);   break;  // groen
                case "*": color = Color.FromArgb(0, 80, 160);  break;  // blauw
                case "-": color = Color.FromArgb(180, 0, 0);   break;  // rood
                default:  color = SystemColors.ControlText;    break;
            }

            int start = rtbDetails.TextLength;
            rtbDetails.AppendText(line + Environment.NewLine);
            rtbDetails.Select(start, line.Length);
            rtbDetails.SelectionColor = color;
            rtbDetails.SelectionLength = 0;
        }
    }
}
