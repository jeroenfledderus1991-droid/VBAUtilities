using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace VBEAddIn
{
    internal sealed class UpdateReleaseForm : Form
    {
        private readonly string _latestVersion;
        private readonly List<GitHubReleaseInfo> _releases;
        private readonly Label _lblSummary;
        private readonly Label _lblCurrentVersion;
        private readonly ComboBox _cmbVersions;
        private readonly Label _lblPublished;
        private readonly TextBox _txtNotes;
        private readonly Button _btnDownload;
        private readonly Button _btnIgnore;
        private readonly Button _btnOpenRelease;

        internal GitHubReleaseInfo SelectedRelease
        {
            get { return _cmbVersions.SelectedItem as GitHubReleaseInfo; }
        }

        internal bool IgnoreLatestVersion { get; private set; }

        internal UpdateReleaseForm(string currentVersion, string latestVersion, IEnumerable<GitHubReleaseInfo> releases, bool hasNewerVersion)
        {
            _latestVersion = latestVersion ?? string.Empty;
            _releases = (releases ?? Enumerable.Empty<GitHubReleaseInfo>()).ToList();

            Text = "VBE Code Tools Update";
            ClientSize = new Size(720, 610);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 96,
                BackColor = ColorTranslator.FromHtml("#0F172A")
            };
            Controls.Add(header);

            Label lblTitle = new Label
            {
                Text = "Update beschikbaar",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 19, FontStyle.Bold),
                AutoSize = false,
                Size = new Size(660, 34),
                Location = new Point(24, 16),
                BackColor = Color.Transparent
            };
            header.Controls.Add(lblTitle);

            _lblSummary = new Label
            {
                Text = hasNewerVersion
                    ? "Kies welke release je wilt downloaden. Je kunt ook een oudere release selecteren."
                    : "Je gebruikt al de nieuwste versie, maar je kunt nog steeds een andere release kiezen.",
                ForeColor = Color.FromArgb(219, 234, 254),
                Font = new Font("Segoe UI", 10),
                AutoSize = false,
                Size = new Size(660, 40),
                Location = new Point(24, 50),
                BackColor = Color.Transparent
            };
            header.Controls.Add(_lblSummary);

            _lblCurrentVersion = new Label
            {
                Text = "Geinstalleerd: " + currentVersion + "    Nieuwste release: " + latestVersion,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#0F172A"),
                AutoSize = false,
                Size = new Size(660, 24),
                Location = new Point(24, 118)
            };
            Controls.Add(_lblCurrentVersion);

            Label lblSelect = new Label
            {
                Text = "Beschikbare releases",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#334155"),
                AutoSize = false,
                Size = new Size(200, 22),
                Location = new Point(24, 156)
            };
            Controls.Add(lblSelect);

            _cmbVersions = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 10),
                Location = new Point(24, 182),
                Size = new Size(420, 28)
            };
            _cmbVersions.SelectedIndexChanged += CmbVersions_SelectedIndexChanged;
            Controls.Add(_cmbVersions);

            _btnOpenRelease = new Button
            {
                Text = "Releasepagina",
                Font = new Font("Segoe UI", 9),
                Size = new Size(120, 30),
                Location = new Point(458, 180),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnOpenRelease.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#CBD5E1");
            _btnOpenRelease.Click += BtnOpenRelease_Click;
            Controls.Add(_btnOpenRelease);

            _lblPublished = new Label
            {
                Font = new Font("Segoe UI", 9),
                ForeColor = ColorTranslator.FromHtml("#64748B"),
                AutoSize = false,
                Size = new Size(660, 22),
                Location = new Point(24, 220)
            };
            Controls.Add(_lblPublished);

            Label lblNotes = new Label
            {
                Text = "Release-opmerkingen",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#334155"),
                AutoSize = false,
                Size = new Size(200, 22),
                Location = new Point(24, 252)
            };
            Controls.Add(lblNotes);

            _txtNotes = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 10),
                Location = new Point(24, 278),
                Size = new Size(660, 230),
                BackColor = ColorTranslator.FromHtml("#F8FAFC"),
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(_txtNotes);

            _btnDownload = new Button
            {
                Text = "Download installer",
                DialogResult = DialogResult.OK,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new Size(170, 38),
                Location = new Point(24, 530),
                BackColor = ColorTranslator.FromHtml("#2563EB"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnDownload.FlatAppearance.BorderSize = 0;
            Controls.Add(_btnDownload);

            _btnIgnore = new Button
            {
                Text = "Negeer nieuwste versie",
                Font = new Font("Segoe UI", 10),
                Size = new Size(180, 38),
                Location = new Point(208, 530),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnIgnore.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#CBD5E1");
            _btnIgnore.Click += BtnIgnore_Click;
            Controls.Add(_btnIgnore);

            Button btnLater = new Button
            {
                Text = "Later",
                DialogResult = DialogResult.Cancel,
                Font = new Font("Segoe UI", 10),
                Size = new Size(110, 38),
                Location = new Point(574, 530),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnLater.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#CBD5E1");
            Controls.Add(btnLater);

            AcceptButton = _btnDownload;
            CancelButton = btnLater;

            foreach (GitHubReleaseInfo release in _releases)
            {
                _cmbVersions.Items.Add(release);
            }

            if (_cmbVersions.Items.Count > 0)
            {
                int defaultIndex = _releases.FindIndex(r => string.Equals(r.Version, _latestVersion, StringComparison.OrdinalIgnoreCase));
                _cmbVersions.SelectedIndex = defaultIndex >= 0 ? defaultIndex : 0;
            }

            _btnDownload.Enabled = _cmbVersions.Items.Count > 0;
            _btnIgnore.Enabled = !string.IsNullOrWhiteSpace(_latestVersion);
        }

        private void CmbVersions_SelectedIndexChanged(object sender, EventArgs e)
        {
            GitHubReleaseInfo release = SelectedRelease;
            if (release == null)
            {
                _lblPublished.Text = string.Empty;
                _txtNotes.Text = string.Empty;
                return;
            }

            _lblPublished.Text = "Publicatiedatum: " + release.PublishedDisplay + "    Bestandsnaam: " + release.InstallerFileName;
            _txtNotes.Text = string.IsNullOrWhiteSpace(release.Body)
                ? "Geen release-opmerkingen beschikbaar."
                : release.Body.Trim();
        }

        private void BtnIgnore_Click(object sender, EventArgs e)
        {
            IgnoreLatestVersion = true;
            DialogResult = DialogResult.No;
            Close();
        }

        private void BtnOpenRelease_Click(object sender, EventArgs e)
        {
            GitHubReleaseInfo release = SelectedRelease;
            if (release == null || string.IsNullOrWhiteSpace(release.ReleaseUrl))
            {
                return;
            }

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = release.ReleaseUrl,
                UseShellExecute = true
            });
        }
    }
}
