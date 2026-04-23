using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VBEAddInInstaller
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm());
        }
    }

    public class InstallerForm : Form
    {
        private static readonly string[] OfficeProcessNames = { "EXCEL", "WINWORD", "POWERPNT", "OUTLOOK", "MSACCESS", "ONENOTE", "VISIO" };
        private static readonly string InstallPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "VBE AddIn");
        private static readonly string VBEAddInDll = "VBEAddIn.dll";
        private static readonly string VBEAddInPath = Path.Combine(InstallPath, VBEAddInDll);
        private static readonly string RegAsmPath = Path.Combine(Environment.GetEnvironmentVariable("windir"), @"Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe");

        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblVersion;
        private Button btnInstall;
        private Button btnUninstall;
        private Button btnClose;
        private ProgressBar progressBar;
        private Label lblStatus;

        public InstallerForm()
        {
            InitializeUI();
            CheckAdminRights();
        }

        private void InitializeUI()
        {
            Text = "VBE Code Tools - Setup";
            Size = new Size(580, 430);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;

            Panel headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 104,
                BackColor = ColorTranslator.FromHtml("#0F172A")
            };
            Controls.Add(headerPanel);

            lblTitle = new Label
            {
                Text = "VBE Code Tools",
                Font = new Font("Segoe UI", 20, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                Size = new Size(320, 42),
                Location = new Point(20, 18),
                BackColor = Color.Transparent
            };
            headerPanel.Controls.Add(lblTitle);

            lblSubtitle = new Label
            {
                Text = "Setup voor de VBA Editor add-in",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(191, 219, 254),
                AutoSize = false,
                Size = new Size(320, 24),
                Location = new Point(20, 62),
                BackColor = Color.Transparent
            };
            headerPanel.Controls.Add(lblSubtitle);

            lblVersion = new Label
            {
                Text = "Versie " + GetBundledVersion(),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                Size = new Size(180, 24),
                Location = new Point(360, 24),
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = Color.Transparent
            };
            headerPanel.Controls.Add(lblVersion);

            btnInstall = new Button
            {
                Text = "Installeren of bijwerken",
                Size = new Size(220, 50),
                Location = new Point(40, 136),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                BackColor = ColorTranslator.FromHtml("#2563EB"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;
            Controls.Add(btnInstall);

            btnUninstall = new Button
            {
                Text = "Verwijderen",
                Size = new Size(220, 50),
                Location = new Point(290, 136),
                Font = new Font("Segoe UI", 12),
                BackColor = ColorTranslator.FromHtml("#DC2626"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnUninstall.FlatAppearance.BorderSize = 0;
            btnUninstall.Click += BtnUninstall_Click;
            Controls.Add(btnUninstall);

            progressBar = new ProgressBar
            {
                Size = new Size(500, 24),
                Location = new Point(30, 220),
                Visible = false,
                Style = ProgressBarStyle.Continuous
            };
            Controls.Add(progressBar);

            lblStatus = new Label
            {
                Text = "Klaar om te installeren",
                Size = new Size(500, 74),
                Location = new Point(30, 256),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.TopLeft
            };
            Controls.Add(lblStatus);

            btnClose = new Button
            {
                Text = "Sluiten",
                Size = new Size(110, 36),
                Location = new Point(420, 346),
                Font = new Font("Segoe UI", 10),
                BackColor = Color.WhiteSmoke,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderColor = Color.Gray;
            btnClose.Click += (s, e) => Close();
            Controls.Add(btnClose);
        }

        private void CheckAdminRights()
        {
            if (!IsRunningAsAdmin())
            {
                lblStatus.Text = "Deze installer moet als administrator worden uitgevoerd.\n\nRechtsklik op het bestand en kies 'Als administrator uitvoeren'.";
                lblStatus.ForeColor = Color.Red;
                btnInstall.Enabled = false;
                btnUninstall.Enabled = false;
            }
            else if (!File.Exists(RegAsmPath))
            {
                lblStatus.Text = ".NET Framework 4.8.x is niet gevonden.\n\nInstalleer eerst het juiste Developer Pack en probeer daarna opnieuw.";
                lblStatus.ForeColor = Color.Red;
                btnInstall.Enabled = false;
                btnUninstall.Enabled = false;
            }
        }

        private void BtnInstall_Click(object sender, EventArgs e)
        {
            SetBusyState(true);

            try
            {
                if (!EnsureOfficeApplicationsClosed("installeren of bijwerken"))
                {
                    progressBar.Visible = false;
                    lblStatus.Text = "Installatie geannuleerd. Sluit eerst alle Office-applicaties en probeer het daarna opnieuw.";
                    lblStatus.ForeColor = Color.DarkOrange;
                    return;
                }

                if (Directory.Exists(InstallPath) || File.Exists(VBEAddInPath))
                {
                    UpdateStatus("Bestaande installatie verwijderen...", 10);
                    try
                    {
                        UnregisterCOM();
                        RemoveRegistryEntries();
                        if (Directory.Exists(InstallPath))
                        {
                            Directory.Delete(InstallPath, true);
                        }
                    }
                    catch
                    {
                    }
                }

                UpdateStatus("Installatiemap voorbereiden...", 25);
                if (!Directory.Exists(InstallPath))
                {
                    Directory.CreateDirectory(InstallPath);
                }

                UpdateStatus("Add-in bestanden uitpakken...", 45);
                ExtractEmbeddedDll();

                UpdateStatus("COM-registratie uitvoeren...", 70);
                RegisterCOM();

                UpdateStatus("VBE add-in registreren...", 88);
                CreateRegistryEntries();

                UpdateStatus("Installatie afronden...", 100);
                lblStatus.ForeColor = Color.Green;

                MessageBox.Show(
                    "VBE Code Tools is succesvol geinstalleerd.\n\nStart Office opnieuw om de nieuwe versie te laden.\nOpen daarna de VBA Editor met Alt+F11.",
                    "Installatie voltooid",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                lblStatus.Text = "Fout tijdens installatie:\n" + GetFriendlyErrorMessage(ex);
                lblStatus.ForeColor = Color.Red;

                MessageBox.Show(
                    "Installatie mislukt:\n\n" + GetFriendlyErrorMessage(ex),
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetBusyState(false);
            }
        }

        private void BtnUninstall_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Weet je zeker dat je VBE Code Tools wilt verwijderen?",
                "Bevestigen",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
            {
                return;
            }

            SetBusyState(true);

            try
            {
                if (!EnsureOfficeApplicationsClosed("verwijderen"))
                {
                    progressBar.Visible = false;
                    lblStatus.Text = "Verwijderen geannuleerd. Sluit eerst alle Office-applicaties en probeer het daarna opnieuw.";
                    lblStatus.ForeColor = Color.DarkOrange;
                    return;
                }

                UpdateStatus("COM-registratie verwijderen...", 30);
                UnregisterCOM();

                UpdateStatus("Registerverwijzingen verwijderen...", 60);
                RemoveRegistryEntries();

                UpdateStatus("Bestanden verwijderen...", 90);
                if (Directory.Exists(InstallPath))
                {
                    Directory.Delete(InstallPath, true);
                }

                UpdateStatus("Verwijdering voltooid.", 100);
                lblStatus.ForeColor = Color.Green;

                MessageBox.Show(
                    "VBE Code Tools is succesvol verwijderd.",
                    "Verwijdering voltooid",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                lblStatus.Text = "Fout tijdens verwijderen:\n" + GetFriendlyErrorMessage(ex);
                lblStatus.ForeColor = Color.Red;

                MessageBox.Show(
                    "Verwijderen mislukt:\n\n" + GetFriendlyErrorMessage(ex),
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetBusyState(false);
            }
        }

        private void SetBusyState(bool isBusy)
        {
            btnInstall.Enabled = !isBusy;
            btnUninstall.Enabled = !isBusy;
            btnClose.Enabled = !isBusy;
            progressBar.Visible = isBusy;
            if (isBusy)
            {
                progressBar.Value = 0;
            }
        }

        private void UpdateStatus(string message, int progress)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = Color.Gray;
            progressBar.Value = progress;
            Application.DoEvents();
            System.Threading.Thread.Sleep(250);
        }

        private void ExtractEmbeddedDll()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream("VBEAddInInstaller.VBEAddIn.dll"))
            {
                if (stream == null)
                {
                    throw new Exception("VBEAddIn.dll niet gevonden in installer.");
                }

                using (FileStream fileStream = new FileStream(VBEAddInPath, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
            }
        }

        private void RegisterCOM()
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = RegAsmPath,
                Arguments = string.Format("\"{0}\" /codebase /tlb", VBEAddInPath),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true
            };

            using (Process process = Process.Start(psi))
            {
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    string error = process.StandardError.ReadToEnd();
                    throw new Exception("Registreren van de add-in is mislukt. " + error.Trim());
                }
            }
        }

        private void UnregisterCOM()
        {
            if (!File.Exists(VBEAddInPath))
            {
                return;
            }

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = RegAsmPath,
                Arguments = string.Format("\"{0}\" /unregister", VBEAddInPath),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true
            };

            using (Process process = Process.Start(psi))
            {
                process.WaitForExit();
            }
        }

        private void CreateRegistryEntries()
        {
            string[] vbeVersions = { "6.0", "7.0", "7.1" };
            string[] addinPaths = { "Addins", "AddIns64" };

            foreach (string version in vbeVersions)
            {
                foreach (string addinPath in addinPaths)
                {
                    string regPath = string.Format(@"Software\Microsoft\VBA\VBE\{0}\{1}\VBEAddIn.Connect", version, addinPath);
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(regPath))
                    {
                        if (key == null)
                        {
                            continue;
                        }

                        key.SetValue("FriendlyName", "VBE Code Tools");
                        key.SetValue("Description", "VBA Editor Add-in voor code formattering");
                        key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                        key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                    }
                }
            }
        }

        private void RemoveRegistryEntries()
        {
            string[] vbeVersions = { "6.0", "7.0", "7.1" };
            string[] addinPaths = { "Addins", "AddIns64" };

            foreach (string version in vbeVersions)
            {
                foreach (string addinPath in addinPaths)
                {
                    try
                    {
                        string regPath = string.Format(@"Software\Microsoft\VBA\VBE\{0}\{1}\VBEAddIn.Connect", version, addinPath);
                        Registry.CurrentUser.DeleteSubKey(regPath, false);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static bool IsRunningAsAdmin()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static string GetBundledVersion()
        {
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            return version == null ? "onbekend" : string.Format("{0}.{1}.{2}", version.Major, version.Minor, version.Build);
        }

        private bool EnsureOfficeApplicationsClosed(string actionDescription)
        {
            List<Process> runningApps = GetRunningOfficeProcesses();
            if (runningApps.Count == 0)
            {
                return true;
            }

            using (OfficeCloseForm prompt = new OfficeCloseForm(runningApps, actionDescription))
            {
                DialogResult result = prompt.ShowDialog(this);
                if (result == DialogResult.Cancel)
                {
                    return false;
                }

                if (result == DialogResult.Yes)
                {
                    CloseOfficeApplications(runningApps);
                }
            }

            while (GetRunningOfficeProcesses().Count > 0)
            {
                DialogResult retry = MessageBox.Show(
                    "Er zijn nog Office-applicaties actief. Sluit deze volledig af en klik daarna op Opnieuw proberen.",
                    "Office nog actief",
                    MessageBoxButtons.RetryCancel,
                    MessageBoxIcon.Information);

                if (retry != DialogResult.Retry)
                {
                    return false;
                }
            }

            return true;
        }

        private static List<Process> GetRunningOfficeProcesses()
        {
            return Process.GetProcesses()
                .Where(process => OfficeProcessNames.Contains(process.ProcessName, StringComparer.OrdinalIgnoreCase))
                .OrderBy(process => process.ProcessName)
                .ToList();
        }

        private static void CloseOfficeApplications(IEnumerable<Process> processes)
        {
            foreach (Process process in processes)
            {
                try
                {
                    if (!process.HasExited && process.CloseMainWindow())
                    {
                        process.WaitForExit(3000);
                    }

                    if (!process.HasExited)
                    {
                        process.Kill();
                        process.WaitForExit(3000);
                    }
                }
                catch
                {
                }
            }
        }

        private static string GetFriendlyErrorMessage(Exception ex)
        {
            string message = ex == null ? string.Empty : ex.Message;
            if (message.IndexOf("VBEAddIn.dll", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("wordt gebruikt door een ander proces", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "De add-in is nog in gebruik. Sluit eerst alle Office-applicaties en probeer het opnieuw.";
            }

            return message;
        }
    }

    internal sealed class OfficeCloseForm : Form
    {
        internal OfficeCloseForm(IEnumerable<Process> processes, string actionDescription)
        {
            Text = "Office-applicaties sluiten";
            Size = new Size(520, 350);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;

            Label lblTitle = new Label
            {
                Text = "Sluit eerst Office-applicaties",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#0F172A"),
                AutoSize = false,
                Size = new Size(460, 30),
                Location = new Point(24, 20)
            };
            Controls.Add(lblTitle);

            Label lblText = new Label
            {
                Text = "Voor het " + actionDescription + " moet Office volledig gesloten zijn. Je kunt de applicaties zelf sluiten of de installer toestemming geven om ze af te sluiten.",
                Font = new Font("Segoe UI", 10),
                ForeColor = ColorTranslator.FromHtml("#475569"),
                AutoSize = false,
                Size = new Size(460, 54),
                Location = new Point(24, 58)
            };
            Controls.Add(lblText);

            ListBox listBox = new ListBox
            {
                Font = new Font("Segoe UI", 10),
                Location = new Point(24, 126),
                Size = new Size(460, 100)
            };
            foreach (string processName in processes.Select(p => p.ProcessName).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                listBox.Items.Add(processName);
            }
            Controls.Add(listBox);

            Button btnAuto = new Button
            {
                Text = "Sluit Office automatisch",
                DialogResult = DialogResult.Yes,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new Size(190, 38),
                Location = new Point(24, 254),
                BackColor = ColorTranslator.FromHtml("#2563EB"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnAuto.FlatAppearance.BorderSize = 0;
            Controls.Add(btnAuto);

            Button btnManual = new Button
            {
                Text = "Ik sluit ze zelf",
                DialogResult = DialogResult.OK,
                Font = new Font("Segoe UI", 10),
                Size = new Size(146, 38),
                Location = new Point(228, 254),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnManual.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#CBD5E1");
            Controls.Add(btnManual);

            Button btnCancel = new Button
            {
                Text = "Annuleren",
                DialogResult = DialogResult.Cancel,
                Font = new Font("Segoe UI", 10),
                Size = new Size(110, 38),
                Location = new Point(390, 254),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#CBD5E1");
            Controls.Add(btnCancel);

            AcceptButton = btnAuto;
            CancelButton = btnCancel;
        }
    }
}
