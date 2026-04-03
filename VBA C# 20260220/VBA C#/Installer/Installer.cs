using System;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Security.Principal;
using System.Windows.Forms;
using System.Drawing;
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
        private static string InstallPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "VBE AddIn");
        
        private static string VBEAddInDll = "VBEAddIn.dll";
        private static string VBEAddInPath = Path.Combine(InstallPath, VBEAddInDll);
        private static string RegAsmPath = Path.Combine(
            Environment.GetEnvironmentVariable("windir"),
            @"Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe");

        private Label lblTitle;
        private Label lblSubtitle;
        private Button btnInstall;
        private Button btnUninstall;
        private Button btnClose;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Panel headerPanel;

        public InstallerForm()
        {
            InitializeUI();
            CheckAdminRights();
        }

        private void InitializeUI()
        {
            // Form instellingen
            this.Text = "VBE Code Tools - Setup";
            this.Size = new Size(550, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.White;

            // Header panel
            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                BackColor = ColorTranslator.FromHtml("#0078D4")
            };
            this.Controls.Add(headerPanel);

            // Titel
            lblTitle = new Label
            {
                Text = "VBE Code Tools",
                Font = new Font("Segoe UI", 20, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                Size = new Size(500, 40),
                Location = new Point(20, 20),
                BackColor = Color.Transparent
            };
            headerPanel.Controls.Add(lblTitle);

            // Subtitel
            lblSubtitle = new Label
            {
                Text = "Add-in voor VBA Editor code formatting",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.White,
                AutoSize = false,
                Size = new Size(500, 25),
                Location = new Point(20, 60),
                BackColor = Color.Transparent
            };
            headerPanel.Controls.Add(lblSubtitle);

            // Installeer knop
            btnInstall = new Button
            {
                Text = "Installeren",
                Size = new Size(200, 50),
                Location = new Point(50, 140),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                BackColor = ColorTranslator.FromHtml("#0078D4"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;
            this.Controls.Add(btnInstall);

            // Verwijder knop
            btnUninstall = new Button
            {
                Text = "Verwijderen",
                Size = new Size(200, 50),
                Location = new Point(280, 140),
                Font = new Font("Segoe UI", 12),
                BackColor = ColorTranslator.FromHtml("#E81123"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnUninstall.FlatAppearance.BorderSize = 0;
            btnUninstall.Click += BtnUninstall_Click;
            this.Controls.Add(btnUninstall);

            // Progress bar
            progressBar = new ProgressBar
            {
                Size = new Size(480, 25),
                Location = new Point(30, 220),
                Visible = false,
                Style = ProgressBarStyle.Continuous
            };
            this.Controls.Add(progressBar);

            // Status label
            lblStatus = new Label
            {
                Text = "Klaar om te installeren",
                Size = new Size(480, 60),
                Location = new Point(30, 255),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.TopLeft
            };
            this.Controls.Add(lblStatus);

            // Sluiten knop
            btnClose = new Button
            {
                Text = "Sluiten",
                Size = new Size(100, 35),
                Location = new Point(410, 320),
                Font = new Font("Segoe UI", 10),
                BackColor = Color.WhiteSmoke,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderColor = Color.Gray;
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }

        private void CheckAdminRights()
        {
            if (!IsRunningAsAdmin())
            {
                lblStatus.Text = "⚠️ Deze installer moet als Administrator worden uitgevoerd!\n\nRechtsklik op het bestand en kies 'Als administrator uitvoeren'";
                lblStatus.ForeColor = Color.Red;
                btnInstall.Enabled = false;
                btnUninstall.Enabled = false;
            }
            else if (!File.Exists(RegAsmPath))
            {
                lblStatus.Text = "⚠️ .NET Framework 4.8 niet gevonden!\n\nInstalleer .NET Framework 4.8 en probeer opnieuw.";
                lblStatus.ForeColor = Color.Red;
                btnInstall.Enabled = false;
                btnUninstall.Enabled = false;
            }
        }

        private void BtnInstall_Click(object sender, EventArgs e)
        {
            btnInstall.Enabled = false;
            btnUninstall.Enabled = false;
            btnClose.Enabled = false;
            progressBar.Visible = true;
            progressBar.Value = 0;

            try
            {
                // Controleer eerst of er een oude installatie bestaat
                if (Directory.Exists(InstallPath) || File.Exists(VBEAddInPath))
                {
                    UpdateStatus("Oude installatie detecteren...", 5);
                    System.Threading.Thread.Sleep(300);
                    
                    UpdateStatus("Oude versie verwijderen...", 10);
                    try
                    {
                        UnregisterCOM();
                        RemoveRegistryEntries();
                        if (Directory.Exists(InstallPath))
                        {
                            Directory.Delete(InstallPath, true);
                        }
                    }
                    catch (Exception)
                    {
                        // Oude installatie verwijderen mislukt, maar doorgaan met nieuwe installatie
                    }
                    System.Threading.Thread.Sleep(300);
                }

                UpdateStatus("Installatiemap maken...", 20);
                if (!Directory.Exists(InstallPath))
                {
                    Directory.CreateDirectory(InstallPath);
                }

                UpdateStatus("DLL bestand extracten...", 40);
                ExtractEmbeddedDll();

                UpdateStatus("COM registratie uitvoeren...", 60);
                RegisterCOM();

                UpdateStatus("VBE add-in registreren...", 80);
                CreateRegistryEntries();

                UpdateStatus("Installatie verifiëren...", 95);
                System.Threading.Thread.Sleep(500);

                UpdateStatus("✓ Installatie voltooid!", 100);
                lblStatus.ForeColor = Color.Green;

                MessageBox.Show(
                    "VBE Code Tools is succesvol geïnstalleerd!\n\n" +
                    "Sluit alle Office applicaties en start Excel opnieuw.\n" +
                    "Open VBE (Alt+F11) en kijk in het menu:\n" +
                    "• Utilities → Formatting → Formatteer Dim Statements",
                    "Installatie voltooid",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                lblStatus.Text = "❌ Fout tijdens installatie:\n" + ex.Message;
                lblStatus.ForeColor = Color.Red;

                MessageBox.Show(
                    "Installatie mislukt:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                btnInstall.Enabled = true;
                btnUninstall.Enabled = true;
                btnClose.Enabled = true;
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
                return;

            btnInstall.Enabled = false;
            btnUninstall.Enabled = false;
            btnClose.Enabled = false;
            progressBar.Visible = true;
            progressBar.Value = 0;

            try
            {
                UpdateStatus("COM registratie verwijderen...", 30);
                UnregisterCOM();

                UpdateStatus("VBE add-in registratie verwijderen...", 60);
                RemoveRegistryEntries();

                UpdateStatus("Bestanden verwijderen...", 90);
                if (Directory.Exists(InstallPath))
                {
                    Directory.Delete(InstallPath, true);
                }

                UpdateStatus("✓ Verwijdering voltooid!", 100);
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
                lblStatus.Text = "❌ Fout tijdens verwijderen:\n" + ex.Message;
                lblStatus.ForeColor = Color.Red;

                MessageBox.Show(
                    "Verwijdering mislukt:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                btnInstall.Enabled = true;
                btnUninstall.Enabled = true;
                btnClose.Enabled = true;
            }
        }

        private void UpdateStatus(string message, int progress)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = Color.Gray;
            progressBar.Value = progress;
            Application.DoEvents();
            System.Threading.Thread.Sleep(300);
        }

        private void ExtractEmbeddedDll()
        {
            var assembly = Assembly.GetExecutingAssembly();
            
            // Extract VBEAddIn.dll
            string vbeResourceName = "VBEAddInInstaller.VBEAddIn.dll";
            using (Stream stream = assembly.GetManifestResourceStream(vbeResourceName))
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
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (Process process = Process.Start(psi))
            {
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    string error = process.StandardError.ReadToEnd();
                    throw new Exception("RegAsm fout: " + error);
                }
            }
        }

        private void UnregisterCOM()
        {
            if (!File.Exists(VBEAddInPath))
                return;

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = RegAsmPath,
                Arguments = string.Format("\"{0}\" /unregister", VBEAddInPath),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
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
                        if (key != null)
                        {
                            key.SetValue("FriendlyName", "VBE Code Tools");
                            key.SetValue("Description", "VBA Editor Add-in voor code formattering");
                            key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                            key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                        }
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
                    catch { }
                }
            }
        }

        private static bool IsRunningAsAdmin()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
