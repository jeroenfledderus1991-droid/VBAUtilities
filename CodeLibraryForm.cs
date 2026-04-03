using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace VBEAddIn
{
    /// <summary>
    /// Form voor het selecteren van VBA modules uit de library
    /// </summary>
    public class CodeLibraryForm : Form
    {
        private SplitContainer splitContainer;
        private CheckedListBox lstModules;
        private TextBox txtPreview;
        private Button btnEdit;
        private Button btnImport;
        private Button btnCancel;
        private Button btnSelectAll;
        private Button btnSelectNone;
        private Button btnOpenFolder;
        private Button btnManagePaths;
        private Label lblInfo;
        private Label lblPath;
        private Label lblPreview;
        private Label lblDescription;
        
        public List<string> SelectedFiles { get; private set; }
        private List<string> libraryPaths;
        private List<string> allFiles;
        
        public CodeLibraryForm(List<string> libraryPaths)
        {
            this.libraryPaths = libraryPaths ?? new List<string>();
            SelectedFiles = new List<string>();
            allFiles = new List<string>();
            InitializeForm();
            LoadModules();
        }
        
        private void InitializeForm()
        {
            this.Text = "VBA Code Library";
            this.Width = 900;
            this.Height = 600;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(700, 500);
            
            // Info label
            lblInfo = new Label();
            lblInfo.Text = "Selecteer modules om toe te voegen aan het actieve VBA project:";
            lblInfo.Left = 20;
            lblInfo.Top = 20;
            lblInfo.Width = this.Width - 50;
            lblInfo.Height = 25;
            lblInfo.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(lblInfo);
            
            // Library path label
            lblPath = new Label();
            lblPath.Text = "Libraries: " + libraryPaths.Count + " pad(en)";
            lblPath.Left = 20;
            lblPath.Top = 45;
            lblPath.Width = this.Width - 50;
            lblPath.Height = 20;
            lblPath.Font = new Font(lblPath.Font, FontStyle.Italic);
            lblPath.ForeColor = Color.DarkGray;
            lblPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(lblPath);
            
            // SplitContainer voor modules lijst en preview
            splitContainer = new SplitContainer();
            splitContainer.Left = 20;
            splitContainer.Top = 150;
            splitContainer.Width = this.Width - 40;
            splitContainer.Height = this.Height - 180;
            splitContainer.Orientation = Orientation.Vertical;
            splitContainer.SplitterDistance = 400;
            splitContainer.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(splitContainer);
            
            // CheckedListBox voor modules (left panel)
            lstModules = new CheckedListBox();
            lstModules.Dock = DockStyle.Fill;
            lstModules.CheckOnClick = true;
            lstModules.SelectedIndexChanged += LstModules_SelectedIndexChanged;
            splitContainer.Panel1.Controls.Add(lstModules);
            
            // Preview panel (right panel)
            lblPreview = new Label();
            lblPreview.Text = "Module Preview";
            lblPreview.Dock = DockStyle.Top;
            lblPreview.Height = 25;
            lblPreview.Font = new Font(lblPreview.Font, FontStyle.Bold);
            lblPreview.TextAlign = ContentAlignment.MiddleLeft;
            lblPreview.Padding = new Padding(5);
            splitContainer.Panel2.Controls.Add(lblPreview);
            
            lblDescription = new Label();
            lblDescription.Text = "";
            lblDescription.Dock = DockStyle.Top;
            lblDescription.Height = 0; // Initially hidden
            lblDescription.ForeColor = Color.DarkBlue;
            lblDescription.BackColor = Color.LightYellow;
            lblDescription.Padding = new Padding(5);
            lblDescription.Font = new Font(lblDescription.Font, FontStyle.Italic);
            lblDescription.AutoSize = false;
            splitContainer.Panel2.Controls.Add(lblDescription);
            
            txtPreview = new TextBox();
            txtPreview.Multiline = true;
            txtPreview.ScrollBars = ScrollBars.Both;
            txtPreview.ReadOnly = true;
            txtPreview.Font = new Font("Consolas", 9);
            txtPreview.WordWrap = false;
            txtPreview.BackColor = Color.White;
            txtPreview.Dock = DockStyle.Fill;
            splitContainer.Panel2.Controls.Add(txtPreview);
            
            btnEdit = new Button();
            btnEdit.Text = "Open in Editor";
            btnEdit.Dock = DockStyle.Bottom;
            btnEdit.Height = 35;
            btnEdit.Click += BtnEdit_Click;
            splitContainer.Panel2.Controls.Add(btnEdit);
            
            // Select All button
            btnSelectAll = new Button();
            btnSelectAll.Text = "Alles";
            btnSelectAll.Width = 80;
            btnSelectAll.Height = 30;
            btnSelectAll.Left = 20;
            btnSelectAll.Top = splitContainer.Bottom + 10;
            btnSelectAll.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSelectAll.Click += BtnSelectAll_Click;
            this.Controls.Add(btnSelectAll);
            
            // Select None button
            btnSelectNone = new Button();
            btnSelectNone.Text = "Geen";
            btnSelectNone.Width = 80;
            btnSelectNone.Height = 30;
            btnSelectNone.Left = btnSelectAll.Right + 10;
            btnSelectNone.Top = splitContainer.Bottom + 10;
            btnSelectNone.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSelectNone.Click += BtnSelectNone_Click;
            this.Controls.Add(btnSelectNone);
            
            // Open Folder button
            btnOpenFolder = new Button();
            btnOpenFolder.Text = "Open Map";
            btnOpenFolder.Width = 100;
            btnOpenFolder.Height = 30;
            btnOpenFolder.Left = btnSelectNone.Right + 10;
            btnOpenFolder.Top = splitContainer.Bottom + 10;
            btnOpenFolder.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnOpenFolder.Click += BtnOpenFolder_Click;
            this.Controls.Add(btnOpenFolder);
            
            // Manage Paths button
            btnManagePaths = new Button();
            btnManagePaths.Text = "Beheer Paden...";
            btnManagePaths.Width = 120;
            btnManagePaths.Height = 30;
            btnManagePaths.Left = btnOpenFolder.Right + 10;
            btnManagePaths.Top = splitContainer.Bottom + 10;
            btnManagePaths.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnManagePaths.Click += BtnManagePaths_Click;
            this.Controls.Add(btnManagePaths);
            
            // Import button
            btnImport = new Button();
            btnImport.Text = "Importeren";
            btnImport.Width = 120;
            btnImport.Height = 30;
            btnImport.Left = this.Width - 260;
            btnImport.Top = splitContainer.Bottom + 10;
            btnImport.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnImport.Click += BtnImport_Click;
            this.Controls.Add(btnImport);
            
            // Cancel button
            btnCancel = new Button();
            btnCancel.Text = "Annuleren";
            btnCancel.Width = 120;
            btnCancel.Height = 30;
            btnCancel.Left = this.Width - 130;
            btnCancel.Top = splitContainer.Bottom + 10;
            btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
            
            this.AcceptButton = btnImport;
            this.CancelButton = btnCancel;
        }
        
        private void LoadModules()
        {
            try
            {
                allFiles.Clear();
                lstModules.Items.Clear();
                
                if (libraryPaths == null || libraryPaths.Count == 0)
                {
                    lblInfo.Text = "Geen library paden geconfigureerd. Gebruik Instellingen om paden toe te voegen.";
                    lblInfo.ForeColor = Color.Orange;
                    txtPreview.Text = "Geen library paden geconfigureerd.";
                    btnEdit.Enabled = false;
                    return;
                }
                
                // Zoek alle VBA bestanden uit alle library paden
                string[] extensions = { "*.bas", "*.cls", "*.frm" };
                
                foreach (string libraryPath in libraryPaths)
                {
                    if (!Directory.Exists(libraryPath))
                    {
                        // Probeer map te maken
                        try
                        {
                            Directory.CreateDirectory(libraryPath);
                        }
                        catch
                        {
                            // Skip als niet bereikbaar (bijv. offline SharePoint)
                            continue;
                        }
                    }
                    
                    foreach (string ext in extensions)
                    {
                        string[] files = Directory.GetFiles(libraryPath, ext, SearchOption.AllDirectories);
                        allFiles.AddRange(files);
                    }
                }
                
                if (allFiles.Count == 0)
                {
                    lblInfo.Text = "Geen modules gevonden. Kopieer .bas/.cls/.frm bestanden naar een library map.";
                    lblInfo.ForeColor = Color.Orange;
                    txtPreview.Text = "Geen module geselecteerd.";
                    btnEdit.Enabled = false;
                    return;
                }
                
                // Sorteer
                allFiles.Sort();
                
                // Voeg toe aan lijst met library path info
                foreach (string file in allFiles)
                {
                    // Zoek uit welk library path dit bestand komt
                    string sourceLibrary = "";
                    foreach (string libPath in libraryPaths)
                    {
                        if (file.StartsWith(libPath))
                        {
                            sourceLibrary = Path.GetFileName(libPath.TrimEnd(new char[] { '\\', '/' }));
                            break;
                        }
                    }
                    
                    string fileName = Path.GetFileName(file);
                    
                    // Voeg type icoon toe
                    string ext = Path.GetExtension(file).ToLower();
                    string icon = "";
                    switch (ext)
                    {
                        case ".bas": icon = "[M] "; break;  // Module
                        case ".cls": icon = "[C] "; break;  // Class
                        case ".frm": icon = "[F] "; break;  // Form
                    }
                    
                    string displayName = icon + fileName;
                    if (!string.IsNullOrEmpty(sourceLibrary))
                    {
                        displayName += " (" + sourceLibrary + ")";
                    }
                    
                    lstModules.Items.Add(displayName, false);
                }
                
                lblInfo.Text = allFiles.Count + " module(s) gevonden uit " + libraryPaths.Count + " library pad(en). Selecteer modules om te importeren:";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij laden library:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lstModules.Items.Count; i++)
            {
                lstModules.SetItemChecked(i, true);
            }
        }
        
        private void BtnSelectNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lstModules.Items.Count; i++)
            {
                lstModules.SetItemChecked(i, false);
            }
        }
        
        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            try
            {
                if (libraryPaths == null || libraryPaths.Count == 0)
                {
                    MessageBox.Show(
                        "Geen library paden geconfigureerd.",
                        "Open Map",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }
                
                // Open eerste library path
                string pathToOpen = libraryPaths[0];
                if (!Directory.Exists(pathToOpen))
                {
                    Directory.CreateDirectory(pathToOpen);
                }
                
                System.Diagnostics.Process.Start("explorer.exe", pathToOpen);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Kan map niet openen:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        private void BtnManagePaths_Click(object sender, EventArgs e)
        {
            using (LibraryPathsForm form = new LibraryPathsForm(libraryPaths))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // Update library paths
                    libraryPaths = form.LibraryPaths;
                    FormatterSettings.CodeLibraryPaths = libraryPaths;
                    FormatterSettings.SaveToRegistry();
                    
                    // Reload modules from all paths
                    LoadModules();
                    
                    MessageBox.Show(
                        string.Format("Library paden bijgewerkt!\n\n{0} pad(en) geconfigureerd.\n{1} module(s) gevonden.", 
                            libraryPaths.Count, allFiles.Count),
                        "Paden Bijgewerkt",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }
        
        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (lstModules.CheckedIndices.Count == 0)
            {
                MessageBox.Show(
                    "Selecteer minimaal één module om te importeren.",
                    "Code Library",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            
            SelectedFiles.Clear();
            
            foreach (int index in lstModules.CheckedIndices)
            {
                SelectedFiles.Add(allFiles[index]);
            }
            
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
        
        private void LstModules_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstModules.SelectedIndex < 0 || lstModules.SelectedIndex >= allFiles.Count)
            {
                txtPreview.Text = "Geen module geselecteerd.";
                lblPreview.Text = "Module Preview";
                lblDescription.Text = "";
                lblDescription.Height = 0;
                btnEdit.Enabled = false;
                return;
            }
            
            try
            {
                string filePath = allFiles[lstModules.SelectedIndex];
                string fileName = Path.GetFileName(filePath);
                
                lblPreview.Text = "Preview: " + fileName;
                txtPreview.Text = File.ReadAllText(filePath);
                btnEdit.Enabled = true;
                
                // Laad beschrijving als die bestaat
                string descriptionPath = Path.ChangeExtension(filePath, ".txt");
                if (File.Exists(descriptionPath))
                {
                    string description = File.ReadAllText(descriptionPath).Trim();
                    if (!string.IsNullOrEmpty(description))
                    {
                        lblDescription.Text = "Beschrijving: " + description;
                        lblDescription.Height = 60;
                    }
                    else
                    {
                        lblDescription.Text = "";
                        lblDescription.Height = 0;
                    }
                }
                else
                {
                    lblDescription.Text = "";
                    lblDescription.Height = 0;
                }
            }
            catch (Exception ex)
            {
                txtPreview.Text = "Fout bij laden preview:\n\n" + ex.Message;
                lblDescription.Text = "";
                lblDescription.Height = 0;
                btnEdit.Enabled = false;
            }
        }
        
        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (lstModules.SelectedIndex < 0 || lstModules.SelectedIndex >= allFiles.Count)
                return;
            
            try
            {
                string filePath = allFiles[lstModules.SelectedIndex];
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Kan bestand niet openen:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
