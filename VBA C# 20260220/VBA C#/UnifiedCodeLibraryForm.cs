using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Unified form voor import EN export van VBA modules
    /// </summary>
    public class UnifiedCodeLibraryForm : Form
    {
        // Left panel - Project modules
        private CheckedListBox lstProjectModules;
        private Label lblProject;
        private Button btnSelectAllProject;
        private Button btnSelectNoneProject;
        
        // Middle panel - Library tree
        private TreeView treeLibrary;
        private Label lblLibrary;
        private Button btnRefresh;
        private Button btnManagePaths;
        private Button btnNewFolder;
        private TextBox txtNewFolder;
        
        // Right panel - Library modules + preview
        private CheckedListBox lstLibraryModules;
        private TextBox txtPreview;
        private Label lblLibraryModules;
        private Label lblPreview;
        private Label lblDescription;
        private Button btnSelectAllLibrary;
        private Button btnSelectNoneLibrary;
        
        // Bottom buttons
        private Button btnImport;
        private Button btnExport;
        private Button btnClose;
        
        private VBProject project;
        private List<string> libraryPaths;
        private Dictionary<int, VBComponent> projectComponentMap;
        private List<string> libraryFiles;
        private string currentLibraryFolder;
        
        public UnifiedCodeLibraryForm(VBProject project, List<string> libraryPaths)
        {
            this.project = project;
            this.libraryPaths = libraryPaths ?? new List<string>();
            this.projectComponentMap = new Dictionary<int, VBComponent>();
            this.libraryFiles = new List<string>();
            this.currentLibraryFolder = "";
            
            InitializeForm();
            this.Load += UnifiedCodeLibraryForm_Load;
        }
        
        private void UnifiedCodeLibraryForm_Load(object sender, EventArgs e)
        {
            LoadProjectModules();
            LoadLibraryTree();
        }
        
        private void BtnOpenInEditor_Click(object sender, EventArgs e)
        {
            if (lstLibraryModules.SelectedIndex < 0 || lstLibraryModules.SelectedIndex >= libraryFiles.Count)
            {
                MessageBox.Show("Selecteer een module om te openen.", "Open in Editor",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            try
            {
                string filePath = libraryFiles[lstLibraryModules.SelectedIndex];
                System.Diagnostics.Process.Start("notepad.exe", filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fout bij openen bestand:\n\n" + ex.Message, "Fout",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void InitializeForm()
        {
            this.Text = "VBA Code Library";
            this.ClientSize = new Size(1200, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(1000, 600);
            
            // Main layout: 3 columns
            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.ColumnCount = 3;
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F)); // Project
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F)); // Library tree
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F)); // Library modules
            mainLayout.RowCount = 2;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            this.Controls.Add(mainLayout);
            
            // LEFT PANEL - Project Modules
            Panel leftPanel = CreateProjectPanel();
            mainLayout.Controls.Add(leftPanel, 0, 0);
            
            // MIDDLE PANEL - Library Tree
            Panel middlePanel = CreateLibraryTreePanel();
            mainLayout.Controls.Add(middlePanel, 1, 0);
            
            // RIGHT PANEL - Library Modules
            Panel rightPanel = CreateLibraryModulesPanel();
            mainLayout.Controls.Add(rightPanel, 2, 0);
            
            // BOTTOM PANEL - Action buttons
            Panel bottomPanel = new Panel();
            bottomPanel.Dock = DockStyle.Fill;
            bottomPanel.Padding = new Padding(10, 5, 10, 5);
            mainLayout.Controls.Add(bottomPanel, 0, 1);
            mainLayout.SetColumnSpan(bottomPanel, 3);
            
            btnImport = new Button();
            btnImport.Text = "← Importeren";
            btnImport.Width = 130;
            btnImport.Height = 35;
            btnImport.Left = 10;
            btnImport.Top = 8;
            btnImport.Click += BtnImport_Click;
            bottomPanel.Controls.Add(btnImport);
            
            btnExport = new Button();
            btnExport.Text = "Exporteren →";
            btnExport.Width = 130;
            btnExport.Height = 35;
            btnExport.Left = 150;
            btnExport.Top = 8;
            btnExport.Click += BtnExport_Click;
            bottomPanel.Controls.Add(btnExport);
            
            btnClose = new Button();
            btnClose.Text = "Sluiten";
            btnClose.Width = 100;
            btnClose.Height = 35;
            btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnClose.Left = bottomPanel.Width - 110;
            btnClose.Top = 8;
            btnClose.Click += (s, e) => this.Close();
            bottomPanel.Controls.Add(btnClose);
            
            this.CancelButton = btnClose;
        }
        
        private Panel CreateProjectPanel()
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.BorderStyle = BorderStyle.FixedSingle;
            panel.Padding = new Padding(5);
            
            lblProject = new Label();
            lblProject.Text = "Project Modules (selecteer voor export)";
            lblProject.Dock = DockStyle.Top;
            lblProject.Height = 30;
            lblProject.Font = new Font(lblProject.Font, FontStyle.Bold);
            lblProject.TextAlign = ContentAlignment.MiddleLeft;
            panel.Controls.Add(lblProject);
            
            Panel buttonPanel = new Panel();
            buttonPanel.Dock = DockStyle.Bottom;
            buttonPanel.Height = 40;
            panel.Controls.Add(buttonPanel);
            
            btnSelectAllProject = new Button();
            btnSelectAllProject.Text = "Alles";
            btnSelectAllProject.Width = 70;
            btnSelectAllProject.Left = 5;
            btnSelectAllProject.Top = 5;
            btnSelectAllProject.Click += (s, e) => {
                for (int i = 0; i < lstProjectModules.Items.Count; i++)
                    lstProjectModules.SetItemChecked(i, true);
            };
            buttonPanel.Controls.Add(btnSelectAllProject);
            
            btnSelectNoneProject = new Button();
            btnSelectNoneProject.Text = "Geen";
            btnSelectNoneProject.Width = 70;
            btnSelectNoneProject.Left = 85;
            btnSelectNoneProject.Top = 5;
            btnSelectNoneProject.Click += (s, e) => {
                for (int i = 0; i < lstProjectModules.Items.Count; i++)
                    lstProjectModules.SetItemChecked(i, false);
            };
            buttonPanel.Controls.Add(btnSelectNoneProject);
            
            lstProjectModules = new CheckedListBox();
            lstProjectModules.Dock = DockStyle.Fill;
            lstProjectModules.CheckOnClick = true;
            lstProjectModules.MouseDown += LstProjectModules_MouseDown;
            panel.Controls.Add(lstProjectModules);
            
            return panel;
        }
        
        private Panel CreateLibraryTreePanel()
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.BorderStyle = BorderStyle.FixedSingle;
            panel.Padding = new Padding(5);
            
            // Add controls in reverse order: Fill first, then Top panels
            treeLibrary = new TreeView();
            treeLibrary.Dock = DockStyle.Fill;
            treeLibrary.AllowDrop = true;
            treeLibrary.AfterSelect += TreeLibrary_AfterSelect;
            treeLibrary.DragEnter += TreeLibrary_DragEnter;
            treeLibrary.DragOver += TreeLibrary_DragOver;
            treeLibrary.DragDrop += TreeLibrary_DragDrop;
            panel.Controls.Add(treeLibrary);
            
            Panel topPanel = new Panel();
            topPanel.Dock = DockStyle.Top;
            topPanel.Height = 70;
            topPanel.Padding = new Padding(5);
            panel.Controls.Add(topPanel);
            
            lblLibrary = new Label();
            lblLibrary.Text = "Library Folders (drop modules hier)";
            lblLibrary.Dock = DockStyle.Top;
            lblLibrary.Height = 30;
            lblLibrary.Font = new Font(lblLibrary.Font, FontStyle.Bold);
            lblLibrary.TextAlign = ContentAlignment.MiddleLeft;
            panel.Controls.Add(lblLibrary);
            
            Label lblNewFolder = new Label();
            lblNewFolder.Text = "Nieuwe map:";
            lblNewFolder.Left = 5;
            lblNewFolder.Top = 5;
            lblNewFolder.Width = 80;
            lblNewFolder.TextAlign = ContentAlignment.MiddleRight;
            topPanel.Controls.Add(lblNewFolder);
            
            txtNewFolder = new TextBox();
            txtNewFolder.Left = 90;
            txtNewFolder.Top = 5;
            txtNewFolder.Width = 200;
            txtNewFolder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            topPanel.Controls.Add(txtNewFolder);
            
            btnNewFolder = new Button();
            btnNewFolder.Text = "Maak aan";
            btnNewFolder.Left = 5;
            btnNewFolder.Top = 35;
            btnNewFolder.Width = 90;
            btnNewFolder.Click += BtnNewFolder_Click;
            topPanel.Controls.Add(btnNewFolder);
            
            btnRefresh = new Button();
            btnRefresh.Text = "Ververs";
            btnRefresh.Left = 105;
            btnRefresh.Top = 35;
            btnRefresh.Width = 90;
            btnRefresh.Click += (s, e) => { LoadLibraryTree(); LoadLibraryModules(); };
            topPanel.Controls.Add(btnRefresh);
            
            btnManagePaths = new Button();
            btnManagePaths.Text = "Beheer Paden...";
            btnManagePaths.Left = 205;
            btnManagePaths.Top = 35;
            btnManagePaths.Width = 120;
            btnManagePaths.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            btnManagePaths.Click += BtnManagePaths_Click;
            topPanel.Controls.Add(btnManagePaths);
            
            return panel;
        }
        
        private Panel CreateLibraryModulesPanel()
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.BorderStyle = BorderStyle.FixedSingle;
            panel.Padding = new Padding(5);
            
            lblLibraryModules = new Label();
            lblLibraryModules.Text = "Library Modules (selecteer folder links)";
            lblLibraryModules.Dock = DockStyle.Top;
            lblLibraryModules.Height = 30;
            lblLibraryModules.Font = new Font(lblLibraryModules.Font, FontStyle.Bold);
            lblLibraryModules.TextAlign = ContentAlignment.MiddleLeft;
            panel.Controls.Add(lblLibraryModules);
            
            SplitContainer splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = Orientation.Horizontal;
            splitContainer.SplitterDistance = 250;
            panel.Controls.Add(splitContainer);
            
            // Top: modules list
            Panel topSection = new Panel();
            topSection.Dock = DockStyle.Fill;
            splitContainer.Panel1.Controls.Add(topSection);
            
            Panel buttonPanel = new Panel();
            buttonPanel.Dock = DockStyle.Bottom;
            buttonPanel.Height = 35;
            topSection.Controls.Add(buttonPanel);
            
            btnSelectAllLibrary = new Button();
            btnSelectAllLibrary.Text = "Alles";
            btnSelectAllLibrary.Width = 70;
            btnSelectAllLibrary.Left = 5;
            btnSelectAllLibrary.Top = 5;
            btnSelectAllLibrary.Click += (s, e) => {
                for (int i = 0; i < lstLibraryModules.Items.Count; i++)
                    lstLibraryModules.SetItemChecked(i, true);
            };
            buttonPanel.Controls.Add(btnSelectAllLibrary);
            
            btnSelectNoneLibrary = new Button();
            btnSelectNoneLibrary.Text = "Geen";
            btnSelectNoneLibrary.Width = 70;
            btnSelectNoneLibrary.Left = 85;
            btnSelectNoneLibrary.Top = 5;
            btnSelectNoneLibrary.Click += (s, e) => {
                for (int i = 0; i < lstLibraryModules.Items.Count; i++)
                    lstLibraryModules.SetItemChecked(i, false);
            };
            buttonPanel.Controls.Add(btnSelectNoneLibrary);
            
            lstLibraryModules = new CheckedListBox();
            lstLibraryModules.Dock = DockStyle.Fill;
            lstLibraryModules.CheckOnClick = true;
            lstLibraryModules.SelectedIndexChanged += LstLibraryModules_SelectedIndexChanged;
            topSection.Controls.Add(lstLibraryModules);
            
            // Bottom: preview
            txtPreview = new TextBox();
            txtPreview.Multiline = true;
            txtPreview.ScrollBars = ScrollBars.Both;
            txtPreview.ReadOnly = true;
            txtPreview.Font = new Font("Consolas", 9);
            txtPreview.WordWrap = false;
            txtPreview.BackColor = Color.White;
            txtPreview.Dock = DockStyle.Fill;
            splitContainer.Panel2.Controls.Add(txtPreview);
            
            Panel previewButtonPanel = new Panel();
            previewButtonPanel.Dock = DockStyle.Bottom;
            previewButtonPanel.Height = 35;
            splitContainer.Panel2.Controls.Add(previewButtonPanel);
            
            Button btnOpenInEditor = new Button();
            btnOpenInEditor.Text = "Open in Editor";
            btnOpenInEditor.Left = 5;
            btnOpenInEditor.Top = 5;
            btnOpenInEditor.Width = 120;
            btnOpenInEditor.Click += BtnOpenInEditor_Click;
            previewButtonPanel.Controls.Add(btnOpenInEditor);
            
            lblDescription = new Label();
            lblDescription.Text = "";
            lblDescription.Dock = DockStyle.Top;
            lblDescription.Height = 0;
            lblDescription.ForeColor = Color.DarkBlue;
            lblDescription.BackColor = Color.LightYellow;
            lblDescription.Padding = new Padding(5);
            lblDescription.Font = new Font(lblDescription.Font, FontStyle.Italic);
            lblDescription.AutoSize = false;
            splitContainer.Panel2.Controls.Add(lblDescription);
            
            lblPreview = new Label();
            lblPreview.Text = "Preview";
            lblPreview.Dock = DockStyle.Top;
            lblPreview.Height = 25;
            lblPreview.Font = new Font(lblPreview.Font, FontStyle.Bold);
            lblPreview.TextAlign = ContentAlignment.MiddleLeft;
            splitContainer.Panel2.Controls.Add(lblPreview);
            
            return panel;
        }
        
        private void LoadProjectModules()
        {
            lstProjectModules.Items.Clear();
            projectComponentMap.Clear();
            
            try
            {
                int index = 0;
                foreach (VBComponent component in project.VBComponents)
                {
                    if (component.Type == vbext_ComponentType.vbext_ct_StdModule ||
                        component.Type == vbext_ComponentType.vbext_ct_ClassModule ||
                        component.Type == vbext_ComponentType.vbext_ct_MSForm)
                    {
                        string icon = "";
                        switch (component.Type)
                        {
                            case vbext_ComponentType.vbext_ct_StdModule: icon = "[M] "; break;
                            case vbext_ComponentType.vbext_ct_ClassModule: icon = "[C] "; break;
                            case vbext_ComponentType.vbext_ct_MSForm: icon = "[F] "; break;
                        }
                        
                        lstProjectModules.Items.Add(icon + component.Name, false);
                        projectComponentMap[index] = component;
                        index++;
                    }
                }
                
                lblProject.Text = string.Format("Project Modules ({0} modules)", lstProjectModules.Items.Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fout bij laden project modules:\n\n" + ex.Message, "Fout", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void LoadLibraryTree()
        {
            try
            {
                treeLibrary.Nodes.Clear();
                
                if (libraryPaths == null || libraryPaths.Count == 0)
                {
                    return;
                }
                
                foreach (string libraryPath in libraryPaths)
                {
                    if (!Directory.Exists(libraryPath))
                    {
                        try { Directory.CreateDirectory(libraryPath); }
                        catch { continue; }
                    }
                    
                    string displayName = Path.GetFileName(libraryPath.TrimEnd(new char[] { '\\', '/' }));
                    if (string.IsNullOrEmpty(displayName))
                        displayName = libraryPath;
                    
                    TreeNode rootNode = new TreeNode(displayName);
                    rootNode.Tag = libraryPath;
                    treeLibrary.Nodes.Add(rootNode);
                    
                    LoadSubfolders(rootNode, libraryPath);
                    rootNode.Expand();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fout bij laden library:\n\n" + ex.Message, "Fout",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void LoadSubfolders(TreeNode parentNode, string path)
        {
            try
            {
                string[] subDirs = Directory.GetDirectories(path);
                foreach (string dir in subDirs)
                {
                    TreeNode node = new TreeNode(Path.GetFileName(dir));
                    node.Tag = dir;
                    parentNode.Nodes.Add(node);
                    LoadSubfolders(node, dir);
                }
            }
            catch { }
        }
        
        private void TreeLibrary_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node != null && e.Node.Tag != null)
            {
                currentLibraryFolder = e.Node.Tag.ToString();
                LoadLibraryModules();
            }
        }
        
        private void LoadLibraryModules()
        {
            lstLibraryModules.Items.Clear();
            libraryFiles.Clear();
            txtPreview.Text = "";
            lblDescription.Text = "";
            lblDescription.Height = 0;
            
            if (string.IsNullOrEmpty(currentLibraryFolder) || !Directory.Exists(currentLibraryFolder))
            {
                lblLibraryModules.Text = "Library Modules (selecteer folder links)";
                return;
            }
            
            try
            {
                string[] extensions = { "*.bas", "*.cls", "*.frm" };
                foreach (string ext in extensions)
                {
                    string[] files = Directory.GetFiles(currentLibraryFolder, ext);
                    libraryFiles.AddRange(files);
                }
                
                libraryFiles.Sort();
                
                foreach (string file in libraryFiles)
                {
                    string fileName = Path.GetFileName(file);
                    string ext = Path.GetExtension(file).ToLower();
                    string icon = "";
                    switch (ext)
                    {
                        case ".bas": icon = "[M] "; break;
                        case ".cls": icon = "[C] "; break;
                        case ".frm": icon = "[F] "; break;
                    }
                    
                    lstLibraryModules.Items.Add(icon + fileName, false);
                }
                
                string folderName = Path.GetFileName(currentLibraryFolder);
                lblLibraryModules.Text = string.Format("Library Modules in '{0}' ({1} modules)", 
                    folderName, libraryFiles.Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fout bij laden modules:\n\n" + ex.Message, "Fout",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void LstLibraryModules_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstLibraryModules.SelectedIndex < 0 || lstLibraryModules.SelectedIndex >= libraryFiles.Count)
            {
                txtPreview.Text = "Geen module geselecteerd.";
                lblPreview.Text = "Preview";
                lblDescription.Text = "";
                lblDescription.Height = 0;
                return;
            }
            
            try
            {
                string filePath = libraryFiles[lstLibraryModules.SelectedIndex];
                string fileName = Path.GetFileName(filePath);
                
                lblPreview.Text = "Preview: " + fileName;
                txtPreview.Text = File.ReadAllText(filePath);
                
                // Laad beschrijving
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
            }
        }
        
        // Drag & Drop for export
        private void LstProjectModules_MouseDown(object sender, MouseEventArgs e)
        {
            if (lstProjectModules.SelectedItem == null)
                return;
            
            int index = lstProjectModules.SelectedIndex;
            if (!projectComponentMap.ContainsKey(index))
                return;
            
            lstProjectModules.DoDragDrop(index, DragDropEffects.Copy);
        }
        
        private void TreeLibrary_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(int)))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }
        
        private void TreeLibrary_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(int)))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }
        
        private void TreeLibrary_DragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(int)))
                return;
            
            int index = (int)e.Data.GetData(typeof(int));
            if (!projectComponentMap.ContainsKey(index))
                return;
            
            VBComponent component = projectComponentMap[index];
            
            Point pt = treeLibrary.PointToClient(new Point(e.X, e.Y));
            TreeNode node = treeLibrary.GetNodeAt(pt);
            
            if (node == null || node.Tag == null)
            {
                MessageBox.Show("Selecteer een folder in de library tree.", "Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            string targetFolder = node.Tag.ToString();
            ExportModule(component, targetFolder);
        }
        
        private void ExportModule(VBComponent component, string targetFolder)
        {
            try
            {
                string extension = "";
                switch (component.Type)
                {
                    case vbext_ComponentType.vbext_ct_StdModule: extension = ".bas"; break;
                    case vbext_ComponentType.vbext_ct_ClassModule: extension = ".cls"; break;
                    case vbext_ComponentType.vbext_ct_MSForm: extension = ".frm"; break;
                }
                
                string targetPath = Path.Combine(targetFolder, component.Name + extension);
                
                if (File.Exists(targetPath))
                {
                    DialogResult result = MessageBox.Show(
                        string.Format("Bestand '{0}{1}' bestaat al.\n\nOverschrijven?", component.Name, extension),
                        "Bestand bestaat",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    
                    if (result != DialogResult.Yes)
                        return;
                    
                    File.Delete(targetPath);
                }
                
                component.Export(targetPath);
                
                // Vraag om beschrijving
                string description = PromptForDescription(component.Name);
                if (!string.IsNullOrEmpty(description))
                {
                    string descriptionPath = Path.ChangeExtension(targetPath, ".txt");
                    File.WriteAllText(descriptionPath, description);
                }
                
                MessageBox.Show(
                    string.Format("Module '{0}' geëxporteerd naar:\n{1}", component.Name, targetPath),
                    "Export Succesvol",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                // Refresh als we in deze folder zitten
                if (currentLibraryFolder == targetFolder)
                    LoadLibraryModules();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fout bij exporteren:\n\n" + ex.Message, "Fout",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private string PromptForDescription(string moduleName)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 250,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Beschrijving toevoegen (optioneel)",
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };
            
            Label lblInfo = new Label()
            {
                Left = 10,
                Top = 10,
                Width = 460,
                Height = 40,
                Text = string.Format("Voeg een beschrijving toe voor module '{0}':\n(Optioneel - klik Overslaan om geen beschrijving toe te voegen)", moduleName)
            };
            
            TextBox txtDescription = new TextBox()
            {
                Left = 10,
                Top = 55,
                Width = 460,
                Height = 100,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            
            Button btnSave = new Button() { Text = "Opslaan", Left = 280, Width = 90, Top = 170, DialogResult = DialogResult.OK };
            Button btnSkip = new Button() { Text = "Overslaan", Left = 380, Width = 90, Top = 170, DialogResult = DialogResult.Cancel };
            
            prompt.Controls.Add(lblInfo);
            prompt.Controls.Add(txtDescription);
            prompt.Controls.Add(btnSave);
            prompt.Controls.Add(btnSkip);
            prompt.AcceptButton = btnSave;
            prompt.CancelButton = btnSkip;
            
            return prompt.ShowDialog() == DialogResult.OK ? txtDescription.Text.Trim() : "";
        }
        
        private void BtnNewFolder_Click(object sender, EventArgs e)
        {
            if (treeLibrary.SelectedNode == null || treeLibrary.SelectedNode.Tag == null)
            {
                MessageBox.Show("Selecteer eerst een parent folder in de tree.", "Nieuwe Map",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            string folderName = txtNewFolder.Text.Trim();
            if (string.IsNullOrEmpty(folderName))
            {
                MessageBox.Show("Voer een naam in voor de nieuwe map.", "Nieuwe Map",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                string parentPath = treeLibrary.SelectedNode.Tag.ToString();
                string newPath = Path.Combine(parentPath, folderName);
                
                if (Directory.Exists(newPath))
                {
                    MessageBox.Show("Deze map bestaat al.", "Nieuwe Map",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                Directory.CreateDirectory(newPath);
                txtNewFolder.Text = "";
                LoadLibraryTree();
                
                MessageBox.Show("Map aangemaakt:\n" + newPath, "Map Aangemaakt",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fout bij aanmaken map:\n\n" + ex.Message, "Fout",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void BtnManagePaths_Click(object sender, EventArgs e)
        {
            using (LibraryPathsForm form = new LibraryPathsForm(libraryPaths))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    libraryPaths = form.LibraryPaths;
                    FormatterSettings.CodeLibraryPaths = libraryPaths;
                    FormatterSettings.SaveToRegistry();
                    
                    LoadLibraryTree();
                    
                    MessageBox.Show(
                        string.Format("Library paden bijgewerkt!\n\n{0} pad(en) geconfigureerd.", libraryPaths.Count),
                        "Paden Bijgewerkt",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }
        
        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (treeLibrary.SelectedNode == null || treeLibrary.SelectedNode.Tag == null)
            {
                MessageBox.Show("Selecteer een target folder in de library tree.", "Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            List<int> checkedIndices = new List<int>();
            for (int i = 0; i < lstProjectModules.CheckedIndices.Count; i++)
                checkedIndices.Add(lstProjectModules.CheckedIndices[i]);
            
            if (checkedIndices.Count == 0)
            {
                MessageBox.Show("Selecteer minimaal één module om te exporteren.", "Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            string targetFolder = treeLibrary.SelectedNode.Tag.ToString();
            int exported = 0;
            
            foreach (int index in checkedIndices)
            {
                if (projectComponentMap.ContainsKey(index))
                {
                    VBComponent component = projectComponentMap[index];
                    
                    try
                    {
                        string extension = "";
                        switch (component.Type)
                        {
                            case vbext_ComponentType.vbext_ct_StdModule: extension = ".bas"; break;
                            case vbext_ComponentType.vbext_ct_ClassModule: extension = ".cls"; break;
                            case vbext_ComponentType.vbext_ct_MSForm: extension = ".frm"; break;
                        }
                        
                        string targetPath = Path.Combine(targetFolder, component.Name + extension);
                        
                        if (File.Exists(targetPath))
                        {
                            DialogResult result = MessageBox.Show(
                                string.Format("Bestand '{0}{1}' bestaat al.\n\nOverschrijven?", component.Name, extension),
                                "Bestand bestaat",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question);
                            
                            if (result == DialogResult.Cancel)
                                break;
                            if (result == DialogResult.No)
                                continue;
                            
                            File.Delete(targetPath);
                        }
                        
                        component.Export(targetPath);
                        exported++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            string.Format("Fout bij exporteren '{0}':\n\n{1}", component.Name, ex.Message),
                            "Export Fout",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
            
            MessageBox.Show(
                string.Format("{0} van {1} modules geëxporteerd naar:\n{2}", exported, checkedIndices.Count, targetFolder),
                "Export Compleet",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            
            if (currentLibraryFolder == targetFolder)
                LoadLibraryModules();
        }
        
        private void BtnImport_Click(object sender, EventArgs e)
        {
            List<int> checkedIndices = new List<int>();
            for (int i = 0; i < lstLibraryModules.CheckedIndices.Count; i++)
                checkedIndices.Add(lstLibraryModules.CheckedIndices[i]);
            
            if (checkedIndices.Count == 0)
            {
                MessageBox.Show("Selecteer minimaal één module om te importeren.", "Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            int imported = 0;
            
            foreach (int index in checkedIndices)
            {
                if (index >= 0 && index < libraryFiles.Count)
                {
                    string filePath = libraryFiles[index];
                    string moduleName = Path.GetFileNameWithoutExtension(filePath);
                    
                    try
                    {
                        // Check if exists
                        bool exists = false;
                        foreach (VBComponent comp in project.VBComponents)
                        {
                            if (comp.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                            {
                                exists = true;
                                break;
                            }
                        }
                        
                        if (exists)
                        {
                            DialogResult result = MessageBox.Show(
                                string.Format("Module '{0}' bestaat al.\n\nOverschrijven?", moduleName),
                                "Module Bestaat",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question);
                            
                            if (result == DialogResult.Cancel)
                                break;
                            if (result == DialogResult.No)
                                continue;
                            
                            // Remove old
                            foreach (VBComponent comp in project.VBComponents)
                            {
                                if (comp.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                                {
                                    project.VBComponents.Remove(comp);
                                    break;
                                }
                            }
                        }
                        
                        project.VBComponents.Import(filePath);
                        imported++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            string.Format("Fout bij importeren '{0}':\n\n{1}", moduleName, ex.Message),
                            "Import Fout",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
            
            MessageBox.Show(
                string.Format("{0} van {1} modules geïmporteerd.", imported, checkedIndices.Count),
                "Import Compleet",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            
            LoadProjectModules();
        }
    }
}
