using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Form voor het exporteren van VBA modules naar de library met drag & drop
    /// </summary>
    public class ExportToLibraryForm : Form
    {
        private SplitContainer splitContainer;
        private CheckedListBox lstModules;
        private TreeView treeLibrary;
        private Label lblModules;
        private Label lblLibrary;
        private Button btnExport;
        private Button btnCancel;
        private Button btnSelectAll;
        private Button btnSelectNone;
        private Button btnRefresh;
        private Button btnManagePaths;
        private TextBox txtNewFolder;
        private Button btnCreateFolder;
        
        private VBProject project;
        private List<string> libraryPaths;
        private Dictionary<int, VBComponent> componentMap;
        
        public ExportToLibraryForm(VBProject project, List<string> libraryPaths)
        {
            this.project = project;
            this.libraryPaths = libraryPaths ?? new List<string>();
            this.componentMap = new Dictionary<int, VBComponent>();
            
            InitializeForm();
            LoadModules();
            LoadLibraryTree();
        }
        
        private void InitializeForm()
        {
            this.Text = "Export naar Library";
            this.ClientSize = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(700, 500);
            
            // Main SplitContainer
            splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = Orientation.Vertical;
            splitContainer.SplitterDistance = 200;
            this.Controls.Add(splitContainer);
            
            // LEFT PANEL - Project modules
            lblModules = new Label();
            lblModules.Text = "Modules uit project (drag naar library folder):";
            lblModules.Dock = DockStyle.Top;
            lblModules.Height = 30;
            lblModules.Font = new Font(lblModules.Font, FontStyle.Bold);
            lblModules.TextAlign = ContentAlignment.MiddleLeft;
            lblModules.Padding = new Padding(5);
            splitContainer.Panel1.Controls.Add(lblModules);
            
            Panel leftButtonPanel = new Panel();
            leftButtonPanel.Dock = DockStyle.Bottom;
            leftButtonPanel.Height = 45;
            splitContainer.Panel1.Controls.Add(leftButtonPanel);
            
            btnSelectAll = new Button();
            btnSelectAll.Text = "Alles";
            btnSelectAll.Width = 80;
            btnSelectAll.Height = 30;
            btnSelectAll.Left = 10;
            btnSelectAll.Top = 8;
            btnSelectAll.Click += BtnSelectAll_Click;
            leftButtonPanel.Controls.Add(btnSelectAll);
            
            btnSelectNone = new Button();
            btnSelectNone.Text = "Geen";
            btnSelectNone.Width = 80;
            btnSelectNone.Height = 30;
            btnSelectNone.Left = 100;
            btnSelectNone.Top = 8;
            btnSelectNone.Click += BtnSelectNone_Click;
            leftButtonPanel.Controls.Add(btnSelectNone);
            
            lstModules = new CheckedListBox();
            lstModules.Dock = DockStyle.Fill;
            lstModules.CheckOnClick = true;
            lstModules.MouseDown += LstModules_MouseDown;
            splitContainer.Panel1.Controls.Add(lstModules);
            
            // RIGHT PANEL - Library folders
            // Create TreeView first (will be Dock.Fill)
            treeLibrary = new TreeView();
            treeLibrary.Dock = DockStyle.Fill;
            treeLibrary.AllowDrop = true;
            treeLibrary.DragEnter += TreeLibrary_DragEnter;
            treeLibrary.DragOver += TreeLibrary_DragOver;
            treeLibrary.DragDrop += TreeLibrary_DragDrop;
            treeLibrary.ImageList = new ImageList();
            treeLibrary.ImageList.Images.Add("folder", SystemIcons.Shield.ToBitmap());
            
            // Bottom panel with buttons
            Panel bottomPanel = new Panel();
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.Height = 50;
            bottomPanel.Padding = new Padding(5);
            
            btnExport = new Button();
            btnExport.Text = "Exporteren";
            btnExport.Width = 120;
            btnExport.Height = 35;
            btnExport.Left = 10;
            btnExport.Top = 8;
            btnExport.Click += BtnExport_Click;
            bottomPanel.Controls.Add(btnExport);
            
            btnCancel = new Button();
            btnCancel.Text = "Annuleren";
            btnCancel.Width = 120;
            btnCancel.Height = 35;
            btnCancel.Left = 140;
            btnCancel.Top = 8;
            btnCancel.Click += BtnCancel_Click;
            bottomPanel.Controls.Add(btnCancel);
            
            // Top panel with folder creation
            Panel topPanel = new Panel();
            topPanel.Dock = DockStyle.Top;
            topPanel.Height = 70;
            topPanel.Padding = new Padding(5);
            
            Label lblNewFolder = new Label();
            lblNewFolder.Text = "Nieuwe map:";
            lblNewFolder.Left = 5;
            lblNewFolder.Top = 8;
            lblNewFolder.Width = 85;
            lblNewFolder.Height = 20;
            lblNewFolder.TextAlign = ContentAlignment.MiddleRight;
            topPanel.Controls.Add(lblNewFolder);
            
            txtNewFolder = new TextBox();
            txtNewFolder.Left = 95;
            txtNewFolder.Top = 5;
            txtNewFolder.Width = 300;
            txtNewFolder.Height = 20;
            txtNewFolder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            topPanel.Controls.Add(txtNewFolder);
            
            btnCreateFolder = new Button();
            btnCreateFolder.Text = "Maak aan";
            btnCreateFolder.Left = 5;
            btnCreateFolder.Top = 35;
            btnCreateFolder.Width = 100;
            btnCreateFolder.Height = 28;
            btnCreateFolder.Click += BtnCreateFolder_Click;
            topPanel.Controls.Add(btnCreateFolder);
            
            btnRefresh = new Button();
            btnRefresh.Text = "Ververs";
            btnRefresh.Left = 115;
            btnRefresh.Top = 35;
            btnRefresh.Width = 100;
            btnRefresh.Height = 28;
            btnRefresh.Click += BtnRefresh_Click;
            topPanel.Controls.Add(btnRefresh);
            
            btnManagePaths = new Button();
            btnManagePaths.Text = "Beheer Paden...";
            btnManagePaths.Left = 225;
            btnManagePaths.Top = 35;
            btnManagePaths.Width = 130;
            btnManagePaths.Height = 28;
            btnManagePaths.Click += BtnManagePaths_Click;
            topPanel.Controls.Add(btnManagePaths);
            
            // Header label
            lblLibrary = new Label();
            lblLibrary.Text = "Library folders (drop modules hier):";
            lblLibrary.Dock = DockStyle.Top;
            lblLibrary.Height = 30;
            lblLibrary.Font = new Font(lblLibrary.Font, FontStyle.Bold);
            lblLibrary.TextAlign = ContentAlignment.MiddleLeft;
            lblLibrary.Padding = new Padding(5);
            
            // Add all controls to Panel2 - Fill MUST be added FIRST!
            splitContainer.Panel2.Controls.Add(treeLibrary);     // Fill - add FIRST (fills available space)
            splitContainer.Panel2.Controls.Add(topPanel);        // Top - add SECOND
            splitContainer.Panel2.Controls.Add(lblLibrary);      // Top - add THIRD (appears at very top)
            splitContainer.Panel2.Controls.Add(bottomPanel);     // Bottom - add LAST
            
            this.AcceptButton = btnExport;
            this.CancelButton = btnCancel;
            
            // Force layout after form is shown
            this.Load += (s, e) => {
                this.ClientSize = new Size(900, 600);
                splitContainer.SplitterDistance = 250;
                lblNewFolder.Top = 8;
                lblNewFolder.Width = 85;
                txtNewFolder.Left = 95;
                txtNewFolder.Top = 5;
                txtNewFolder.Width = 300;
            };
        }
        
        private void LoadModules()
        {
            try
            {
                lstModules.Items.Clear();
                componentMap.Clear();
                int index = 0;
                
                foreach (VBComponent component in project.VBComponents)
                {
                    // Only exportable components
                    if (component.Type == vbext_ComponentType.vbext_ct_StdModule ||
                        component.Type == vbext_ComponentType.vbext_ct_ClassModule ||
                        component.Type == vbext_ComponentType.vbext_ct_MSForm)
                    {
                        string icon = "";
                        switch (component.Type)
                        {
                            case vbext_ComponentType.vbext_ct_StdModule:
                                icon = "[M] ";
                                break;
                            case vbext_ComponentType.vbext_ct_ClassModule:
                                icon = "[C] ";
                                break;
                            case vbext_ComponentType.vbext_ct_MSForm:
                                icon = "[F] ";
                                break;
                        }
                        
                        lstModules.Items.Add(icon + component.Name, false);
                        componentMap[index] = component;
                        index++;
                    }
                }
                
                lblModules.Text = string.Format("Modules uit project ({0} modules):", lstModules.Items.Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij laden modules:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        private void LoadLibraryTree()
        {
            try
            {
                treeLibrary.Nodes.Clear();
                
                if (libraryPaths == null || libraryPaths.Count == 0)
                {
                    MessageBox.Show(
                        "Geen library paden geconfigureerd.\nGebruik Instellingen om paden toe te voegen.",
                        "Export to Library",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }
                
                // Voeg elke library path toe als root node
                foreach (string libraryPath in libraryPaths)
                {
                    if (!Directory.Exists(libraryPath))
                    {
                        try
                        {
                            Directory.CreateDirectory(libraryPath);
                        }
                        catch
                        {
                            // Skip if can't create (e.g., offline network path)
                            continue;
                        }
                    }
                    
                    string displayName = Path.GetFileName(libraryPath.TrimEnd('\\', '/'));
                    if (string.IsNullOrEmpty(displayName))
                    {
                        displayName = libraryPath; // Voor root paths zoals "C:\\"
                    }
                    
                    TreeNode rootNode = new TreeNode(displayName);
                    rootNode.Tag = libraryPath;
                    rootNode.ImageIndex = 0;
                    rootNode.SelectedImageIndex = 0;
                    treeLibrary.Nodes.Add(rootNode);
                    
                    LoadSubfolders(rootNode, libraryPath);
                    
                    rootNode.Expand();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij laden library folders:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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
                    node.ImageIndex = 0;
                    node.SelectedImageIndex = 0;
                    parentNode.Nodes.Add(node);
                    
                    LoadSubfolders(node, dir);
                }
            }
            catch
            {
                // Skip folders without access
            }
        }
        
        // Drag & Drop implementation
        private void LstModules_MouseDown(object sender, MouseEventArgs e)
        {
            if (lstModules.SelectedItem == null)
                return;
            
            int index = lstModules.SelectedIndex;
            if (!componentMap.ContainsKey(index))
                return;
            
            // Gebruik index in plaats van COM object voor drag & drop
            lstModules.DoDragDrop(index, DragDropEffects.Copy);
        }
        
        private void TreeLibrary_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(int)))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        
        private void TreeLibrary_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(int)))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        
        private void TreeLibrary_DragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(int)))
                return;
            
            int index = (int)e.Data.GetData(typeof(int));
            if (!componentMap.ContainsKey(index))
                return;
            
            VBComponent component = componentMap[index];
            
            // Get drop location node
            Point pt = treeLibrary.PointToClient(new Point(e.X, e.Y));
            TreeNode node = treeLibrary.GetNodeAt(pt);
            
            if (node == null || node.Tag == null)
            {
                MessageBox.Show(
                    "Selecteer een folder in de library tree.",
                    "Export to Library",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            
            string targetFolder = node.Tag.ToString();
            
            try
            {
                string extension = "";
                switch (component.Type)
                {
                    case vbext_ComponentType.vbext_ct_StdModule:
                        extension = ".bas";
                        break;
                    case vbext_ComponentType.vbext_ct_ClassModule:
                        extension = ".cls";
                        break;
                    case vbext_ComponentType.vbext_ct_MSForm:
                        extension = ".frm";
                        break;
                }
                
                string targetPath = Path.Combine(targetFolder, component.Name + extension);
                
                if (File.Exists(targetPath))
                {
                    DialogResult result = MessageBox.Show(
                        string.Format("Bestand '{0}{1}' bestaat al in deze folder.\n\nOverschrijven?", component.Name, extension),
                        "Bestand bestaat",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    
                    if (result != DialogResult.Yes)
                        return;
                    
                    File.Delete(targetPath);
                }
                
                component.Export(targetPath);
                
                // Vraag om optionele beschrijving
                string description = PromptForDescription(component.Name);
                if (!string.IsNullOrEmpty(description))
                {
                    SaveModuleDescription(targetPath, description);
                }
                
                MessageBox.Show(
                    string.Format("Module '{0}' geëxporteerd naar:\n{1}", component.Name, targetPath),
                    "Export Succesvol",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Fout bij exporteren module:\n\n{0}", ex.Message),
                    "Export Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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
            
            Button btnSave = new Button()
            {
                Text = "Opslaan",
                Left = 280,
                Width = 90,
                Top = 170,
                DialogResult = DialogResult.OK
            };
            
            Button btnSkip = new Button()
            {
                Text = "Overslaan",
                Left = 380,
                Width = 90,
                Top = 170,
                DialogResult = DialogResult.Cancel
            };
            
            prompt.Controls.Add(lblInfo);
            prompt.Controls.Add(txtDescription);
            prompt.Controls.Add(btnSave);
            prompt.Controls.Add(btnSkip);
            prompt.AcceptButton = btnSave;
            prompt.CancelButton = btnSkip;
            
            return prompt.ShowDialog() == DialogResult.OK ? txtDescription.Text.Trim() : "";
        }
        
        private void SaveModuleDescription(string modulePath, string description)
        {
            try
            {
                string descriptionPath = Path.ChangeExtension(modulePath, ".txt");
                File.WriteAllText(descriptionPath, description);
            }
            catch
            {
                // Stille fout - beschrijving is optioneel
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
        
        private void BtnCreateFolder_Click(object sender, EventArgs e)
        {
            string folderName = txtNewFolder.Text.Trim();
            if (string.IsNullOrEmpty(folderName))
            {
                MessageBox.Show(
                    "Voer een folder naam in.",
                    "Folder Maken",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            
            try
            {
                TreeNode selectedNode = treeLibrary.SelectedNode ?? treeLibrary.Nodes[0];
                string parentPath = selectedNode.Tag.ToString();
                string newFolderPath = Path.Combine(parentPath, folderName);
                
                if (Directory.Exists(newFolderPath))
                {
                    MessageBox.Show(
                        "Folder bestaat al.",
                        "Folder Maken",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
                
                Directory.CreateDirectory(newFolderPath);
                txtNewFolder.Clear();
                LoadLibraryTree();
                
                MessageBox.Show(
                    string.Format("Folder '{0}' aangemaakt.", folderName),
                    "Folder Gemaakt",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij maken folder:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            LoadLibraryTree();
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
                    
                    // Refresh tree view
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
            if (lstModules.CheckedIndices.Count == 0)
            {
                MessageBox.Show(
                    "Selecteer minimaal één module om te exporteren.\n\n" +
                    "Tip: Gebruik drag & drop om individuele modules naar specifieke folders te exporteren.",
                    "Export to Library",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            
            TreeNode selectedNode = treeLibrary.SelectedNode;
            if (selectedNode == null || selectedNode.Tag == null)
            {
                MessageBox.Show(
                    "Selecteer eerst een folder in de library tree waar je naartoe wilt exporteren.",
                    "Export to Library",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            
            string targetFolder = selectedNode.Tag.ToString();
            
            try
            {
                int exported = 0;
                int skipped = 0;
                System.Text.StringBuilder result = new System.Text.StringBuilder();
                
                foreach (int index in lstModules.CheckedIndices)
                {
                    if (!componentMap.ContainsKey(index))
                        continue;
                    
                    VBComponent component = componentMap[index];
                    
                    string extension = "";
                    switch (component.Type)
                    {
                        case vbext_ComponentType.vbext_ct_StdModule:
                            extension = ".bas";
                            break;
                        case vbext_ComponentType.vbext_ct_ClassModule:
                            extension = ".cls";
                            break;
                        case vbext_ComponentType.vbext_ct_MSForm:
                            extension = ".frm";
                            break;
                    }
                    
                    string targetPath = Path.Combine(targetFolder, component.Name + extension);
                    
                    if (File.Exists(targetPath))
                    {
                        DialogResult overwrite = MessageBox.Show(
                            string.Format("Bestand '{0}{1}' bestaat al.\n\nOverschrijven?", component.Name, extension),
                            "Bestand bestaat",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);
                        
                        if (overwrite == DialogResult.Cancel)
                        {
                            break;
                        }
                        else if (overwrite == DialogResult.No)
                        {
                            result.AppendLine(string.Format("○ Overgeslagen: {0}{1}", component.Name, extension));
                            skipped++;
                            continue;
                        }
                        
                        File.Delete(targetPath);
                    }
                    
                    component.Export(targetPath);
                    result.AppendLine(string.Format("✓ Geëxporteerd: {0}{1}", component.Name, extension));
                    exported++;
                }
                
                MessageBox.Show(
                    string.Format("Modules geëxporteerd naar:\n{0}\n\n", targetFolder) +
                    result.ToString() + string.Format("\n{0} van {1} modules geëxporteerd", exported, lstModules.CheckedIndices.Count),
                    "Export Compleet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij exporteren:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
