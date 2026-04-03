using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace VBEAddIn
{
    /// <summary>
    /// Form voor het beheren van meerdere library paden
    /// </summary>
    public class LibraryPathsForm : Form
    {
        private ListBox lstPaths;
        private Button btnAdd;
        private Button btnRemove;
        private Button btnMoveUp;
        private Button btnMoveDown;
        private Button btnOK;
        private Button btnCancel;
        private Label lblInfo;
        
        public List<string> LibraryPaths { get; private set; }
        
        public LibraryPathsForm(List<string> currentPaths)
        {
            LibraryPaths = new List<string>(currentPaths ?? new List<string>());
            InitializeForm();
            LoadPaths();
        }
        
        private void InitializeForm()
        {
            this.Text = "Code Library Paden";
            this.Width = 600;
            this.Height = 450;
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            lblInfo = new Label();
            lblInfo.Text = "Beheer je code library paden (persoonlijk + gedeeld):\nVolgorde bepaalt prioriteit bij zoeken.";
            lblInfo.Left = 20;
            lblInfo.Top = 20;
            lblInfo.Width = 540;
            lblInfo.Height = 40;
            this.Controls.Add(lblInfo);
            
            lstPaths = new ListBox();
            lstPaths.Left = 20;
            lstPaths.Top = 70;
            lstPaths.Width = 540;
            lstPaths.Height = 220;
            lstPaths.SelectedIndexChanged += LstPaths_SelectedIndexChanged;
            this.Controls.Add(lstPaths);
            
            // Buttons panel
            int buttonTop = 300;
            int buttonLeft = 20;
            
            btnAdd = new Button();
            btnAdd.Text = "Pad Toevoegen...";
            btnAdd.Left = buttonLeft;
            btnAdd.Top = buttonTop;
            btnAdd.Width = 120;
            btnAdd.Click += BtnAdd_Click;
            this.Controls.Add(btnAdd);
            
            btnRemove = new Button();
            btnRemove.Text = "Verwijderen";
            btnRemove.Left = buttonLeft + 130;
            btnRemove.Top = buttonTop;
            btnRemove.Width = 100;
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);
            
            btnMoveUp = new Button();
            btnMoveUp.Text = "▲ Omhoog";
            btnMoveUp.Left = buttonLeft + 240;
            btnMoveUp.Top = buttonTop;
            btnMoveUp.Width = 90;
            btnMoveUp.Click += BtnMoveUp_Click;
            this.Controls.Add(btnMoveUp);
            
            btnMoveDown = new Button();
            btnMoveDown.Text = "▼ Omlaag";
            btnMoveDown.Left = buttonLeft + 340;
            btnMoveDown.Top = buttonTop;
            btnMoveDown.Width = 90;
            btnMoveDown.Click += BtnMoveDown_Click;
            this.Controls.Add(btnMoveDown);
            
            btnOK = new Button();
            btnOK.Text = "OK";
            btnOK.Left = 350;
            btnOK.Top = 360;
            btnOK.Width = 100;
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);
            
            btnCancel = new Button();
            btnCancel.Text = "Annuleren";
            btnCancel.Left = 460;
            btnCancel.Top = 360;
            btnCancel.Width = 100;
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
            
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
            
            UpdateButtonStates();
        }
        
        private void LoadPaths()
        {
            lstPaths.Items.Clear();
            foreach (string path in LibraryPaths)
            {
                lstPaths.Items.Add(path);
            }
        }
        
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Selecteer een code library map:";
                dialog.ShowNewFolderButton = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string path = dialog.SelectedPath;
                    
                    if (LibraryPaths.Contains(path))
                    {
                        MessageBox.Show(
                            "Dit pad is al toegevoegd.",
                            "Dubbel pad",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }
                    
                    LibraryPaths.Add(path);
                    lstPaths.Items.Add(path);
                    lstPaths.SelectedIndex = lstPaths.Items.Count - 1;
                    UpdateButtonStates();
                }
            }
        }
        
        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (lstPaths.SelectedIndex >= 0)
            {
                int index = lstPaths.SelectedIndex;
                LibraryPaths.RemoveAt(index);
                lstPaths.Items.RemoveAt(index);
                
                if (lstPaths.Items.Count > 0)
                {
                    lstPaths.SelectedIndex = Math.Min(index, lstPaths.Items.Count - 1);
                }
                
                UpdateButtonStates();
            }
        }
        
        private void BtnMoveUp_Click(object sender, EventArgs e)
        {
            int index = lstPaths.SelectedIndex;
            if (index > 0)
            {
                string temp = LibraryPaths[index];
                LibraryPaths[index] = LibraryPaths[index - 1];
                LibraryPaths[index - 1] = temp;
                
                LoadPaths();
                lstPaths.SelectedIndex = index - 1;
                UpdateButtonStates();
            }
        }
        
        private void BtnMoveDown_Click(object sender, EventArgs e)
        {
            int index = lstPaths.SelectedIndex;
            if (index >= 0 && index < LibraryPaths.Count - 1)
            {
                string temp = LibraryPaths[index];
                LibraryPaths[index] = LibraryPaths[index + 1];
                LibraryPaths[index + 1] = temp;
                
                LoadPaths();
                lstPaths.SelectedIndex = index + 1;
                UpdateButtonStates();
            }
        }
        
        private void LstPaths_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateButtonStates();
        }
        
        private void UpdateButtonStates()
        {
            bool hasSelection = lstPaths.SelectedIndex >= 0;
            int index = lstPaths.SelectedIndex;
            
            btnRemove.Enabled = hasSelection;
            btnMoveUp.Enabled = hasSelection && index > 0;
            btnMoveDown.Enabled = hasSelection && index < lstPaths.Items.Count - 1;
        }
        
        private void BtnOK_Click(object sender, EventArgs e)
        {
            // Valideer dat alle paden bestaan
            List<string> invalidPaths = new List<string>();
            foreach (string path in LibraryPaths)
            {
                if (!Directory.Exists(path))
                {
                    invalidPaths.Add(path);
                }
            }
            
            if (invalidPaths.Count > 0)
            {
                DialogResult result = MessageBox.Show(
                    string.Format("De volgende paden bestaan niet:\n\n{0}\n\nToch opslaan?",
                        string.Join("\n", invalidPaths.ToArray())),
                    "Ongeldige paden",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                
                if (result != DialogResult.Yes)
                    return;
            }
            
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
