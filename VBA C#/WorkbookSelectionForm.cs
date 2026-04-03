using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace VBEAddIn
{
    /// <summary>
    /// Form voor het selecteren van een workbook met radio buttons
    /// </summary>
    public class WorkbookSelectionForm : Form
    {
        private List<RadioButton> radioButtons;
        private Button btnOK;
        private Button btnCancel;
        private Label lblInfo;
        
        public int SelectedIndex { get; private set; }
        
        public WorkbookSelectionForm(List<string> workbookNames, int activeIndex)
        {
            SelectedIndex = -1;
            radioButtons = new List<RadioButton>();
            
            InitializeForm(workbookNames, activeIndex);
        }
        
        private void InitializeForm(List<string> workbookNames, int activeIndex)
        {
            // Form properties
            this.Text = "Selecteer Workbook";
            this.Width = 550;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Info label
            lblInfo = new Label();
            lblInfo.Text = "Selecteer de workbook/add-in waarvan je de VBA componenten wilt exporteren:";
            lblInfo.Left = 20;
            lblInfo.Top = 20;
            lblInfo.Width = this.Width - 50;
            lblInfo.Height = 40;
            this.Controls.Add(lblInfo);
            
            // Panel voor radio buttons met scrollbar
            Panel panel = new Panel();
            panel.Left = 20;
            panel.Top = 70;
            panel.Width = this.Width - 60;
            panel.Height = Math.Min(workbookNames.Count * 30, 300); // Max 300 pixels, daarna scroll
            panel.AutoScroll = true;
            panel.BorderStyle = BorderStyle.FixedSingle;
            this.Controls.Add(panel);
            
            // Radio buttons in panel
            int yPosition = 5;
            for (int i = 0; i < workbookNames.Count; i++)
            {
                RadioButton rb = new RadioButton();
                rb.Text = workbookNames[i];
                rb.Left = 10;
                rb.Top = yPosition;
                rb.Width = panel.Width - 40;
                rb.Tag = i; // Store index
                rb.AutoSize = false;
                
                // Markeer actieve workbook
                if (i == activeIndex)
                {
                    rb.Text += "  (actief)";
                    rb.Checked = true;
                    rb.Font = new Font(rb.Font, FontStyle.Bold);
                }
                
                radioButtons.Add(rb);
                panel.Controls.Add(rb);
                
                yPosition += 30;
            }
            
            // Buttons positie onder panel
            int buttonTop = panel.Top + panel.Height + 20;
            
            // OK button
            btnOK = new Button();
            btnOK.Text = "OK";
            btnOK.Width = 100;
            btnOK.Height = 30;
            btnOK.Left = this.Width - 250;
            btnOK.Top = buttonTop;
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);
            
            // Cancel button
            btnCancel = new Button();
            btnCancel.Text = "Annuleren";
            btnCancel.Width = 100;
            btnCancel.Height = 30;
            btnCancel.Left = this.Width - 140;
            btnCancel.Top = buttonTop;
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
            
            // Set Accept and Cancel buttons
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
            
            // Set form height based on content
            this.Height = buttonTop + 80; // Button height + margins
        }
        
        private void BtnOK_Click(object sender, EventArgs e)
        {
            // Find which radio button is selected
            foreach (RadioButton rb in radioButtons)
            {
                if (rb.Checked)
                {
                    SelectedIndex = (int)rb.Tag;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                    return;
                }
            }
            
            // Fallback - geen selectie
            MessageBox.Show(
                "Selecteer een workbook.",
                "Selectie vereist",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
