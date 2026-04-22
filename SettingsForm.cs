using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace VBEAddIn
{
    public class SettingsForm : Form
    {
        private TabControl tabControl;
        private TabPage tabFormatting;
        private TabPage tabCodeFormatter;
        private TabPage tabGeneral;
        private TabPage tabReferences;
        private TabPage tabComments;

        // Code Formatter tab controls
        private Panel panelFormatterScroll;
        private ComboBox cmbIndentType;
        private ComboBox cmbIndentSize;
        private ComboBox cmbIndentLabelStyle;
        private ComboBox cmbKeywordsCase;
        private NumericUpDown numBlankLinesBetweenProcs;
        private NumericUpDown numBlankLinesAfterDecl;
        private NumericUpDown numKeepBlankLinesMax;
        private CheckBox chkSpacingOperators;
        private CheckBox chkSpacingComma;
        private CheckBox chkSpacingParens;
        private ComboBox cmbCommentStyle;
        private ComboBox cmbOptionExplicit;
        private CheckBox chkTrailingWhitespace;
        private CheckBox chkFinalNewline;
        private Button btnPresetMinimal;
        private Button btnPresetStandard;
        private Button btnPresetStrict;

        // CommandBar – nieuwe formatter-knoppen
        private CheckBox chkCmdFormatProcedure;
        private CheckBox chkCmdFormatFile;
        
        private ListBox lstDimTypes;
        private Button btnUp;
        private Button btnDown;
        private Button btnAdd;
        private Button btnRemove;
        private Button btnRename;
        private Button btnSave;
        private Button btnCancel;
        private CheckBox chkSortDims;
        private CheckBox chkAlignTypes;
        private NumericUpDown numMinSpaces;
        private Label lblMinSpaces;
        
        // CommandBar settings
        private CheckBox chkShowCommandBar;
        private CheckBox chkCmdWhoAmI;
        private CheckBox chkCmdOptUit;
        private CheckBox chkCmdOptAan;
        private CheckBox chkCmdFormatDim;
        private CheckBox chkCmdFormatComplete;
        private CheckBox chkCmdSettings;
        private CheckBox chkCmdExportVBA;
        private CheckBox chkCmdReferenceManager;
        private CheckBox chkCmdCodeLibrary;
        private CheckBox chkCmdExportToLibrary;
        private CheckBox chkCmdInsertComment;
        
        // Comment settings
        private TextBox txtCommentUserName;
        private TextBox txtCommentTemplateNormal;
        private TextBox txtCommentTemplateShift;
        private TextBox txtCommentTemplate;
        private NumericUpDown numCommentLineLength;
        
        // Reference settings
        private CheckBox chkRefMSCOMCTL;
        private CheckBox chkRefMSScriptControl;
        private CheckBox chkRefScriptingRuntime;
        private CheckBox chkRefRegExp;
        private CheckBox chkRefShellControls;
        private CheckBox chkRefMSForms;

        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.Text = "VBE Add-in Instellingen";
            this.Size = new Size(700, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // TabControl
            tabControl = new TabControl
            {
                Location = new Point(10, 10),
                Size = new Size(665, 490),
                Font = new Font("Segoe UI", 9)
            };
            this.Controls.Add(tabControl);

            // Tab: Formatting (Dim sorting)
            tabFormatting = new TabPage("Formatting");
            tabControl.TabPages.Add(tabFormatting);

            // Tab: Code Formatter opties
            tabCodeFormatter = new TabPage("Code Formatter");
            tabControl.TabPages.Add(tabCodeFormatter);

            // Tab: Comments
            tabComments = new TabPage("Comments");
            tabControl.TabPages.Add(tabComments);

            // Tab: Commandbar
            tabGeneral = new TabPage("Commandbar");
            tabControl.TabPages.Add(tabGeneral);

            // Tab: References
            tabReferences = new TabPage("References");
            tabControl.TabPages.Add(tabReferences);

            // === FORMATTING TAB CONTENT ===
            InitializeFormattingTab();

            // === CODE FORMATTER TAB CONTENT ===
            InitializeCodeFormatterTab();

            // === COMMENTS TAB CONTENT ===
            InitializeCommentsTab();

            // === GENERAL TAB CONTENT ===
            InitializeGeneralTab();

            // === REFERENCES TAB CONTENT ===
            InitializeReferencesTab();

            // Save button (onder tabs)
            btnSave = new Button
            {
                Text = "Opslaan",
                Location = new Point(460, 515),
                Size = new Size(100, 35),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                BackColor = ColorTranslator.FromHtml("#0078D4"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            // Cancel button (onder tabs)
            btnCancel = new Button
            {
                Text = "Annuleren",
                Location = new Point(570, 515),
                Size = new Size(100, 35),
                Font = new Font("Segoe UI", 9),
                BackColor = Color.LightGray,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
        }

        private void InitializeFormattingTab()
        {
            // Sorteer Dims checkbox
            chkSortDims = new CheckBox
            {
                Text = "Sorteer Dims op type",
                Location = new Point(20, 20),
                Size = new Size(240, 20),
                Checked = FormatterSettings.SortDimsByType
            };
            tabFormatting.Controls.Add(chkSortDims);

            // Align Types checkbox
            chkAlignTypes = new CheckBox
            {
                Text = "Lijn 'As Type' uit op kolom",
                Location = new Point(20, 50),
                Size = new Size(240, 20),
                Checked = FormatterSettings.AlignAsTypes
            };
            tabFormatting.Controls.Add(chkAlignTypes);

            // Minimum spaces
            lblMinSpaces = new Label
            {
                Text = "Minimaal aantal spaties voor 'As':",
                Location = new Point(20, 80),
                Size = new Size(240, 30)
            };
            tabFormatting.Controls.Add(lblMinSpaces);

            numMinSpaces = new NumericUpDown
            {
                Location = new Point(260, 78),
                Size = new Size(60, 20),
                Minimum = 1,
                Maximum = 10,
                Value = FormatterSettings.MinimumSpaceBeforeAsType
            };
            tabFormatting.Controls.Add(numMinSpaces);

            // Label voor type volgorde
            Label lblTypeOrder = new Label
            {
                Text = "Dim Type Sorteer Volgorde:",
                Location = new Point(20, 120),
                Size = new Size(240, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            tabFormatting.Controls.Add(lblTypeOrder);

            // ListBox voor type volgorde
            lstDimTypes = new ListBox
            {
                Location = new Point(20, 150),
                Size = new Size(300, 280),
                Font = new Font("Consolas", 9)
            };
            tabFormatting.Controls.Add(lstDimTypes);

            // Up button
            btnUp = new Button
            {
                Text = "▲ Omhoog",
                Location = new Point(340, 150),
                Size = new Size(120, 35),
                Font = new Font("Segoe UI", 9)
            };
            btnUp.Click += BtnUp_Click;
            tabFormatting.Controls.Add(btnUp);

            // Down button
            btnDown = new Button
            {
                Text = "▼ Omlaag",
                Location = new Point(340, 195),
                Size = new Size(120, 35),
                Font = new Font("Segoe UI", 9)
            };
            btnDown.Click += BtnDown_Click;
            tabFormatting.Controls.Add(btnDown);

            // Add button
            btnAdd = new Button
            {
                Text = "+ Toevoegen",
                Location = new Point(340, 250),
                Size = new Size(120, 35),
                Font = new Font("Segoe UI", 9),
                BackColor = ColorTranslator.FromHtml("#107C10"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnAdd.FlatAppearance.BorderSize = 0;
            btnAdd.Click += BtnAdd_Click;
            tabFormatting.Controls.Add(btnAdd);

            // Remove button
            btnRemove = new Button
            {
                Text = "− Verwijderen",
                Location = new Point(340, 295),
                Size = new Size(120, 35),
                Font = new Font("Segoe UI", 9),
                BackColor = ColorTranslator.FromHtml("#E81123"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnRemove.FlatAppearance.BorderSize = 0;
            btnRemove.Click += BtnRemove_Click;
            tabFormatting.Controls.Add(btnRemove);

            // Rename button
            btnRename = new Button
            {
                Text = "✎ Hernoemen",
                Location = new Point(340, 340),
                Size = new Size(120, 35),
                Font = new Font("Segoe UI", 9),
                BackColor = ColorTranslator.FromHtml("#FFB900"),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            btnRename.FlatAppearance.BorderSize = 0;
            btnRename.Click += BtnRename_Click;
            tabFormatting.Controls.Add(btnRename);
        }

        private void InitializeCodeFormatterTab()
        {
            panelFormatterScroll = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(658, 462),
                AutoScroll = true
            };
            tabCodeFormatter.Controls.Add(panelFormatterScroll);

            int y = 10;
            int labelW = 220;
            int ctrlX = 240;
            int ctrlW = 310;
            Font boldFont = new Font("Segoe UI", 9, FontStyle.Bold);
            Color sectionColor = Color.FromArgb(0, 78, 140);

            // Helper: sectie-header
            Action<string> addHeader = title =>
            {
                panelFormatterScroll.Controls.Add(new Label
                {
                    Text = title,
                    Location = new Point(10, y),
                    Size = new Size(600, 20),
                    Font = boldFont,
                    ForeColor = sectionColor
                });
                y += 25;
            };
            // Helper: label + combobox op één rij
            Func<string, string[], string, ComboBox> addCombo = (lbl, items, selected) =>
            {
                panelFormatterScroll.Controls.Add(new Label
                {
                    Text = lbl,
                    Location = new Point(20, y + 3),
                    Size = new Size(labelW, 20)
                });
                var cmb = new ComboBox
                {
                    Location = new Point(ctrlX, y),
                    Size = new Size(ctrlW, 23),
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = new Font("Segoe UI", 9)
                };
                cmb.Items.AddRange(items);
                cmb.SelectedItem = selected;
                if (cmb.SelectedIndex < 0 && cmb.Items.Count > 0) cmb.SelectedIndex = 0;
                panelFormatterScroll.Controls.Add(cmb);
                y += 30;
                return cmb;
            };
            // Helper: label + numericupdown
            Func<string, int, int, int, NumericUpDown> addNum = (lbl, val, min, max) =>
            {
                panelFormatterScroll.Controls.Add(new Label
                {
                    Text = lbl,
                    Location = new Point(20, y + 3),
                    Size = new Size(labelW, 20)
                });
                var num = new NumericUpDown
                {
                    Location = new Point(ctrlX, y),
                    Size = new Size(70, 23),
                    Minimum = min,
                    Maximum = max,
                    Value = val,
                    Font = new Font("Segoe UI", 9)
                };
                panelFormatterScroll.Controls.Add(num);
                y += 30;
                return num;
            };
            // Helper: checkbox
            Func<string, bool, CheckBox> addCheck = (lbl, val) =>
            {
                var chk = new CheckBox
                {
                    Text = lbl,
                    Location = new Point(20, y),
                    Size = new Size(580, 20),
                    Checked = val,
                    Font = new Font("Segoe UI", 9)
                };
                panelFormatterScroll.Controls.Add(chk);
                y += 28;
                return chk;
            };

            // === INDENTATIE ===
            addHeader("Indentatie");
            cmbIndentType = addCombo("Inspringing type:", new[] { "spaces", "tabs" }, FormatterSettings.IndentType);
            cmbIndentSize = addCombo("Grootte (spaties):", new[] { "2", "4", "8" }, FormatterSettings.IndentSize.ToString());
            cmbIndentLabelStyle = addCombo("GoTo labels:", new[] { "flush_left", "indent_with_code" }, FormatterSettings.IndentLabelStyle);
            y += 5;

            // === KEYWORDS ===
            addHeader("Keywords");
            cmbKeywordsCase = addCombo("Keyword opmaak:", new[] { "preserve", "pascal", "uppercase", "lowercase" }, FormatterSettings.KeywordsCase);
            y += 5;

            // === BLOKKEN & LEGE REGELS ===
            addHeader("Blokken & Lege Regels");
            numBlankLinesBetweenProcs = addNum("Lege regels tussen procedures:", FormatterSettings.BlockBlankLinesBetweenProcedures, 0, 2);
            numBlankLinesAfterDecl = addNum("Lege regels na declaraties:", FormatterSettings.BlockBlankLinesAfterDeclarations, 0, 2);
            numKeepBlankLinesMax = addNum("Max opeenvolgende lege regels:", FormatterSettings.MiscKeepBlankLinesMax, 0, 3);
            y += 5;

            // === SPATIES ===
            addHeader("Spaties");
            chkSpacingOperators = addCheck("Spaties rond operatoren  (a + b, x = y)", FormatterSettings.SpacingAroundOperators);
            chkSpacingComma = addCheck("Spatie na komma's  (Foo(a, b))", FormatterSettings.SpacingAfterComma);
            chkSpacingParens = addCheck("Spaties binnen haakjes  ( a + b )", FormatterSettings.SpacingInsideParentheses);
            y += 5;

            // === COMMENTAAR ===
            addHeader("Commentaar");
            cmbCommentStyle = addCombo("Commentaar stijl:", new[] { "preserve", "apostrophe", "rem" }, FormatterSettings.CommentStyle);
            y += 5;

            // === DECLARATIES ===
            addHeader("Declaraties");
            cmbOptionExplicit = addCombo("Option Explicit:", new[] { "preserve", "require", "remove" }, FormatterSettings.DeclarationsOptionExplicit);
            y += 5;

            // === DIVERSEN ===
            addHeader("Diversen");
            chkTrailingWhitespace = addCheck("Verwijder trailing spaties (regeleindes)", FormatterSettings.MiscRemoveTrailingWhitespace);
            chkFinalNewline = addCheck("Zorg voor lege regel aan einde van module", FormatterSettings.MiscEnsureFinalNewline);
            y += 10;

            // === PRESETS ===
            addHeader("Presets – snel instellen");
            panelFormatterScroll.Controls.Add(new Label
            {
                Text = "Let op: preset overschrijft bovenstaande opties.",
                Location = new Point(20, y),
                Size = new Size(600, 18),
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.Gray
            });
            y += 22;

            btnPresetMinimal = new Button
            {
                Text = "Minimaal",
                Location = new Point(20, y),
                Size = new Size(110, 30),
                Font = new Font("Segoe UI", 9),
                FlatStyle = FlatStyle.Flat
            };
            btnPresetMinimal.FlatAppearance.BorderSize = 1;
            btnPresetMinimal.Click += (s, e) => ApplyPreset("minimal");
            panelFormatterScroll.Controls.Add(btnPresetMinimal);

            btnPresetStandard = new Button
            {
                Text = "Standaard",
                Location = new Point(140, y),
                Size = new Size(110, 30),
                Font = new Font("Segoe UI", 9),
                BackColor = ColorTranslator.FromHtml("#0078D4"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnPresetStandard.FlatAppearance.BorderSize = 0;
            btnPresetStandard.Click += (s, e) => ApplyPreset("standard");
            panelFormatterScroll.Controls.Add(btnPresetStandard);

            btnPresetStrict = new Button
            {
                Text = "Strikt",
                Location = new Point(260, y),
                Size = new Size(110, 30),
                Font = new Font("Segoe UI", 9),
                BackColor = ColorTranslator.FromHtml("#107C10"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnPresetStrict.FlatAppearance.BorderSize = 0;
            btnPresetStrict.Click += (s, e) => ApplyPreset("strict");
            panelFormatterScroll.Controls.Add(btnPresetStrict);
        }

        private void ApplyPreset(string preset)
        {
            // Reset naar veilige defaults
            cmbIndentType.SelectedItem = "spaces";
            cmbIndentSize.SelectedItem = "4";
            cmbIndentLabelStyle.SelectedItem = "flush_left";
            cmbKeywordsCase.SelectedItem = "preserve";
            numBlankLinesBetweenProcs.Value = 1;
            numBlankLinesAfterDecl.Value = 0;
            numKeepBlankLinesMax.Value = 2;
            chkSpacingOperators.Checked = false;
            chkSpacingComma.Checked = false;
            chkSpacingParens.Checked = false;
            cmbCommentStyle.SelectedItem = "preserve";
            cmbOptionExplicit.SelectedItem = "preserve";
            chkTrailingWhitespace.Checked = false;
            chkFinalNewline.Checked = false;

            if (preset == "standard" || preset == "strict")
            {
                chkSpacingComma.Checked = true;
                chkTrailingWhitespace.Checked = true;
                cmbKeywordsCase.SelectedItem = "pascal";
            }
            if (preset == "strict")
            {
                chkSpacingOperators.Checked = true;
                cmbOptionExplicit.SelectedItem = "require";
                chkFinalNewline.Checked = true;
            }
        }

        private void InitializeGeneralTab()
        {
            // Titel
            Label lblCommandBarTitle = new Label
            {
                Text = "CommandBar Snelkoppelingen:",
                Location = new Point(20, 20),
                Size = new Size(400, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            tabGeneral.Controls.Add(lblCommandBarTitle);

            // Toon CommandBar checkbox
            chkShowCommandBar = new CheckBox
            {
                Text = "Toon CommandBar (toolbar) met geselecteerde items",
                Location = new Point(20, 50),
                Size = new Size(400, 20),
                Checked = FormatterSettings.ShowCommandBar
            };
            chkShowCommandBar.CheckedChanged += ChkShowCommandBar_CheckedChanged;
            tabGeneral.Controls.Add(chkShowCommandBar);

            // Separator
            Label lblSelectItems = new Label
            {
                Text = "Selecteer items voor CommandBar:",
                Location = new Point(20, 85),
                Size = new Size(300, 20),
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.Gray
            };
            tabGeneral.Controls.Add(lblSelectItems);

            // WhoAmI
            chkCmdWhoAmI = new CheckBox
            {
                Text = "WhoAmI",
                Location = new Point(40, 110),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowWhoAmI,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdWhoAmI);

            // Optimalisatie UIT
            chkCmdOptUit = new CheckBox
            {
                Text = "Optimalisatie UIT",
                Location = new Point(40, 135),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowOptUit,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdOptUit);

            // Optimalisatie AAN
            chkCmdOptAan = new CheckBox
            {
                Text = "Optimalisatie AAN",
                Location = new Point(40, 160),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowOptAan,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdOptAan);

            // Formatteer Dim
            chkCmdFormatDim = new CheckBox
            {
                Text = "Formatteer Dim Statements",
                Location = new Point(40, 185),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowFormatDim,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdFormatDim);

            // Formatteer Module (was "Complete Code")
            chkCmdFormatComplete = new CheckBox
            {
                Text = "Formatteer Module",
                Location = new Point(40, 210),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowFormatComplete,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdFormatComplete);

            // Instellingen
            chkCmdSettings = new CheckBox
            {
                Text = "Instellingen",
                Location = new Point(40, 235),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowSettings,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdSettings);

            // Export VBA
            chkCmdExportVBA = new CheckBox
            {
                Text = "Export VBA Componenten",
                Location = new Point(40, 260),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowExportVBA,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdExportVBA);

            // Reference Manager
            chkCmdReferenceManager = new CheckBox
            {
                Text = "Reference Manager",
                Location = new Point(40, 285),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowReferenceManager,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdReferenceManager);

            // Code Library
            chkCmdCodeLibrary = new CheckBox
            {
                Text = "Code Library",
                Location = new Point(40, 310),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowCodeLibrary,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdCodeLibrary);

            // Export to Library - VERBORGEN (functionaliteit in Code Library geïntegreerd)
            chkCmdExportToLibrary = new CheckBox
            {
                Text = "Export to Library",
                Location = new Point(40, 335),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowExportToLibrary,
                Enabled = false,
                Visible = false  // Verborgen - niet meer in gebruik
            };
            tabGeneral.Controls.Add(chkCmdExportToLibrary);

            // Insert Comment
            chkCmdInsertComment = new CheckBox
            {
                Text = "Insert Comment",
                Location = new Point(40, 335),  // Verplaatst naar positie van Export to Library
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowInsertComment,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdInsertComment);

            // Formatteer Procedure
            chkCmdFormatProcedure = new CheckBox
            {
                Text = "Formatteer Procedure",
                Location = new Point(40, 360),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowFormatProcedure,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdFormatProcedure);

            // Formatteer Volledig VBA-bestand
            chkCmdFormatFile = new CheckBox
            {
                Text = "Formatteer Volledig VBA-bestand",
                Location = new Point(40, 385),
                Size = new Size(300, 20),
                Checked = FormatterSettings.CommandBarShowFormatFile,
                Enabled = FormatterSettings.ShowCommandBar
            };
            tabGeneral.Controls.Add(chkCmdFormatFile);

            // Info label
            Label lblInfo = new Label
            {
                Text = "CommandBar wordt direct bijgewerkt na opslaan (geen herstart nodig).",
                Location = new Point(20, 420),
                Size = new Size(450, 40),
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.DarkGreen
            };
            tabGeneral.Controls.Add(lblInfo);
        }

        private void ChkShowCommandBar_CheckedChanged(object sender, EventArgs e)
        {
            bool enabled = chkShowCommandBar.Checked;
            chkCmdWhoAmI.Enabled = enabled;
            chkCmdOptUit.Enabled = enabled;
            chkCmdOptAan.Enabled = enabled;
            chkCmdFormatDim.Enabled = enabled;
            chkCmdFormatComplete.Enabled = enabled;
            chkCmdSettings.Enabled = enabled;
            chkCmdExportVBA.Enabled = enabled;
            chkCmdReferenceManager.Enabled = enabled;
            chkCmdCodeLibrary.Enabled = enabled;
            chkCmdExportToLibrary.Enabled = enabled;
            chkCmdInsertComment.Enabled = enabled;
            chkCmdFormatProcedure.Enabled = enabled;
            chkCmdFormatFile.Enabled = enabled;
        }

        private void InitializeReferencesTab()
        {
            // Header label
            Label lblHeader = new Label
            {
                Text = "Selecteer welke references standaard toegevoegd moeten worden:",
                Location = new Point(20, 20),
                Size = new Size(450, 40),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            tabReferences.Controls.Add(lblHeader);

            int yPosition = 70;

            // MSCOMCTL.OCX
            chkRefMSCOMCTL = new CheckBox
            {
                Text = "MSCOMCTL.OCX (Windows Common Controls)",
                Location = new Point(40, yPosition),
                Size = new Size(450, 20),
                Checked = FormatterSettings.RefEnableMSCOMCTL
            };
            tabReferences.Controls.Add(chkRefMSCOMCTL);
            yPosition += 30;

            // MSScriptControl
            chkRefMSScriptControl = new CheckBox
            {
                Text = "MSScriptControl (Script Control)",
                Location = new Point(40, yPosition),
                Size = new Size(450, 20),
                Checked = FormatterSettings.RefEnableMSScriptControl
            };
            tabReferences.Controls.Add(chkRefMSScriptControl);
            yPosition += 30;

            // Scripting Runtime
            chkRefScriptingRuntime = new CheckBox
            {
                Text = "Microsoft Scripting Runtime (FileSystemObject, Dictionary)",
                Location = new Point(40, yPosition),
                Size = new Size(450, 20),
                Checked = FormatterSettings.RefEnableScriptingRuntime
            };
            tabReferences.Controls.Add(chkRefScriptingRuntime);
            yPosition += 30;

            // VBScript RegExp
            chkRefRegExp = new CheckBox
            {
                Text = "Microsoft VBScript Regular Expressions",
                Location = new Point(40, yPosition),
                Size = new Size(450, 20),
                Checked = FormatterSettings.RefEnableRegExp
            };
            tabReferences.Controls.Add(chkRefRegExp);
            yPosition += 30;

            // Shell Controls
            chkRefShellControls = new CheckBox
            {
                Text = "Microsoft Shell Controls And Automation",
                Location = new Point(40, yPosition),
                Size = new Size(450, 20),
                Checked = FormatterSettings.RefEnableShellControls
            };
            tabReferences.Controls.Add(chkRefShellControls);
            yPosition += 30;

            // MS Forms 2.0
            chkRefMSForms = new CheckBox
            {
                Text = "Microsoft Forms 2.0 Object Library (FM20.DLL)",
                Location = new Point(40, yPosition),
                Size = new Size(450, 20),
                Checked = FormatterSettings.RefEnableMSForms
            };
            tabReferences.Controls.Add(chkRefMSForms);
            yPosition += 40;

            // Info label
            Label lblInfo = new Label
            {
                Text = "Deze instellingen bepalen welke references toegevoegd worden via\n" +
                       "Utilities > Reference Manager.\n\n" +
                       "Al aanwezige references worden overgeslagen.",
                Location = new Point(20, yPosition),
                Size = new Size(450, 80),
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.DarkGreen
            };
            tabReferences.Controls.Add(lblInfo);
        }

        private void InitializeCommentsTab()
        {
            // Header label
            Label lblHeader = new Label
            {
                Text = "Insert Comment Instellingen:",
                Location = new Point(20, 20),
                Size = new Size(450, 30),
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            tabComments.Controls.Add(lblHeader);

            // Info over functionaliteit
            Label lblInfo = new Label
            {
                Text = "Voeg automatisch commentaren met timestamp en uw naam toe aan VBA code.\n" +
                       "Gebruik: Normaal = simpel commentaar | SHIFT = met asterisks | CTRL = START/END block",
                Location = new Point(20, 55),
                Size = new Size(460, 40),
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.DarkBlue
            };
            tabComments.Controls.Add(lblInfo);

            int yPos = 110;

            // Gebruikersnaam
            Label lblUserName = new Label
            {
                Text = "Uw naam:",
                Location = new Point(20, yPos),
                Size = new Size(140, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            tabComments.Controls.Add(lblUserName);

            txtCommentUserName = new TextBox
            {
                Location = new Point(170, yPos - 3),
                Size = new Size(440, 23),
                Font = new Font("Segoe UI", 9),
                Text = FormatterSettings.CommentUserName
            };
            tabComments.Controls.Add(txtCommentUserName);
            yPos += 40;

            // Regel lengte
            Label lblLineLength = new Label
            {
                Text = "Regel lengte (voor * filler):",
                Location = new Point(20, yPos),
                Size = new Size(180, 20)
            };
            tabComments.Controls.Add(lblLineLength);

            numCommentLineLength = new NumericUpDown
            {
                Location = new Point(200, yPos - 3),
                Size = new Size(80, 23),
                Minimum = 60,
                Maximum = 200,
                Value = FormatterSettings.CommentLineLength
            };
            tabComments.Controls.Add(numCommentLineLength);
            yPos += 50;

            // Templates header
            Label lblTemplates = new Label
            {
                Text = "Commentaar Templates:",
                Location = new Point(20, yPos),
                Size = new Size(400, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            tabComments.Controls.Add(lblTemplates);
            yPos += 30;

            // Info over placeholders
            Label lblPlaceholders = new Label
            {
                Text = "Placeholders: {TIMESTAMP} = datum/tijd | {USERNAME} = uw naam | {FILLER} = asterisks | {TYPE} = START/END",
                Location = new Point(20, yPos),
                Size = new Size(600, 30),
                Font = new Font("Segoe UI", 7.5f, FontStyle.Italic),
                ForeColor = Color.Gray
            };
            tabComments.Controls.Add(lblPlaceholders);
            yPos += 35;

            // Template Normal
            Label lblTemplateNormal = new Label
            {
                Text = "Normaal:",
                Location = new Point(20, yPos),
                Size = new Size(140, 20)
            };
            tabComments.Controls.Add(lblTemplateNormal);

            txtCommentTemplateNormal = new TextBox
            {
                Location = new Point(170, yPos - 3),
                Size = new Size(440, 23),
                Font = new Font("Consolas", 9),
                Text = FormatterSettings.CommentTemplateNormal
            };
            tabComments.Controls.Add(txtCommentTemplateNormal);
            yPos += 35;

            // Template Shift
            Label lblTemplateShift = new Label
            {
                Text = "SHIFT (asterisks):",
                Location = new Point(20, yPos),
                Size = new Size(140, 20)
            };
            tabComments.Controls.Add(lblTemplateShift);

            txtCommentTemplateShift = new TextBox
            {
                Location = new Point(170, yPos - 3),
                Size = new Size(440, 23),
                Font = new Font("Consolas", 9),
                Text = FormatterSettings.CommentTemplateShift
            };
            tabComments.Controls.Add(txtCommentTemplateShift);
            yPos += 35;

            // Template CTRL (START/END)
            Label lblTemplate = new Label
            {
                Text = "CTRL (START/END):",
                Location = new Point(20, yPos),
                Size = new Size(140, 20)
            };
            tabComments.Controls.Add(lblTemplate);

            txtCommentTemplate = new TextBox
            {
                Location = new Point(170, yPos - 3),
                Size = new Size(440, 23),
                Font = new Font("Consolas", 9),
                Text = FormatterSettings.CommentTemplate
            };
            tabComments.Controls.Add(txtCommentTemplate);
            yPos += 45;

            // Preview header
            Label lblPreview = new Label
            {
                Text = "Voorbeeld output:",
                Location = new Point(20, yPos),
                Size = new Size(400, 20),
                Font = new Font("Segoe UI", 8, FontStyle.Bold)
            };
            tabComments.Controls.Add(lblPreview);
            yPos += 25;

            // Preview examples
            string exampleTime = DateTime.Now.ToString("yyyyMMdd-HHmm");
            string exampleName = string.IsNullOrWhiteSpace(FormatterSettings.CommentUserName) ? "Jeroen Fledderus" : FormatterSettings.CommentUserName;
            
            Label lblPreviewText = new Label
            {
                Text = string.Format("Normaal: Dim x As Long    '{0} {1} - \n", exampleTime, exampleName) +
                       string.Format("SHIFT: Dim x As Long    '{0} {1} ***...\n", exampleTime, exampleName) +
                       string.Format("CTRL: ' ### START {0} {1} | ***...", exampleTime, exampleName),
                Location = new Point(20, yPos),
                Size = new Size(460, 60),
                Font = new Font("Consolas", 8),
                ForeColor = Color.DarkGreen,
                BackColor = Color.FromArgb(240, 255, 240)
            };
            tabComments.Controls.Add(lblPreviewText);
        }

        private void LoadSettings()
        {
            lstDimTypes.Items.Clear();
            foreach (string type in FormatterSettings.DimTypeSortOrder)
            {
                lstDimTypes.Items.Add(type);
            }
        }

        private void BtnUp_Click(object sender, EventArgs e)
        {
            int selectedIndex = lstDimTypes.SelectedIndex;
            if (selectedIndex > 0)
            {
                string item = lstDimTypes.SelectedItem.ToString();
                lstDimTypes.Items.RemoveAt(selectedIndex);
                lstDimTypes.Items.Insert(selectedIndex - 1, item);
                lstDimTypes.SelectedIndex = selectedIndex - 1;
            }
        }

        private void BtnDown_Click(object sender, EventArgs e)
        {
            int selectedIndex = lstDimTypes.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < lstDimTypes.Items.Count - 1)
            {
                string item = lstDimTypes.SelectedItem.ToString();
                lstDimTypes.Items.RemoveAt(selectedIndex);
                lstDimTypes.Items.Insert(selectedIndex + 1, item);
                lstDimTypes.SelectedIndex = selectedIndex + 1;
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string newType = ShowInputDialog("Type toevoegen", "Voer nieuwe type naam in (bijv. COLLECTION):");
            if (!string.IsNullOrWhiteSpace(newType))
            {
                string upperType = newType.ToUpper().Trim();
                
                // Check for duplicates
                foreach (string existingType in lstDimTypes.Items)
                {
                    if (existingType.Equals(upperType, StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show(
                            "Type '" + upperType + "' bestaat al in de lijst.",
                            "Duplicaat",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }
                }
                
                lstDimTypes.Items.Add(upperType);
                lstDimTypes.SelectedIndex = lstDimTypes.Items.Count - 1;
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            int selectedIndex = lstDimTypes.SelectedIndex;
            if (selectedIndex >= 0)
            {
                string typeName = lstDimTypes.SelectedItem.ToString();
                DialogResult result = MessageBox.Show(
                    "Weet je zeker dat je '" + typeName + "' wilt verwijderen?",
                    "Type verwijderen",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    lstDimTypes.Items.RemoveAt(selectedIndex);
                    
                    // Select next item if available
                    if (lstDimTypes.Items.Count > 0)
                    {
                        if (selectedIndex < lstDimTypes.Items.Count)
                            lstDimTypes.SelectedIndex = selectedIndex;
                        else
                            lstDimTypes.SelectedIndex = lstDimTypes.Items.Count - 1;
                    }
                }
            }
            else
            {
                MessageBox.Show(
                    "Selecteer eerst een type om te verwijderen.",
                    "Geen selectie",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void BtnRename_Click(object sender, EventArgs e)
        {
            int selectedIndex = lstDimTypes.SelectedIndex;
            if (selectedIndex >= 0)
            {
                string oldName = lstDimTypes.SelectedItem.ToString();
                string newName = ShowInputDialog("Type hernoemen", "Nieuwe naam voor '" + oldName + "':", oldName);
                
                if (!string.IsNullOrWhiteSpace(newName))
                {
                    string upperName = newName.ToUpper().Trim();
                    
                    // Check for duplicates (excluding current item)
                    for (int i = 0; i < lstDimTypes.Items.Count; i++)
                    {
                        if (i != selectedIndex && lstDimTypes.Items[i].ToString().Equals(upperName, StringComparison.OrdinalIgnoreCase))
                        {
                            MessageBox.Show(
                                "Type '" + upperName + "' bestaat al in de lijst.",
                                "Duplicaat",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    
                    lstDimTypes.Items[selectedIndex] = upperName;
                    lstDimTypes.SelectedIndex = selectedIndex;
                }
            }
            else
            {
                MessageBox.Show(
                    "Selecteer eerst een type om te hernoemen.",
                    "Geen selectie",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private string ShowInputDialog(string title, string prompt, string defaultValue = "")
        {
            Form inputForm = new Form
            {
                Width = 400,
                Height = 180,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = title,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };

            Label lblPrompt = new Label
            {
                Left = 20,
                Top = 20,
                Width = 340,
                Height = 40,
                Text = prompt
            };

            TextBox txtInput = new TextBox
            {
                Left = 20,
                Top = 60,
                Width = 340,
                Text = defaultValue,
                Font = new Font("Segoe UI", 10)
            };

            Button btnOk = new Button
            {
                Text = "OK",
                Left = 200,
                Width = 75,
                Top = 100,
                DialogResult = DialogResult.OK,
                BackColor = ColorTranslator.FromHtml("#0078D4"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnOk.FlatAppearance.BorderSize = 0;

            Button btnCancelDialog = new Button
            {
                Text = "Annuleren",
                Left = 285,
                Width = 75,
                Top = 100,
                DialogResult = DialogResult.Cancel,
                BackColor = Color.LightGray,
                FlatStyle = FlatStyle.Flat
            };
            btnCancelDialog.FlatAppearance.BorderSize = 0;

            inputForm.Controls.Add(lblPrompt);
            inputForm.Controls.Add(txtInput);
            inputForm.Controls.Add(btnOk);
            inputForm.Controls.Add(btnCancelDialog);
            inputForm.AcceptButton = btnOk;
            inputForm.CancelButton = btnCancelDialog;

            txtInput.Select();
            txtInput.SelectAll();

            return inputForm.ShowDialog(this) == DialogResult.OK ? txtInput.Text : null;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // Save formatter settings
                FormatterSettings.SortDimsByType = chkSortDims.Checked;
                FormatterSettings.AlignAsTypes = chkAlignTypes.Checked;
                FormatterSettings.MinimumSpaceBeforeAsType = (int)numMinSpaces.Value;

                // Save type order
                FormatterSettings.DimTypeSortOrder.Clear();
                foreach (string item in lstDimTypes.Items)
                {
                    FormatterSettings.DimTypeSortOrder.Add(item);
                }

                // Save CommandBar settings
                FormatterSettings.ShowCommandBar = chkShowCommandBar.Checked;
                FormatterSettings.CommandBarShowWhoAmI = chkCmdWhoAmI.Checked;
                FormatterSettings.CommandBarShowOptUit = chkCmdOptUit.Checked;
                FormatterSettings.CommandBarShowOptAan = chkCmdOptAan.Checked;
                FormatterSettings.CommandBarShowFormatDim = chkCmdFormatDim.Checked;
                FormatterSettings.CommandBarShowFormatComplete = chkCmdFormatComplete.Checked;
                FormatterSettings.CommandBarShowSettings = chkCmdSettings.Checked;
                FormatterSettings.CommandBarShowExportVBA = chkCmdExportVBA.Checked;
                FormatterSettings.CommandBarShowReferenceManager = chkCmdReferenceManager.Checked;
                FormatterSettings.CommandBarShowCodeLibrary = chkCmdCodeLibrary.Checked;
                FormatterSettings.CommandBarShowExportToLibrary = chkCmdExportToLibrary.Checked;
                FormatterSettings.CommandBarShowInsertComment = chkCmdInsertComment.Checked;
                FormatterSettings.CommandBarShowFormatProcedure = chkCmdFormatProcedure.Checked;
                FormatterSettings.CommandBarShowFormatFile = chkCmdFormatFile.Checked;

                // Code Formatter instellingen
                FormatterSettings.IndentType = cmbIndentType.SelectedItem?.ToString() ?? "spaces";
                FormatterSettings.IndentSize = int.TryParse(cmbIndentSize.SelectedItem?.ToString(), out int sz) ? sz : 4;
                FormatterSettings.IndentLabelStyle = cmbIndentLabelStyle.SelectedItem?.ToString() ?? "flush_left";
                FormatterSettings.KeywordsCase = cmbKeywordsCase.SelectedItem?.ToString() ?? "preserve";
                FormatterSettings.BlockBlankLinesBetweenProcedures = (int)numBlankLinesBetweenProcs.Value;
                FormatterSettings.BlockBlankLinesAfterDeclarations = (int)numBlankLinesAfterDecl.Value;
                FormatterSettings.MiscKeepBlankLinesMax = (int)numKeepBlankLinesMax.Value;
                FormatterSettings.SpacingAroundOperators = chkSpacingOperators.Checked;
                FormatterSettings.SpacingAfterComma = chkSpacingComma.Checked;
                FormatterSettings.SpacingInsideParentheses = chkSpacingParens.Checked;
                FormatterSettings.CommentStyle = cmbCommentStyle.SelectedItem?.ToString() ?? "preserve";
                FormatterSettings.DeclarationsOptionExplicit = cmbOptionExplicit.SelectedItem?.ToString() ?? "preserve";
                FormatterSettings.MiscRemoveTrailingWhitespace = chkTrailingWhitespace.Checked;
                FormatterSettings.MiscEnsureFinalNewline = chkFinalNewline.Checked;

                // Save Comment settings (alleen als controls bestaan)
                if (txtCommentUserName != null)
                    FormatterSettings.CommentUserName = txtCommentUserName.Text.Trim();
                if (txtCommentTemplateNormal != null)
                    FormatterSettings.CommentTemplateNormal = txtCommentTemplateNormal.Text;
                if (txtCommentTemplateShift != null)
                    FormatterSettings.CommentTemplateShift = txtCommentTemplateShift.Text;
                if (txtCommentTemplate != null)
                    FormatterSettings.CommentTemplate = txtCommentTemplate.Text;
                if (numCommentLineLength != null)
                    FormatterSettings.CommentLineLength = (int)numCommentLineLength.Value;

                // Save Reference settings
                FormatterSettings.RefEnableMSCOMCTL = chkRefMSCOMCTL.Checked;
                FormatterSettings.RefEnableMSScriptControl = chkRefMSScriptControl.Checked;
                FormatterSettings.RefEnableScriptingRuntime = chkRefScriptingRuntime.Checked;
                FormatterSettings.RefEnableRegExp = chkRefRegExp.Checked;
                FormatterSettings.RefEnableShellControls = chkRefShellControls.Checked;
                FormatterSettings.RefEnableMSForms = chkRefMSForms.Checked;

                // Save to registry
                FormatterSettings.SaveToRegistry();

                // Refresh CommandBar live
                if (Connect.Instance != null)
                {
                    Connect.Instance.RefreshCommandBar();
                }

                MessageBox.Show(
                    "Instellingen opgeslagen!\n\n" +
                    "• Formatter instellingen worden direct gebruikt.\n" +
                    "• CommandBar is direct bijgewerkt.",
                    "Opgeslagen",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij opslaan: " + ex.Message,
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
