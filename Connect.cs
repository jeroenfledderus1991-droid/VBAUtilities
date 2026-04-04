using System;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using Extensibility;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Core;

namespace VBEAddIn
{
    /// <summary>
    /// The main Add-in class that implements IDTExtensibility2 for VBE integration
    /// </summary>
    [ComVisible(true)]
    [Guid("B1C2D3E4-F5A6-4B78-C901-D234E5678F90")]
    [ProgId("VBEAddIn.Connect")]
    public class Connect : IDTExtensibility2
    {
        // Singleton instance voor toegang vanuit SettingsForm
        private static Connect _instance;
        public static Connect Instance { get { return _instance; } }
        
        private VBE _vbe;
        private AddIn _addInInstance;
        private CommandBarButton _menuButton;
        private CommandBarButton _menuButtonComplete;
        private CommandBarButton _menuButtonSettings;
        private CommandBarButton _menuButtonWhoAmI;
        private CommandBarButton _menuButtonOptUit;
        private CommandBarButton _menuButtonOptAan;
        private CommandBarButton _menuButtonExportVBA;
        private CommandBarButton _menuButtonReferenceManager;
        private CommandBarButton _menuButtonCodeLibrary;
        private CommandBarButton _menuButtonExportToLibrary;
        private CommandBarButton _menuButtonInsertComment;
        private CommandBarButton _menuButtonPasswordRemover;
        private CommandBarButton _menuButtonChangelog;
        private CommandBarPopup _utilitiesMenu;
        private CommandBarPopup _formattingMenu;
        
        // CommandBar (toolbar) fields
        private CommandBar _commandBar;
        private CommandBarButton _cmdWhoAmI;
        private CommandBarButton _cmdOptUit;
        private CommandBarButton _cmdOptAan;
        private CommandBarButton _cmdFormatDim;
        private CommandBarButton _cmdFormatComplete;
        private CommandBarButton _cmdSettings;
        private CommandBarButton _cmdExportVBA;
        private CommandBarButton _cmdReferenceManager;
        private CommandBarButton _cmdCodeLibrary;
        private CommandBarButton _cmdInsertComment;

        public Connect()
        {
            _instance = this;
        }

        #region IDTExtensibility2 Implementation

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _vbe = (VBE)Application;
                _addInInstance = (AddIn)AddInInst;
                
                // Load settings from registry
                FormatterSettings.LoadFromRegistry();
                
                // Maak menu knop
                CreateMenuButton();
                
                // Maak CommandBar als ingeschakeld
                if (FormatterSettings.ShowCommandBar)
                {
                    CreateCommandBar();
                }

                // Startup-notificaties
                CheckForNewVersion();
                System.Threading.ThreadPool.QueueUserWorkItem(delegate
                {
                    System.Threading.Thread.Sleep(1500);
                    CheckForGitHubUpdate();
                });
            }
            catch (Exception ex)
            {
                string errorMsg = "Fout bij laden: " + ex.Message + "\n\nStack: " + ex.StackTrace;
                System.Diagnostics.Debug.WriteLine("=== VBE AddIn OnConnection Fout ===");
                System.Diagnostics.Debug.WriteLine(errorMsg);
                System.Windows.Forms.MessageBox.Show(errorMsg, "Fout");
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                RemoveMenuButton();
                
                _vbe = null;
                _addInInstance = null;
                _menuButton = null;
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                string errorMsg = "Fout bij afsluiten: " + ex.Message;
                System.Diagnostics.Debug.WriteLine("=== VBE AddIn OnDisconnection Fout ===");
                System.Diagnostics.Debug.WriteLine(errorMsg);
                System.Windows.Forms.MessageBox.Show(errorMsg, "Fout");
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion

        #region Menu Management

        private void CreateMenuButton()
        {
            try
            {
                WriteDebug("=== CreateMenuButton START ===");
                
                // Haal de CommandBars
                CommandBars commandBars = (CommandBars)_vbe.CommandBars;
                WriteDebug("CommandBars opgehaald, aantal: " + commandBars.Count);
                
                // Zoek de menu bar
                CommandBar menuBar = null;
                for (int i = 1; i <= commandBars.Count; i++)
                {
                    try
                    {
                        CommandBar bar = commandBars[i];
                        if (bar.Type == MsoBarType.msoBarTypeMenuBar)
                        {
                            menuBar = bar;
                            WriteDebug("MenuBar gevonden: " + bar.Name);
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteDebug("Fout bij lezen CommandBar " + i + ": " + ex.Message);
                    }
                }
                
                if (menuBar == null)
                {
                    WriteDebug("FOUT: MenuBar niet gevonden");
                    System.Windows.Forms.MessageBox.Show(
                        "Kon menubalk niet vinden in VBE.",
                        "VBE AddIn",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }
                
                // Controleer of Utilities menu al bestaat en verwijder deze
                CommandBarPopup utilitiesMenu = null;
                try
                {
                    for (int i = 1; i <= menuBar.Controls.Count; i++)
                    {
                        try
                        {
                            CommandBarControl ctrl = menuBar.Controls[i];
                            if (ctrl.Tag == "VBEAddIn_UtilitiesMenu")
                            {
                                WriteDebug("Oude Utilities menu gevonden, verwijderen");
                                ctrl.Delete(true);
                                break;
                            }
                        }
                        catch { }
                    }
                }
                catch { }
                
                // Maak nieuw Utilities menu in de menubalk
                WriteDebug("Nieuw Utilities menu maken...");
                utilitiesMenu = (CommandBarPopup)menuBar.Controls.Add(
                    MsoControlType.msoControlPopup,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                utilitiesMenu.Caption = "&Utilities";
                utilitiesMenu.Tag = "VBEAddIn_UtilitiesMenu";
                _utilitiesMenu = utilitiesMenu;
                WriteDebug("Utilities menu toegevoegd aan menubalk");
                
                // Maak Formatting submenu binnen Utilities
                WriteDebug("Formatting submenu maken binnen Utilities...");
                CommandBarPopup formattingMenu = (CommandBarPopup)utilitiesMenu.Controls.Add(
                    MsoControlType.msoControlPopup,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                formattingMenu.Caption = "&Formatting";
                formattingMenu.Tag = "VBEAddIn_FormattingMenu";
                _formattingMenu = formattingMenu;
                WriteDebug("Formatting submenu toegevoegd");
                
                // Voeg knop 1 toe: Formatteer Dim Statements (in Formatting submenu)
                WriteDebug("Knop 1 toevoegen aan Formatting submenu...");
                _menuButton = (CommandBarButton)formattingMenu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                _menuButton.Caption = "Formatteer Dim Statements";
                _menuButton.Tag = "VBEAddIn_FormatDim";
                _menuButton.TooltipText = "Sorteer en lijn Dims uit in huidige procedure";
                _menuButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                
                try
                {
                    _menuButton.FaceId = 3495;
                    WriteDebug("Icon gezet");
                }
                catch (Exception iconEx)
                {
                    WriteDebug("Icon error (geen probleem): " + iconEx.Message);
                }
                
                _menuButton.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonClick);
                WriteDebug("Event handler gekoppeld voor Dim formatter");
                
                // Voeg knop 2 toe: Formatteer Complete Code (in Formatting submenu)
                WriteDebug("Knop 2 toevoegen aan Formatting submenu...");
                _menuButtonComplete = (CommandBarButton)formattingMenu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                _menuButtonComplete.Caption = "Formatteer Complete Code";
                _menuButtonComplete.Tag = "VBEAddIn_FormatComplete";
                _menuButtonComplete.TooltipText = "Formatteer hele module: indentatie, Dims, blank lines";
                _menuButtonComplete.Style = MsoButtonStyle.msoButtonIconAndCaption;
                
                try
                {
                    _menuButtonComplete.FaceId = 2151;
                    WriteDebug("Icon gezet voor complete formatter");
                }
                catch (Exception iconEx)
                {
                    WriteDebug("Icon error (geen probleem): " + iconEx.Message);
                }
                
                _menuButtonComplete.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonCompleteClick);
                WriteDebug("Event handler gekoppeld voor complete formatter");
                
                // Voeg Instellingen toe direct aan Utilities (niet in submenu)
                WriteDebug("Instellingen knop toevoegen aan Utilities menu...");
                _menuButtonSettings = (CommandBarButton)utilitiesMenu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                _menuButtonSettings.Caption = "Instellingen...";
                _menuButtonSettings.Tag = "VBEAddIn_Settings";
                _menuButtonSettings.TooltipText = "Pas add-in instellingen aan";
                _menuButtonSettings.Style = MsoButtonStyle.msoButtonIconAndCaption;
                _menuButtonSettings.BeginGroup = true; // Scheidingslijn voor deze knop
                
                try
                {
                    _menuButtonSettings.FaceId = 642; // Settings icon
                    WriteDebug("Icon gezet voor instellingen");
                }
                catch (Exception iconEx)
                {
                    WriteDebug("Icon error (geen probleem): " + iconEx.Message);
                }
                
                _menuButtonSettings.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonSettingsClick);
                WriteDebug("Event handler gekoppeld voor instellingen");
                
                // Voeg WhoAmI knop toe
                WriteDebug("WhoAmI knop toevoegen...");
                _menuButtonWhoAmI = (CommandBarButton)utilitiesMenu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                _menuButtonWhoAmI.Caption = "WhoAmI";
                _menuButtonWhoAmI.Tag = "VBEAddIn_WhoAmI";
                _menuButtonWhoAmI.TooltipText = "Toon workbook info (FullName en ReadOnly status)";
                _menuButtonWhoAmI.Style = MsoButtonStyle.msoButtonIconAndCaption;
                _menuButtonWhoAmI.BeginGroup = true;
                
                try
                {
                    _menuButtonWhoAmI.FaceId = 1954; // Info icon
                }
                catch { }
                
                _menuButtonWhoAmI.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonWhoAmIClick);
                WriteDebug("Event handler gekoppeld voor WhoAmI");
                
                // Voeg Optimalisatie Uit knop toe
                WriteDebug("Optimalisatie Uit knop toevoegen...");
                _menuButtonOptUit = (CommandBarButton)utilitiesMenu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                _menuButtonOptUit.Caption = "Optimalisatie UIT";
                _menuButtonOptUit.Tag = "VBEAddIn_OptUit";
                _menuButtonOptUit.TooltipText = "Zet Excel optimalisaties uit (events, screenupdating, alerts, etc.)";
                _menuButtonOptUit.Style = MsoButtonStyle.msoButtonIconAndCaption;
                
                try
                {
                    _menuButtonOptUit.FaceId = 211; // Stop icon
                }
                catch { }
                
                _menuButtonOptUit.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonOptUitClick);
                WriteDebug("Event handler gekoppeld voor Optimalisatie Uit");
                
                // Voeg Optimalisatie Aan knop toe
                WriteDebug("Optimalisatie Aan knop toevoegen...");
                _menuButtonOptAan = (CommandBarButton)utilitiesMenu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                _menuButtonOptAan.Caption = "Optimalisatie AAN";
                _menuButtonOptAan.Tag = "VBEAddIn_OptAan";
                _menuButtonOptAan.TooltipText = "Zet Excel optimalisaties weer aan";
                _menuButtonOptAan.Style = MsoButtonStyle.msoButtonIconAndCaption;
                
                try
                {
                    _menuButtonOptAan.FaceId = 210; // Play icon
                }
                catch { }
                
                _menuButtonOptAan.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonOptAanClick);
                WriteDebug("Event handler gekoppeld voor Optimalisatie Aan");
                
                // Voeg Export VBA knop toe
                WriteDebug("Export VBA knop toevoegen...");
                _menuButtonExportVBA = (CommandBarButton)utilitiesMenu.Controls.Add(
                    MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                
                _menuButtonExportVBA.Caption = "Export VBA Componenten";
                _menuButtonExportVBA.Tag = "VBEAddIn_ExportVBA";
                _menuButtonExportVBA.TooltipText = "Exporteer alle VBA modules naar bestanden";
                _menuButtonExportVBA.Style = MsoButtonStyle.msoButtonIconAndCaption;
                
                try
                {
                    _menuButtonExportVBA.FaceId = 620; // Export icon
                }
                catch { }
                
                _menuButtonExportVBA.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonExportVBAClick);
                WriteDebug("Event handler gekoppeld voor Export VBA");

                // Reference Manager button
                _menuButtonReferenceManager = (CommandBarButton)_utilitiesMenu.Controls.Add(
                    Type: MsoControlType.msoControlButton,
                    Id: Type.Missing,
                    Parameter: Type.Missing,
                    Before: Type.Missing,
                    Temporary: true);

                _menuButtonReferenceManager.Caption = "Reference Manager";
                _menuButtonReferenceManager.Tag = "VBEAddIn_RefManager";
                _menuButtonReferenceManager.TooltipText = "Beheer VBA references";
                _menuButtonReferenceManager.Style = MsoButtonStyle.msoButtonIconAndCaption;

                try
                {
                    _menuButtonReferenceManager.FaceId = 433; // References icon
                }
                catch { }

                _menuButtonReferenceManager.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonReferenceManagerClick);
                WriteDebug("Event handler gekoppeld voor Reference Manager");

                // Code Library button
                _menuButtonCodeLibrary = (CommandBarButton)_utilitiesMenu.Controls.Add(
                    Type: MsoControlType.msoControlButton,
                    Id: Type.Missing,
                    Parameter: Type.Missing,
                    Before: Type.Missing,
                    Temporary: true);

                _menuButtonCodeLibrary.Caption = "Code Library";
                _menuButtonCodeLibrary.Tag = "VBEAddIn_CodeLibrary";
                _menuButtonCodeLibrary.TooltipText = "Importeer modules uit code library";
                _menuButtonCodeLibrary.Style = MsoButtonStyle.msoButtonIconAndCaption;

                try
                {
                    _menuButtonCodeLibrary.FaceId = 3059; // Library/templates icon
                }
                catch { }

                _menuButtonCodeLibrary.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonCodeLibraryClick);
                WriteDebug("Event handler gekoppeld voor Code Library");

                // Export to Library menu item is verwijderd - functionaliteit zit nu in Code Library

                // Insert Comment button
                _menuButtonInsertComment = (CommandBarButton)_utilitiesMenu.Controls.Add(
                    Type: MsoControlType.msoControlButton,
                    Id: Type.Missing,
                    Parameter: Type.Missing,
                    Before: Type.Missing,
                    Temporary: true);

                _menuButtonInsertComment.Caption = "Insert Comment";
                _menuButtonInsertComment.Tag = "VBEAddIn_InsertComment";
                _menuButtonInsertComment.TooltipText = "Voeg commentaar met timestamp toe\nNormaal | SHIFT (asterisks) | CTRL (START/END)";
                _menuButtonInsertComment.Style = MsoButtonStyle.msoButtonIconAndCaption;

                try
                {
                    _menuButtonInsertComment.FaceId = 2152; // Comment icon
                }
                catch { }

                _menuButtonInsertComment.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonInsertCommentClick);
                WriteDebug("Event handler gekoppeld voor Insert Comment");

                // Password Remover button
                _menuButtonPasswordRemover = (CommandBarButton)_utilitiesMenu.Controls.Add(
                    Type: MsoControlType.msoControlButton,
                    Id: Type.Missing,
                    Parameter: Type.Missing,
                    Before: Type.Missing,
                    Temporary: true);

                _menuButtonPasswordRemover.Caption = "VBA Wachtwoord Verwijderen";
                _menuButtonPasswordRemover.Tag = "VBEAddIn_PasswordRemover";
                _menuButtonPasswordRemover.TooltipText = "Verwijder wachtwoord van het actieve VBA project";
                _menuButtonPasswordRemover.Style = MsoButtonStyle.msoButtonIconAndCaption;
                _menuButtonPasswordRemover.BeginGroup = true;

                try { _menuButtonPasswordRemover.FaceId = 2162; } catch { }

                _menuButtonPasswordRemover.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonPasswordRemoverClick);
                WriteDebug("Event handler gekoppeld voor Password Remover");

                // Versiegeschiedenis button
                _menuButtonChangelog = (CommandBarButton)_utilitiesMenu.Controls.Add(
                    Type: MsoControlType.msoControlButton,
                    Id: Type.Missing,
                    Parameter: Type.Missing,
                    Before: Type.Missing,
                    Temporary: true);

                _menuButtonChangelog.Caption = "Versiegeschiedenis";
                _menuButtonChangelog.Tag = "VBEAddIn_Changelog";
                _menuButtonChangelog.TooltipText = "Toon versiegeschiedenis van de add-in (v" + ChangelogData.CurrentVersion + ")";
                _menuButtonChangelog.Style = MsoButtonStyle.msoButtonIconAndCaption;

                try { _menuButtonChangelog.FaceId = 433; } catch { }

                _menuButtonChangelog.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonChangelogClick);
                WriteDebug("Event handler gekoppeld voor Versiegeschiedenis");

                WriteDebug("=== CreateMenuButton SUCCESS ===");
            }
            catch (Exception ex)
            {
                string errorMsg = "FOUT: " + ex.Message + 
                    "\n\nStackTrace: " + ex.StackTrace;
                
                WriteDebug("=== VBE AddIn CreateMenuButton FOUT ===");
                WriteDebug(errorMsg);
                System.Diagnostics.Debug.WriteLine(errorMsg);
                
                System.Windows.Forms.MessageBox.Show(
                    "Kon menu knop niet aanmaken: " + ex.Message + "\n\nCheck Immediate Window (Ctrl+G) voor details.\n\nDe add-in werkt nog steeds via VBA code.",
                    "VBE AddIn",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void CreateCommandBar()
        {
            try
            {
                WriteDebug("=== CreateCommandBar START ===");
                
                CommandBars commandBars = (CommandBars)_vbe.CommandBars;
                
                // Verwijder oude CommandBar als deze bestaat
                try
                {
                    CommandBar oldBar = commandBars["VBE Tools"];
                    oldBar.Delete();
                    WriteDebug("Oude CommandBar verwijderd");
                }
                catch { }
                
                // Maak nieuwe CommandBar
                _commandBar = commandBars.Add(
                    "VBE Tools",
                    MsoBarPosition.msoBarTop,
                    Type.Missing,
                    true);
                
                _commandBar.Visible = true;
                WriteDebug("CommandBar aangemaakt");
                
                // Voeg knoppen toe op basis van settings
                if (FormatterSettings.CommandBarShowWhoAmI)
                {
                    _cmdWhoAmI = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdWhoAmI.Caption = "WhoAmI";
                    _cmdWhoAmI.TooltipText = "Toon workbook info";
                    _cmdWhoAmI.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdWhoAmI.FaceId = 1954; } catch { }
                    _cmdWhoAmI.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonWhoAmIClick);
                    WriteDebug("CommandBar: WhoAmI toegevoegd");
                }
                
                if (FormatterSettings.CommandBarShowOptUit)
                {
                    _cmdOptUit = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdOptUit.Caption = "Opt UIT";
                    _cmdOptUit.TooltipText = "Optimalisatie UIT";
                    _cmdOptUit.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdOptUit.FaceId = 211; } catch { }
                    _cmdOptUit.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonOptUitClick);
                    WriteDebug("CommandBar: Opt UIT toegevoegd");
                }
                
                if (FormatterSettings.CommandBarShowOptAan)
                {
                    _cmdOptAan = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdOptAan.Caption = "Opt AAN";
                    _cmdOptAan.TooltipText = "Optimalisatie AAN";
                    _cmdOptAan.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdOptAan.FaceId = 210; } catch { }
                    _cmdOptAan.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonOptAanClick);
                    WriteDebug("CommandBar: Opt AAN toegevoegd");
                }
                
                if (FormatterSettings.CommandBarShowFormatDim)
                {
                    _cmdFormatDim = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdFormatDim.Caption = "Dim";
                    _cmdFormatDim.TooltipText = "Formatteer Dim Statements";
                    _cmdFormatDim.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdFormatDim.FaceId = 3495; } catch { }
                    _cmdFormatDim.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonClick);
                    WriteDebug("CommandBar: Format Dim toegevoegd");
                }
                
                if (FormatterSettings.CommandBarShowFormatComplete)
                {
                    _cmdFormatComplete = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdFormatComplete.Caption = "Format";
                    _cmdFormatComplete.TooltipText = "Formatteer Complete Code";
                    _cmdFormatComplete.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdFormatComplete.FaceId = 2151; } catch { }
                    _cmdFormatComplete.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonCompleteClick);
                    WriteDebug("CommandBar: Format Complete toegevoegd");
                }
                
                if (FormatterSettings.CommandBarShowSettings)
                {
                    _cmdSettings = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdSettings.Caption = "Inst.";
                    _cmdSettings.TooltipText = "Instellingen";
                    _cmdSettings.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdSettings.FaceId = 642; } catch { }
                    _cmdSettings.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonSettingsClick);
                    WriteDebug("CommandBar: Instellingen toegevoegd");
                }
                
                if (FormatterSettings.CommandBarShowExportVBA)
                {
                    _cmdExportVBA = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdExportVBA.Caption = "Export";
                    _cmdExportVBA.TooltipText = "Export VBA Componenten";
                    _cmdExportVBA.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdExportVBA.FaceId = 620; } catch { }
                    _cmdExportVBA.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonExportVBAClick);
                    WriteDebug("CommandBar: Export VBA toegevoegd");
                }
                
                if (FormatterSettings.CommandBarShowReferenceManager)
                {
                    _cmdReferenceManager = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdReferenceManager.Caption = "Ref";
                    _cmdReferenceManager.TooltipText = "Reference Manager";
                    _cmdReferenceManager.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdReferenceManager.FaceId = 433; } catch { }
                    _cmdReferenceManager.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonReferenceManagerClick);
                    WriteDebug("CommandBar: Reference Manager toegevoegd");
                }

                if (FormatterSettings.CommandBarShowCodeLibrary)
                {
                    _cmdCodeLibrary = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdCodeLibrary.Caption = "Library";
                    _cmdCodeLibrary.TooltipText = "Code Library";
                    _cmdCodeLibrary.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdCodeLibrary.FaceId = 3059; } catch { }
                    _cmdCodeLibrary.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonCodeLibraryClick);
                    WriteDebug("CommandBar: Code Library toegevoegd");
                }

                // Export to Library commandbar button is verwijderd - functionaliteit zit nu in Code Library

                if (FormatterSettings.CommandBarShowInsertComment)
                {
                    _cmdInsertComment = (CommandBarButton)_commandBar.Controls.Add(
                        MsoControlType.msoControlButton,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    _cmdInsertComment.Caption = "Comment";
                    _cmdInsertComment.TooltipText = "1Insert Comment(Normaal | SHIFT | CTRL)";
                    _cmdInsertComment.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    try { _cmdInsertComment.FaceId = 2152; } catch { }
                    _cmdInsertComment.Click += new _CommandBarButtonEvents_ClickEventHandler(OnMenuButtonInsertCommentClick);
                    WriteDebug("CommandBar: Insert Comment toegevoegd");
                }
                
                WriteDebug("=== CreateCommandBar SUCCESS ===");
            }
            catch (Exception ex)
            {
                WriteDebug("=== CreateCommandBar FOUT: " + ex.Message);
                System.Diagnostics.Debug.WriteLine("CommandBar fout: " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Refresh de CommandBar op basis van huidige settings (live update)
        /// </summary>
        public void RefreshCommandBar()
        {
            try
            {
                WriteDebug("=== RefreshCommandBar START ===");
                
                // Verwijder oude CommandBar
                RemoveCommandBar();
                
                // Maak nieuwe CommandBar als ingeschakeld
                if (FormatterSettings.ShowCommandBar)
                {
                    CreateCommandBar();
                }
                
                WriteDebug("=== RefreshCommandBar SUCCESS ===");
            }
            catch (Exception ex)
            {
                WriteDebug("=== RefreshCommandBar FOUT: " + ex.Message);
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij verversen CommandBar: " + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void RemoveCommandBar()
        {
            try
            {
                if (_commandBar != null)
                {
                    _commandBar.Delete();
                    _commandBar = null;
                }
                
                // Ook cleanup van individuele buttons
                _cmdWhoAmI = null;
                _cmdOptUit = null;
                _cmdOptAan = null;
                _cmdFormatDim = null;
                _cmdFormatComplete = null;
                _cmdSettings = null;
                _cmdExportVBA = null;
                _cmdReferenceManager = null;
                _cmdCodeLibrary = null;
                _cmdInsertComment = null;
            }
            catch (Exception ex)
            {
                WriteDebug("RemoveCommandBar fout: " + ex.Message);
            }
        }

        private void RemoveMenuButton()
        {
            try
            {
                if (_menuButton != null)
                {
                    _menuButton.Delete(true);
                    _menuButton = null;
                }
                if (_menuButtonComplete != null)
                {
                    _menuButtonComplete.Delete(true);
                    _menuButtonComplete = null;
                }
                if (_menuButtonSettings != null)
                {
                    _menuButtonSettings.Delete(true);
                    _menuButtonSettings = null;
                }
                if (_menuButtonWhoAmI != null)
                {
                    _menuButtonWhoAmI.Delete(true);
                    _menuButtonWhoAmI = null;
                }
                if (_menuButtonOptUit != null)
                {
                    _menuButtonOptUit.Delete(true);
                    _menuButtonOptUit = null;
                }
                if (_menuButtonOptAan != null)
                {
                    _menuButtonOptAan.Delete(true);
                    _menuButtonOptAan = null;
                }
                if (_menuButtonExportVBA != null)
                {
                    _menuButtonExportVBA.Delete(true);
                    _menuButtonExportVBA = null;
                }
                if (_menuButtonReferenceManager != null)
                {
                    _menuButtonReferenceManager.Delete(true);
                    _menuButtonReferenceManager = null;
                }
                if (_menuButtonCodeLibrary != null)
                {
                    _menuButtonCodeLibrary.Delete(true);
                    _menuButtonCodeLibrary = null;
                }
                if (_menuButtonExportToLibrary != null)
                {
                    _menuButtonExportToLibrary.Delete(true);
                    _menuButtonExportToLibrary = null;
                }
                if (_menuButtonInsertComment != null)
                {
                    _menuButtonInsertComment.Delete(true);
                    _menuButtonInsertComment = null;
                }
                if (_menuButtonPasswordRemover != null)
                {
                    _menuButtonPasswordRemover.Delete(true);
                    _menuButtonPasswordRemover = null;
                }
                if (_menuButtonChangelog != null)
                {
                    _menuButtonChangelog.Delete(true);
                    _menuButtonChangelog = null;
                }
                if (_formattingMenu != null)
                {
                    _formattingMenu.Delete(true);
                    _formattingMenu = null;
                }
                if (_utilitiesMenu != null)
                {
                    _utilitiesMenu.Delete(true);
                    _utilitiesMenu = null;
                }
                
                // Cleanup CommandBar
                RemoveCommandBar();
            }
            catch
            {
                // Silently ignore errors during cleanup
            }
        }

        private void OnMenuButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FormatDimStatements();
        }

        private void OnMenuButtonCompleteClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FormatCompleteCode();
        }

        private void OnMenuButtonSettingsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OpenSettings();
        }

        private void OnMenuButtonWhoAmIClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            WhoAmI();
        }

        private void OnMenuButtonOptUitClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OptimalisatieUit();
        }

        private void OnMenuButtonOptAanClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OptimalisatieAan();
        }

        private void OnMenuButtonExportVBAClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ExportVBAComponents();
        }

        private void OnMenuButtonReferenceManagerClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ManageReferences();
        }

        private void OnMenuButtonCodeLibraryClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OpenCodeLibrary();
        }

        private void OnMenuButtonExportToLibraryClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ExportToLibrary();
        }

        private void OnMenuButtonInsertCommentClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            InsertComment();
        }

        private void OnMenuButtonPasswordRemoverClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            RemoveVBAPassword();
        }

        private void OnMenuButtonChangelogClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ShowChangelog();
        }

        #endregion

        #region Public Methods (accessible via COM)

        /// <summary>
        /// Public method that can be called from VBA
        /// </summary>
        [ComVisible(true)]
        public void FormatDimStatements()
        {
            try
            {
                if (_vbe == null || _vbe.ActiveCodePane == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Open eerst een code module in de VBA Editor.", 
                        "Geen actieve code", 
                        System.Windows.Forms.MessageBoxButtons.OK, 
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                var codeModule = _vbe.ActiveCodePane.CodeModule;
                var formatter = new DimFormatter();
                var result = formatter.FormatDimStatements(codeModule);

                System.Windows.Forms.MessageBox.Show(
                    result, 
                    "Dim Formatting Resultaat", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = "Fout bij formatteren: " + ex.Message + "\n\nStackTrace: " + ex.StackTrace;
                System.Diagnostics.Debug.WriteLine("=== VBE AddIn FormatDimStatements Fout ===");
                System.Diagnostics.Debug.WriteLine(errorMsg);
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij formatteren: " + ex.Message, 
                    "Fout", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Formatteer complete code met alle structuren
        /// </summary>
        [ComVisible(true)]
        public void FormatCompleteCode()
        {
            try
            {
                if (_vbe == null || _vbe.ActiveCodePane == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Open eerst een code module in de VBA Editor.", 
                        "Geen actieve code", 
                        System.Windows.Forms.MessageBoxButtons.OK, 
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                var codeModule = _vbe.ActiveCodePane.CodeModule;
                var formatter = new CompleteCodeFormatter();
                var result = formatter.FormatCode(codeModule);

                System.Windows.Forms.MessageBox.Show(
                    result, 
                    "Complete Code Formatting Resultaat", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = "Fout bij formatteren: " + ex.Message + "\n\nStackTrace: " + ex.StackTrace;
                System.Diagnostics.Debug.WriteLine("=== VBE AddIn FormatCompleteCode Fout ===");
                System.Diagnostics.Debug.WriteLine(errorMsg);
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij formatteren: " + ex.Message, 
                    "Fout", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Open instellingen dialog
        /// </summary>
        [ComVisible(true)]
        public void OpenSettings()
        {
            try
            {
                SettingsForm settingsForm = new SettingsForm();
                settingsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                string errorMsg = "Fout bij openen instellingen: " + ex.Message + "\n\nStackTrace: " + ex.StackTrace;
                System.Diagnostics.Debug.WriteLine("=== VBE AddIn Settings Fout ===");
                System.Diagnostics.Debug.WriteLine(errorMsg);
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij openen instellingen: " + ex.Message, 
                    "Fout", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Toon workbook info (FullName en ReadOnly status)
        /// </summary>
        [ComVisible(true)]
        public void WhoAmI()
        {
            WhoAmIUtility.Execute(_vbe);
        }

        /// <summary>
        /// Zet Excel optimalisaties UIT (events, screenupdating, alerts, etc.)
        /// </summary>
        [ComVisible(true)]
        public void OptimalisatieUit()
        {
            OptimalisatieUtility.ZetUit();
        }

        /// <summary>
        /// Zet Excel optimalisaties weer AAN
        /// </summary>
        [ComVisible(true)]
        public void OptimalisatieAan()
        {
            OptimalisatieUtility.ZetAan();
        }

        /// <summary>
        /// Exporteer alle VBA componenten naar bestanden
        /// </summary>
        [ComVisible(true)]
        public void ExportVBAComponents()
        {
            ExportVBAUtility.Execute(_vbe);
        }

        /// <summary>
        /// Public method that can be called from VBA
        /// </summary>
        [ComVisible(true)]
        public void ManageReferences()
        {
            ReferenceManagerUtility.Execute(_vbe);
        }

        /// <summary>
        /// Open code library en importeer modules
        /// </summary>
        [ComVisible(true)]
        public void OpenCodeLibrary()
        {
            CodeLibraryUtility.Execute(_vbe);
        }

        /// <summary>
        /// Exporteer modules naar code library
        /// </summary>
        [ComVisible(true)]
        public void ExportToLibrary()
        {
            ExportToLibraryUtility.Execute(_vbe);
        }

        /// <summary>
        /// Voeg commentaar met timestamp en gebruikersnaam toe
        /// Normaal: simpel commentaar | SHIFT: met asterisks | CTRL: START/END block
        /// </summary>
        [ComVisible(true)]
        public void InsertComment()
        {
            InsertCommentUtility.Execute(_vbe);
        }

        /// <summary>
        /// Verwijder wachtwoord van het actieve VBA project
        /// </summary>
        [ComVisible(true)]
        public void RemoveVBAPassword()
        {
            VBAPasswordRemoverUtility.Execute(_vbe);
        }

        /// <summary>
        /// Toon de versiegeschiedenis van de add-in
        /// </summary>
        [ComVisible(true)]
        public void ShowChangelog()
        {
            try
            {
                new ChangelogForm().ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij openen versiegeschiedenis: " + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helper Methods

        private void WriteDebug(string message)
        {
            // Alleen schrijven naar debug output (geen SendKeys meer)
            System.Diagnostics.Debug.WriteLine(message);
        }

        private void CheckForNewVersion()
        {
            try
            {
                string current = ChangelogData.CurrentVersion;
                string lastSeen = FormatterSettings.LastSeenVersion;

                if (string.Equals(current, lastSeen, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                // Zoek de entry voor de huidige versie
                ChangelogEntry entry = null;
                foreach (ChangelogEntry e in ChangelogData.Entries)
                {
                    if (string.Equals(e.Version, current, StringComparison.OrdinalIgnoreCase))
                    {
                        entry = e;
                        break;
                    }
                }

                if (entry != null)
                {
                    new WhatsNewForm(entry).Show();
                }

                // Markeer als gezien en sla op
                FormatterSettings.LastSeenVersion = current;
                FormatterSettings.SaveToRegistry();
            }
            catch (Exception ex)
            {
                WriteDebug("CheckForNewVersion fout: " + ex.Message);
            }
        }

        private void CheckForGitHubUpdate()
        {
            try
            {
                const int remindEveryDays = 7;
                string current = ChangelogData.CurrentVersion;
                string ignored = FormatterSettings.IgnoredGitHubVersion;
                string lastPromptVersion = FormatterSettings.LastGitHubPromptVersion;
                string lastPromptUtcRaw = FormatterSettings.LastGitHubPromptUtc;

                string latest;
                string releaseUrl;
                string installerUrl;
                string failureReason;

                if (!GitHubReleaseChecker.TryGetLatestRelease(out latest, out releaseUrl, out installerUrl, out failureReason))
                {
                    WriteDebug("GitHub update-check overgeslagen: " + failureReason);
                    return;
                }

                if (!GitHubReleaseChecker.IsRemoteNewer(current, latest))
                {
                    return;
                }

                if (string.Equals(ignored, latest, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                DateTime lastPromptUtc;
                if (string.Equals(lastPromptVersion, latest, StringComparison.OrdinalIgnoreCase)
                    && DateTime.TryParse(lastPromptUtcRaw, null, System.Globalization.DateTimeStyles.RoundtripKind, out lastPromptUtc))
                {
                    if ((DateTime.UtcNow - lastPromptUtc).TotalDays < remindEveryDays)
                    {
                        return;
                    }
                }

                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
                    "Er is een nieuwe versie beschikbaar op GitHub." + Environment.NewLine +
                    "Huidige versie: " + current + Environment.NewLine +
                    "Nieuwste versie: " + latest + Environment.NewLine + Environment.NewLine +
                    "Ja = download installer" + Environment.NewLine +
                    "Nee = niet opnieuw tonen voor versie " + latest + Environment.NewLine +
                    "Annuleren = herinner me over " + remindEveryDays + " dagen",
                    "Update beschikbaar",
                    System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                    System.Windows.Forms.MessageBoxIcon.Information);

                // Registreer dat deze versie nu getoond is.
                FormatterSettings.LastGitHubPromptVersion = latest;
                FormatterSettings.LastGitHubPromptUtc = DateTime.UtcNow.ToString("o");

                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    string urlToOpen = !string.IsNullOrWhiteSpace(installerUrl) ? installerUrl : releaseUrl;
                    if (!string.IsNullOrWhiteSpace(urlToOpen))
                    {
                        Process.Start(urlToOpen);
                    }

                    FormatterSettings.SaveToRegistry();
                }
                else if (result == System.Windows.Forms.DialogResult.No)
                {
                    FormatterSettings.IgnoredGitHubVersion = latest;
                    FormatterSettings.SaveToRegistry();
                }
                else
                {
                    FormatterSettings.SaveToRegistry();
                }
            }
            catch (Exception ex)
            {
                WriteDebug("CheckForGitHubUpdate fout: " + ex.Message);
            }
        }

        #endregion

        #region COM Registration

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // Register in VBE 6.0, 7.0, 7.1 voor zowel 32-bit als 64-bit
                string[] vbeVersions = { "6.0", "7.0", "7.1" };
                string[] addinPaths = { "Addins", "AddIns64" }; // 32-bit en 64-bit
                
                foreach (string version in vbeVersions)
                {
                    foreach (string addinPath in addinPaths)
                    {
                        string regPath = string.Format(@"Software\Microsoft\VBA\VBE\{0}\{1}\VBEAddIn.Connect", 
                            version, addinPath);
                        
                        Microsoft.Win32.Registry.CurrentUser.CreateSubKey(regPath);
                        
                        using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(regPath, true))
                        {
                            if (key != null)
                            {
                                key.SetValue("FriendlyName", "VBE Code Tools");
                                key.SetValue("Description", "VBA Editor Add-in voor code formattering");
                                key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                                key.SetValue("CommandLineSafe", 0, Microsoft.Win32.RegistryValueKind.DWord);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = "Registratie fout: " + ex.Message + "\n\nStackTrace: " + ex.StackTrace;
                System.Diagnostics.Debug.WriteLine("=== VBE AddIn RegisterFunction Fout ===");
                System.Diagnostics.Debug.WriteLine(errorMsg);
                System.Windows.Forms.MessageBox.Show(
                    "Registratie fout: " + ex.Message,
                    "VBE AddIn Registratie",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                string[] vbeVersions = { "6.0", "7.0", "7.1" };
                string[] addinPaths = { "Addins", "AddIns64" };
                
                foreach (string version in vbeVersions)
                {
                    foreach (string addinPath in addinPaths)
                    {
                        try
                        {
                            Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(
                                string.Format(@"Software\Microsoft\VBA\VBE\{0}\{1}\VBEAddIn.Connect", 
                                    version, addinPath),
                                false);
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        #endregion
    }
}
