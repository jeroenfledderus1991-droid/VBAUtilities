using System;
using System.Collections.Generic;

namespace VBEAddIn
{
    /// <summary>
    /// Configuratie voor code formatting
    /// </summary>
    public class FormatterSettings
    {
        /// <summary>
        /// Volgorde waarin Dim types gesorteerd moeten worden
        /// Types die niet in deze lijst staan komen onderaan (alfabetisch)
        /// </summary>
        public static List<string> DimTypeSortOrder = new List<string>
        {
            // Basis types eerst
            "BOOLEAN",
            "BYTE",
            "INTEGER",
            "LONG",
            "LONGLONG",
            "SINGLE",
            "DOUBLE",
            "CURRENCY",
            "DECIMAL",
            "DATE",
            "STRING",
            
            // Object types
            "OBJECT",
            "VARIANT",
            
            // Excel specifiek
            "WORKSHEET",
            "WORKBOOK",
            "RANGE",
            "COLLECTION",
            "DICTIONARY",
            
            // Algemeen
            "VARIANT"
        };

        /// <summary>
        /// Minimale spaties tussen variabele naam en "As Type"
        /// </summary>
        public static int MinimumSpaceBeforeAsType = 1;

        /// <summary>
        /// Of Dim statements gesorteerd moeten worden op type
        /// </summary>
        public static bool SortDimsByType = true;

        /// <summary>
        /// Of "As Type" moet worden uitgelijnd op dezelfde kolom
        /// </summary>
        public static bool AlignAsTypes = true;

        // CommandBar Settings
        /// <summary>
        /// Of de CommandBar (toolbar) getoond moet worden
        /// </summary>
        public static bool ShowCommandBar = false;

        /// <summary>
        /// Of WhoAmI in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowWhoAmI = false;

        /// <summary>
        /// Of Optimalisatie UIT in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowOptUit = false;

        /// <summary>
        /// Of Optimalisatie AAN in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowOptAan = false;

        /// <summary>
        /// Of Formatteer Dim Statements in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowFormatDim = false;

        /// <summary>
        /// Of Formatteer Complete Code in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowFormatComplete = false;

        /// <summary>
        /// Of Instellingen in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowSettings = false;

        /// <summary>
        /// Of Export VBA in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowExportVBA = false;

        /// <summary>
        /// Of Reference Manager in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowReferenceManager = false;

        /// <summary>
        /// Of Code Library in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowCodeLibrary = false;

        /// <summary>
        /// Of Export to Library in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowExportToLibrary = false;

        /// <summary>
        /// Of Insert Comment in de CommandBar moet staan
        /// </summary>
        public static bool CommandBarShowInsertComment = false;

        // Insert Comment Settings
        /// <summary>
        /// Gebruikersnaam voor commentaren
        /// </summary>
        public static string CommentUserName = "";

        /// <summary>
        /// Template voor normale commentaren
        /// Placeholders: {TIMESTAMP}, {USERNAME}
        /// </summary>
        public static string CommentTemplateNormal = "'{TIMESTAMP} {USERNAME} - ";

        /// <summary>
        /// Template voor SHIFT commentaren (met asterisks)
        /// Placeholders: {TIMESTAMP}, {USERNAME}, {FILLER}
        /// </summary>
        public static string CommentTemplateShift = "'{TIMESTAMP} {USERNAME} {FILLER}";

        /// <summary>
        /// Template voor CTRL START/END blokken
        /// Placeholders: {TIMESTAMP}, {USERNAME}, {TYPE}, {FILLER}
        /// </summary>
        public static string CommentTemplate = "' ### {TYPE} {TIMESTAMP} {USERNAME} | {FILLER}";

        /// <summary>
        /// Lengte van commentaar regel voor filler berekening
        /// </summary>
        public static int CommentLineLength = 100;

        // Reference Manager Settings
        /// <summary>
        /// Of MSCOMCTL.OCX standaard toegevoegd moet worden
        /// </summary>
        public static bool RefEnableMSCOMCTL = false;

        /// <summary>
        /// Of MSScriptControl standaard toegevoegd moet worden
        /// </summary>
        public static bool RefEnableMSScriptControl = false;

        /// <summary>
        /// Of Scripting Runtime standaard toegevoegd moet worden
        /// </summary>
        public static bool RefEnableScriptingRuntime = false;

        /// <summary>
        /// Of VBScript RegExp standaard toegevoegd moet worden
        /// </summary>
        public static bool RefEnableRegExp = false;

        /// <summary>
        /// Of Shell Controls standaard toegevoegd moet worden
        /// </summary>
        public static bool RefEnableShellControls = false;

        /// <summary>
        /// Of MS Forms 2.0 standaard toegevoegd moet worden
        /// </summary>
        public static bool RefEnableMSForms = false;

        // Code Library Settings
        /// <summary>
        /// Paden naar VBA code library mappen (bijv. persoonlijk + gedeeld)
        /// </summary>
        public static List<string> CodeLibraryPaths = new List<string>();
        
        /// <summary>
        /// Backwards compatibility: oude single path property
        /// </summary>
        public static string CodeLibraryPath
        {
            get { return CodeLibraryPaths.Count > 0 ? CodeLibraryPaths[0] : ""; }
            set
            {
                if (CodeLibraryPaths.Count == 0)
                    CodeLibraryPaths.Add(value);
                else
                    CodeLibraryPaths[0] = value;
            }
        }

        #region Registry Persistence

        private const string RegistryPath = @"Software\VBEAddIn\Settings";

        /// <summary>
        /// Laatste versie die de gebruiker heeft gezien bij opstarten.
        /// Wordt gebruikt om de "Wat is er nieuw?" melding te tonen.
        /// </summary>
        public static string LastSeenVersion = string.Empty;

        /// <summary>
        /// Laad settings uit registry
        /// </summary>
        public static void LoadFromRegistry()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    if (key != null)
                    {
                        // Formatter settings
                        MinimumSpaceBeforeAsType = (int)key.GetValue("MinimumSpaceBeforeAsType", 1);
                        SortDimsByType = ((int)key.GetValue("SortDimsByType", 1)) == 1;
                        AlignAsTypes = ((int)key.GetValue("AlignAsTypes", 1)) == 1;

                        // CommandBar settings
                        ShowCommandBar = ((int)key.GetValue("ShowCommandBar", 0)) == 1;
                        CommandBarShowWhoAmI = ((int)key.GetValue("CommandBarShowWhoAmI", 0)) == 1;
                        CommandBarShowOptUit = ((int)key.GetValue("CommandBarShowOptUit", 0)) == 1;
                        CommandBarShowOptAan = ((int)key.GetValue("CommandBarShowOptAan", 0)) == 1;
                        CommandBarShowFormatDim = ((int)key.GetValue("CommandBarShowFormatDim", 0)) == 1;
                        CommandBarShowFormatComplete = ((int)key.GetValue("CommandBarShowFormatComplete", 0)) == 1;
                        CommandBarShowSettings = ((int)key.GetValue("CommandBarShowSettings", 0)) == 1;
                        CommandBarShowExportVBA = ((int)key.GetValue("CommandBarShowExportVBA", 0)) == 1;
                        CommandBarShowReferenceManager = ((int)key.GetValue("CommandBarShowReferenceManager", 0)) == 1;
                        CommandBarShowCodeLibrary = ((int)key.GetValue("CommandBarShowCodeLibrary", 0)) == 1;
                        CommandBarShowExportToLibrary = ((int)key.GetValue("CommandBarShowExportToLibrary", 0)) == 1;
                        CommandBarShowInsertComment = ((int)key.GetValue("CommandBarShowInsertComment", 0)) == 1;

                        // Insert Comment settings
                        CommentUserName = (string)key.GetValue("CommentUserName", "");
                        CommentTemplateNormal = (string)key.GetValue("CommentTemplateNormal", "'{TIMESTAMP} {USERNAME} - ");
                        CommentTemplateShift = (string)key.GetValue("CommentTemplateShift", "'{TIMESTAMP} {USERNAME} {FILLER}");
                        CommentTemplate = (string)key.GetValue("CommentTemplate", "' ### {TYPE} {TIMESTAMP} {USERNAME} | {FILLER}");
                        CommentLineLength = (int)key.GetValue("CommentLineLength", 100);

                        // Reference Manager settings
                        RefEnableMSCOMCTL = ((int)key.GetValue("RefEnableMSCOMCTL", 0)) == 1;
                        RefEnableMSScriptControl = ((int)key.GetValue("RefEnableMSScriptControl", 0)) == 1;
                        RefEnableScriptingRuntime = ((int)key.GetValue("RefEnableScriptingRuntime", 0)) == 1;
                        RefEnableRegExp = ((int)key.GetValue("RefEnableRegExp", 0)) == 1;
                        RefEnableShellControls = ((int)key.GetValue("RefEnableShellControls", 0)) == 1;
                        RefEnableMSForms = ((int)key.GetValue("RefEnableMSForms", 0)) == 1;

                        // Code Library settings
                        string pathsData = (string)key.GetValue("CodeLibraryPaths", "");
                        if (!string.IsNullOrEmpty(pathsData))
                        {
                            CodeLibraryPaths = new List<string>(pathsData.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries));
                        }
                        else
                        {
                            // Backwards compatibility: migreer oude single path
                            string oldPath = (string)key.GetValue("CodeLibraryPath", "");
                            if (!string.IsNullOrEmpty(oldPath))
                            {
                                CodeLibraryPaths = new List<string> { oldPath };
                            }
                        }

                        // DimTypeSortOrder wordt niet in registry opgeslagen (te complex)

                        // Versienotificatie
                        LastSeenVersion = (string)key.GetValue("LastSeenVersion", "");
                    }
                }
            }
            catch
            {
                // Als registry lezen mislukt, gebruik defaults
            }
        }

        /// <summary>
        /// Sla settings op naar registry
        /// </summary>
        public static void SaveToRegistry()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegistryPath))
                {
                    if (key != null)
                    {
                        // Formatter settings
                        key.SetValue("MinimumSpaceBeforeAsType", MinimumSpaceBeforeAsType, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("SortDimsByType", SortDimsByType ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("AlignAsTypes", AlignAsTypes ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);

                        // CommandBar settings
                        key.SetValue("ShowCommandBar", ShowCommandBar ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowWhoAmI", CommandBarShowWhoAmI ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowOptUit", CommandBarShowOptUit ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowOptAan", CommandBarShowOptAan ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowFormatDim", CommandBarShowFormatDim ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowFormatComplete", CommandBarShowFormatComplete ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowSettings", CommandBarShowSettings ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowExportVBA", CommandBarShowExportVBA ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowReferenceManager", CommandBarShowReferenceManager ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowCodeLibrary", CommandBarShowCodeLibrary ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowExportToLibrary", CommandBarShowExportToLibrary ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("CommandBarShowInsertComment", CommandBarShowInsertComment ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);

                        // Insert Comment settings
                        key.SetValue("CommentUserName", CommentUserName, Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("CommentTemplateNormal", CommentTemplateNormal, Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("CommentTemplateShift", CommentTemplateShift, Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("CommentTemplate", CommentTemplate, Microsoft.Win32.RegistryValueKind.String);
                        key.SetValue("CommentLineLength", CommentLineLength, Microsoft.Win32.RegistryValueKind.DWord);

                        // Reference Manager settings
                        key.SetValue("RefEnableMSCOMCTL", RefEnableMSCOMCTL ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("RefEnableMSScriptControl", RefEnableMSScriptControl ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("RefEnableScriptingRuntime", RefEnableScriptingRuntime ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("RefEnableRegExp", RefEnableRegExp ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("RefEnableShellControls", RefEnableShellControls ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("RefEnableMSForms", RefEnableMSForms ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);

                        // Code Library settings
                        string pathsData = string.Join("|", CodeLibraryPaths.ToArray());
                        key.SetValue("CodeLibraryPaths", pathsData, Microsoft.Win32.RegistryValueKind.String);

                        // Versienotificatie
                        key.SetValue("LastSeenVersion", LastSeenVersion, Microsoft.Win32.RegistryValueKind.String);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Kan settings niet opslaan: " + ex.Message);
            }
        }

        #endregion
    }
}
