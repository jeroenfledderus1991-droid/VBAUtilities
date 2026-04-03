using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Utility voor het beheren van VBA code library
    /// </summary>
    public static class CodeLibraryUtility
    {
        /// <summary>
        /// Open code library en importeer geselecteerde modules
        /// </summary>
        public static void Execute(VBE vbe)
        {
            try
            {
                if (vbe == null || vbe.ActiveVBProject == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Geen actief VBA project gevonden.",
                        "Code Library",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                VBProject project = vbe.ActiveVBProject;
                
                // Haal library paths op
                List<string> libraryPaths = FormatterSettings.CodeLibraryPaths;
                
                if (libraryPaths == null || libraryPaths.Count == 0)
                {
                    // Gebruik default path
                    string defaultPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        "VBA Code Library");
                    
                    libraryPaths = new List<string> { defaultPath };
                    FormatterSettings.CodeLibraryPaths = libraryPaths;
                    FormatterSettings.SaveToRegistry();
                }
                
                // Create library folders if they don't exist
                foreach (string path in libraryPaths)
                {
                    if (!Directory.Exists(path))
                    {
                        try
                        {
                            Directory.CreateDirectory(path);
                        }
                        catch
                        {
                            // Skip if can't create (e.g., network path offline)
                        }
                    }
                }
                
                // Toon unified code library form
                using (UnifiedCodeLibraryForm form = new UnifiedCodeLibraryForm(project, libraryPaths))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij code library:\n\n" + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
