using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Utility voor het exporteren van VBA modules naar de code library
    /// </summary>
    public static class ExportToLibraryUtility
    {
        public static void Execute(VBE vbe)
        {
            try
            {
                // Get active project
                VBProject project = vbe.ActiveVBProject;
                if (project == null)
                {
                    MessageBox.Show(
                        "Geen actief VBA project gevonden.",
                        "Export to Library",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
                
                // Check if project is locked
                try
                {
                    var _ = project.VBComponents.Count;
                }
                catch
                {
                    MessageBox.Show(
                        "Het actieve project is vergrendeld met een wachtwoord.\n\n" +
                        "Ontgrendel het project eerst voordat je modules kunt exporteren.",
                        "Project Vergrendeld",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
                
                // Get library paths from settings
                var libraryPaths = FormatterSettings.CodeLibraryPaths;
                if (libraryPaths == null || libraryPaths.Count == 0)
                {
                    string defaultPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        "VBA Code Library"
                    );
                    libraryPaths = new System.Collections.Generic.List<string> { defaultPath };
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
                
                // Show unified code library form (same as Code Library)
                using (UnifiedCodeLibraryForm form = new UnifiedCodeLibraryForm(project, libraryPaths))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij exporteren naar library:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
