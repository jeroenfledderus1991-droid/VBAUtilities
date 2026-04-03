using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Utility voor het exporteren van VBA componenten
    /// </summary>
    public static class ExportVBAUtility
    {
        /// <summary>
        /// Exporteer alle VBA componenten uit het actieve project
        /// </summary>
        public static void Execute(VBE vbe)
        {
            try
            {
                if (vbe == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "VBE niet beschikbaar.",
                        "Export VBA",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Verzamel ALLE VBProjects direct uit VBE
                List<VBProject> allProjects = new List<VBProject>();
                List<string> projectNames = new List<string>();
                int activeIndex = -1;
                
                // Get Excel Application voor actieve workbook detectie
                Microsoft.Office.Interop.Excel.Application excelApp = null;
                try
                {
                    excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    // Excel app niet nodig als fallback
                }
                
                // Loop door alle VBProjects in VBE (inclusief add-ins!)
                foreach (VBProject proj in vbe.VBProjects)
                {
                    try
                    {
                        // Probeer toegang tot project (test of het niet vergrendeld is)
                        int componentCount = proj.VBComponents.Count;
                        
                        // Voeg project toe aan lijst
                        allProjects.Add(proj);
                        
                        // Maak beschrijvende naam
                        string displayName = proj.Name;
                        
                        // Probeer bestandsnaam toe te voegen als beschikbaar
                        try
                        {
                            if (!string.IsNullOrEmpty(proj.FileName))
                            {
                                string fileName = System.IO.Path.GetFileName(proj.FileName);
                                if (!string.IsNullOrEmpty(fileName))
                                {
                                    displayName = proj.Name + " (" + fileName + ")";
                                }
                            }
                        }
                        catch
                        {
                            // Filename niet beschikbaar, gebruik alleen project naam
                        }
                        
                        projectNames.Add(displayName);
                        
                        // Check of dit het actieve project is
                        if (vbe.ActiveVBProject != null && proj.Name == vbe.ActiveVBProject.Name)
                        {
                            activeIndex = allProjects.Count - 1; // 0-based
                        }
                    }
                    catch
                    {
                        // Skip vergrendelde of ontoegankelijke projecten
                    }
                }
                
                if (allProjects.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Geen toegankelijke VBA projecten gevonden.\n\n" +
                        "Controleer of 'Toegang tot het VBA-objectmodel vertrouwen' is ingeschakeld in Excel opties.",
                        "Export VBA",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Kies welk project te exporteren
                VBProject selectedProject = null;
                int selectedProjectIndex = -1;
                
                // Altijd selectie dialoog tonen
                using (WorkbookSelectionForm form = new WorkbookSelectionForm(projectNames, activeIndex))
                {
                    if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        selectedProjectIndex = form.SelectedIndex;
                        selectedProject = allProjects[selectedProjectIndex];
                    }
                    else
                    {
                        // Gebruiker heeft geannuleerd
                        return;
                    }
                }

                if (selectedProject == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Geen VBA project geselecteerd.",
                        "Export VBA",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Kies export folder
                using (var folderDialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    folderDialog.Description = "Selecteer een map om VBA-componenten te exporteren";
                    folderDialog.ShowNewFolderButton = true;

                    if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Export geannuleerd.",
                            "Export VBA",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return;
                    }

                    string exportPath = folderDialog.SelectedPath;
                    if (!exportPath.EndsWith("\\"))
                        exportPath += "\\";

                    // Ensure directory exists
                    if (!Directory.Exists(exportPath))
                        Directory.CreateDirectory(exportPath);

                    // Export components
                    int exportCount = 0;
                    string exportedFiles = "";

                    foreach (VBComponent component in selectedProject.VBComponents)
                    {
                        string fileName = "";
                        string extension = "";

                        switch (component.Type)
                        {
                            case vbext_ComponentType.vbext_ct_StdModule:
                                extension = ".bas";
                                fileName = component.Name + extension;
                                break;

                            case vbext_ComponentType.vbext_ct_ClassModule:
                                extension = ".cls";
                                fileName = component.Name + extension;
                                break;

                            case vbext_ComponentType.vbext_ct_MSForm:
                                extension = ".frm";
                                fileName = component.Name + extension;
                                break;

                            case vbext_ComponentType.vbext_ct_Document:
                                // Sheet modules and ThisWorkbook
                                extension = ".bas";
                                fileName = component.Name + extension;
                                break;

                            default:
                                continue; // Skip unknown types
                        }

                        if (!string.IsNullOrEmpty(fileName))
                        {
                            string fullPath = exportPath + fileName;
                            component.Export(fullPath);
                            exportCount++;
                            exportedFiles += "  • " + fileName + "\n";
                        }
                    }

                    if (exportCount > 0)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Export geslaagd!\n\n" +
                            "Project: " + selectedProject.Name + "\n" +
                            "Locatie: " + exportPath + "\n\n" +
                            "Geëxporteerde bestanden (" + exportCount + "):\n" + exportedFiles,
                            "Export VBA",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Geen VBA componenten gevonden om te exporteren in " + selectedProject.Name + ".",
                            "Export VBA",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij exporteren VBA componenten:\n\n" + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
