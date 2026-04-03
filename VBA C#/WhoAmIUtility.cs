using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Utility voor het tonen van workbook informatie
    /// </summary>
    public static class WhoAmIUtility
    {
        /// <summary>
        /// Toon workbook info (FullName en ReadOnly status) - werkt ook voor XLAM, XLSM, etc.
        /// Toont info van het actieve VBA project in VBE (niet per se de ActiveWorkbook)
        /// </summary>
        public static void Execute(VBE vbe)
        {
            try
            {
                // Get Excel Application via COM
                var excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                
                if (excelApp == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Excel applicatie niet gevonden.",
                        "WhoAmI",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                var sb = new StringBuilder();
                sb.AppendLine("=== WhoAmI - Actief VBA Project ===");
                sb.AppendLine();

                bool foundActiveProject = false;

                // PRIMAIR: Info van actieve VBE project (waar je NU in werkt)
                try
                {
                    if (vbe != null && vbe.ActiveVBProject != null)
                    {
                        var project = vbe.ActiveVBProject;
                        foundActiveProject = true;
                        
                        sb.AppendLine(">>> ACTIEF IN VBE <<<");
                        sb.AppendLine("Project: " + project.Name);
                        
                        // Haal FileName direct op (net als ExportVBAUtility doet)
                        string projectFileName = "";
                        try
                        {
                            projectFileName = project.FileName;
                        }
                        catch { }
                        
                        if (!string.IsNullOrEmpty(projectFileName))
                        {
                            sb.AppendLine("Bestand: " + System.IO.Path.GetFileName(projectFileName));
                            sb.AppendLine("Volledig pad: " + projectFileName);
                            
                            // Bepaal type op basis van extensie
                            string ext = System.IO.Path.GetExtension(projectFileName).ToUpper();
                            string fileType = "Onbekend";
                            switch (ext)
                            {
                                case ".XLAM": fileType = "Excel Add-in (macro-enabled)"; break;
                                case ".XLSM": fileType = "Excel Workbook (macro-enabled)"; break;
                                case ".XLSX": fileType = "Excel Workbook"; break;
                                case ".XLTM": fileType = "Excel Template (macro-enabled)"; break;
                                case ".XLS": fileType = "Excel 97-2003 Workbook"; break;
                                case ".XLA": fileType = "Excel 97-2003 Add-in"; break;
                            }
                            sb.AppendLine("Type: " + fileType);
                            sb.AppendLine("Map: " + System.IO.Path.GetDirectoryName(projectFileName));
                        }
                        else
                        {
                            sb.AppendLine("(Pad niet beschikbaar - mogelijk niet opgeslagen)");
                        }
                        
                        // Check protection
                        try
                        {
                            if (project.Protection == Microsoft.Vbe.Interop.vbext_ProjectProtection.vbext_pp_locked)
                            {
                                sb.AppendLine("Status: LOCKED (beveiligd met wachtwoord)");
                            }
                            else
                            {
                                sb.AppendLine("Status: Open voor bewerking");
                            }
                        }
                        catch { }
                        
                        sb.AppendLine();
                    }
                }
                catch { }

                if (!foundActiveProject)
                {
                    sb.AppendLine("Geen actief VBA project gevonden in VBE.");
                    sb.AppendLine();
                }

                // SECUNDAIR: Info van ActiveWorkbook (voor context)
                if (excelApp.ActiveWorkbook != null)
                {
                    var workbook = excelApp.ActiveWorkbook;
                    sb.AppendLine("Excel Active Workbook:");
                    sb.AppendLine("  Name: " + workbook.Name);
                    
                    if (!workbook.IsAddin)
                    {
                        sb.AppendLine("  Path: " + workbook.FullName);
                        sb.AppendLine("  Read-Only: " + (workbook.ReadOnly ? "Ja" : "Nee"));
                        sb.AppendLine("  Saved: " + (workbook.Saved ? "Ja" : "Nee (onopgeslagen)"));
                    }
                    else
                    {
                        sb.AppendLine("  (Dit is een Add-in)");
                    }
                    sb.AppendLine();
                }

                // Toon alle geopende workbooks/add-ins (voor overzicht)
                if (excelApp.Workbooks.Count > 0)
                {
                    sb.AppendLine("Alle geopende bestanden (" + excelApp.Workbooks.Count + "):");
                    foreach (Microsoft.Office.Interop.Excel.Workbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            string wbInfo = "  • " + wb.Name;
                            if (wb.IsAddin)
                            {
                                wbInfo += " (Add-in)";
                            }
                            if (wb.ReadOnly)
                            {
                                wbInfo += " [R/O]";
                            }
                            sb.AppendLine(wbInfo);
                        }
                        catch { }
                    }
                }

                string message = sb.ToString();

                System.Windows.Forms.MessageBox.Show(
                    message,
                    "WhoAmI - Actief VBA Project",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij WhoAmI: " + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
