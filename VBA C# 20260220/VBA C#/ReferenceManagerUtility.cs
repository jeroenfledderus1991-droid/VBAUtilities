using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Utility voor het beheren van VBA references
    /// </summary>
    public static class ReferenceManagerUtility
    {
        /// <summary>
        /// Voeg geselecteerde references toe aan het actieve VBA project
        /// </summary>
        public static void Execute(VBE vbe)
        {
            try
            {
                if (vbe == null || vbe.ActiveVBProject == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Geen actief VBA project gevonden.",
                        "Reference Manager",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                VBProject project = vbe.ActiveVBProject;
                
                int addedCount = 0;
                int skippedCount = 0;
                StringBuilder resultMessage = new StringBuilder();
                resultMessage.AppendLine("Reference Manager resultaat:\n");

                // 1. MSCOMCTL.OCX
                if (FormatterSettings.RefEnableMSCOMCTL)
                {
                    string refPath = @"C:\Windows\SysWOW64\MSCOMCTL.OCX";
                    if (!System.IO.File.Exists(refPath))
                        refPath = @"C:\Windows\System32\MSCOMCTL.OCX";

                    if (System.IO.File.Exists(refPath))
                    {
                        if (AddReferenceFromFile(project, refPath, "MSCOMCTL.OCX"))
                        {
                            addedCount++;
                            resultMessage.AppendLine("✓ MSCOMCTL.OCX toegevoegd");
                        }
                        else
                        {
                            skippedCount++;
                            resultMessage.AppendLine("○ MSCOMCTL.OCX reeds toegevoegd");
                        }
                    }
                    else
                    {
                        resultMessage.AppendLine("✗ MSCOMCTL.OCX niet gevonden");
                    }
                }

                // 2. MSScriptControl
                if (FormatterSettings.RefEnableMSScriptControl)
                {
                    if (AddReferenceFromGuid(project, "{00000534-0000-0010-8000-00AA006D2EA4}", 6, 1, "MSScriptControl"))
                    {
                        addedCount++;
                        resultMessage.AppendLine("✓ MSScriptControl toegevoegd");
                    }
                    else
                    {
                        skippedCount++;
                        resultMessage.AppendLine("○ MSScriptControl reeds toegevoegd");
                    }
                }

                // 3. Scripting Runtime
                if (FormatterSettings.RefEnableScriptingRuntime)
                {
                    if (AddReferenceFromGuid(project, "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0, "Scripting Runtime"))
                    {
                        addedCount++;
                        resultMessage.AppendLine("✓ Scripting Runtime toegevoegd");
                    }
                    else
                    {
                        skippedCount++;
                        resultMessage.AppendLine("○ Scripting Runtime reeds toegevoegd");
                    }
                }

                // 4. VBScript Regular Expressions
                if (FormatterSettings.RefEnableRegExp)
                {
                    if (AddReferenceFromGuid(project, "{0002E157-0000-0000-C000-000000000046}", 5, 3, "VBScript RegExp"))
                    {
                        addedCount++;
                        resultMessage.AppendLine("✓ VBScript RegExp toegevoegd");
                    }
                    else
                    {
                        skippedCount++;
                        resultMessage.AppendLine("○ VBScript RegExp reeds toegevoegd");
                    }
                }

                // 5. Microsoft Shell Controls
                if (FormatterSettings.RefEnableShellControls)
                {
                    if (AddReferenceFromGuid(project, "{B691E011-1797-432E-907A-4D8C69339129}", 6, 1, "Shell Controls"))
                    {
                        addedCount++;
                        resultMessage.AppendLine("✓ Shell Controls toegevoegd");
                    }
                    else
                    {
                        skippedCount++;
                        resultMessage.AppendLine("○ Shell Controls reeds toegevoegd");
                    }
                }

                // 6. Microsoft Forms 2.0 (FM20.DLL)
                if (FormatterSettings.RefEnableMSForms)
                {
                    string refPath = @"C:\Windows\SysWOW64\FM20.DLL";
                    if (!System.IO.File.Exists(refPath))
                        refPath = @"C:\Windows\System32\FM20.DLL";

                    if (System.IO.File.Exists(refPath))
                    {
                        if (AddReferenceFromFile(project, refPath, "MS Forms 2.0"))
                        {
                            addedCount++;
                            resultMessage.AppendLine("✓ MS Forms 2.0 toegevoegd");
                        }
                        else
                        {
                            skippedCount++;
                            resultMessage.AppendLine("○ MS Forms 2.0 reeds toegevoegd");
                        }
                    }
                    else
                    {
                        resultMessage.AppendLine("✗ FM20.DLL niet gevonden");
                    }
                }

                // Toon resultaat
                resultMessage.AppendLine();
                resultMessage.AppendLine("Totaal toegevoegd: " + addedCount);
                resultMessage.AppendLine("Reeds aanwezig: " + skippedCount);
                resultMessage.AppendLine();
                resultMessage.AppendLine("Project: " + project.Name);

                System.Windows.Forms.MessageBox.Show(
                    resultMessage.ToString(),
                    "Reference Manager",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    addedCount > 0 ? System.Windows.Forms.MessageBoxIcon.Information : System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij toevoegen references:\n\n" + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private static bool AddReferenceFromGuid(VBProject project, string guid, int major, int minor, string name)
        {
            try
            {
                // Check of reference al bestaat
                foreach (Reference reference in project.References)
                {
                    if (reference.Guid.Equals(guid, StringComparison.OrdinalIgnoreCase))
                    {
                        return false; // Al toegevoegd
                    }
                }

                // Voeg reference toe
                project.References.AddFromGuid(guid, major, minor);
                return true;
            }
            catch
            {
                return false; // Fout of al toegevoegd
            }
        }

        private static bool AddReferenceFromFile(VBProject project, string filePath, string name)
        {
            try
            {
                // Check of reference al bestaat op basis van path
                foreach (Reference reference in project.References)
                {
                    if (reference.FullPath != null && 
                        reference.FullPath.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        return false; // Al toegevoegd
                    }
                }

                // Voeg reference toe
                project.References.AddFromFile(filePath);
                return true;
            }
            catch
            {
                return false; // Fout of al toegevoegd
            }
        }
    }
}
