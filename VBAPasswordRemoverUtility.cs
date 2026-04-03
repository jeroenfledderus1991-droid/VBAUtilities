using System;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace VBEAddIn
{
    /// <summary>
    /// Probeert eerst binnen de huidige Excel sessie een tijdelijk VBA-project te gebruiken
    /// om de bekende hook-code uit te voeren. Als dat niet lukt, valt de utility terug op
    /// het patchen van de DPB-hash in het bestand zelf.
    /// </summary>
    public static class VBAPasswordRemoverUtility
    {
        private const string TempModuleName = "TempVbaUnlocker";

        private static readonly string HookCode = @"Option Explicit

Private Const PAGE_EXECUTE_READWRITE = &H40

Private Declare PtrSafe Sub MoveMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" _
    (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function VirtualProtect Lib ""kernel32"" (lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleHandleA Lib ""kernel32"" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib ""kernel32"" (ByVal hModule As LongPtr, _
    ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function DialogBoxParam Lib ""user32"" Alias ""DialogBoxParamA"" (ByVal hInstance As LongPtr, _
    ByVal pTemplateName As LongPtr, ByVal hwndParent As LongPtr, _
    ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
Dim HookBytes(0 To 11) As Byte
Dim OriginBytes(0 To 11) As Byte
Dim pFunc As LongPtr
Dim Flag As Boolean

Sub unprotected()
    If Hook Then
        MsgBox ""VBA Project is unprotected!"", vbInformation, String(Len(""VBA Project is unprotected!"") * 1.5, ""*"")
    End If
End Sub

Sub unprotectVBA()
    If Hook Then
    End If
End Sub

Public Sub InstallHook()
    unprotectVBA
End Sub

Public Sub RemoveHook()
    RecoverBytes
End Sub

Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
    GetPtr = Value
End Function

Public Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 12
End Sub

Public Function Hook() As Boolean
    Dim TmpBytes(0 To 11) As Byte
    Dim p As LongPtr, osi As Byte
    Dim OriginProtect As LongPtr

    Hook = False

    #If Win64 Then
    osi = 1
    #Else
    osi = 0
    #End If

    pFunc = GetProcAddress(GetModuleHandleA(""user32.dll""), ""DialogBoxParamA"")
    If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, osi + 1
        If TmpBytes(osi) <> &HB8 Then
            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 12
            p = GetPtr(AddressOf MyDialogBoxParam)
            If osi Then HookBytes(0) = &H48
            HookBytes(osi) = &HB8
            osi = osi + 1
            MoveMemory ByVal VarPtr(HookBytes(osi)), ByVal VarPtr(p), 4 * osi
            HookBytes(osi + 4 * osi) = &HFF
            HookBytes(osi + 4 * osi + 1) = &HE0
            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 12
            Flag = True
            Hook = True
        End If
    End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As LongPtr, _
    ByVal pTemplateName As LongPtr, ByVal hwndParent As LongPtr, _
    ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
            hwndParent, lpDialogFunc, dwInitParam)
        Hook
    End If
End Function";

        public static void Execute(VBE vbe)
        {
            string suggestedPath = GetSuggestedPath(vbe);
            string targetProjectName = GetTargetProjectName(vbe);

            string inExcelFailureReason;
            if (TryUnlockInsideExcel(vbe, targetProjectName, out inExcelFailureReason))
            {
                MessageBox.Show(
                    "Het VBA project lijkt binnen de huidige Excel sessie ontgrendeld te zijn.\n\n" +
                    "Als je het wachtwoord permanent wilt verwijderen, open dan VBAProject Properties > Protection, " +
                    "haal het vinkje weg en sla het bestand opnieuw op.",
                    "VBA Wachtwoord Verwijderen",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string message = "Ontgrendelen binnen Excel is niet gelukt.";
            if (!string.IsNullOrEmpty(inExcelFailureReason))
            {
                message += "\n\nReden:\n" + inExcelFailureReason;
            }

            MessageBox.Show(
                message + "\n\nDe utility schakelt nu over op de bestandsmethode.",
                "VBA Wachtwoord Verwijderen",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            PromptAndPatchFile(suggestedPath);
        }

        private static string GetSuggestedPath(VBE vbe)
        {
            try
            {
                if (vbe != null && vbe.ActiveVBProject != null)
                {
                    return vbe.ActiveVBProject.FileName;
                }
            }
            catch { }

            return string.Empty;
        }

        private static string GetTargetProjectName(VBE vbe)
        {
            try
            {
                if (vbe != null && vbe.ActiveVBProject != null)
                {
                    return vbe.ActiveVBProject.Name;
                }
            }
            catch { }

            return string.Empty;
        }

        private static bool TryUnlockInsideExcel(VBE vbe, string targetProjectName, out string failureReason)
        {
            failureReason = string.Empty;

            if (vbe == null || vbe.ActiveVBProject == null)
            {
                failureReason = "Geen actief VBA project gevonden.";
                return false;
            }

            if (string.IsNullOrEmpty(targetProjectName))
            {
                failureReason = "Kon de projectnaam van het actieve VBA project niet bepalen.";
                return false;
            }

            Excel.Application excelApp = null;
            Excel.Workbook tempWorkbook = null;
            VBProject tempProject = null;
            VBProject targetProject = vbe.ActiveVBProject;
            string macroPrefix = string.Empty;
            bool previousDisplayAlerts = true;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                previousDisplayAlerts = excelApp.DisplayAlerts;
                excelApp.DisplayAlerts = false;
                tempWorkbook = excelApp.Workbooks.Add();
                if (tempWorkbook.Windows.Count > 0)
                {
                    tempWorkbook.Windows[1].Visible = false;
                }
                tempProject = tempWorkbook.VBProject;

                VBComponent module = tempProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                module.Name = TempModuleName;
                WriteModuleCode(module.CodeModule, HookCode);

                string injectedCode = module.CodeModule.Lines[1, module.CodeModule.CountOfLines];
                if (!IsInjectedCodeValid(injectedCode))
                {
                    failureReason = "De tijdelijke VBA module is niet correct gevuld. De code in het tijdelijke project wijkt af van de broncode.";
                    return false;
                }

                macroPrefix = QuoteWorkbookName(tempWorkbook.Name) + TempModuleName + ".";
                excelApp.Run(macroPrefix + "InstallHook");

                // Trigger toegang tot het OORSPRONKELIJKE beveiligde project terwijl de hook actief is.
                int componentCount = targetProject.VBComponents.Count;

                try
                {
                    excelApp.Run(macroPrefix + "RemoveHook");
                }
                catch { }

                return componentCount >= 0;
            }
            catch (COMException ex)
            {
                failureReason = BuildComFailureReason(ex);
                return false;
            }
            catch (Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
            finally
            {
                if (excelApp != null && tempWorkbook != null && !string.IsNullOrEmpty(macroPrefix))
                {
                    try
                    {
                        excelApp.Run(macroPrefix + "RemoveHook");
                    }
                    catch { }
                }

                try
                {
                    if (tempWorkbook != null)
                    {
                        tempWorkbook.Close(false);
                    }
                }
                catch { }

                if (excelApp != null)
                {
                    try
                    {
                        excelApp.DisplayAlerts = previousDisplayAlerts;
                    }
                    catch { }
                }
            }
        }

        private static string QuoteWorkbookName(string workbookName)
        {
            return "'" + workbookName.Replace("'", "''") + "'!";
        }

        private static string BuildComFailureReason(COMException ex)
        {
            if (ex.HResult == unchecked((int)0x800A17B4))
            {
                return "Programmatic access to the VBA project is waarschijnlijk uitgeschakeld in Excel Trust Center.";
            }

            if (ex.HResult == unchecked((int)0x800A03EC) || ex.HResult == unchecked((int)0x800AC3D4))
            {
                return "Excel gaf nog steeds een VBA project fout terug tijdens het ontgrendelen.";
            }

            return ex.Message;
        }

        private static void WriteModuleCode(CodeModule codeModule, string code)
        {
            string normalizedCode = NormalizeCode(code);

            if (codeModule.CountOfLines > 0)
            {
                codeModule.DeleteLines(1, codeModule.CountOfLines);
            }

            string[] lines = normalizedCode.Split(new[] { "\r\n" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                codeModule.InsertLines(i + 1, lines[i]);
            }
        }

        private static bool IsInjectedCodeValid(string injectedCode)
        {
            string normalizedInjected = NormalizeCode(injectedCode);
            string normalizedSource = NormalizeCode(HookCode);

            return normalizedInjected == normalizedSource
                && normalizedInjected.Contains("Sub unprotectVBA()")
                && normalizedInjected.Contains("Private Function MyDialogBoxParam")
                && !normalizedInjected.EndsWith("()", StringComparison.Ordinal);
        }

        private static string NormalizeCode(string code)
        {
            return (code ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Trim('\n')
                .Replace("\n", "\r\n");
        }

        private static void PromptAndPatchFile(string suggestedPath)
        {
            var warn = MessageBox.Show(
                "LET OP: Voor de bestandsmethode moet het bestand gesloten zijn in Excel.\n\n" +
                (string.IsNullOrEmpty(suggestedPath) ? string.Empty : "Actief project:\n" + suggestedPath + "\n\n") +
                "Sluit het bestand in Excel, klik daarna OK en selecteer het bestand.",
                "VBA Wachtwoord Verwijderen",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);

            if (warn != DialogResult.OK)
            {
                return;
            }

            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Selecteer het Excel bestand";
                dlg.Filter = "Excel bestanden (*.xlsm;*.xlam;*.xltm;*.xlsb;*.xls;*.xla)|*.xlsm;*.xlam;*.xltm;*.xlsb;*.xls;*.xla|Alle bestanden (*.*)|*.*";

                if (!string.IsNullOrEmpty(suggestedPath) && File.Exists(suggestedPath))
                {
                    dlg.InitialDirectory = Path.GetDirectoryName(suggestedPath);
                    dlg.FileName = Path.GetFileName(suggestedPath);
                }

                if (dlg.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                RemovePasswordFromFile(dlg.FileName);
            }
        }

        private static void RemovePasswordFromFile(string filePath)
        {
            try
            {
                string ext = Path.GetExtension(filePath).ToLower();
                string backupPath = filePath + ".bak";
                File.Copy(filePath, backupPath, true);

                if (ext == ".xlsm" || ext == ".xlam" || ext == ".xltm" || ext == ".xlsb")
                {
                    RemovePasswordZip(filePath);
                }
                else
                {
                    RemovePasswordRawBinary(filePath);
                }

                MessageBox.Show(
                    "Wachtwoord succesvol via bestandspatch verwijderd.\n\n" +
                    "Backup opgeslagen als:\n" + backupPath + "\n\n" +
                    "Open het bestand opnieuw in Excel.\n" +
                    "Als Excel vraagt of het VBA project gereset moet worden, klik dan op 'Ja'.",
                    "Geslaagd",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij verwijderen wachtwoord:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static void RemovePasswordZip(string filePath)
        {
            string tempPath = filePath + ".tmp";
            try
            {
                using (var inputStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                using (var inputArchive = new ZipArchive(inputStream, ZipArchiveMode.Read))
                using (var outputStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                using (var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create))
                {
                    foreach (ZipArchiveEntry entry in inputArchive.Entries)
                    {
                        byte[] data;
                        using (var entryStream = entry.Open())
                        using (var ms = new MemoryStream())
                        {
                            entryStream.CopyTo(ms);
                            data = ms.ToArray();
                        }

                        if (entry.FullName == "xl/vbaProject.bin")
                        {
                            data = PatchDpb(data);
                        }

                        ZipArchiveEntry newEntry = outputArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                        newEntry.LastWriteTime = entry.LastWriteTime;
                        using (var newEntryStream = newEntry.Open())
                        {
                            newEntryStream.Write(data, 0, data.Length);
                        }
                    }
                }

                File.Delete(filePath);
                File.Move(tempPath, filePath);
            }
            catch
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
                throw;
            }
        }

        private static void RemovePasswordRawBinary(string filePath)
        {
            byte[] data = File.ReadAllBytes(filePath);
            data = PatchDpb(data);
            File.WriteAllBytes(filePath, data);
        }

        private static byte[] PatchDpb(byte[] data)
        {
            byte[] search = Encoding.ASCII.GetBytes("DPB=\"");

            for (int i = 0; i <= data.Length - search.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < search.Length; j++)
                {
                    if (data[i + j] != search[j])
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                {
                    data[i + 2] = (byte)'x';
                    return data;
                }
            }

            throw new Exception(
                "Geen DPB-wachtwoordhash gevonden in het bestand.\n\n" +
                "Mogelijke oorzaken:\n" +
                "• Het bestand is niet VBA-wachtwoord beveiligd\n" +
                "• Het bestandsformaat wordt niet ondersteund");
        }
    }
}