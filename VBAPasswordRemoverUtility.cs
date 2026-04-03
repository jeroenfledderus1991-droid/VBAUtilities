using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace VBEAddIn
{
    /// <summary>
    /// Voert tijdelijk een VBA module uit vanuit een open, schrijfbaar werkboek.
    /// De geïnjecteerde module blijft staan voor hergebruik.
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

''zet deze code in module 2
'Sub unprotected(Optional strBook$)
Sub unprotected()
    If Hook Then
        MsgBox ""VBA Project is unprotected!"", vbInformation, String(Len(""VBA Project is unprotected!"") * 1.5, ""*"")
    End If
End Sub




Sub unprotectVBA()
    If Hook Then
        ' MsgBox ""VBA Project is unprotected!"", vbInformation, String(Len(""VBA Project is unprotected!"") * 1.5, ""*"")
    End If
End Sub

Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
    GetPtr = Value
End Function

Public Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal varPtr(OriginBytes(0)), 12
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
        MoveMemory ByVal varPtr(TmpBytes(0)), ByVal pFunc, osi + 1
        If TmpBytes(osi) <> &HB8 Then
            MoveMemory ByVal varPtr(OriginBytes(0)), ByVal pFunc, 12
            p = GetPtr(AddressOf MyDialogBoxParam)
            If osi Then HookBytes(0) = &H48
            HookBytes(osi) = &HB8
            osi = osi + 1
            MoveMemory ByVal varPtr(HookBytes(osi)), ByVal varPtr(p), 4 * osi
            HookBytes(osi + 4 * osi) = &HFF
            HookBytes(osi + 4 * osi + 1) = &HE0
            MoveMemory ByVal pFunc, ByVal varPtr(HookBytes(0)), 12
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
End Function

";

        public static void Execute(VBE vbe)
        {
            string targetProjectName = GetTargetProjectName(vbe);
            string inExcelFailureReason;
            TryUnlockInsideExcel(vbe, targetProjectName, out inExcelFailureReason);
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

            if (string.IsNullOrWhiteSpace(targetProjectName))
            {
                failureReason = "Kon de projectnaam van het actieve VBA project niet bepalen.";
                return false;
            }

            Excel.Application excelApp = null;
            Excel.Workbook hostWorkbook = null;
            VBProject hostProject = null;
            VBComponent injectedModule = null;
            string macroPrefix = string.Empty;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                hostWorkbook = FindWritableHostWorkbook(excelApp);
                if (hostWorkbook == null)
                {
                    failureReason = "Open eerst een onbeveiligd Excel-werkboek waarin tijdelijk een module geplaatst mag worden.";
                    return false;
                }

                hostProject = hostWorkbook.VBProject;

                injectedModule = hostProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                injectedModule.Name = BuildInjectedModuleName(hostProject);
                WriteModuleCode(injectedModule.CodeModule, HookCode);

                string injectedCode = injectedModule.CodeModule.Lines[1, injectedModule.CodeModule.CountOfLines];
                if (!IsInjectedCodeValid(injectedCode))
                {
                    failureReason = "De tijdelijke VBA module is niet correct gevuld. De code in het host-project wijkt af van de broncode.";
                    return false;
                }

                macroPrefix = QuoteWorkbookName(hostWorkbook.Name) + injectedModule.Name + ".";
                excelApp.Run(macroPrefix + "unprotected");
                return true;
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
                ReleaseComObject(injectedModule);
                ReleaseComObject(hostProject);
                ReleaseComObject(hostWorkbook);
                ReleaseComObject(excelApp);
            }
        }

        private static Excel.Workbook FindWritableHostWorkbook(Excel.Application excelApp)
        {
            if (excelApp == null)
            {
                return null;
            }

            Excel.Workbook activeWorkbook = null;
            Excel.Workbooks openWorkbooks = null;

            try
            {
                activeWorkbook = excelApp.ActiveWorkbook;
                if (CanWriteToVbaProject(activeWorkbook))
                {
                    return activeWorkbook;
                }

                openWorkbooks = excelApp.Workbooks;
                for (int index = 1; index <= openWorkbooks.Count; index++)
                {
                    Excel.Workbook workbook = null;

                    try
                    {
                        workbook = openWorkbooks[index];
                        if (workbook != null && !SameWorkbook(workbook, activeWorkbook) && CanWriteToVbaProject(workbook))
                        {
                            return workbook;
                        }
                    }
                    catch
                    {
                        ReleaseComObject(workbook);
                    }
                }
            }
            finally
            {
                ReleaseComObject(openWorkbooks);
            }

            return null;
        }

        private static bool CanWriteToVbaProject(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            VBProject project = null;

            try
            {
                project = workbook.VBProject;
                return project != null && project.Protection == vbext_ProjectProtection.vbext_pp_none;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComObject(project);
            }
        }

        private static bool SameWorkbook(Excel.Workbook left, Excel.Workbook right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            try
            {
                return string.Equals(left.FullName, right.FullName, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static string BuildInjectedModuleName(VBProject project)
        {
            string moduleName = TempModuleName;
            int suffix = 1;

            while (ModuleExists(project, moduleName))
            {
                suffix++;
                moduleName = TempModuleName + suffix;
            }

            return moduleName;
        }

        private static bool ModuleExists(VBProject project, string moduleName)
        {
            if (project == null || string.IsNullOrWhiteSpace(moduleName))
            {
                return false;
            }

            try
            {
                foreach (VBComponent component in project.VBComponents)
                {
                    try
                    {
                        if (string.Equals(component.Name, moduleName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(component);
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static string QuoteWorkbookName(string workbookName)
        {
            return "'" + workbookName.Replace("'", "''") + "'!";
        }

        private static void ReleaseComObject(object comObject)
        {
            try
            {
                if (comObject is Excel.Application)
                {
                    return;
                }

                if (comObject != null
                    && RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                    && Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch { }
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

            return string.Equals(normalizedInjected, normalizedSource, StringComparison.OrdinalIgnoreCase)
                && normalizedInjected.IndexOf("Sub unprotectVBA()", StringComparison.OrdinalIgnoreCase) >= 0
                && normalizedInjected.IndexOf("Private Function MyDialogBoxParam", StringComparison.OrdinalIgnoreCase) >= 0
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

    }
}