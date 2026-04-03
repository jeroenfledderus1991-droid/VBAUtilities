using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Utility om VBA project wachtwoorden te verwijderen via een VBA macro snippet.
    /// De methode injecteert tijdelijk een module in het actieve project dat de
    /// hook-gebaseerde unlock procedure bevat, voert deze uit, en verwijdert de module daarna.
    /// </summary>
    public static class VBAPasswordRemoverUtility
    {
        private const string TempModuleName = "_TempUnlocker";

        /// <summary>
        /// VBA code die de hook-techniek gebruikt om het wachtwoord van het huidige project te verwijderen.
        /// </summary>
        private static readonly string VbaUnlockCode = @"
Option Explicit

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

Sub RunUnlocker()
    If Hook() Then
        MsgBox ""VBA Project is unprotected!"", vbInformation, String(Len(""VBA Project is unprotected!"") * 1.5, ""*"")
    End If
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
End Function
";

        /// <summary>
        /// Injecteert een tijdelijk VBA-module in het actieve VBProject, voert de unlock-macro uit,
        /// en verwijdert de module daarna weer.
        /// </summary>
        public static void Execute(VBE vbe)
        {
            VBProject project = null;
            VBComponent tempModule = null;

            try
            {
                if (vbe == null || vbe.ActiveVBProject == null)
                {
                    MessageBox.Show(
                        "Geen actief VBA project gevonden.",
                        "VBA Wachtwoord Verwijderen",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                project = vbe.ActiveVBProject;

                // Check of het project überhaupt bereikbaar is (b.v. vergrendeld maar leesbaar voor injectie)
                try
                {
                    var _ = project.Name;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Het VBA project is niet beschikbaar:\n\n" + ex.Message,
                        "VBA Wachtwoord Verwijderen",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Verwijder bestaande temp module als die nog bestaat
                RemoveTempModule(project);

                // Voeg tijdelijke module toe
                tempModule = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                tempModule.Name = TempModuleName;
                tempModule.CodeModule.AddFromString(VbaUnlockCode);

                // Voer de unlock macro uit via Excel Application.Run
                var excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                excelApp.Run(project.Name + "." + TempModuleName + ".RunUnlocker");
            }
            catch (COMException comEx) when (comEx.HResult == unchecked((int)0x800A03EC))
            {
                MessageBox.Show(
                    "Het VBA project is wachtwoord-beveiligd.\n\n" +
                    "Zorg dat het project open is (verwijder het wachtwoord handmatig via Tools > VBAProject Properties) " +
                    "of gebruik de hook-methode in een unlocked project.",
                    "Project Vergrendeld",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij verwijderen wachtwoord:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // Verwijder altijd de temp module
                if (project != null)
                {
                    RemoveTempModule(project);
                }
            }
        }

        private static void RemoveTempModule(VBProject project)
        {
            try
            {
                foreach (VBComponent comp in project.VBComponents)
                {
                    if (comp.Name == TempModuleName)
                    {
                        project.VBComponents.Remove(comp);
                        break;
                    }
                }
            }
            catch { }
        }
    }
}
