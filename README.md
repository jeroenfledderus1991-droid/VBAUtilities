# VBE Code Tools Add-In

Een C# COM Add-in voor de Visual Basic Editor (VBE) die code formatterings functionaliteit en Excel utilities biedt.

## Functionaliteit

### Code Formattering
- **Formatteer Dim Statements**: Zorgt ervoor dat alle `Dim` statements in je VBA code minimaal 1 tab (4 spaties) indentatie hebben.
- **Formatteer Complete Code**: Complete code formatting met indentatie, Dims en blank lines.

### Excel Utilities
- **WhoAmI**: Toon informatie over het actieve workbook (FullName en ReadOnly status).
- **Optimalisatie UIT/AAN**: Schakel Excel optimalisaties uit/aan (events, screenupdating, alerts).
- **Export VBA**: Exporteer alle VBA componenten naar bestanden.
- **Reference Manager**: Beheer VBA references in je project.

## Vereisten
- .NET Framework 4.8.1 (of hoger)
- Microsoft Office met VBA ondersteuning (Excel, Word, Access, etc.)
- Visual Studio 2019 of nieuwer (alleen voor development/compilatie)

## Installatie

### Stap 1: Project Compileren (indien nodig)

Het project is al gecompileerd. De DLL bevindt zich in `bin\Release\VBEAddIn.dll`.

Als je het opnieuw wilt compileren:
1. Open Command Prompt
2. Navigeer naar de project folder
3. Run:
```cmd
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe" VBEAddIn.csproj /p:Configuration=Release
```

### Stap 2: COM Registratie

**BELANGRIJK: Run Command Prompt als Administrator!**

1. Open Command Prompt als Administrator (Rechtermuisknop → "Als administrator uitvoeren")
2. Navigeer naar de bin\Release folder:
```cmd
cd "C:\Users\SBH118\OneDrive - VMI Holland B.V\Documents\VBA C#\bin\Release"
```
3. Registreer de COM DLL:
```cmd
regasm VBEAddIn.dll /codebase
```

Je zou een bericht moeten zien: "Types registered successfully"

### Stap 3: Registry Entries Toevoegen

1. Dubbelklik op `Register.reg` in de project folder
2. Bevestig de registry wijziging door op "Ja" te klikken
3. Je ziet een bevestigingsmelding

### Stap 4: Office Herstarten

Sluit alle Office applicaties volledig af en start ze opnieuw op.

## Gebruik

De add-in is toegankelijk via het menu "Utilities" in de VBA Editor menu balk.

### Menu Structuur

**Utilities Menu:**
- **Formatting** (submenu)
  - Formatteer Dim Statements
  - Formatteer Complete Code
- Instellingen...
- WhoAmI
- Optimalisatie UIT
- Optimalisatie AAN
- Export VBA Componenten
- Reference Manager

### Methode 1: Via VBE Menu (Aanbevolen)

Gebruik het **Utilities** menu in de VBA Editor menubar voor toegang tot alle functies.

### Methode 2: Vanuit VBA Code

Voeg deze code toe aan een VBA module:

```vb
' Formatteer Dim Statements
Sub FormateerAlleDimStatements()
    Dim addin As Object
    On Error Resume Next
    Set addin = Application.COMAddIns("VBEAddIn.Connect").Object
    If Not addin Is Nothing Then
        addin.FormatDimStatements
    Else
        MsgBox "VBE AddIn niet geladen", vbExclamation
    End If
End Sub

' PDF Export
Sub ExporteerNaarPDF()
    Dim addin As Object
    On Error Resume Next
    Set addin = Application.COMAddIns("VBEAddIn.Connect").Object
    If Not addin Is Nothing Then
        addin.ExportToPDF
    Else
        MsgBox "VBE AddIn niet geladen", vbExclamation
    End If
End Sub

' WhoAmI
Sub ToonWorkbookInfo()
    Dim addin As Object
    On Error Resume Next
    Set addin = Application.COMAddIns("VBEAddIn.Connect").Object
    If Not addin Is Nothing Then
        addin.WhoAmI
    Else
        MsgBox "VBE AddIn niet geladen", vbExclamation
    End If
End Sub
```
### Methode 3: Direct vanuit Immediate Window

1. Open de VBA Editor (`Alt+F11`)
2. Open de Immediate Window (`Ctrl+G`)
3. Type een van deze commando's:

```vb
' Formatteer Dim Statements
Application.COMAddIns("VBEAddIn.Connect").Object.FormatDimStatements()

' PDF Export
Application.COMAddIns("VBEAddIn.Connect").Object.ExportToPDF()

' WhoAmI
Application.COMAddIns("VBEAddIn.Connect").Object.WhoAmI()

' Optimalisatie UIT
Application.COMAddIns("VBEAddIn.Connect").Object.OptimalisatieUit()

' Optimalisatie AAN
Application.COMAddIns("VBEAddIn.Connect").Object.OptimalisatieAan()
```

## De-installatie

### Stap 1: Registry Verwijderen
Dubbelklik op `Unregister.reg` om de registry entries te verwijderen

### Stap 2: COM De-registreren
Open Command Prompt als Administrator en run:
```cmd
cd "C:\Users\SBH118\OneDrive - VMI Holland B.V\Documents\VBA C#\bin\Release"
regasm /u VBEAddIn.dll
```
## Licentie

Dit project is vrij te gebruiken en aan te passen voor eigen doeleinden.

## Ondersteuning

Voor vragen of problemen:
1. Check de Troubleshooting sectie
2. Kijk in Windows Event Viewer voor COM/VBA fouten
3. Test de add-in in een nieuw Excel workbook met simpele VBA code

