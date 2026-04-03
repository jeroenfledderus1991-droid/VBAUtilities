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
- **PDF Export met Inhoudsopgave**: Exporteer geselecteerde sheets naar PDF met automatische inhoudsopgave en bookmarks voor navigatie.

## Vereisten

- .NET Framework 4.8 (of hoger)
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
- **PDF Export met Inhoudsopgave**

### PDF Export Functionaliteit

De PDF Export functie stelt je in staat om meerdere worksheets te exporteren naar één PDF bestand met een automatische inhoudsopgave en navigeerbare bookmarks.

#### Gebruik:
1. Open een Excel workbook in de VBA Editor
2. Ga naar menu: **Utilities → PDF Export met Inhoudsopgave**
3. In het selectie venster:
   - Vink de sheets aan die je wilt exporteren
   - (Optioneel) Stel voor elk sheet een specifieke printrange in
   - (Optioneel) Geef sheets een aangepaste display naam
4. Klik op **PDF Genereren**
5. Kies een locatie om de PDF op te slaan

#### Features:
- **Selectie van sheets**: Kies welke worksheets je wilt exporteren
- **PrintRange instelling**: Stel voor elk sheet een specifieke printrange in (bijv. `A1:G50`)
- **Automatische detectie**: Detecteer automatisch de gebruikte range van een sheet
- **Display namen**: Geef sheets een aangepaste naam in de PDF inhoudsopgave
- **Inhoudsopgave**: Automatisch gegenereerde eerste pagina met alle sheets
- **Navigeerbare hyperlinks**: Klik in de inhoudsopgave om naar de juiste sheet te navigeren
- **Batch export**: Exporteer meerdere sheets in één keer naar één PDF

#### Voorbeeld Workflow:
```
1. Je hebt een workbook met sheets: "Data", "Rapport Q1", "Rapport Q2", "Analyse"
2. Je wilt alleen de rapporten exporteren met aangepaste ranges
3. Open PDF Export dialoog
4. Selecteer "Rapport Q1" en stel printrange in: A1:H30
5. Selecteer "Rapport Q2" en stel printrange in: A1:H30  
6. Wijzig display namen naar "Q1 2026" en "Q2 2026"
7. Genereer PDF
8. Resultaat: PDF met inhoudsopgave + 2 pagina's (Q1 en Q2)
```

### Methode 1: Via VBE Menu (Aanbevolen)

Gebruik het **Utilities** menu in de VBA Editor menubar voor toegang tot alle functies.

### Methode 2: Vanuit VBA Code

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

4. Druk op Enter

### Voorbeeld

**Voor formattering:**
```vb
Sub Test()
Dim x As Integer
  Dim y As String
Dim z As Boolean
End Sub
```

**Na formattering:**
```vb
Sub Test()
    Dim x As Integer
    Dim y As String
    Dim z As Boolean
End Sub
```

## Verificatie

Om te controleren of de add-in correct geregistreerd is:

1. Open Registry Editor (`Win+R`, type `regedit`)
2. Navigeer naar: `HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\VBEAddIn.Connect`
3. Je zou entries moeten zien voor Description, FriendlyName, en LoadBehavior

Of gebruik PowerShell:
```powershell
Get-ItemProperty -Path "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins\VBEAddIn.Connect" -ErrorAction SilentlyContinue
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

## Projectstructuur

```
VBA C#/
├── Connect.cs                      # Hoofd add-in klasse (IDTExtensibility2)
├── DimFormatter.cs                 # Dim statement formattering logica
├── CompleteCodeFormatter.cs        # Complete code formattering
├── FormatterSettings.cs            # Instellingen configuratie
├── SettingsForm.cs                 # Instellingen dialoog
├── WorkbookSelectionForm.cs        # Workbook selectie dialoog
├── WhoAmIUtility.cs               # Workbook info utility
├── OptimalisatieUtility.cs        # Excel optimalisatie utility
├── ExportVBAUtility.cs            # VBA export utility
├── ReferenceManagerUtility.cs     # Reference manager utility
├── PDFExportUtility.cs            # PDF export logica (NIEUW)
├── PrintRangeSelectionForm.cs     # Sheet/Range selectie dialoog (NIEUW)
├── Properties/
│   └── AssemblyInfo.cs            # Assembly informatie en COM settings
├── VBEAddIn.csproj                # Project bestand
├── Register.reg                   # Registry bestand voor installatie
├── Unregister.reg                 # Registry bestand voor de-installatie
├── bin/Release/
│   └── VBEAddIn.dll              # Gecompileerde COM DLL
└── README.md                      # Deze file
```

## Troubleshooting

### "Add-in niet gevonden" fout
- Controleer of de DLL correct geregistreerd is met `regasm`
- Controleer of de registry entries bestaan
- Herstart de Office applicatie volledig

### "Toegang geweigerd" bij regasm
- Run Command Prompt als Administrator
- Zorg ervoor dat het pad naar de DLL correct is

### Add-in laadt niet
- Controleer LoadBehavior in registry (moet 3 zijn voor automatisch laden)
- Check Windows Event Viewer voor foutmeldingen
- Zorg ervoor dat .NET Framework 4.8 geïnstalleerd is

### Formattering werkt niet
- Zorg ervoor dat je een code module hebt geopend in VBE
- Controleer of de module niet read-only is
- Check de Immediate Window voor foutmeldingen

## Technische Details

- **Framework**: .NET Framework 4.8
- **COM Interop**: Ja, geregistreerd via RegAsm
- **VBE Integration**: IDTExtensibility2 interface
- **Excel Integration**: Microsoft.Office.Interop.Excel
- **Public Methods**: 
  - `FormatDimStatements()` - Formatteer Dim statements
  - `FormatCompleteCode()` - Complete code formattering  
  - `WhoAmI()` - Toon workbook info
  - `OptimalisatieUit()` - Schakel Excel optimalisaties uit
  - `OptimalisatieAan()` - Schakel Excel optimalisaties aan
  - `ExportVBAComponents()` - Exporteer VBA code
  - `ManageReferences()` - Beheer VBA references
  - `ExportToPDF()` - PDF export met inhoudsopgave (NIEUW)
- **GUID**: B1C2D3E4-F5A6-4B78-C901-D234E5678F90
- **ProgID**: VBEAddIn.Connect

## Uitbreidingsmogelijkheden

Je kunt deze add-in uitbreiden met extra functionaliteit:

- **Code Beautification**: Complete code formatting
- **Variable Naming**: Enforcing naming conventions
- **Comment Formatting**: Standardize comments
- **Code Analysis**: Detect common issues
- **Snippet Insertion**: Quick code templates
- **Refactoring Tools**: Rename variables, extract methods

Voeg gewoon nieuwe publieke methoden toe aan de `Connect` klasse en maak ze `[ComVisible(true)]`.

## Licentie

Dit project is vrij te gebruiken en aan te passen voor eigen doeleinden.

## Ondersteuning

Voor vragen of problemen:
1. Check de Troubleshooting sectie
2. Kijk in Windows Event Viewer voor COM/VBA fouten
3. Test de add-in in een nieuw Excel workbook met simpele VBA code

---

**Laatst bijgewerkt**: 11 februari 2026

**Nieuwe features in deze versie:**
- PDF Export met automatische inhoudsopgave en navigeerbare bookmarks
- Mogelijkheid om specifieke printranges per sheet in te stellen
- Automatische detectie van gebruikte ranges
- Aangepaste display namen voor sheets in PDF
