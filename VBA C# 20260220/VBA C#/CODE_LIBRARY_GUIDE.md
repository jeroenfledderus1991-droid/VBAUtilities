# Code Library Systeem - Gebruikershandleiding

## Overzicht
Het Code Library systeem stelt je in staat om een verzameling VBA modules (.bas), classes (.cls) en forms (.frm) bij te houden en deze eenvoudig te importeren in je VBA projecten.

## Installatie
1. Run: `.\Installer\bin\Release\VBEAddIn-Installer.exe`
2. Herstart Excel/VBE

## Library Folder Setup

### Standaard Locatie
De library wordt standaard opgeslagen in:
```
C:\Users\[Gebruiker]\Documents\VBA Code Library\
```

### Folder Structuur
Je kunt subfolders gebruiken voor organisatie:
```
VBA Code Library\
├── Utilities\
│   ├── StringHelper.bas
│   └── FileUtilities.bas
├── Classes\
│   ├── CustomCollection.cls
│   └── DataProcessor.cls
└── Forms\
    └── InputDialog.frm
```

## Gebruik

### 1. Modules Toevoegen aan Library
Plaats je VBA bestanden in de library folder:
- **Modules**: `.bas` files
- **Classes**: `.cls` files  
- **Forms**: `.frm` files (met bijbehorende .frx indien nodig)

### 2. Modules Importeren

#### Via Menu
1. Open VBE (Alt+F11 in Excel)
2. Klik **Utilities** → **Code Library**
3. Selecteer gewenste modules (checkboxes)
4. Klik **Import**

#### Via CommandBar
Als de CommandBar is ingeschakeld:
1. Klik op de **Library** button in de toolbar
2. Selecteer modules
3. Klik **Import**

### 3. Module Selectie Form

#### Controls
- **CheckedListBox**: Toont alle beschikbare modules
  - `[M]` = Module (.bas)
  - `[C]` = Class (.cls)
  - `[F]` = Form (.frm)
  - Relatieve paden tonen subfolder structuur

- **Select All**: Selecteer alle modules
- **Select None**: Deselecteer alle modules
- **Open Folder**: Open library folder in Windows Explorer
- **Import**: Importeer geselecteerde modules
- **Cancel**: Annuleer zonder importeren

#### Duplicate Handling
Als een module al bestaat in je project:
- Dialog vraagt: **Overschrijven? (Yes/No/Cancel)**
  - **Yes**: Module wordt vervangen
  - **No**: Module wordt overgeslagen
  - **Cancel**: Import wordt gestopt

### 4. Import Resultaat
Na import zie je een samenvatting:
```
Modules geïmporteerd:
✓ Geïmporteerd: StringHelper.bas
✓ Geïmporteerd: CustomCollection.cls
○ Overgeslagen: FileUtilities.bas (bestaat al)

2 van 3 modules succesvol geïmporteerd
```

## Instellingen

### CommandBar Button Inschakelen
1. Open **Utilities** → **Instellingen**
2. Ga naar tab **Commandbar**
3. Vink aan: ☑ **Code Library**
4. Klik **Opslaan**
5. De **Library** button verschijnt direct in de toolbar

### Library Path Wijzigen
⚠️ Momenteel gebruikt het systeem de standaard locatie.
Voor een custom path moet je `FormatterSettings.CodeLibraryPath` direct in de registry aanpassen:
```
HKEY_CURRENT_USER\Software\VBEAddIn\Settings
  CodeLibraryPath = "C:\Jouw\Custom\Path"
```

## Technische Details

### Ondersteunde Bestandstypes
- `.bas` - Standard modules
- `.cls` - Class modules
- `.frm` - UserForms (met .frx indien nodig)

### Import Methode
Het systeem gebruikt `VBProject.VBComponents.Import()`:
- Behoudt originele module naam
- Behoudt alle code en metadata
- Importeert forms inclusief controls

### Duplicate Detection
Controleert op `VBComponent.Name` matching:
- Case-insensitive vergelijking
- Per module een overwrite prompt
- Cancel stopt verdere imports

### Recursive Search
Library folder wordt recursief doorzocht:
- Alle subfolders worden gescand
- Relatieve paden worden getoond in lijst
- Geen limiet op folder depth

## Tips & Best Practices

### 1. Organisatie
Gebruik subfolders voor categorieën:
```
VBA Code Library\
├── Database\
├── FileIO\
├── StringManipulation\
└── UserInterface\
```

### 2. Naamgeving
Gebruik beschrijvende namen:
- ✓ `StringHelper.bas`
- ✓ `DatabaseConnection.cls`
- ✗ `Module1.bas`
- ✗ `Class1.cls`

### 3. Versie Controle
Overweeg een backup/versioning systeem:
- Git repository voor library folder
- Datum in bestandsnaam: `Utility_v2024-01-15.bas`
- Changelog in module comments

### 4. Code Standaarden
Zorg dat library modules:
- Geen hardcoded paden bevatten
- Geen workbook-specifieke references hebben
- Goed gedocumenteerd zijn
- Onafhankelijk van andere modules zijn (of dependencies duidelijk vermelden)

### 5. Testing
Test geïmporteerde modules altijd:
- Controleer of alle references beschikbaar zijn
- Test functionaliteit in target workbook
- Let op conflicterende module namen

## Troubleshooting

### "Library folder niet gevonden"
- Controleer of folder bestaat: `Documents\VBA Code Library`
- Folder wordt automatisch aangemaakt bij eerste gebruik
- Zorg dat je schrijfrechten hebt

### "Module kon niet worden geïmporteerd"
Mogelijke oorzaken:
- Bestand is in gebruik door ander proces
- Bestand is corrupt of niet geldig VBA formaat
- Onvoldoende rechten

### "CommandBar Library button verdwijnt"
- Open Instellingen
- Controleer: ☑ **CommandBar tonen** (master checkbox)
- Controleer: ☑ **Code Library** checkbox
- Klik Opslaan

### Forms met .frx files
Forms met controls hebben .frx binary files nodig:
- Zorg dat beide files (.frm en .frx) in library staan
- .frx moet dezelfde naam hebben als .frm
- Import haalt automatisch beide files op

## Keyboard Shortcuts
In Code Library Form:
- **Spacebar**: Toggle checkbox van geselecteerde item
- **Ctrl+A**: Select All (focus in list)
- **Enter**: Import (als Import button focus heeft)
- **Escape**: Cancel

## Toekomstige Features (Optioneel)
Mogelijke uitbreidingen:
- [ ] Library path configureren via Settings UI
- [ ] Export to Library functie (reverse)
- [ ] Module preview (code tonen voor import)
- [ ] Categorieën/tags systeem
- [ ] Search/filter functionaliteit
- [ ] Favorites/recent items
- [ ] Module update detection
- [ ] Batch export van actief project naar library

## Voorbeeld Workflow

### Scenario: Nieuwe Helper Module Toevoegen
1. **Ontwikkel module** in test workbook
2. **Export naar library**:
   ```
   - Export module: StringHelper.bas
   - Kopieer naar: Documents\VBA Code Library\Utilities\
   ```
3. **Gebruik in ander project**:
   ```
   - Open target workbook in VBE
   - Utilities → Code Library
   - Zoek: Utilities\StringHelper.bas [M]
   - Vink aan en Import
   ```
4. **Module is nu beschikbaar** in target project

### Scenario: Project Template Setup
1. **Maak standard modules**:
   ```
   - ErrorHandler.bas
   - Logger.bas
   - ConfigManager.bas
   ```
2. **Plaats in library** subfolder "Template"
3. **Bij nieuw project**:
   ```
   - Code Library openen
   - Select All in Template folder
   - Import in nieuw project
   ```
4. **Standard modules** direct beschikbaar

## Support
Voor vragen of problemen:
- Check dit document
- Test in schone workbook
- Controleer VBE Immediate window voor error messages

---

**Versie**: 1.0  
**Datum**: 2024  
**Onderdeel van**: VBE AddIn C# Project
