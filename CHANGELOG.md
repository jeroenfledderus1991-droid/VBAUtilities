# Changelog — VBE Code Tools Add-In

Alle noemenswaardige wijzigingen worden bijgehouden in dit bestand.  
Formaat gebaseerd op [Keep a Changelog](https://keepachangelog.com/nl/1.0.0/).

---

## [1.3.2] — 2026-04-04

### Fixes
- Update-check wordt nu betrouwbaar uitgevoerd vanuit `OnConnection` (werkt in VBE-sessie)
- Bij een GitHub Release met installer-asset opent `Ja` nu direct de installer-download
- Zichtbare versie in de add-in bijgewerkt naar **1.3.2**

---

## [1.3.0] — 2026-04-04

### Toegevoegd
- **VBA Wachtwoord Verwijderen** — nieuwe utility om VBA projectwachtwoorden te omzeilen via `DialogBoxParamA` API hooking; de geïnjecteerde module blijft staan voor hergebruik
- **Versiegeschiedenis** — bekijk de changelog direct vanuit de add-in via `Utilities → Versiegeschiedenis`

### Fixes
- VBA password remover: quote escaping hersteld in embedded VBA string (ontbrekende sluit-`""`)
- VBA password remover: code-validatie is nu hoofdletterongevoelig (VBE past automatisch hoofdletter aan bij `varPtr` → `VarPtr`)
- VBA password remover: module werd onterecht verwijderd vóór het uitvoeren van het macro — opgelost door cleanup te verwijderen uit `finally`-blok

---

## [1.2.0] — 2026-02-20

### Toegevoegd
- **Code Library** — importeer en beheer VBA modules via een centrale bibliotheek (`CodeLibraryForm`, `CodeLibraryUtility`)
- **Export naar Library** — exporteer modules vanuit de VBE naar de code library (`ExportToLibraryForm`, `ExportToLibraryUtility`)
- **UnifiedCodeLibraryForm** — gecombineerde interface voor library beheer
- **LibraryPathsForm** — configureer paden voor de code library
- **Insert Comment** — voeg een commentaarregel met timestamp en gebruikersnaam toe; modifier keys: normaal / SHIFT (asterisks) / CTRL (START–END blok)

---

## [1.1.0] — 2026-01-15

### Toegevoegd
- **Reference Manager** — beheer VBA references in het actieve project via een overzichtelijk formulier
- **WhoAmI** — toon workbook info (FullName en ReadOnly status) van het actieve werkboek
- **Optimalisatie UIT** — schakel Excel optimalisaties uit (events, screenupdating, alerts, calculatie)
- **Optimalisatie AAN** — herstel Excel optimalisaties na een run
- **Export VBA Componenten** — exporteer alle VBA modules/formulieren/klassen naar schijf

---

## [1.0.0] — 2026-01-01

### Initieel
- VBE COM Add-in structuur opgezet (.NET 4.8, COM Interop)
- **Formatteer Dim Statements** — sorteer en lijn `Dim` statements minimaal 1 tab uit binnen de actieve procedure
- **Formatteer Complete Code** — volledige code formatting: indentatie, Dims, lege regels
- **Instellingen** — form voor add-in configuratie (bewaard in registry)
- CommandBar (toolbar) met instelbare knoppen per functie
