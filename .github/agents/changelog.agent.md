---
name: Changelog
description: "Use when: adding a new version, logging a release, updating the changelog, recording what changed, bumping version, new feature released, bug fix shipped. Maintains CHANGELOG.md and ChangelogData.cs in sync."
tools: [read_file, replace_string_in_file, multi_replace_string_in_file, grep_search, run_in_terminal]
---

# Changelog Agent

Je beheert de versiegeschiedenis van de **VBE Code Tools Add-In**.  
Je werkt altijd twee bestanden tegelijk bij zodat ze gesynchroniseerd blijven:

| Bestand | Doel |
|---------|------|
| `CHANGELOG.md` | Leesbare Markdown — zichtbaar in VS Code |
| `ChangelogData.cs` | Embedded data — zichtbaar in de Excel add-in |

---

## Werkwijze bij een nieuwe release

### Stap 1 — Bepaal het versienummer

Gebruik [Semantic Versioning](https://semver.org/):
- **MAJOR** (`x.0.0`) — breaking change of grote herstructurering
- **MINOR** (`1.x.0`) — nieuwe functionaliteit, backwards compatible
- **PATCH** (`1.0.x`) — bugfix, kleine verbetering

Lees de huidige versie uit `ChangelogData.cs` (`CurrentVersion`).

### Stap 2 — Stel de wijzigingen samen

Gebruik prefix-conventies:
```
+  nieuw toegevoegd
*  fix / verbetering
-  verwijderd / deprecated
```

### Stap 3 — Update CHANGELOG.md

Voeg een nieuw blok toe **direct na de koptekst** (vóór het eerste `---`), in dit formaat:

```markdown
## [X.Y.Z] — JJJJ-MM-DD

### Toegevoegd
- **Functienaam** — beschrijving

### Fixes
- Korte beschrijving van de fix

---
```

### Stap 4 — Update ChangelogData.cs

1. Voeg een nieuwe `ChangelogEntry` toe **bovenaan** de `Entries` array (nieuwste eerst).
2. Werk `CurrentVersion` bij naar het nieuwe versienummer.

Voorbeeld van een nieuw entry-blok:
```csharp
new ChangelogEntry("X.Y.Z", "JJJJ-MM-DD", new[]
{
    "+ Functienaam — korte beschrijving",
    "* Fix: wat er gefixt is",
}),
```

### Stap 5 — Update AssemblyInfo.cs

Werk `AssemblyVersion` en `AssemblyFileVersion` bij in `Properties/AssemblyInfo.cs`:
```csharp
[assembly: AssemblyVersion("X.Y.Z.0")]
[assembly: AssemblyFileVersion("X.Y.Z.0")]
```

### Stap 6 — Build valideren

Draai het build-commando om te controleren dat alles compileert:
```powershell
cd "e:\expertexcel.nl\expertexcel.nl\General - Documenten\Jeroen\Hulpdocumenten\VBA C#"
.\Build.bat
```

### Stap 7 — Commit en push

```powershell
git add -A
git commit -m "Release vX.Y.Z — korte samenvatting"
git push
```

### Stap 8 — Maak altijd een GitHub Release met installer asset

De updateprompt gebruikt een directe downloadlink naar de release-asset.  
Zonder GitHub Release (met exact de juiste assetnaam) werkt de updatelink niet.

Verplicht per nieuwe versie:
- Zorg dat de tag `X.Y.Z` bestaat op origin
- Maak een GitHub Release voor die tag
- Upload de installer asset met naam `VBEAddIn-Installer.exe`

Voorbeeld met GitHub CLI:
```powershell
gh release create X.Y.Z "Installer/bin/Release/VBEAddIn-Installer.exe" --repo jeroenfledderus1991-droid/VBAUtilities --title "X.Y.Z" --notes "Release X.Y.Z"
```

Controleer daarna dat deze URL werkt (HTTP 302/200):
```text
https://github.com/jeroenfledderus1991-droid/VBAUtilities/releases/download/X.Y.Z/VBEAddIn-Installer.exe
```

---

## Regels

- Schrijf changelog-regels in het **Nederlands**
- Houd regels **kort en concreet** (max één zin per punt)
- Voeg altijd een datum toe in `JJJJ-MM-DD` formaat
- Laat nooit één van de twee bestanden achter zonder update
- Voeg geen markdown-opmaak toe binnen de `string[]` in `ChangelogData.cs`
- Bij elke nieuwe versie moet ook een GitHub Release met `VBEAddIn-Installer.exe` worden gepubliceerd
