# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Memory
- Lees bij start altijd: ~/.claude/projects/VBA C#/memory/MEMORY.md
- Volg de index daarin en laad alle genoemde files

## Project Overview

A C# COM add-in for the Visual Basic Editor (VBE) in Microsoft Office applications. It adds a custom "Utilities" menu to the VBE with code formatting, Excel helpers, a code library system, and developer tools.

- **Framework:** .NET Framework 4.8.1
- **COM ProgID:** `VBEAddIn.Connect` (GUID: `B1C2D3E4-F5A6-4B78-C901-D234E5678F90`)
- **Two projects:** `VBEAddIn.csproj` (main DLL) and `Installer/Installer.csproj` (WinForms installer)

## Build Commands

**Full build (add-in + installer):**
```powershell
.\Build-VBEAddIn.ps1
```

**MSBuild directly (add-in only):**
```
"C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" VBEAddIn.csproj /p:Configuration=Release /t:Clean,Build /nologo /v:minimal
```

**Debug build:**
```powershell
.\Build-VBEAddIn.ps1 -Debug
```

Output: `bin\Release\VBEAddIn.dll` and `Installer\bin\Release\VBEAddIn-Installer.exe`

There are no automated tests in this project.

## Architecture

### Entry Point: `Connect.cs`
Implements `IDTExtensibility2` — the VBE COM add-in interface. This is the central class (~1,500 lines) that:
- Builds the "Utilities" VBE menu and optional CommandBar on `OnConnection`
- Tears everything down on `OnDisconnection`
- Exposes a static `Instance` property used by all forms to reach the VBE context
- Instantiates and calls all utility classes on menu item click

### Utility Classes
Each feature lives in its own class file (e.g., `DimFormatter.cs`, `ExportVBAUtility.cs`, `VBAPasswordRemoverUtility.cs`). Utilities receive the VBE reference from `Connect.Instance` and interact with it through `Microsoft.Vbe.Interop`.

### Configuration: `FormatterSettings.cs`
All settings are persisted to the Windows Registry at:
```
HKEY_CURRENT_USER\Software\VBEAddIn\Settings
```

### Versioning: `ChangelogData.cs`
`ChangelogData.cs` is the **single source of truth** for the current version number (`CurrentVersion` field). The build script reads it to name versioned installer artifacts. Always update `ChangelogData.cs` and `CHANGELOG.md` together when bumping the version.

### CI/CD: `.github/workflows/release.yml`
Triggers on push of semantic version tags (`*.*.*`). Builds both projects and publishes a GitHub release with the versioned installer as an asset. Auto-removes releases beyond the last 9.

## Releasing a New Version

1. Update `CurrentVersion` in `ChangelogData.cs`
2. Add an entry to `CHANGELOG.md`
3. Add the version entry to the `Entries` list in `ChangelogData.cs` (shown in-app changelog)
4. Commit and tag: `git tag x.y.z && git push origin x.y.z`
   — GitHub Actions handles the rest.

## Hard Rules
- Run tests/verificatie vóór "done".
- Max 50 regels per bugfix.
- Eén fix per commit.
- Raak auth/stripe/middleware niet aan zonder expliciete toestemming.
- Claude mag niet "done/fixed/ready" zeggen zonder het verify-commando te draaien en output te tonen.