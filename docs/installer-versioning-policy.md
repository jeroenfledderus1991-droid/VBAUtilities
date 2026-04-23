# Installer and Versioning Policy

This document replaces earlier ad-hoc instructions about installer naming and version packaging.

## Fixed installer name

- The public installer filename must always be:
  - `VBEAddIn-Installer.exe`
- Do not publish versioned installer filenames like `VBEAddIn-Installer-1.4.0.exe`.

## Version folders

- Build output must keep version-separated copies in folders:
  - `Installer/bin/Release/versions/<version>/VBEAddIn-Installer.exe`
- Example:
  - `Installer/bin/Release/versions/1.4.0/VBEAddIn-Installer.exe`

## GitHub Releases

- Every release uploads exactly one public installer asset:
  - `VBEAddIn-Installer.exe`
- Older versions are installed by selecting an older GitHub release, not by different installer filenames.

## Updater UX expectation

- The updater can show multiple release versions.
- Selecting an older release should open that release's installer download URL.
- The filename remains the same; the selected release determines installer contents.
