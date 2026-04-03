# Code Library System - Technical Documentation

## Architecture Overview

The Code Library system extends the VBEAddIn with module import functionality, allowing users to maintain a reusable VBA code repository and selectively import modules into active VBA projects.

## Components

### 1. CodeLibraryForm.cs
**Purpose**: Windows Form for module selection from library folder

**Key Features**:
- CheckedListBox with type indicators [M]/[C]/[F]
- Recursive file search through library folder and subfolders
- Select All/None buttons for quick selection
- Open Folder button (launches Windows Explorer)
- Relative path display for subfolder structure

**Implementation Details**:
```csharp
public partial class CodeLibraryForm : Form
{
    private CheckedListBox lstModules;
    private List<string> allFiles;  // Full paths to all found modules
    
    public List<string> SelectedFiles {
        get {
            var selected = new List<string>();
            for (int i = 0; i < lstModules.CheckedItems.Count; i++) {
                int index = lstModules.Items.IndexOf(lstModules.CheckedItems[i]);
                selected.Add(allFiles[index]);
            }
            return selected;
        }
    }
    
    private void LoadFiles(string libraryPath) {
        // Recursive search for .bas/.cls/.frm
        var basFiles = Directory.GetFiles(libraryPath, "*.bas", SearchOption.AllDirectories);
        var clsFiles = Directory.GetFiles(libraryPath, "*.cls", SearchOption.AllDirectories);
        var frmFiles = Directory.GetFiles(libraryPath, "*.frm", SearchOption.AllDirectories);
        
        // Display with type indicators and relative paths
        foreach (string file in allFiles) {
            string relativePath = file.Substring(libraryPath.Length + 1);
            string extension = Path.GetExtension(file).ToLower();
            string icon = extension == ".bas" ? "[M]" : 
                         extension == ".cls" ? "[C]" : "[F]";
            lstModules.Items.Add($"{icon} {relativePath}");
        }
    }
}
```

**Event Handlers**:
- `BtnSelectAll_Click`: Checks all items in list
- `BtnSelectNone_Click`: Unchecks all items
- `BtnOpenFolder_Click`: Opens library folder in Explorer
- `BtnImport_Click`: Sets DialogResult.OK and closes form
- `BtnCancel_Click`: Sets DialogResult.Cancel and closes form

**Form Properties**:
- Size: 600x500
- Resizable: true
- StartPosition: CenterScreen
- Controls anchored for dynamic resizing

### 2. CodeLibraryUtility.cs
**Purpose**: Static utility class for module import logic

**Key Methods**:
```csharp
public static class CodeLibraryUtility
{
    public static void Execute(VBE vbe)
    {
        // 1. Get library path from settings (default: Documents\VBA Code Library)
        string libraryPath = FormatterSettings.CodeLibraryPath;
        if (string.IsNullOrEmpty(libraryPath)) {
            libraryPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "VBA Code Library"
            );
        }
        
        // 2. Create folder if doesn't exist
        if (!Directory.Exists(libraryPath)) {
            Directory.CreateDirectory(libraryPath);
        }
        
        // 3. Show file selection form
        using (CodeLibraryForm form = new CodeLibraryForm(libraryPath)) {
            if (form.ShowDialog() != DialogResult.OK) return;
            
            // 4. Import selected files
            VBProject project = vbe.ActiveVBProject;
            StringBuilder result = new StringBuilder();
            int imported = 0, skipped = 0;
            
            foreach (string filePath in form.SelectedFiles) {
                string fileName = Path.GetFileName(filePath);
                string moduleName = Path.GetFileNameWithoutExtension(filePath);
                
                // Check for duplicate
                bool exists = false;
                foreach (VBComponent component in project.VBComponents) {
                    if (component.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase)) {
                        exists = true;
                        
                        // Ask to overwrite
                        DialogResult overwrite = MessageBox.Show(
                            $"Module '{moduleName}' bestaat al.\n\nOverschrijven?",
                            "Duplicaat gevonden",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question
                        );
                        
                        if (overwrite == DialogResult.Yes) {
                            project.VBComponents.Remove(component);
                            exists = false;
                        } else if (overwrite == DialogResult.Cancel) {
                            return;  // Stop entire import
                        }
                        break;
                    }
                }
                
                if (!exists) {
                    try {
                        project.VBComponents.Import(filePath);
                        result.AppendLine($"✓ Geïmporteerd: {fileName}");
                        imported++;
                    } catch (Exception ex) {
                        result.AppendLine($"✗ Fout bij {fileName}: {ex.Message}");
                    }
                } else {
                    result.AppendLine($"○ Overgeslagen: {fileName}");
                    skipped++;
                }
            }
            
            // 5. Show results
            MessageBox.Show(
                $"Modules geïmporteerd:\n\n{result}\n" +
                $"{imported} van {form.SelectedFiles.Count} modules succesvol geïmporteerd",
                "Import Compleet",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }
}
```

**Error Handling**:
- Try-catch per file prevents single error from stopping entire import
- Locked projects show specific error message
- Missing files or corrupt files caught per-import

**Duplicate Strategy**:
- Check `VBComponent.Name` case-insensitive
- Per-module overwrite prompt (Yes/No/Cancel)
- Cancel stops remaining imports
- Remove existing component before import (VBComponents.Remove)

### 3. FormatterSettings.cs Extensions
**Added Properties**:
```csharp
public static class FormatterSettings
{
    // CommandBar visibility flag
    public static bool CommandBarShowCodeLibrary { get; set; }
    
    // Library folder path (configurable)
    public static string CodeLibraryPath { get; set; }
    
    // Registry persistence
    public static void LoadFromRegistry() {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\VBEAddIn\Settings")) {
            if (key != null) {
                // ... other settings ...
                CommandBarShowCodeLibrary = key.GetValue("CommandBarShowCodeLibrary", "0").ToString() == "1";
                CodeLibraryPath = key.GetValue("CodeLibraryPath", "").ToString();
            }
        }
    }
    
    public static void SaveToRegistry() {
        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\VBEAddIn\Settings")) {
            // ... other settings ...
            key.SetValue("CommandBarShowCodeLibrary", CommandBarShowCodeLibrary ? "1" : "0");
            key.SetValue("CodeLibraryPath", CodeLibraryPath ?? "");
        }
    }
}
```

**Registry Location**:
```
HKEY_CURRENT_USER\Software\VBEAddIn\Settings
    CommandBarShowCodeLibrary = REG_SZ ("1" or "0")
    CodeLibraryPath = REG_SZ (full path or empty string)
```

### 4. Connect.cs Integration
**Added Fields**:
```csharp
private CommandBarButton _menuButtonCodeLibrary;
private CommandBarButton _cmdCodeLibrary;
```

**Menu Button Creation** (in CreateMenuButton):
```csharp
// Code Library button
_menuButtonCodeLibrary = (CommandBarButton)utilitiesPopup.Controls.Add(
    Type: Microsoft.Office.Core.MsoControlType.msoControlButton,
    Temporary: true
);
_menuButtonCodeLibrary.Caption = "Code Library";
_menuButtonCodeLibrary.Tag = "VBEAddIn_CodeLibrary";
_menuButtonCodeLibrary.TooltipText = "Importeer modules uit code library";
_menuButtonCodeLibrary.FaceId = 3059;  // Library/templates icon
_menuButtonCodeLibrary.Click += OnMenuButtonCodeLibraryClick;
```

**CommandBar Button Creation** (in CreateCommandBar):
```csharp
if (FormatterSettings.CommandBarShowCodeLibrary) {
    _cmdCodeLibrary = (CommandBarButton)_commandBar.Controls.Add(
        Type: Microsoft.Office.Core.MsoControlType.msoControlButton,
        Temporary: true
    );
    _cmdCodeLibrary.Style = MsoButtonStyle.msoButtonIconAndCaption;
    _cmdCodeLibrary.Caption = "Library";
    _cmdCodeLibrary.FaceId = 3059;
    _cmdCodeLibrary.Tag = "VBEAddIn_CMD_CodeLibrary";
    _cmdCodeLibrary.TooltipText = "Code Library";
    _cmdCodeLibrary.Click += OnMenuButtonCodeLibraryClick;
}
```

**Event Handler**:
```csharp
private void OnMenuButtonCodeLibraryClick(CommandBarButton Ctrl, ref bool CancelDefault)
{
    OpenCodeLibrary();
}
```

**Public COM Method**:
```csharp
[ComVisible(true)]
public void OpenCodeLibrary()
{
    CodeLibraryUtility.Execute(_vbe);
}
```

**Cleanup** (in RemoveCommandBar and RemoveMenuButton):
```csharp
if (_cmdCodeLibrary != null) {
    _cmdCodeLibrary.Delete(true);
    _cmdCodeLibrary = null;
}

if (_menuButtonCodeLibrary != null) {
    _menuButtonCodeLibrary.Delete(true);
    _menuButtonCodeLibrary = null;
}
```

### 5. SettingsForm.cs Integration
**Added Controls**:
```csharp
private CheckBox chkCmdCodeLibrary;  // Field declaration

// In InitializeGeneralTab() - after Reference Manager checkbox
chkCmdCodeLibrary = new CheckBox {
    Text = "Code Library",
    Location = new Point(40, 310),
    Size = new Size(300, 20),
    Checked = FormatterSettings.CommandBarShowCodeLibrary,
    Enabled = FormatterSettings.ShowCommandBar
};
tabGeneral.Controls.Add(chkCmdCodeLibrary);
```

**Enable/Disable Logic** (in ChkShowCommandBar_CheckedChanged):
```csharp
chkCmdCodeLibrary.Enabled = enabled;
```

**Save Logic** (in BtnSave_Click):
```csharp
FormatterSettings.CommandBarShowCodeLibrary = chkCmdCodeLibrary.Checked;
```

**UI Layout Adjustments**:
- Code Library checkbox: y=310
- Insert Comment checkbox: moved from y=310 to y=335
- Info label: moved from y=345 to y=370

### 6. VBEAddIn.csproj Updates
**Added Compilation Units**:
```xml
<Compile Include="CodeLibraryForm.cs">
  <SubType>Form</SubType>
</Compile>
<Compile Include="CodeLibraryUtility.cs" />
```

## Data Flow

1. **User Action**: Clicks Utilities → Code Library (or CommandBar button)
2. **Event Handler**: `OnMenuButtonCodeLibraryClick` → `OpenCodeLibrary()`
3. **Utility Entry**: `CodeLibraryUtility.Execute(vbe)` called
4. **Path Resolution**:
   - Read `FormatterSettings.CodeLibraryPath` from registry
   - Default: `Documents\VBA Code Library`
   - Create folder if missing
5. **Form Display**: `CodeLibraryForm` shown with file list
6. **File Discovery**: Recursive search for .bas/.cls/.frm
7. **User Selection**: Check modules to import via CheckedListBox
8. **Import Loop**: For each selected file:
   - Check if module name exists in active VBProject
   - If duplicate: prompt Yes/No/Cancel
   - If Yes: remove existing, then import
   - If No: skip this module
   - If Cancel: abort entire operation
   - Use `project.VBComponents.Import(filePath)`
9. **Result Display**: Show summary with ✓/○/✗ icons

## VBA Interop Details

### VBComponents.Import()
```csharp
VBComponent imported = project.VBComponents.Import(string FileName);
```
- **Parameters**: Full path to .bas/.cls/.frm file
- **Returns**: VBComponent object of imported module
- **Behavior**:
  - Module keeps original name from file
  - If .frm file, automatically imports associated .frx
  - Throws exception if file not found or corrupt
  - Adds to VBComponents collection

### VBComponents.Remove()
```csharp
project.VBComponents.Remove(VBComponent component);
```
- **Purpose**: Delete existing component before reimport
- **Note**: Cannot remove certain protected modules (ThisWorkbook, Sheet modules)

### Component Type Detection
```csharp
switch (component.Type) {
    case vbext_ComponentType.vbext_ct_StdModule:    // .bas
    case vbext_ComponentType.vbext_ct_ClassModule:  // .cls
    case vbext_ComponentType.vbext_ct_MSForm:       // .frm
}
```

## File System Operations

### Directory Structure
```
Documents\
└── VBA Code Library\          (root library folder)
    ├── *.bas                  (top-level modules)
    ├── *.cls
    ├── *.frm (+ .frx)
    └── Subfolder\             (category folders)
        ├── *.bas
        └── *.cls
```

### File Search
```csharp
string[] basFiles = Directory.GetFiles(libraryPath, "*.bas", SearchOption.AllDirectories);
string[] clsFiles = Directory.GetFiles(libraryPath, "*.cls", SearchOption.AllDirectories);
string[] frmFiles = Directory.GetFiles(libraryPath, "*.frm", SearchOption.AllDirectories);
```
- **SearchOption.AllDirectories**: Recursive search
- **Extension filtering**: Only .bas/.cls/.frm recognized
- **No depth limit**: Searches entire folder tree

### Path Display
```csharp
string relativePath = file.Substring(libraryPath.Length + 1);
// Example: "Utilities\StringHelper.bas" instead of full path
```

## UI Design Patterns

### Type Indicators
- `[M]` - Module (.bas) - Standard code module
- `[C]` - Class (.cls) - Class module
- `[F]` - Form (.frm) - UserForm with controls

### Result Icons
- `✓` - Successfully imported
- `○` - Skipped (user choice or duplicate without overwrite)
- `✗` - Error occurred during import

### Dialog Strategy
- **Selection Form**: Modal dialog (ShowDialog)
- **Overwrite Prompts**: Per-module YesNoCancel
- **Results**: Informational MessageBox with summary

## Settings Persistence

### Registry Schema
```
HKCU\Software\VBEAddIn\Settings\
    CommandBarShowCodeLibrary    REG_SZ    "1" = visible, "0" = hidden
    CodeLibraryPath              REG_SZ    full path or empty (use default)
```

### Default Values
- `CommandBarShowCodeLibrary`: false (hidden by default)
- `CodeLibraryPath`: empty string → uses Documents\VBA Code Library

### Live Refresh
Settings form calls `Connect.Instance.RefreshCommandBar()` after save:
- CommandBar rebuilds with updated settings
- Changes visible immediately (no VBE restart required)

## Error Handling Strategy

### Per-Module Try-Catch
```csharp
foreach (string filePath in selectedFiles) {
    try {
        project.VBComponents.Import(filePath);
        result.AppendLine($"✓ Geïmporteerd: {fileName}");
        imported++;
    } catch (Exception ex) {
        result.AppendLine($"✗ Fout bij {fileName}: {ex.Message}");
        // Continue with next file
    }
}
```
**Rationale**: One corrupt file shouldn't block all imports

### Locked Project Detection
```csharp
try {
    VBProject project = vbe.ActiveVBProject;
    var components = project.VBComponents;  // Triggers error if locked
} catch (Exception) {
    MessageBox.Show("Project is vergrendeld. Kan geen modules importeren.");
    return;
}
```

## Performance Considerations

### File Search
- **Recursive Directory.GetFiles()**: Fast for typical library size (<1000 files)
- **Lazy loading**: Search only when form opens
- **No caching**: Always finds latest files (add-edit-import workflow)

### Import Speed
- **VBComponents.Import()**: ~50-200ms per module (COM overhead)
- **Sequential processing**: Import one-by-one (no batch API available)
- **UI responsiveness**: Form blocks during import (acceptable for typical batch size)

## Testing Scenarios

### 1. Empty Library
- **Expected**: Form shows "Geen modules gevonden in library"
- **Behavior**: Import button disabled

### 2. Mixed File Types
- **Input**: 5 .bas, 3 .cls, 2 .frm files
- **Expected**: All shown with correct type indicators
- **Verify**: Relative paths correct for subfolders

### 3. Duplicate Handling
- **Setup**: Import module A twice
- **Expected**: 
  - First import: succeeds
  - Second import: overwrite prompt shown
  - Yes: module replaced
  - No: original kept
  - Cancel: second import aborted

### 4. Locked Project
- **Setup**: Protect VBProject with password
- **Expected**: Error message shown, no import attempted

### 5. Corrupt File
- **Setup**: Create invalid .bas file (wrong format)
- **Expected**: Error shown for that file, other files still imported

### 6. Forms with .frx
- **Setup**: UserForm with controls (needs .frx binary)
- **Expected**: Both .frm and .frx imported, form functional

## Future Enhancement Ideas

### 1. Library Path Configuration in UI
Add to Settings form (new tab or section):
```csharp
// Settings tab "Code Library"
TextBox txtLibraryPath;
Button btnBrowseLibraryPath;

private void BtnBrowseLibraryPath_Click(object sender, EventArgs e) {
    using (FolderBrowserDialog dialog = new FolderBrowserDialog()) {
        dialog.SelectedPath = FormatterSettings.CodeLibraryPath;
        if (dialog.ShowDialog() == DialogResult.OK) {
            txtLibraryPath.Text = dialog.SelectedPath;
        }
    }
}

// In BtnSave_Click:
FormatterSettings.CodeLibraryPath = txtLibraryPath.Text;
```

### 2. Export to Library
Reverse functionality - export modules FROM project TO library:
```csharp
public static void ExportToLibrary(VBE vbe) {
    // Select modules from active project
    // Export to library folder
    // Option to categorize (subfolder selection)
}
```

### 3. Module Preview
Show code before import:
```csharp
private void LstModules_SelectedIndexChanged(object sender, EventArgs e) {
    string filePath = allFiles[lstModules.SelectedIndex];
    string code = File.ReadAllText(filePath);
    txtPreview.Text = code;  // Syntax highlighted if possible
}
```

### 4. Favorites/Recent
Track frequently used modules:
```
HKCU\Software\VBEAddIn\Library\
    Recent    REG_MULTI_SZ    list of recently imported paths
```

### 5. Search/Filter
RealTime filtering in CheckedListBox:
```csharp
TextBox txtSearch;

private void TxtSearch_TextChanged(object sender, EventArgs e) {
    string filter = txtSearch.Text.ToLower();
    lstModules.Items.Clear();
    foreach (var file in allFiles.Where(f => f.ToLower().Contains(filter))) {
        lstModules.Items.Add(FormatFileName(file));
    }
}
```

### 6. Module Metadata
Store description/tags per module:
```vba
'@Description: "String manipulation utilities"
'@Tags: string, helper, text
```
Parse in C# and show in UI.

## Known Limitations

1. **Forms with .frx**: Must be in same folder as .frm
2. **No version control**: No built-in diff/merge
3. **Name-based duplicate detection**: Doesn't check content
4. **No dependency management**: If Module A uses Module B, both must be imported
5. **Protected modules**: Cannot overwrite Sheet/ThisWorkbook modules
6. **Single project**: Always imports to active project (no multi-project support)

## Dependencies

### .NET Framework
- System.IO (Directory, Path, File operations)
- System.Windows.Forms (CheckedListBox, FolderBrowserDialog)
- System.Collections.Generic (List<string>)

### COM Interop
- Microsoft.Vbe.Interop (VBE, VBProject, VBComponents, VBComponent)

### Internal Dependencies
- FormatterSettings (configuration storage)
- Connect (menu/CommandBar integration)
- SettingsForm (UI configuration)

## Deployment

### Installation
1. Compile VBEAddIn.dll with CodeLibrary* files included
2. Installer embeds DLL and registers COM add-in
3. First run: Registry keys created with default values
4. User opens Settings: CommandBar checkboxes available
5. Default library folder auto-created on first usage

### Uninstallation
1. Installer removes COM registration
2. Registry keys remain (preserve user settings)
3. Library folder remains (user data)

## Maintenance Notes

### Code Locations
- **CodeLibraryForm.cs**: ~220 lines - module selection UI
- **CodeLibraryUtility.cs**: ~150 lines - import logic
- **FormatterSettings.cs**: ~299 lines - +2 properties for library
- **Connect.cs**: ~1105 lines - +2 fields, +2 buttons, +2 methods
- **SettingsForm.cs**: ~985 lines - +1 checkbox in Commandbar tab

### Testing Checklist
- [ ] Empty library folder
- [ ] Single module import
- [ ] Multiple module import (mixed types)
- [ ] Duplicate overwrite: Yes
- [ ] Duplicate overwrite: No
- [ ] Duplicate overwrite: Cancel
- [ ] Subfolder structure display
- [ ] Select All/None buttons
- [ ] Open Folder button
- [ ] CommandBar button visibility toggle
- [ ] Settings persistence (close/reopen VBE)
- [ ] Forms with .frx files
- [ ] Locked project handling
- [ ] Corrupt file handling

### Common Issues
- **"Automation error" during import**: Usually locked project
- **Form controls missing after import**: .frx file not found
- **Module name conflict**: Case-sensitive file system vs case-insensitive VBA
- **Path length limit**: Windows MAX_PATH (260 chars) - use relative paths

---

**Last Updated**: January 2024  
**Version**: 1.0  
**Author**: VBE AddIn Development Team
