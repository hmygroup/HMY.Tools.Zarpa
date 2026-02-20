# CopyAsInsert - SQL INSERT Generator

A PowerToys-style system tray app that converts Excel table data to SQL Server INSERT statements with automatic temporal table generation.

## Quick Start

### Running the Application

```bash
cd CopyAsInsert
dotnet run
```

The app will launch and minimize to the system tray.

### Usage Workflow

#### Quick Mode (Recommended for Prototyping)

1. **Copy Excel Data**: Select and copy a table from Excel (Ctrl+C)
2. **Press Alt+Shift+O**: Instantly generates `#temp` table - no dialogs!
3. **Paste into SSMS**: SQL is on clipboard, ready to execute

#### Standard Mode (Full Control)

1. **Copy Excel Data**: Select and copy a table from Excel (Ctrl+C)
   - Supports both tab-separated (TSV) and comma-separated (CSV) formats
   - Minimum: 1 header row + 1 data row

2. **Trigger Conversion** (Alt+Shift+I):
   - Press the global hotkey anywhere in Windows
   - If data is detected, a config dialog appears

3. **Configure & Generate**:
   - Enter table name (e.g., `Orders`)
   - Optionally change schema (default: `dbo`)
   - Enable/disable temporal table support (default: ON)
   - Click "Generate"

4. **Copy to SSMS**:
   - SQL is automatically copied to clipboard
   - Paste into SQL Server Management Studio and execute
   - A success notification shows row count

### Features

- **Quick Mode (Alt+Shift+O)**: Instant `#temp` table generation - no prompts, all NVARCHAR(100)
- **Auto Type Detection**: INT, FLOAT, DATETIME2, BIT, VARCHAR (70% lenient matching)
- **Primary Key Detection**: Auto-detects first INT or ID-named column as identity PK
- **Temporal Tables**: Includes SysStartTime/SysEndTime with system versioning enabled
- **History Tracking**: View recent conversions (max 10) from tray menu
- **Excel Support**: Both clipboard and .xlsx drag-drop parsing (ClosedXML)
- **Error Handling**: Validates table names, SQL keywords, data formats

### Hotkey Details

- **Standard Mode**: Alt+Shift+I
  - Opens configuration dialog for full customization
  - Allows table naming, schema selection, type overrides, temporal table options
  - Best for production use where you need control over the generated SQL

- **Quick Mode**: Alt+Shift+O (NEW in v2.2.0)
  - **No prompts** - instantly generates SQL without any dialogs
  - Creates temporary table `#temp` with all columns as `NVARCHAR(100)`
  - Perfect for rapid prototyping and quick data exploration
  - Assumes first row contains headers
  - No temporal table features (pure temporary table)

- **Why These Keys**: Non-intrusive, unlikely to conflict with other apps
- **When Available**: Global - works even when other windows are active

### Generated SQL

Each conversion produces:
```sql
CREATE TABLE [dbo].[TableName_Temporal] (
    [ColumnName] INT IDENTITY(1,1) PRIMARY KEY,
    [ColumnName] VARCHAR(255),
    ...
    SysStartTime DATETIME2 GENERATED ALWAYS AS ROW START NOT NULL,
    SysEndTime DATETIME2 GENERATED ALWAYS AS ROW END NOT NULL,
    PERIOD FOR SYSTEM_TIME (SysStartTime, SysEndTime)
)
WITH (SYSTEM_VERSIONING = ON (HISTORY_TABLE = [dbo].[TableName_History], DATA_CONSISTENCY_CHECK = ON));

INSERT INTO [dbo].[TableName_Temporal] (...)
VALUES
    (...),
    (...),
    ...;
```

### Troubleshooting

#### Hotkey doesn't register
- **Symptom**: Balloon tip shows error message on startup
- **Fix**: Hotkey may be in use by another app. Try different app combinations or use Settings to change default hotkey in future version.

#### Clipboard data not detected
- **Symptom**: "No Table Data" balloon after pressing hotkey
- **Fix**: 
  - Ensure you copied from Excel (not screenshot)
  - Data must be tab or comma-delimited
  - Minimum: header row + 1 data row
  - Avoid empty cells in header row

#### SQL generation fails
- **Symptom**: Error message in balloon tip
- **Fix**:
  - Table name must be valid SQL identifier (alphanumeric + underscore)
  - Cannot contain reserved keywords (SELECT, INSERT, etc.)
  - Max 128 characters

### Project Structure

```
CopyAsInsert/
├── Models/
│   ├── ColumnTypeInfo.cs          # Column metadata
│   ├── DataTableSchema.cs         # Parsed table structure
│   └── ConversionResult.cs        # Generation results
├── Services/
│   ├── ClipboardInterceptor.cs    # Windows API hotkey + clipboard
│   ├── TableDataParser.cs         # TSV/CSV/XLSX parsing
│   ├── TypeInferenceEngine.cs     # Auto type detection
│   └── SqlServerGenerator.cs      # CREATE + INSERT generation
├── Forms/
│   ├── MainForm.cs                # Tray app + hotkey listener
│   ├── TableConfigForm.cs         # Config dialog
│   └── SettingsForm.cs            # Settings dialog
└── Program.cs                     # Entry point
```

### Dependencies

- **EPPlus** (7.1.1+): Excel file parsing
- **ClosedXML** (0.55.0+): Excel workbook handling
- **.NET 8.0**: Runtime

### Build & Release

```bash
# Debug build
dotnet build

# Release build (self-contained, single file)
dotnet publish -p:PublishProfile=Portable -c Release

# Output: CopyAsInsert/bin/Release/net8.0-windows/publish/CopyAsInsert.exe
```

### Continuous Integration & Releases

The project uses **GitHub Actions** for automatic builds and releases:

#### How It Works
1. **Automatic Trigger**: Every push to the `main` branch automatically:
   - Builds the project
   - Publishes the portable executable
   - Creates a GitHub release
   - Uploads the `.exe` as a release asset

2. **Version Management**:
   - Version is read from `CopyAsInsert.csproj` (`<Version>` tag)
   - Release tag is auto-generated from version (v1.0.0, v1.0.1, etc.)
   - GitHub releases page is updated with the new executable

3. **Release Assets**:
   - `CopyAsInsert-v1.0.0.exe` - The portable executable
   - `Group-3.ico` - Application icon

#### To Release a New Version
1. Update the version in [CopyAsInsert/CopyAsInsert.csproj](CopyAsInsert/CopyAsInsert.csproj):
   ```xml
   <Version>1.1.0</Version>
   ```
2. Push to main branch
3. GitHub Actions automatically:
   - Builds and publishes
   - Creates release v1.1.0
   - Uploads (`CopyAsInsert-v1.1.0.exe`)
4. Users can download from [**GitHub Releases**](https://github.com/hmygroup/HMY.Tools.Zarpa/releases)

#### Update Notifications
- The app checks for new versions on startup
- Users see a balloon notification if an update is available
- Users can manually check via tray menu: **Check for Update**
- Clicking opens the GitHub releases page to download the new version

#### Applying Updates
```bash
# Run the update script with the new executable
UpdateCopyAsInsert.bat "C:\Downloads\CopyAsInsert-v1.1.0.exe"
```
The script will:
- Find your existing installation
- Create a backup
- Replace the executable
- Restore backup if anything fails (safe rollback)

### Future Enhancements

- [ ] Custom hotkey configuration
- [ ] Multi-row batch INSERT optimization
- [ ] Column name/type manual override in dialog
- [ ] Export history to file
- [ ] Support for other databases (MySQL, PostgreSQL)
- [ ] Drag-drop .xlsx files to process
- [ ] Undo/redo conversion history

---

**Version**: 1.0  
**Last Updated**: February 5, 2026
