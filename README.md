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

- **Auto Type Detection**: INT, FLOAT, DATETIME2, BIT, VARCHAR (70% lenient matching)
- **Primary Key Detection**: Auto-detects first INT or ID-named column as identity PK
- **Temporal Tables**: Includes SysStartTime/SysEndTime with system versioning enabled
- **History Tracking**: View recent conversions (max 10) from tray menu
- **Excel Support**: Both clipboard and .xlsx drag-drop parsing (ClosedXML)
- **Error Handling**: Validates table names, SQL keywords, data formats

### Hotkey Details

- **Default Hotkey**: Alt+Shift+I
- **Why This Key**: Non-intrusive, unlikely to conflict with other apps
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

# Release build
dotnet publish -c Release -o ./publish

# Run from release
./publish/CopyAsInsert.exe
```

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
