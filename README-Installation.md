# CopyAsInsert - Installation Guide

## Quick Start

**CopyAsInsert** is a portable application that includes everything needed to run. No installation or prerequisites required!

### System Requirements
- Windows 10 or Windows 11
- At least 100 MB free disk space
- No additional software needed (runtime is included)

---

## Installation Methods

### Method 1: Simple Portable Execution (Recommended)

1. **Download** `CopyAsInsert.exe` from the releases page
2. **Place** the file anywhere you prefer:
   - Desktop for quick access
   - A dedicated folder (e.g., `C:\Apps\CopyAsInsert\`)
   - Create a new folder for organization
3. **Run** by double-clicking `CopyAsInsert.exe`
4. **Allow** the Windows Defender/Security prompt if it appears (first run only)

The application will:
- Launch and minimize to the system tray
- Register the global hotkey: **Alt+Shift+I**
- Run in the background, monitoring your clipboard

### Method 2: Enable Auto-Start (Optional)

To have CopyAsInsert automatically start when Windows boots:

1. **Download** and extract `CopyAsInsert.exe` to your chosen location
2. **Open** Command Prompt (or PowerShell)
3. **Navigate** to the folder (or use the script directly):
   ```cmd
   cd Scripts
   SetupAutoStart.bat
   ```
4. **Choose option 1** to enable auto-start
5. **Restart** your computer to verify it enables automatically

**To disable auto-start:**
- Run `SetupAutoStart.bat` again and choose option 2

---

## Usage

### Accessing the Application

Once running, CopyAsInsert operates silently in the system tray:

1. **System Tray Icon**: Look for the CopyAsInsert icon in the bottom-right corner of your taskbar
2. **Right-click** the icon to see options:
   - **Show** - Display the main window
   - **Settings** - Configure application behavior
   - **View History** - See processed clipboard entries
   - **Exit** - Stop the application

### Global Hotkey

Press **Alt+Shift+I** from anywhere to trigger the primary function (configured in Settings).

### Main Features

- **Clipboard Monitor**: Automatically detects copied data
- **Table Detection**: Recognizes when you copy table data
- **SQL Generation**: Converts table data to INSERT statements
- **Excel Support**: Handles Excel files and spreadsheets

---

## Troubleshooting

### Application Won't Start

**Issue**: Double-clicking `CopyAsInsert.exe` doesn't launch the app

**Solutions**:
1. Try running as Administrator:
   - Right-click `CopyAsInsert.exe` → "Run as administrator"
2. Check Windows Defender SmartScreen:
   - Click "More info" → "Run anyway"
3. Verify Windows 10/11 is up to date:
   - Settings → Update & Security → Windows Update
4. Check antivirus software:
   - Some antivirus tools may block unknown applications
   - Add `CopyAsInsert.exe` to your antivirus whitelist

### Hotkey Not Working

**Issue**: Alt+Shift+I doesn't trigger the function

**Solutions**:
1. Verify the app is running by checking the system tray
2. Some applications may override this hotkey:
   - Check Settings → Keyboard Shortcuts
   - Try a different hotkey in CopyAsInsert Settings
3. Restart the application after changing hotkey settings

### Application Keeps Minimizing

**Issue**: Windows keeps showing the CopyAsInsert window instead of staying hidden

**Expected Behavior**: CopyAsInsert is designed to run minimized in the system tray. You can access it through:
- Right-click the system tray icon
- Press the global hotkey
- Click "Show" from the system tray menu

### Auto-Start Script Failing

**Issue**: `SetupAutoStart.bat` produces an error

**Solutions**:
1. **Run as Administrator**:
   - Right-click `SetupAutoStart.bat` → "Run as administrator"
2. **Windows Defender/SmartScreen**:
   - Click "More info" → "Run anyway"
3. **Manual Setup** (alternative):
   - Press `Win+R` and type: `shell:startup`
   - Create a shortcut to `CopyAsInsert.exe` in this folder
   - Restart Windows to test

---

## Performance & Resources

- **Startup Time**: ~1-2 seconds
- **Memory Usage**: ~50-100 MB (typical)
- **CPU Usage**: Minimal when idle (monitors clipboard only)
- **Disk Space**: ~60-80 MB (single file with Python runtime)

---

## Getting Help

If you encounter issues not covered above:

1. Check the application logs:
   - Look in the `%APPDATA%\CopyAsInsert\` folder for log files
2. Review Settings:
   - Click "Show" on system tray icon → Settings tab
   - Verify configuration matches your needs
3. Report issues:
   - Include Windows version and error messages
   - Description of what you were doing when the issue occurred

---

## Uninstallation

Since CopyAsInsert is portable, uninstallation is simple:

1. **Disable Auto-Start** (if enabled):
   - Run `SetupAutoStart.bat` → Choose option 2
2. **Delete the File**:
   - Simply delete `CopyAsInsert.exe`
   - No registry entries or leftover files
3. **Optional**: Delete the configuration folder:
   - Press `Win+R` and type: `%APPDATA%\CopyAsInsert`
   - Delete the folder

That's it! The application leaves no footprint on your system.

---

## Tips & Best Practices

✅ **DO:**
- Keep the executable in a permanent location if using auto-start
- Write down your Settings configuration in case you need to reinstall
- Check for updates regularly at the releases page

❌ **DON'T:**
- Move the executable while it's running
- Delete the `.exe` file while it's actively running
- Disable notifications if you want to be aware of clipboard monitoring

---

## Version & Updates

**Current Version**: `1.0.0`  
**Release Date**: February 2026

### Checking for Updates

CopyAsInsert checks for updates automatically on startup. You'll be notified if a new version is available.

### Manual Update

To manually update to a new version:

1. **Download** the new `CopyAsInsert.exe` from the [GitHub Releases](https://github.com/hmygroup/HMY.Tools.Zarpa/releases)
2. **Run the update script**:
   ```cmd
   UpdateCopyAsInsert.bat "path-to-new-CopyAsInsert.exe"
   ```
   Example: `UpdateCopyAsInsert.bat "C:\Downloads\CopyAsInsert.exe"`
3. **Script Features**:
   - Automatically detects existing installation
   - Creates backup of current version
   - Stops the running app safely
   - Replaces the executable
   - Restores backup if update fails (rollback protection)

### Update Locations

The script searches for installed CopyAsInsert in:
- `Program Files\CopyAsInsert\`
- `AppData\Local\CopyAsInsert\`
- Current directory

If none found, you can manually copy to your preferred location.

### Version History

- **1.0.0** (February 2026) - Initial Release
  - System tray integration
  - Global hotkey support (Alt+Shift+I)
  - Clipboard monitoring
  - SQL INSERT generation
  - Excel file support

Check the releases page for:
- Bug fixes
- New features
- Performance improvements
- Security updates

---

**Questions or feedback?** Let us know!
