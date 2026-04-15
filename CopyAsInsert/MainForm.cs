using CopyAsInsert.Models;
using CopyAsInsert.Services;
using CopyAsInsert.Forms;
using System.Runtime.InteropServices;

namespace CopyAsInsert;

/// <summary>
/// Main application form with system tray integration
/// </summary>
public partial class MainForm : Form
{
    [DllImport("user32.dll")]
    private static extern IntPtr SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int WM_HOTKEY = 0x0312;
    private const int CLIPBOARD_HOTKEY_ID = 9000;
    private const int EXCEL_IMPORT_HOTKEY_ID = 9001;
    private const int MOD_ALT = 0x0001;
    private const int MOD_SHIFT = 0x0004;
    private const int MOD_CTRL = 0x0002;

    private NotifyIcon? _trayIcon;
    private ContextMenuStrip? _contextMenu;
    private ClipboardInterceptor? _clipboardInterceptor;
    private UpdateChecker? _updateChecker;
    private readonly List<ConversionResult> _conversionHistory = new();
    private const int MAX_HISTORY = 10;

    private string _defaultSchema = "dbo";
    private bool _temporalTableByDefault = true;
    private bool _runOnStartup = false;
    private bool _autoAppendTemporalSuffix = false;
    private bool _showFormOnStartup = false;
    private int _hotKeyModifier = 0x0001 | 0x0004; // MOD_ALT | MOD_SHIFT
    private int _hotKeyVirtualKey = 0x49; // 'I'
    private bool _hotkeyRegistered = false;
    
    // Excel import hotkey settings
    private int _excelImportHotKeyModifier = MOD_ALT | MOD_SHIFT;
    private int _excelImportHotKeyVirtualKey = 0x45; // 'E'
    private bool _excelImportHotkeyRegistered = false;

    public MainForm()
    {
        InitializeComponent();
        // Initialize logger early
        Logger.Initialize();
        Logger.LogInfo("MainForm constructor called");
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        // Load settings from file
        var settings = SettingsManager.LoadSettings();
        _defaultSchema = settings.DefaultSchema;
        _temporalTableByDefault = settings.TemporalTableByDefault;
        _runOnStartup = settings.RunOnStartup;
        _autoAppendTemporalSuffix = settings.AutoAppendTemporalSuffix;
        _showFormOnStartup = settings.ShowFormOnStartup;
        _hotKeyModifier = settings.HotKeyModifier;
        _hotKeyVirtualKey = settings.HotKeyVirtualKey;
        _excelImportHotKeyModifier = settings.ExcelImportHotKeyModifier;
        _excelImportHotKeyVirtualKey = settings.ExcelImportHotKeyVirtualKey;
        // Load history from file
        var loadedHistory = HistoryManager.LoadHistory();
        _conversionHistory.Clear();
        _conversionHistory.AddRange(loadedHistory);
        // Set up tray icon first
        SetupTrayIcon();
        // Register hotkeys after form is fully created and has window handle
        SetupHotkey();
        // Check for updates asynchronously
        CheckForUpdatesAsync();
        // Show or hide the form based on settings
        if (_showFormOnStartup)
        {
            this.WindowState = FormWindowState.Normal;
            this.Show();
            SetForegroundWindow(this.Handle);
        }
        else
        {
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
        }
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        // Window handle is now available for hotkey registration
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();

        var iconPath = Path.Combine(AppContext.BaseDirectory, "Group-3.ico");
        this.Text = "ZARPA";
        this.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
        this.Width = 500;
        this.Height = 400;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.ShowIcon = false;
        this.ShowInTaskbar = false;
        this.WindowState = FormWindowState.Normal;
        this.Opacity = 1.0;

        // Status bar
        var statusBar = new StatusStrip();
        var statusLabel = new ToolStripStatusLabel
        {
            Name = "statusLabel",
            Text = "Ready"
        };
        statusBar.Items.Add(statusLabel);

        // Main content area
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(10)
        };

        var lblInfo = new Label
        {
            Text = "ZARPA - SQL INSERT Generator\r\n\r\nPress Alt+Shift+I after copying Excel table data to generate SQL statements.\r\n\r\nSupported formats:\r\n• Clipboard TSV/CSV\r\n• Excel files (.xlsx) drag-drop",
            Left = 10,
            Top = 10,
            Width = 480,
            Height = 150,
            AutoSize = false
        };

        panel.Controls.Add(lblInfo);

        this.Controls.Add(panel);
        this.Controls.Add(statusBar);

        this.Resize += (s, e) =>
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
            }
        };

        this.FormClosing += (s, e) =>
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            }
        };

        this.ResumeLayout(false);
        this.PerformLayout();
    }

    private void SetupTrayIcon()
    {
        if (_trayIcon != null)
            return; // Already initialized

        var iconPath = Path.Combine(AppContext.BaseDirectory, "Group-3.ico");
        _trayIcon = new NotifyIcon
        {
            Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application,
            Visible = true,
            Text = $"ZARPA - {FormatHotkey(_hotKeyModifier, _hotKeyVirtualKey)}"
        };

        _contextMenu = new ContextMenuStrip();

        var exitItem = new ToolStripMenuItem("Exit", null, (s, e) => Application.Exit());
        var settingsItem = new ToolStripMenuItem("Settings", null, (s, e) => ShowSettings());
        var historyItem = new ToolStripMenuItem("View History", null, (s, e) => ShowHistory());
        var updateItem = new ToolStripMenuItem("Check for Update", null, (s, e) => CheckForUpdateManually());
        var restoreItem = new ToolStripMenuItem("Show", null, (s, e) => ShowMainWindow());

        _contextMenu.Items.Add(restoreItem);
        _contextMenu.Items.Add(new ToolStripSeparator());
        _contextMenu.Items.Add(settingsItem);
        _contextMenu.Items.Add(historyItem);
        _contextMenu.Items.Add(new ToolStripSeparator());
        _contextMenu.Items.Add(updateItem);
        _contextMenu.Items.Add(new ToolStripSeparator());
        _contextMenu.Items.Add(exitItem);

        _trayIcon.ContextMenuStrip = _contextMenu;
        _trayIcon.DoubleClick += (s, e) => ShowMainWindow();
        
        // Show that tray icon is ready
        // _trayIcon.ShowBalloonTip(1500, "ZARPA", "Starting...", ToolTipIcon.Info);
    }

    private string FormatHotkey(int modifiers, int vKey)
    {
        var keys = new List<string>();

        if ((modifiers & 0x0002) != 0) // MOD_CTRL
            keys.Add("Ctrl");
        if ((modifiers & 0x0001) != 0) // MOD_ALT
            keys.Add("Alt");
        if ((modifiers & 0x0004) != 0) // MOD_SHIFT
            keys.Add("Shift");

        // Convert virtual key to character
        if (vKey >= 0x41 && vKey <= 0x5A) // A-Z
        {
            keys.Add(((char)vKey).ToString());
        }
        else if (vKey >= 0x30 && vKey <= 0x39) // 0-9
        {
            keys.Add(((char)vKey).ToString());
        }
        else
        {
            keys.Add($"0x{vKey:X}");
        }

        return string.Join("+", keys);
    }

    private void SetupHotkey()
    {
        if (_hotkeyRegistered && _excelImportHotkeyRegistered)
            return; // Both already registered

        // Set up clipboard interceptor for clipboard hotkey
        if (!_hotkeyRegistered)
        {
            _clipboardInterceptor = new ClipboardInterceptor();
            _clipboardInterceptor.HotKeyPressed += OnHotKeyPressed;

            try
            {
                _clipboardInterceptor.InitializeHotKey(this.Handle, _hotKeyModifier, _hotKeyVirtualKey);
                _hotkeyRegistered = true;
                
                Logger.LogInfo($"Hotkey {FormatHotkey(_hotKeyModifier, _hotKeyVirtualKey)} registered successfully");
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to register hotkey", ex);
                // Try to fall back to Alt+Shift+I
                try
                {
                    Logger.LogInfo("Attempting fallback to Alt+Shift+I...");
                    _clipboardInterceptor.InitializeHotKey(this.Handle, ClipboardInterceptor.MOD_ALT | ClipboardInterceptor.MOD_SHIFT, 0x49);
                    _hotkeyRegistered = true;
                    _hotKeyModifier = ClipboardInterceptor.MOD_ALT | ClipboardInterceptor.MOD_SHIFT;
                    _hotKeyVirtualKey = 0x49;
                    Logger.LogInfo("Fallback hotkey Alt+Shift+I registered successfully");
                }
                catch (Exception ex2)
                {
                    Logger.LogError("Failed to register fallback hotkey", ex2);
                }
            }
        }

        // Register Excel import hotkey directly
        if (!_excelImportHotkeyRegistered)
        {
            try
            {
                if (RegisterHotKey(this.Handle, EXCEL_IMPORT_HOTKEY_ID, _excelImportHotKeyModifier, _excelImportHotKeyVirtualKey))
                {
                    _excelImportHotkeyRegistered = true;
                    Logger.LogInfo($"Excel import hotkey {FormatHotkey(_excelImportHotKeyModifier, _excelImportHotKeyVirtualKey)} registered successfully");
                }
                else
                {
                    Logger.LogWarning("Failed to register Excel import hotkey");
                    // Try fallback: Alt+Shift+E
                    try
                    {
                        Logger.LogInfo("Attempting fallback to Alt+Shift+E for Excel import...");
                        if (RegisterHotKey(this.Handle, EXCEL_IMPORT_HOTKEY_ID, MOD_ALT | MOD_SHIFT, 0x45))
                        {
                            _excelImportHotkeyRegistered = true;
                            _excelImportHotKeyModifier = MOD_ALT | MOD_SHIFT;
                            _excelImportHotKeyVirtualKey = 0x45;
                            Logger.LogInfo("Fallback Excel import hotkey Alt+Shift+E registered successfully");
                        }
                    }
                    catch (Exception ex2)
                    {
                        Logger.LogError("Failed to register fallback Excel import hotkey", ex2);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to register Excel import hotkey", ex);
            }
        }
    }

    protected override void WndProc(ref Message m)
    {
        // Handle Excel import hotkey directly
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == EXCEL_IMPORT_HOTKEY_ID)
        {
            OnExcelImportHotKeyPressed();
            return;
        }

        _clipboardInterceptor?.ProcessWindowMessage(ref m);
        base.WndProc(ref m);
    }

    private void OnHotKeyPressed(object? sender, EventArgs e)
    {
        ProcessClipboard();
    }

    private void OnExcelImportHotKeyPressed()
    {
        try
        {
            Logger.LogDebug("Excel import hotkey pressed");
            var excelForm = new ExcelImportForm();
            excelForm.ShowDialog();
        }
        catch (Exception ex)
        {
            Logger.LogError("Error opening Excel import form", ex);
            MessageBox.Show($"Error opening Excel import form: {ex.Message}", "Error", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ProcessClipboard()
    {
        try
        {
            Logger.LogDebug("ProcessClipboard started");
            var clipboardText = ClipboardInterceptor.GetClipboardText();

            if (string.IsNullOrEmpty(clipboardText) || !ClipboardInterceptor.IsClipboardTabularData())
            {
                Logger.LogInfo("Clipboard does not contain tabular data");
                // _trayIcon?.ShowBalloonTip(2000, "No Table Data", "Clipboard does not contain tabular data (TSV/CSV)", ToolTipIcon.Info);
                return;
            }

            Logger.LogDebug("Clipboard contains tabular data, showing header check form");

            // Ask user if data has headers
            var headerCheckForm = new HeaderCheckForm();
            var headerCheckResult = headerCheckForm.ShowDialog();
            if (headerCheckResult != DialogResult.Yes && headerCheckResult != DialogResult.No)
            {
                Logger.LogDebug("User cancelled header check form");
                return; // User cancelled
            }

            bool hasHeaders = headerCheckForm.HasHeaders;
            Logger.LogInfo($"Header check result: hasHeaders={hasHeaders}");

            // Parse clipboard data with header check
            var schema = TableDataParser.ParseClipboardText(clipboardText, hasHeaders);
            Logger.LogInfo($"Clipboard parsed successfully: {schema.Columns.Count} columns, {schema.DataRows.Count} rows");

            // Infer column types
            TypeInferenceEngine.InferColumnTypes(schema);
            Logger.LogDebug("Column types inferred");

            // Show config dialog with schema for type override
            var configForm = new TableConfigForm(_defaultSchema);
            configForm.SetSchema(schema);  // Load schema into type override control
            
            if (configForm.ShowDialog() == DialogResult.OK)
            {
                Logger.LogInfo($"Config form accepted: TableName={configForm.TableName}, Schema={configForm.SchemaName}");

                // Use the potentially modified schema from config form
                var finalSchema = configForm.Schema ?? schema;

                // Generate SQL with final schema (including any user-overridden types)
                var result = SqlServerGenerator.GenerateSql(finalSchema, configForm.TableName, configForm.SchemaName, false, configForm.IsTemporaryTable, _autoAppendTemporalSuffix);

                if (result.Success)
                {
                    Logger.LogInfo($"SQL generated successfully: {result.Summary}");

                    // Copy SQL to clipboard
                    ClipboardInterceptor.SetClipboardText(result.GeneratedSql);

                    // Add to history
                    AddToHistory(result);

                    // Show success notification
                    var message = $"SQL generated for {result.Summary}\n{result.RowCount} rows inserted";
                    _trayIcon?.ShowBalloonTip(3000, "Success", message, ToolTipIcon.Info);
                }
                else
                {
                    Logger.LogError($"SQL generation failed: {result.ErrorMessage}");
                    _trayIcon?.ShowBalloonTip(3000, "Error", $"Failed to generate SQL: {result.ErrorMessage}", ToolTipIcon.Error);
                }
            }
            else
            {
                Logger.LogDebug("User cancelled config form");
            }
        }
        catch (Exception ex)
        {
            Logger.LogError("Error processing clipboard", ex);
            // _trayIcon?.ShowBalloonTip(3000, "Error", $"Error processing clipboard: {ex.Message}", ToolTipIcon.Error);
        }
    }

    private void AddToHistory(ConversionResult result)
    {
        _conversionHistory.Insert(0, result);
        if (_conversionHistory.Count > MAX_HISTORY)
        {
            _conversionHistory.RemoveAt(_conversionHistory.Count - 1);
        }
        // Save history to file
        HistoryManager.SaveHistory(_conversionHistory);
    }

    private void ShowMainWindow()
    {
        this.WindowState = FormWindowState.Normal;
        this.Show();
        SetForegroundWindow(this.Handle);
    }

    private void ShowSettings()
    {
        var settingsForm = new SettingsForm
        {
            DefaultSchema = _defaultSchema,
            TemporalTableByDefault = _temporalTableByDefault,
            RunOnStartup = _runOnStartup,
            AutoAppendTemporalSuffix = _autoAppendTemporalSuffix,
            ShowFormOnStartup = _showFormOnStartup,
            HotKeyModifier = _hotKeyModifier,
            HotKeyVirtualKey = _hotKeyVirtualKey,
            ExcelImportHotKeyModifier = _excelImportHotKeyModifier,
            ExcelImportHotKeyVirtualKey = _excelImportHotKeyVirtualKey
        };

        if (settingsForm.ShowDialog() == DialogResult.OK)
        {
            _defaultSchema = settingsForm.DefaultSchema;
            _temporalTableByDefault = settingsForm.TemporalTableByDefault;
            _runOnStartup = settingsForm.RunOnStartup;
            _autoAppendTemporalSuffix = settingsForm.AutoAppendTemporalSuffix;
            _showFormOnStartup = settingsForm.ShowFormOnStartup;

            bool clipboardHotkeyChanged = _hotKeyModifier != settingsForm.HotKeyModifier || 
                                _hotKeyVirtualKey != settingsForm.HotKeyVirtualKey;
            
            bool excelHotkeyChanged = _excelImportHotKeyModifier != settingsForm.ExcelImportHotKeyModifier || 
                                      _excelImportHotKeyVirtualKey != settingsForm.ExcelImportHotKeyVirtualKey;

            // Save settings to file
            var settings = new SettingsManager.ApplicationSettings
            {
                DefaultSchema = _defaultSchema,
                AutoCreateHistoryTable = true, // This wasn't being used, keeping as default
                TemporalTableByDefault = _temporalTableByDefault,
                RunOnStartup = _runOnStartup,
                AutoAppendTemporalSuffix = _autoAppendTemporalSuffix,
                ShowFormOnStartup = _showFormOnStartup,
                HotKeyModifier = settingsForm.HotKeyModifier,
                HotKeyVirtualKey = settingsForm.HotKeyVirtualKey,
                ExcelImportHotKeyModifier = settingsForm.ExcelImportHotKeyModifier,
                ExcelImportHotKeyVirtualKey = settingsForm.ExcelImportHotKeyVirtualKey
            };
            SettingsManager.SaveSettings(settings);

            // Update clipboard hotkey if it changed
            if (clipboardHotkeyChanged && _clipboardInterceptor != null)
            {
                _hotKeyModifier = settingsForm.HotKeyModifier;
                _hotKeyVirtualKey = settingsForm.HotKeyVirtualKey;

                bool hotkeyUpdated = _clipboardInterceptor.UpdateHotKey(_hotKeyModifier, _hotKeyVirtualKey);

                if (!hotkeyUpdated)
                {
                    MessageBox.Show(
                        $"Could not update clipboard hotkey to {FormatHotkey(_hotKeyModifier, _hotKeyVirtualKey)}.\n\n" +
                        "It may be in use by another application. Please try a different combination.",
                        "Hotkey Registration Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    // Revert to previous settings in file
                    settings.HotKeyModifier = _hotKeyModifier;
                    settings.HotKeyVirtualKey = _hotKeyVirtualKey;
                }
                else
                {
                    // Update tray icon text
                    if (_trayIcon != null)
                    {
                        _trayIcon.Text = $"ZARPA - {FormatHotkey(_hotKeyModifier, _hotKeyVirtualKey)}";
                    }
                    Logger.LogInfo($"Clipboard hotkey updated to {FormatHotkey(_hotKeyModifier, _hotKeyVirtualKey)}");
                }
            }

            // Update Excel import hotkey if it changed
            if (excelHotkeyChanged && _excelImportHotkeyRegistered)
            {
                _excelImportHotKeyModifier = settingsForm.ExcelImportHotKeyModifier;
                _excelImportHotKeyVirtualKey = settingsForm.ExcelImportHotKeyVirtualKey;

                // Unregister old hotkey
                UnregisterHotKey(this.Handle, EXCEL_IMPORT_HOTKEY_ID);
                _excelImportHotkeyRegistered = false;

                // Register new hotkey
                if (RegisterHotKey(this.Handle, EXCEL_IMPORT_HOTKEY_ID, _excelImportHotKeyModifier, _excelImportHotKeyVirtualKey))
                {
                    _excelImportHotkeyRegistered = true;
                    Logger.LogInfo($"Excel import hotkey updated to {FormatHotkey(_excelImportHotKeyModifier, _excelImportHotKeyVirtualKey)}");
                }
                else
                {
                    MessageBox.Show(
                        $"Could not update Excel import hotkey to {FormatHotkey(_excelImportHotKeyModifier, _excelImportHotKeyVirtualKey)}.\n\n" +
                        "It may be in use by another application. Please try a different combination.",
                        "Excel Import Hotkey Registration Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    // Revert to previous settings
                    _excelImportHotKeyModifier = settings.ExcelImportHotKeyModifier;
                    _excelImportHotKeyVirtualKey = settings.ExcelImportHotKeyVirtualKey;
                    
                    // Re-register old hotkey
                    if (RegisterHotKey(this.Handle, EXCEL_IMPORT_HOTKEY_ID, _excelImportHotKeyModifier, _excelImportHotKeyVirtualKey))
                    {
                        _excelImportHotkeyRegistered = true;
                    }
                }
            }

            // Update Registry for startup
            if (_runOnStartup)
            {
                StartupManager.EnableStartup();
            }
            else
            {
                StartupManager.DisableStartup();
            }
        }
    }

    private void ShowHistory()
    {
        if (_conversionHistory.Count == 0)
        {
            MessageBox.Show("No conversion history", "History", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var historyForm = new HistoryForm(_conversionHistory);
        if (historyForm.ShowDialog() == DialogResult.OK)
        {
            // History may have been modified (items deleted), save it
            HistoryManager.SaveHistory(_conversionHistory);
        }
    }

    private async void CheckForUpdatesAsync()
    {
        try
        {
            _updateChecker = new UpdateChecker();
            var result = await _updateChecker.CheckForUpdatesAsync();

            if (result.IsUpdateAvailable)
            {
                // Show notification asynchronously (on UI thread)
                this.Invoke(() =>
                {
                    _trayIcon?.ShowBalloonTip(5000, "Update Available",
                        $"CopyAsInsert {result.AvailableVersion} is available.\nClick 'Check for Update' to download.",
                        ToolTipIcon.Info);
                });
            }
        }
        catch (Exception ex)
        {
            // Silently fail - don't bother user with update check errors
            System.Diagnostics.Debug.WriteLine($"Update check failed: {ex.Message}");
        }
    }

    private async void CheckForUpdateManually()
    {
        try
        {
            _updateChecker ??= new UpdateChecker();
            var result = await _updateChecker.CheckForUpdatesAsync();

            if (result.IsUpdateAvailable)
            {
                var message = $"Update Available!\n\n" +
                    $"Current Version: {result.CurrentVersion}\n" +
                    $"Latest Version: {result.AvailableVersion}\n\n" +
                    $"Would you like to download and install the update?";

                var dialogResult = MessageBox.Show(message, "Update Available",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dialogResult == DialogResult.Yes)
                {
                    // Launch the updater with download URL and version
                    LaunchUpdateProcess(result);
                }
            }
            else if (string.IsNullOrEmpty(result.ErrorMessage))
            {
                // No error and no update available - show appropriate message
                MessageBox.Show($"You are using version {result.CurrentVersion}.\n\n" +
                    $"No newer releases are currently available.",
                    "No Update Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show($"Could not check for updates:\n{result.ErrorMessage}",
                    "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error checking for updates:\n{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LaunchUpdateProcess(UpdateCheckResult result)
    {
        try
        {
            if (string.IsNullOrEmpty(result.DownloadUrl) || string.IsNullOrEmpty(result.AvailableVersion))
            {
                MessageBox.Show("Could not determine update information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _updateChecker ??= new UpdateChecker();
            var updaterPath = _updateChecker.GetUpdaterPath();

            if (!File.Exists(updaterPath))
            {
                MessageBox.Show($"Updater not found: {updaterPath}\n\nPlease download and reinstall the application.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Logger.LogInfo($"Launching updater: {updaterPath}");
            Logger.LogInfo($"Download URL: {result.DownloadUrl}");
            Logger.LogInfo($"Version: {result.AvailableVersion}");

            // Build arguments for the updater
            var appPath = AppContext.BaseDirectory.TrimEnd('\\');
            var args = $"--version {result.AvailableVersion} --url \"{result.DownloadUrl}\" --app-path \"{appPath}\"";

            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = updaterPath,
                Arguments = args,
                UseShellExecute = true,
                CreateNoWindow = false
            };

            System.Diagnostics.Process.Start(startInfo);

            // Force immediate exit to release file locks for updater
            Logger.LogInfo("Update approved. Closing application immediately for file replacement.");
            Logger.CloseAndFlush();
            Environment.Exit(0);
        }
        catch (Exception ex)
        {
            Logger.LogError("Error launching update process", ex);
            MessageBox.Show($"Error starting update:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }


    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Logger.LogInfo("MainForm disposing");
            _clipboardInterceptor?.Dispose();
            _trayIcon?.Dispose();
            _contextMenu?.Dispose();
            _updateChecker = null;
            Logger.CloseAndFlush();
        }
        base.Dispose(disposing);
    }
}
