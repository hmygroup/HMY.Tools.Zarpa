using CopyAsInsert.Models;
using CopyAsInsert.Services;
using CopyAsInsert.Forms;

namespace CopyAsInsert;

/// <summary>
/// Main application form with system tray integration
/// </summary>
public partial class MainForm : Form
{
    private NotifyIcon? _trayIcon;
    private ContextMenuStrip? _contextMenu;
    private ClipboardInterceptor? _clipboardInterceptor;
    private UpdateChecker? _updateChecker;
    private readonly List<ConversionResult> _conversionHistory = new();
    private const int MAX_HISTORY = 10;

    private string _defaultSchema = "dbo";
    private bool _temporalTableByDefault = true;
    private bool _runOnStartup = false;
    private bool _hotkeyRegistered = false;

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
        // Load history from file
        var loadedHistory = HistoryManager.LoadHistory();
        _conversionHistory.Clear();
        _conversionHistory.AddRange(loadedHistory);
        // Set up tray icon first
        SetupTrayIcon();
        // Register hotkey after form is fully created and has window handle
        SetupHotkey();
        // Check for updates asynchronously
        CheckForUpdatesAsync();
        // Now hide the form (it stays in message loop but invisible)
        this.WindowState = FormWindowState.Minimized;
        this.Hide();
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
            Text = "ZARPA - Alt+Shift+I"
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
        _trayIcon.ShowBalloonTip(1500, "ZARPA", "Starting...", ToolTipIcon.Info);
    }

    private void SetupHotkey()
    {
        if (_hotkeyRegistered)
            return; // Already registered

        _clipboardInterceptor = new ClipboardInterceptor();
        _clipboardInterceptor.HotKeyPressed += OnHotKeyPressed;

        try
        {
            _clipboardInterceptor.InitializeHotKey(this.Handle);
            _hotkeyRegistered = true;
            
            Logger.LogInfo("Hotkey Alt+Shift+I registered successfully");
            
            // Show confirmation
            if (_trayIcon != null)
                _trayIcon.ShowBalloonTip(2000, "Ready", "Alt+Shift+I registered successfully", ToolTipIcon.Info);
        }
        catch (Exception ex)
        {
            Logger.LogError("Failed to register hotkey", ex);
            if (_trayIcon != null)
                _trayIcon.ShowBalloonTip(3000, "Error", $"Hotkey registration failed: {ex.Message}", ToolTipIcon.Error);
        }
    }

    protected override void WndProc(ref Message m)
    {
        _clipboardInterceptor?.ProcessWindowMessage(ref m);
        base.WndProc(ref m);
    }

    private void OnHotKeyPressed(object? sender, EventArgs e)
    {
        ProcessClipboard();
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
                _trayIcon?.ShowBalloonTip(2000, "No Table Data", "Clipboard does not contain tabular data (TSV/CSV)", ToolTipIcon.Info);
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

            // Show config dialog
            var configForm = new TableConfigForm();
            if (configForm.ShowDialog() == DialogResult.OK)
            {
                Logger.LogInfo($"Config form accepted: TableName={configForm.TableName}, Schema={configForm.SchemaName}");

                // Generate SQL
                var result = SqlServerGenerator.GenerateSql(schema, configForm.TableName, configForm.SchemaName, configForm.IsTemporalTable, configForm.IsTemporaryTable);

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
            _trayIcon?.ShowBalloonTip(3000, "Error", $"Error processing clipboard: {ex.Message}", ToolTipIcon.Error);
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
        this.Activate();
    }

    private void ShowSettings()
    {
        var settingsForm = new SettingsForm
        {
            DefaultSchema = _defaultSchema,
            TemporalTableByDefault = _temporalTableByDefault,
            RunOnStartup = _runOnStartup
        };

        if (settingsForm.ShowDialog() == DialogResult.OK)
        {
            _defaultSchema = settingsForm.DefaultSchema;
            _temporalTableByDefault = settingsForm.TemporalTableByDefault;
            _runOnStartup = settingsForm.RunOnStartup;

            // Save settings to file
            var settings = new SettingsManager.ApplicationSettings
            {
                DefaultSchema = _defaultSchema,
                AutoCreateHistoryTable = true, // This wasn't being used, keeping as default
                TemporalTableByDefault = _temporalTableByDefault,
                RunOnStartup = _runOnStartup
            };
            SettingsManager.SaveSettings(settings);

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
