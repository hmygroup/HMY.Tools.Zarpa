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
    private bool _hotkeyRegistered = false;

    public MainForm()
    {
        InitializeComponent();
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
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
            
            // Show confirmation
            if (_trayIcon != null)
                _trayIcon.ShowBalloonTip(2000, "Ready", "Alt+Shift+I registered successfully", ToolTipIcon.Info);
        }
        catch (Exception ex)
        {
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
            var clipboardText = ClipboardInterceptor.GetClipboardText();

            if (string.IsNullOrEmpty(clipboardText) || !ClipboardInterceptor.IsClipboardTabularData())
            {
                _trayIcon?.ShowBalloonTip(2000, "No Table Data", "Clipboard does not contain tabular data (TSV/CSV)", ToolTipIcon.Info);
                return;
            }

            // Ask user if data has headers
            var headerCheckForm = new HeaderCheckForm();
            var headerCheckResult = headerCheckForm.ShowDialog();
            if (headerCheckResult != DialogResult.Yes && headerCheckResult != DialogResult.No)
                return; // User cancelled

            bool hasHeaders = headerCheckForm.HasHeaders;

            // Parse clipboard data with header check
            var schema = TableDataParser.ParseClipboardText(clipboardText, hasHeaders);

            // Infer column types
            TypeInferenceEngine.InferColumnTypes(schema);

            // Show config dialog
            var configForm = new TableConfigForm();
            if (configForm.ShowDialog() == DialogResult.OK)
            {
                // Generate SQL
                var result = SqlServerGenerator.GenerateSql(schema, configForm.TableName, configForm.SchemaName, configForm.IsTemporalTable, configForm.IsTemporaryTable);

                if (result.Success)
                {
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
                    _trayIcon?.ShowBalloonTip(3000, "Error", $"Failed to generate SQL: {result.ErrorMessage}", ToolTipIcon.Error);
                }
            }
        }
        catch (Exception ex)
        {
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
            TemporalTableByDefault = _temporalTableByDefault
        };

        if (settingsForm.ShowDialog() == DialogResult.OK)
        {
            _defaultSchema = settingsForm.DefaultSchema;
            _temporalTableByDefault = settingsForm.TemporalTableByDefault;
        }
    }

    private void ShowHistory()
    {
        if (_conversionHistory.Count == 0)
        {
            MessageBox.Show("No conversion history", "History", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var historyText = "Recent Conversions:\r\n\r\n";
        for (int i = 0; i < _conversionHistory.Count; i++)
        {
            var result = _conversionHistory[i];
            historyText += $"{i + 1}. {result.Summary} - {result.ConversionTime:g}\r\n";
        }

        MessageBox.Show(historyText, "Conversion History", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    $"Would you like to download the update?";

                var dialogResult = MessageBox.Show(message, "Update Available",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dialogResult == DialogResult.Yes)
                {
                    // Open the download URL or GitHub releases page
                    var downloadUrl = result.DownloadUrl ?? "https://github.com/hmygroup/HMY.Tools.Zarpa/releases";
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = downloadUrl,
                        UseShellExecute = true
                    });

                    // Show instructions
                    MessageBox.Show(
                        "After downloading:\n\n" +
                        "1. Run UpdateCopyAsInsert.bat\n" +
                        "2. Select the downloaded CopyAsInsert.exe\n\n" +
                        "The script will handle the update automatically.",
                        "Update Instructions",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
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

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _clipboardInterceptor?.Dispose();
            _trayIcon?.Dispose();
            _contextMenu?.Dispose();
            _updateChecker = null;
        }
        base.Dispose(disposing);
    }
}
