using CopyAsInsert.Services;
using System.Runtime.InteropServices;
using CopyAsInsert.Models;
using System.Threading;
using System.Threading.Tasks;
using System;
using System.Windows.Forms;

namespace CopyAsInsert.Forms;

public class ExcelImportForm : Form
{
    private TextBox serverTextBox = new();
    private TextBox databaseTextBox = new();
    private Button importButton = new();
    private Button cancelButton = new();
    private RichTextBox logTextBox = new();
    private string? clipboardQuery;
    private bool _importRunning = false;

    public ExcelImportForm()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        this.Text = "Import SQL Query to Excel";
        this.Size = new Size(500, 380);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // Server label and textbox
        Label serverLabel = new()
        {
            Text = "SQL Server:",
            Location = new Point(20, 20),
            Size = new Size(100, 25),
            AutoSize = false
        };
        this.Controls.Add(serverLabel);

        serverTextBox.Location = new Point(150, 20);
        serverTextBox.Size = new Size(300, 25);
        serverTextBox.Text = "(localhost)";
        this.Controls.Add(serverTextBox);

        // Database label and textbox
        Label databaseLabel = new()
        {
            Text = "Database:",
            Location = new Point(20, 55),
            Size = new Size(100, 25),
            AutoSize = false
        };
        this.Controls.Add(databaseLabel);

        databaseTextBox.Location = new Point(150, 55);
        databaseTextBox.Size = new Size(300, 25);
        this.Controls.Add(databaseTextBox);

        // Log text area (replaces progress bar)
        logTextBox.Location = new Point(20, 100);
        logTextBox.Size = new Size(430, 180);
        logTextBox.ReadOnly = true;
        logTextBox.Multiline = true;
        logTextBox.ScrollBars = RichTextBoxScrollBars.Vertical;
        logTextBox.Visible = false;
        this.Controls.Add(logTextBox);

        // Import Button
        importButton.Text = "Import to Excel";
        importButton.Location = new Point(150, 300);
        importButton.Size = new Size(120, 30);
        importButton.Click += ImportButton_Click;
        this.Controls.Add(importButton);

        // Cancel Button
        cancelButton.Text = "Cancel";
        cancelButton.Location = new Point(280, 300);
        cancelButton.Size = new Size(100, 30);
        cancelButton.Click += (s, e) => this.Close();
        this.Controls.Add(cancelButton);
    }

    private void LoadSettings()
    {
        try
        {
            var settings = SettingsManager.LoadSettings();
            serverTextBox.Text = string.IsNullOrEmpty(settings.ExcelImportServer) 
                ? "(localhost)" 
                : settings.ExcelImportServer;
            databaseTextBox.Text = settings.ExcelImportDatabase;
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error loading settings: {ex.Message}");
        }
    }

    private void SaveSettings()
    {
        try
        {
            var settings = SettingsManager.LoadSettings();
            settings.ExcelImportServer = serverTextBox.Text;
            settings.ExcelImportDatabase = databaseTextBox.Text;
            SettingsManager.SaveSettings(settings);
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error saving settings: {ex.Message}");
        }
    }

    private async void ImportButton_Click(object? sender, EventArgs e)
    {
        // Get query from clipboard
        clipboardQuery = GetClipboardText();
        if (string.IsNullOrWhiteSpace(clipboardQuery))
        {
            MessageBox.Show("Clipboard is empty. Please copy a SQL query first.",
                "No Query", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string server = serverTextBox.Text.Trim();
        string database = databaseTextBox.Text.Trim();

        if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(database))
        {
            MessageBox.Show("Please enter server and database names.",
                "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        importButton.Enabled = false;
        cancelButton.Enabled = false;
        logTextBox.Visible = true;
        logTextBox.Clear();

        // Ensure logger is initialized and subscribe to its messages
        Logger.Initialize();
        Action<string, string> onLog = (lvl, msg) => AppendLog(lvl, msg);
        Logger.MessageLogged += onLog;
        _importRunning = true;

        try
        {
            Logger.LogInfo($"Starting Excel import to {server}/{database}");

            var tcs = new TaskCompletionSource<ImportResult>();
            var staThread = new Thread(() =>
            {
                try
                {
                    var res = ExcelInteropManager.InjectQueryIntoExcel(server, database, clipboardQuery);
                    tcs.SetResult(res);
                }
                catch (Exception ex)
                {
                    tcs.SetResult(new ImportResult { Success = false, ErrorMessage = ex.Message, ErrorStackTrace = ex.StackTrace });
                }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.IsBackground = true;
            staThread.Start();

            var result = await tcs.Task;

            if (result.Success)
            {
                SaveSettings();
                Logger.LogInfo($"Query imported successfully: {result.RowCount} rows");
                // Clear import-running flag before closing so OnFormClosing doesn't prompt
                _importRunning = false;
                this.Close();

                // Try to set focus to Excel if hwnd was returned
                try
                {
                    if (result.ExcelHwnd.HasValue)
                    {
                        SetForegroundWindow(new IntPtr(result.ExcelHwnd.Value));
                    }
                }
                catch { }
            }
            else
            {
                string errorMsg = result.ErrorMessage ?? "Unknown error";
                string stackTrace = result.ErrorStackTrace ?? string.Empty;
                Logger.LogError($"Import failed: {errorMsg}\n{stackTrace}");

                // Bring this form to front so the error dialog is reachable
                try { SetForegroundWindow(this.Handle); this.BringToFront(); this.Activate(); } catch { }

                var errorForm = new ErrorDetailForm(errorMsg, stackTrace);
                errorForm.ShowDialog(this);
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Unexpected error during import: {ex.Message}\n{ex.StackTrace}");
            var errorForm = new ErrorDetailForm(
                $"Unexpected error: {ex.Message}",
                ex.StackTrace ?? "No stack trace available"
            );
            errorForm.ShowDialog(this);
        }
        finally
        {
            _importRunning = false;
            Logger.MessageLogged -= onLog;
            importButton.Enabled = true;
            cancelButton.Enabled = true;
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_importRunning)
        {
            try { SetForegroundWindow(this.Handle); this.BringToFront(); this.Activate(); } catch { }
            var dr = MessageBox.Show("An import is running. Close anyway?", "Import in progress", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }
        }
        base.OnFormClosing(e);
    }

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    private void AppendLog(string level, string message)
    {
        if (logTextBox.InvokeRequired)
        {
            logTextBox.BeginInvoke(new Action(() => AppendLog(level, message)));
            return;
        }

        logTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] [{level}] {message}{Environment.NewLine}");
        logTextBox.SelectionStart = logTextBox.Text.Length;
        logTextBox.ScrollToCaret();
    }

    private string GetClipboardText()
    {
        try
        {
            if (Clipboard.ContainsText())
            {
                return Clipboard.GetText();
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error reading clipboard: {ex.Message}");
        }
        return string.Empty;
    }
}
