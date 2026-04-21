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
    private CheckBox useOpenWorkbookCheckBox = new();
    private ComboBox workbookComboBox = new();
    private ComboBox sheetComboBox = new();
    private Button refreshWorkbooksButton = new();
    private Button importButton = new();
    private Button cancelButton = new();
    private RichTextBox logTextBox = new();
    private Label workbookLabel = new();
    private Label sheetLabel = new();
    private Label openWorkbookHintLabel = new();
    private string? clipboardQuery;
    private bool _importRunning = false;

    private sealed class SheetSelectionItem
    {
        public string? SheetName { get; init; }
        public bool CreateNewSheet { get; init; }
        public string DisplayText { get; init; } = string.Empty;

        public override string ToString() => DisplayText;
    }

    public ExcelImportForm()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        this.Text = "Import SQL Query to Excel";
        this.Size = new Size(540, 470);
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

        useOpenWorkbookCheckBox.Text = "Usar un libro Excel ya abierto";
        useOpenWorkbookCheckBox.Location = new Point(20, 95);
        useOpenWorkbookCheckBox.Size = new Size(260, 25);
        useOpenWorkbookCheckBox.CheckedChanged += UseOpenWorkbookCheckBox_CheckedChanged;
        this.Controls.Add(useOpenWorkbookCheckBox);

        workbookLabel = new Label
        {
            Text = "Libro abierto:",
            Location = new Point(20, 130),
            Size = new Size(120, 25),
            AutoSize = false
        };
        this.Controls.Add(workbookLabel);

        workbookComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        workbookComboBox.Location = new Point(150, 130);
        workbookComboBox.Size = new Size(260, 25);
        workbookComboBox.SelectedIndexChanged += WorkbookComboBox_SelectedIndexChanged;
        this.Controls.Add(workbookComboBox);

        refreshWorkbooksButton.Text = "Refresh";
        refreshWorkbooksButton.Location = new Point(420, 130);
        refreshWorkbooksButton.Size = new Size(80, 25);
        refreshWorkbooksButton.Click += RefreshWorkbooksButton_Click;
        this.Controls.Add(refreshWorkbooksButton);

        sheetLabel = new Label
        {
            Text = "Hoja destino:",
            Location = new Point(20, 165),
            Size = new Size(120, 25),
            AutoSize = false
        };
        this.Controls.Add(sheetLabel);

        sheetComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        sheetComboBox.Location = new Point(150, 165);
        sheetComboBox.Size = new Size(350, 25);
        this.Controls.Add(sheetComboBox);

        openWorkbookHintLabel = new Label
        {
            Text = "Si eliges una hoja existente, la importación se añadirá debajo del contenido actual.",
            Location = new Point(150, 195),
            Size = new Size(350, 35),
            AutoSize = false
        };
        this.Controls.Add(openWorkbookHintLabel);

        // Log text area (replaces progress bar)
        logTextBox.Location = new Point(20, 240);
        logTextBox.Size = new Size(480, 150);
        logTextBox.ReadOnly = true;
        logTextBox.Multiline = true;
        logTextBox.ScrollBars = RichTextBoxScrollBars.Vertical;
        logTextBox.Visible = false;
        this.Controls.Add(logTextBox);

        // Import Button
        importButton.Text = "Import to Excel";
        importButton.Location = new Point(170, 405);
        importButton.Size = new Size(120, 30);
        importButton.Click += ImportButton_Click;
        this.Controls.Add(importButton);

        // Cancel Button
        cancelButton.Text = "Cancel";
        cancelButton.Location = new Point(300, 405);
        cancelButton.Size = new Size(100, 30);
        cancelButton.Click += (s, e) => this.Close();
        this.Controls.Add(cancelButton);

        UpdateOpenWorkbookControls();
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
        ExcelInteropManager.ImportTargetOptions? targetOptions = null;

        if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(database))
        {
            MessageBox.Show("Please enter server and database names.",
                "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (useOpenWorkbookCheckBox.Checked)
        {
            if (workbookComboBox.SelectedItem is not ExcelInteropManager.OpenWorkbookInfo selectedWorkbook)
            {
                MessageBox.Show("No se ha encontrado ningún libro de Excel abierto. Desmarca la opción o abre un libro antes de importar.",
                    "Libro no disponible", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (sheetComboBox.SelectedItem is not SheetSelectionItem selectedSheet)
            {
                MessageBox.Show("Selecciona una hoja destino o crea una nueva antes de importar.",
                    "Hoja no seleccionada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            targetOptions = new ExcelInteropManager.ImportTargetOptions
            {
                UseOpenWorkbook = true,
                WorkbookKey = selectedWorkbook.WorkbookKey,
                WorkbookName = selectedWorkbook.Name,
                WorksheetName = selectedSheet.SheetName,
                CreateNewWorksheet = selectedSheet.CreateNewSheet
            };
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
            if (targetOptions?.UseOpenWorkbook == true)
            {
                if (targetOptions.CreateNewWorksheet)
                {
                    Logger.LogInfo($"Targeting open workbook '{targetOptions.WorkbookName}' and creating a new worksheet.");
                }
                else
                {
                    Logger.LogInfo($"Targeting open workbook '{targetOptions.WorkbookName}', worksheet '{targetOptions.WorksheetName}'.");
                }
            }

            var tcs = new TaskCompletionSource<ImportResult>();
            var staThread = new Thread(() =>
            {
                try
                {
                    var res = ExcelInteropManager.InjectQueryIntoExcel(server, database, clipboardQuery, targetOptions);
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

    private void UseOpenWorkbookCheckBox_CheckedChanged(object? sender, EventArgs e)
    {
        UpdateOpenWorkbookControls();
        if (useOpenWorkbookCheckBox.Checked)
        {
            RefreshOpenWorkbookOptions();
        }
    }

    private void RefreshWorkbooksButton_Click(object? sender, EventArgs e)
    {
        RefreshOpenWorkbookOptions();
    }

    private void WorkbookComboBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        PopulateSheetOptions();
        UpdateOpenWorkbookControls();
    }

    private void RefreshOpenWorkbookOptions()
    {
        workbookComboBox.Items.Clear();
        sheetComboBox.Items.Clear();

        var openWorkbooks = ExcelInteropManager.GetOpenWorkbooks();
        foreach (var workbook in openWorkbooks)
        {
            workbookComboBox.Items.Add(workbook);
        }

        if (workbookComboBox.Items.Count > 0)
        {
            workbookComboBox.SelectedIndex = 0;
            openWorkbookHintLabel.Text = "Si eliges una hoja existente, la importación se añadirá debajo del contenido actual.";
        }
        else
        {
            openWorkbookHintLabel.Text = "No hay libros de Excel abiertos. Desmarca esta opción para usar el flujo normal.";
        }

        UpdateOpenWorkbookControls();
    }

    private void PopulateSheetOptions()
    {
        sheetComboBox.Items.Clear();

        if (workbookComboBox.SelectedItem is not ExcelInteropManager.OpenWorkbookInfo selectedWorkbook)
        {
            return;
        }

        sheetComboBox.Items.Add(new SheetSelectionItem
        {
            CreateNewSheet = true,
            DisplayText = "<Crear hoja nueva>"
        });

        foreach (string worksheetName in selectedWorkbook.WorksheetNames)
        {
            sheetComboBox.Items.Add(new SheetSelectionItem
            {
                SheetName = worksheetName,
                CreateNewSheet = false,
                DisplayText = worksheetName
            });
        }

        if (sheetComboBox.Items.Count > 0)
        {
            sheetComboBox.SelectedIndex = 0;
        }
    }

    private void UpdateOpenWorkbookControls()
    {
        bool useOpenWorkbook = useOpenWorkbookCheckBox.Checked;
        bool hasWorkbookSelection = useOpenWorkbook && workbookComboBox.Items.Count > 0;

        workbookLabel.Enabled = useOpenWorkbook;
        workbookComboBox.Enabled = hasWorkbookSelection;
        refreshWorkbooksButton.Enabled = useOpenWorkbook;
        sheetLabel.Enabled = useOpenWorkbook;
        sheetComboBox.Enabled = hasWorkbookSelection;
        openWorkbookHintLabel.Enabled = useOpenWorkbook;
    }
}
