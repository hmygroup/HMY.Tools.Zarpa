using CopyAsInsert.Services;

namespace CopyAsInsert.Forms;

public class ExcelImportForm : Form
{
    private TextBox serverTextBox = new();
    private TextBox databaseTextBox = new();
    private Button testConnectionButton = new();
    private Label connectionStatusLabel = new();
    private Button importButton = new();
    private Button cancelButton = new();
    private ProgressBar progressBar = new();
    private string? clipboardQuery;
    private bool connectionTested = false;

    public ExcelImportForm()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        this.Text = "Import SQL Query to Excel";
        this.Size = new Size(500, 300);
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

        // Test Connection Button
        testConnectionButton.Text = "Test Connection";
        testConnectionButton.Location = new Point(150, 90);
        testConnectionButton.Size = new Size(120, 30);
        testConnectionButton.Click += TestConnectionButton_Click;
        this.Controls.Add(testConnectionButton);

        // Connection Status Label
        connectionStatusLabel.Location = new Point(280, 95);
        connectionStatusLabel.Size = new Size(170, 20);
        connectionStatusLabel.ForeColor = Color.Red;
        connectionStatusLabel.Text = "Not tested";
        this.Controls.Add(connectionStatusLabel);

        // Progress Bar
        progressBar.Location = new Point(20, 130);
        progressBar.Size = new Size(430, 20);
        progressBar.Visible = false;
        this.Controls.Add(progressBar);

        // Import Button
        importButton.Text = "Import to Excel";
        importButton.Location = new Point(150, 160);
        importButton.Size = new Size(120, 30);
        importButton.Enabled = false;
        importButton.Click += ImportButton_Click;
        this.Controls.Add(importButton);

        // Cancel Button
        cancelButton.Text = "Cancel";
        cancelButton.Location = new Point(280, 160);
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

    private void TestConnectionButton_Click(object? sender, EventArgs e)
    {
        string server = serverTextBox.Text.Trim();
        string database = databaseTextBox.Text.Trim();

        if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(database))
        {
            connectionStatusLabel.Text = "Enter server and database";
            connectionStatusLabel.ForeColor = Color.Red;
            connectionTested = false;
            importButton.Enabled = false;
            return;
        }

        testConnectionButton.Enabled = false;
        progressBar.Visible = true;
        progressBar.Style = ProgressBarStyle.Marquee;

        try
        {
            var (success, errorMessage) = ExcelInteropManager.TestConnection(server, database);
            
            if (success)
            {
                connectionStatusLabel.Text = "✓ Connection OK";
                connectionStatusLabel.ForeColor = Color.Green;
                connectionTested = true;
                importButton.Enabled = true;
            }
            else
            {
                connectionStatusLabel.Text = "✗ Connection Failed";
                connectionStatusLabel.ForeColor = Color.Red;
                connectionTested = false;
                importButton.Enabled = false;
                
                MessageBox.Show($"Connection test failed:\n{errorMessage}", 
                    "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        catch (Exception ex)
        {
            connectionStatusLabel.Text = "✗ Error";
            connectionStatusLabel.ForeColor = Color.Red;
            connectionTested = false;
            importButton.Enabled = false;
            
            MessageBox.Show($"Error testing connection:\n{ex.Message}", 
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            testConnectionButton.Enabled = true;
            progressBar.Visible = false;
        }
    }

    private void ImportButton_Click(object? sender, EventArgs e)
    {
        if (!connectionTested)
        {
            MessageBox.Show("Please test connection first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

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

        importButton.Enabled = false;
        testConnectionButton.Enabled = false;
        progressBar.Visible = true;
        progressBar.Style = ProgressBarStyle.Marquee;

        try
        {
            // Import query to Excel
            var result = ExcelInteropManager.InjectQueryIntoExcel(server, database, clipboardQuery);
            
            if (result.Success)
            {
                SaveSettings();
                
                Logger.LogDebug($"Query imported successfully: {result.RowCount} rows");
                
                // Close this form - Excel now has the data
                this.Close();
            }
            else
            {
                string errorMsg = result.ErrorMessage ?? "Unknown error";
                string stackTrace = result.ErrorStackTrace ?? "";
                
                Logger.LogError($"Import failed: {errorMsg}\n{stackTrace}");
                
                // Show error detail form
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
            importButton.Enabled = true;
            testConnectionButton.Enabled = true;
            progressBar.Visible = false;
        }
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
