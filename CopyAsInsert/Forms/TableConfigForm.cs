namespace CopyAsInsert.Forms;

using CopyAsInsert.Models;
using CopyAsInsert.Services;
using System.Runtime.InteropServices;

/// <summary>
/// Configuration dialog for table name, schema, and column type overrides
/// All configuration on a single form with type grid below settings
/// </summary>
public partial class TableConfigForm : Form
{
    [DllImport("user32.dll")]
    private static extern IntPtr SetForegroundWindow(IntPtr hWnd);

    public string TableName { get; set; } = string.Empty;
    public string SchemaName { get; set; }
    public bool IsTemporaryTable { get; set; }
    public DataTableSchema? Schema { get; set; }

    private TypeOverrideControl? _typeOverrideControl;
    private string _defaultSchema = "dbo";
    private TextBox? _txtTableName;

    public TableConfigForm()
    {
        InitializeComponent();
        SchemaName = _defaultSchema;
        IsTemporaryTable = true;
    }

    public TableConfigForm(string defaultSchema)
    {
        _defaultSchema = defaultSchema;
        InitializeComponent();
        SchemaName = _defaultSchema;
        IsTemporaryTable = true;
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        SetForegroundWindow(this.Handle);

        // Set focus to table name field with text pre-selected
        if (_txtTableName != null)
        {
            _txtTableName.Focus();
            _txtTableName.SelectAll();
        }
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();

        var iconPath = Path.Combine(AppContext.BaseDirectory, "Group-3.ico");
        // Form properties
        this.Text = "SQL Table Configuration";
        this.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
        this.Width = 1000;
        this.Height = 650;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ShowIcon = true;
        this.TopMost = true;

        // ============ Top Configuration Panel ============
        var configPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 150,
            Padding = new Padding(20)
        };

        // Table Name Label
        var lblTableName = new Label
        {
            Text = "Table Name:",
            Location = new Point(20, 15),
            Width = 100,
            Height = 20
        };

        // Table Name TextBox
        var txtTableName = new TextBox
        {
            Name = "txtTableName",
            Location = new Point(130, 15),
            Width = 240,
            Height = 20
        };
        _txtTableName = txtTableName;

        // Schema Label
        var lblSchema = new Label
        {
            Text = "Schema:",
            Location = new Point(20, 50),
            Width = 100,
            Height = 20
        };

        // Schema TextBox
        var txtSchema = new TextBox
        {
            Name = "txtSchema",
            Location = new Point(130, 50),
            Width = 240,
            Height = 20,
            Text = _defaultSchema
        };

        // Temporary Table Checkbox
        var chkTemporary = new CheckBox
        {
            Name = "chkTemporary",
            Text = "Create as Temporary Table (#)",
            Location = new Point(20, 85),
            Width = 350,
            Height = 20,
            Checked = true
        };

        configPanel.Controls.Add(lblTableName);
        configPanel.Controls.Add(txtTableName);
        configPanel.Controls.Add(lblSchema);
        configPanel.Controls.Add(txtSchema);
        configPanel.Controls.Add(chkTemporary);

        // ============ Type Override Control ============
        _typeOverrideControl = new TypeOverrideControl
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(10)
        };

        // ============ Column Types Info Label ============
        var typeInfoPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            Padding = new Padding(20, 10, 20, 10),
            BackColor = SystemColors.Control
        };

        var lblTypeInfo = new Label
        {
            Text = "Review and adjust inferred column types. Columns with <85% confidence are highlighted.",
            Dock = DockStyle.Fill,
            AutoSize = true,
            TextAlign = ContentAlignment.MiddleLeft
        };

        typeInfoPanel.Controls.Add(lblTypeInfo);

        // ============ Buttons ============
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            Padding = new Padding(10),
            BackColor = SystemColors.Control
        };

        var btnAutoDetect = new Button
        {
            Text = "Auto-Detect Types",
            Location = new Point(10, 10),
            Width = 120,
            Height = 30
        };

        btnAutoDetect.Click += (s, e) =>
        {
            if (_typeOverrideControl != null && Schema != null)
            {
                TypeInferenceEngine.InferColumnTypes(Schema);
                _typeOverrideControl.LoadSchema(Schema);
                MessageBox.Show("Type inference completed. Review the results below.", "Auto-Detect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        };

        var btnSetAllNvarchar = new Button
        {
            Text = "Set All to NVARCHAR",
            Location = new Point(140, 10),
            Width = 140,
            Height = 30
        };

        btnSetAllNvarchar.Click += (s, e) =>
        {
            if (_typeOverrideControl != null)
            {
                _typeOverrideControl.SetAllColumnsToNvarchar();
                MessageBox.Show("All columns set to NVARCHAR type.", "Set NVARCHAR", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        };

        var btnGenerate = new Button
        {
            Text = "Generate",
            Location = new Point(780, 10),
            Width = 100,
            Height = 30,
            DialogResult = DialogResult.OK
        };

        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(890, 10),
            Width = 100,
            Height = 30,
            DialogResult = DialogResult.Cancel
        };

        buttonPanel.Controls.Add(btnAutoDetect);
        buttonPanel.Controls.Add(btnSetAllNvarchar);
        buttonPanel.Controls.Add(btnGenerate);
        buttonPanel.Controls.Add(btnCancel);

        // ============ Main Layout ============
        this.Controls.Add(_typeOverrideControl);      // Fill middle
        this.Controls.Add(typeInfoPanel);             // Below config
        this.Controls.Add(buttonPanel);               // Bottom
        this.Controls.Add(configPanel);               // Top

        this.AcceptButton = btnGenerate;
        this.CancelButton = btnCancel;

        this.FormClosing += (s, e) =>
        {
            if (this.DialogResult == DialogResult.OK)
            {
                TableName = txtTableName.Text.Trim();
                SchemaName = txtSchema.Text.Trim();
                IsTemporaryTable = chkTemporary.Checked;

                // Get modified schema from type override control
                if (_typeOverrideControl != null)
                {
                    Schema = _typeOverrideControl.GetModifiedSchema();
                }

                if (string.IsNullOrWhiteSpace(TableName))
                {
                    MessageBox.Show("Table name is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                }
            }
        };

        this.ResumeLayout(false);
    }

    /// <summary>
    /// Set the schema to display in the type override control
    /// </summary>
    public void SetSchema(DataTableSchema schema)
    {
        Schema = schema;
        if (_typeOverrideControl != null)
        {
            _typeOverrideControl.LoadSchema(schema);
        }
    }
}

