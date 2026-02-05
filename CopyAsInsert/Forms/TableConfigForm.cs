namespace CopyAsInsert.Forms;

/// <summary>
/// Configuration dialog for table name, schema, and temporal table options
/// </summary>
public partial class TableConfigForm : Form
{
    public string TableName { get; set; } = string.Empty;
    public string SchemaName { get; set; }
    public bool IsTemporalTable { get; set; }
    public bool IsTemporaryTable { get; set; }

    public TableConfigForm()
    {
        InitializeComponent();
        SchemaName = "dbo";
        IsTemporalTable = false;
        IsTemporaryTable = false;
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();

        var iconPath = Path.Combine(AppContext.BaseDirectory, "Group-3.ico");
        // Form properties
        this.Text = "SQL Table Configuration";
        this.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
        this.Width = 420;
        this.Height = 310;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ShowIcon = true;

        // Table Name Label
        var lblTableName = new Label
        {
            Text = "Table Name:",
            Left = 20,
            Top = 20,
            Width = 100,
            Height = 20
        };

        // Table Name TextBox
        var txtTableName = new TextBox
        {
            Name = "txtTableName",
            Left = 130,
            Top = 20,
            Width = 240,
            Height = 20
        };

        // Schema Label
        var lblSchema = new Label
        {
            Text = "Schema:",
            Left = 20,
            Top = 60,
            Width = 100,
            Height = 20
        };

        // Schema TextBox
        var txtSchema = new TextBox
        {
            Name = "txtSchema",
            Left = 130,
            Top = 60,
            Width = 240,
            Height = 20,
            Text = "dbo"
        };

        // Temporal Table Checkbox
        var chkTemporal = new CheckBox
        {
            Name = "chkTemporal",
            Text = "Create Temporal Table with System Versioning",
            Left = 20,
            Top = 100,
            Width = 350,
            Height = 20,
            Checked = false
        };

        // Temporary Table Checkbox
        var chkTemporary = new CheckBox
        {
            Name = "chkTemporary",
            Text = "Create as Temporary Table (#)",
            Left = 20,
            Top = 130,
            Width = 350,
            Height = 20,
            Checked = false
        };

        // Generate Button
        var btnGenerate = new Button
        {
            Text = "Generate",
            Left = 210,
            Top = 220,
            Width = 80,
            Height = 30,
            DialogResult = DialogResult.OK
        };

        // Cancel Button
        var btnCancel = new Button
        {
            Text = "Cancel",
            Left = 300,
            Top = 220,
            Width = 80,
            Height = 30,
            DialogResult = DialogResult.Cancel
        };

        this.Controls.Add(lblTableName);
        this.Controls.Add(txtTableName);
        this.Controls.Add(lblSchema);
        this.Controls.Add(txtSchema);
        this.Controls.Add(chkTemporal);
        this.Controls.Add(chkTemporary);
        this.Controls.Add(btnGenerate);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnGenerate;
        this.CancelButton = btnCancel;

        this.FormClosing += (s, e) =>
        {
            if (this.DialogResult == DialogResult.OK)
            {
                TableName = txtTableName.Text.Trim();
                SchemaName = txtSchema.Text.Trim();
                IsTemporalTable = chkTemporal.Checked;
                IsTemporaryTable = chkTemporary.Checked;

                if (string.IsNullOrWhiteSpace(TableName))
                {
                    MessageBox.Show("Table name is required", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                }
            }
        };

        this.ResumeLayout(false);
    }
}
