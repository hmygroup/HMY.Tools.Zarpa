namespace CopyAsInsert.Forms;

/// <summary>
/// Settings dialog for application configuration
/// </summary>
public partial class SettingsForm : Form
{
    public string DefaultSchema { get; set; }
    public bool AutoCreateHistoryTable { get; set; }
    public bool TemporalTableByDefault { get; set; }

    public SettingsForm()
    {
        InitializeComponent();
        DefaultSchema = "dbo";
        AutoCreateHistoryTable = true;
        TemporalTableByDefault = true;
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();

        // Form properties
        this.Text = "Settings";
        this.Width = 400;
        this.Height = 300;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ShowIcon = false;

        // Default Schema Label
        var lblSchema = new Label
        {
            Text = "Default Schema:",
            Left = 20,
            Top = 20,
            Width = 150,
            Height = 20
        };

        // Default Schema TextBox
        var txtSchema = new TextBox
        {
            Name = "txtSchema",
            Left = 170,
            Top = 20,
            Width = 200,
            Height = 20,
            Text = DefaultSchema
        };

        // Auto-create History Table Checkbox
        var chkAutoHistory = new CheckBox
        {
            Name = "chkAutoHistory",
            Text = "Auto-create history table",
            Left = 20,
            Top = 70,
            Width = 350,
            Height = 20,
            Checked = AutoCreateHistoryTable
        };

        // Temporal Table by Default Checkbox
        var chkTemporal = new CheckBox
        {
            Name = "chkTemporal",
            Text = "Create temporal tables by default",
            Left = 20,
            Top = 110,
            Width = 350,
            Height = 20,
            Checked = TemporalTableByDefault
        };

        // OK Button
        var btnOK = new Button
        {
            Text = "OK",
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

        this.Controls.Add(lblSchema);
        this.Controls.Add(txtSchema);
        this.Controls.Add(chkAutoHistory);
        this.Controls.Add(chkTemporal);
        this.Controls.Add(btnOK);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;

        this.FormClosing += (s, e) =>
        {
            if (this.DialogResult == DialogResult.OK)
            {
                DefaultSchema = txtSchema.Text.Trim();
                AutoCreateHistoryTable = chkAutoHistory.Checked;
                TemporalTableByDefault = chkTemporal.Checked;
            }
        };

        this.ResumeLayout(false);
    }
}
