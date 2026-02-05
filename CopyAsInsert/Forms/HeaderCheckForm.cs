namespace CopyAsInsert.Forms;

/// <summary>
/// Dialog to ask user if their data includes header row
/// </summary>
public partial class HeaderCheckForm : Form
{
    public bool HasHeaders { get; set; } = true;

    public HeaderCheckForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();

        var iconPath = Path.Combine(AppContext.BaseDirectory, "Group-3.ico");
        // Form properties
        this.Text = "Data Format";
        this.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
        this.Width = 350;
        this.Height = 150;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ShowIcon = true;
        this.TopMost = true;
        this.ControlBox = false;

        // Question Label
        var lblQuestion = new Label
        {
            Text = "Does your data include a header row?",
            Left = 20,
            Top = 20,
            Width = 300,
            Height = 40,
            AutoSize = false,
            Font = new Font("Segoe UI", 11, FontStyle.Regular)
        };

        // Yes Button
        var btnYes = new Button
        {
            Text = "Yes (First row is headers)",
            Left = 20,
            Top = 75,
            Width = 140,
            Height = 30,
            DialogResult = DialogResult.Yes
        };

        // No Button
        var btnNo = new Button
        {
            Text = "No (All rows are data)",
            Left = 170,
            Top = 75,
            Width = 140,
            Height = 30,
            DialogResult = DialogResult.No
        };

        this.Controls.Add(lblQuestion);
        this.Controls.Add(btnYes);
        this.Controls.Add(btnNo);

        this.FormClosing += (s, e) =>
        {
            if (this.DialogResult == DialogResult.Yes)
            {
                HasHeaders = true;
            }
            else if (this.DialogResult == DialogResult.No)
            {
                HasHeaders = false;
                e.Cancel = false;
            }
        };

        this.ResumeLayout(false);
    }
}
