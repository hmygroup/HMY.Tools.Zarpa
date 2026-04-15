namespace CopyAsInsert.Forms;

public class ErrorDetailForm : Form
{
    private TextBox errorTextBox = new();
    private Button copyButton = new();
    private Button closeButton = new();

    public ErrorDetailForm(string errorMessage, string? stackTrace = null)
    {
        InitializeComponent();
        
        string fullError = errorMessage;
        if (!string.IsNullOrEmpty(stackTrace))
        {
            fullError += $"\n\n--- Stack Trace ---\n{stackTrace}";
        }
        
        errorTextBox.Text = fullError;
    }

    private void InitializeComponent()
    {
        this.Text = "Error Details";
        this.Size = new Size(600, 400);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // Error TextBox
        errorTextBox.Location = new Point(10, 10);
        errorTextBox.Size = new Size(570, 330);
        errorTextBox.Multiline = true;
        errorTextBox.ReadOnly = true;
        errorTextBox.ScrollBars = ScrollBars.Vertical;
        errorTextBox.Font = new Font("Courier New", 9);
        this.Controls.Add(errorTextBox);

        // Copy Button
        copyButton.Text = "Copy to Clipboard";
        copyButton.Location = new Point(250, 350);
        copyButton.Size = new Size(130, 30);
        copyButton.Click += CopyButton_Click;
        this.Controls.Add(copyButton);

        // Close Button
        closeButton.Text = "Close";
        closeButton.Location = new Point(390, 350);
        closeButton.Size = new Size(100, 30);
        closeButton.Click += (s, e) => this.Close();
        this.Controls.Add(closeButton);
    }

    private void CopyButton_Click(object? sender, EventArgs e)
    {
        try
        {
            Clipboard.SetText(errorTextBox.Text);
            MessageBox.Show("Error details copied to clipboard", "Success", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to copy to clipboard: {ex.Message}", 
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
