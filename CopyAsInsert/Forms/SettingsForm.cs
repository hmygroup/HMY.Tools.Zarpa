namespace CopyAsInsert.Forms;

/// <summary>
/// Settings dialog for application configuration
/// </summary>
public partial class SettingsForm : Form
{
    public string DefaultSchema { get; set; }
    public bool AutoCreateHistoryTable { get; set; }
    public bool TemporalTableByDefault { get; set; }
    public bool RunOnStartup { get; set; }
    public int HotKeyModifier { get; set; }
    public int HotKeyVirtualKey { get; set; }
    public bool AutoAppendTemporalSuffix { get; set; }
    public bool ShowFormOnStartup { get; set; }

    private int _pendingModifier; // Stores modifier during key capture
    private int _pendingVirtualKey; // Stores virtual key during capture
    
    // Control references
    private TextBox? _txtSchema;
    private TextBox? _txtHotkey; // Read-only display of current hotkey
    private TextBox? _txtHotkeyCapture; // Captures hotkey input
    private CheckBox? _chkAutoHistory;
    private CheckBox? _chkTemporal;
    private CheckBox? _chkRunOnStartup;
    private CheckBox? _chkAutoAppendTemporalSuffix;
    private CheckBox? _chkShowFormOnStartup;

    public SettingsForm()
    {
        InitializeComponent();
        DefaultSchema = "dbo";
        AutoCreateHistoryTable = true;
        TemporalTableByDefault = true;
        RunOnStartup = false;
        HotKeyModifier = 0x0001 | 0x0004; // MOD_ALT | MOD_SHIFT
        HotKeyVirtualKey = 0x49; // 'I'
        AutoAppendTemporalSuffix = false;
        ShowFormOnStartup = false;
        _pendingModifier = HotKeyModifier;
        _pendingVirtualKey = HotKeyVirtualKey;
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        
        // CRITICAL: Populate ALL controls with settings AFTER MainForm has set the properties
        // This ensures we display loaded settings, not defaults
        
        // Schema
        if (_txtSchema != null)
        {
            _txtSchema.Text = DefaultSchema;
        }
        
        // Hotkey - sync pending values with public properties set by MainForm
        _pendingModifier = HotKeyModifier;
        _pendingVirtualKey = HotKeyVirtualKey;
        UpdateHotkeyDisplay();
        
        // Checkboxes
        if (_chkAutoHistory != null)
            _chkAutoHistory.Checked = AutoCreateHistoryTable;
        if (_chkTemporal != null)
            _chkTemporal.Checked = TemporalTableByDefault;
        if (_chkRunOnStartup != null)
            _chkRunOnStartup.Checked = RunOnStartup;
        if (_chkAutoAppendTemporalSuffix != null)
            _chkAutoAppendTemporalSuffix.Checked = AutoAppendTemporalSuffix;
        if (_chkShowFormOnStartup != null)
            _chkShowFormOnStartup.Checked = ShowFormOnStartup;
        
        // Set focus to hotkey capture textbox so KeyDown events will fire
        if (_txtHotkeyCapture != null)
        {
            _txtHotkeyCapture.Focus();
            _txtHotkeyCapture.SelectAll();
        }
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();

        var iconPath = Path.Combine(AppContext.BaseDirectory, "Group-3.ico");
        // Form properties
        this.Text = "Settings";
        this.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
        this.Width = 530;
        this.Height = 560;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ShowIcon = true;

        // ============ Schema Section ============
        var lblSchema = new Label
        {
            Text = "Default Schema:",
            Left = 20,
            Top = 20,
            Width = 150,
            Height = 20,
            Font = new Font(this.Font, FontStyle.Bold)
        };

        _txtSchema = new TextBox
        {
            Name = "txtSchema",
            Left = 180,
            Top = 20,
            Width = 320,
            Height = 20,
            Text = ""
        };

        // ============ Hotkey Section ============
        var lblHotkey = new Label
        {
            Text = "Custom Hotkey:",
            Left = 20,
            Top = 70,
            Width = 150,
            Height = 20,
            Font = new Font(this.Font, FontStyle.Bold)
        };

        _txtHotkey = new TextBox
        {
            Name = "txtHotkey",
            Left = 180,
            Top = 70,
            Width = 320,
            Height = 20,
            ReadOnly = true,
            BackColor = SystemColors.ControlLight,
            ForeColor = SystemColors.WindowText,
            TabStop = false,
            Text = ""
        };

        var lblHotkeyInfo = new Label
        {
            Text = "Click the field below and press your desired key combination:",
            Left = 20,
            Top = 100,
            Width = 480,
            Height = 20,
            ForeColor = SystemColors.GrayText,
            Font = new Font(this.Font, FontStyle.Italic)
        };

        _txtHotkeyCapture = new TextBox
        {
            Name = "txtHotkeyCapture",
            Left = 180,
            Top = 130,
            Width = 320,
            Height = 20,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Text = "Click here and press hotkey...",
            ForeColor = SystemColors.GrayText
        };

        // ============ Options Section ============
        var lblOptions = new Label
        {
            Text = "Options:",
            Left = 20,
            Top = 175,
            Width = 150,
            Height = 20,
            Font = new Font(this.Font, FontStyle.Bold)
        };

        _chkAutoHistory = new CheckBox
        {
            Name = "chkAutoHistory",
            Text = "Auto-create history table",
            Left = 20,
            Top = 205,
            Width = 480,
            Height = 20,
            Checked = false
        };

        _chkTemporal = new CheckBox
        {
            Name = "chkTemporal",
            Text = "Create temporal tables by default",
            Left = 20,
            Top = 235,
            Width = 480,
            Height = 20,
            Checked = false
        };

        _chkRunOnStartup = new CheckBox
        {
            Name = "chkRunOnStartup",
            Text = "Run on Windows startup",
            Left = 20,
            Top = 265,
            Width = 480,
            Height = 20,
            Checked = false
        };

        // ============ Buttons ============
        var btnOK = new Button
        {
            Text = "OK",
            Left = 340,
            Top = 380,
            Width = 80,
            Height = 30,
            DialogResult = DialogResult.OK
        };

        var btnCancel = new Button
        {
            Text = "Cancel",
            Left = 430,
            Top = 380,
            Width = 80,
            Height = 30,
            DialogResult = DialogResult.Cancel
        };

        this.Controls.Add(lblSchema);
        this.Controls.Add(_txtSchema);
        
        this.Controls.Add(lblHotkey);
        this.Controls.Add(_txtHotkey);
        this.Controls.Add(lblHotkeyInfo);
        this.Controls.Add(_txtHotkeyCapture);
        
        this.Controls.Add(lblOptions);
        this.Controls.Add(_chkAutoHistory);
        this.Controls.Add(_chkTemporal);
        this.Controls.Add(_chkRunOnStartup);
        this.Controls.Add(_chkAutoAppendTemporalSuffix);
        this.Controls.Add(_chkShowFormOnStartup);
        
        this.Controls.Add(btnOK);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;

        // ============ Event Handlers ============
        
        // Hotkey capture event handler - fires when user presses a key in the capture field
        _txtHotkeyCapture.KeyDown += (s, e) =>
        {
            e.Handled = true;
            e.SuppressKeyPress = true;

            // Clear placeholder text on first key press
            if (_txtHotkeyCapture.Text == "Click here and press hotkey...")
            {
                _txtHotkeyCapture.Text = "";
            }

            // Extract modifiers
            var modifiers = 0;

            if ((e.Modifiers & Keys.Control) == Keys.Control)
                modifiers |= 0x0002; // MOD_CTRL

            if ((e.Modifiers & Keys.Alt) == Keys.Alt)
                modifiers |= 0x0001; // MOD_ALT

            if ((e.Modifiers & Keys.Shift) == Keys.Shift)
                modifiers |= 0x0004; // MOD_SHIFT

            // Get the virtual key
            int vKey = (int)e.KeyCode;

            // Validate hotkey - reject modifier keys only
            var reservedKeys = new[] { Keys.LControlKey, Keys.RControlKey, Keys.LShiftKey, Keys.RShiftKey, 
                                       Keys.LMenu, Keys.RMenu, Keys.Control, Keys.Shift, Keys.Alt };

            if (reservedKeys.Contains(e.KeyCode))
            {
                MessageBox.Show("Modifier keys alone are not allowed.\nPlease press a regular key (A-Z, 0-9, F1-F12, etc.) along with modifiers.", 
                    "Invalid Hotkey", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtHotkeyCapture.Text = "Click here and press hotkey...";
                return;
            }

            if (modifiers == 0)
            {
                MessageBox.Show("Please use at least one modifier key:\nHold Ctrl, Alt, or Shift while pressing your key.", 
                    "Invalid Hotkey", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtHotkeyCapture.Text = "Click here and press hotkey...";
                return;
            }

            // Store the new hotkey
            _pendingModifier = modifiers;
            _pendingVirtualKey = vKey;

            // Update the display
            UpdateHotkeyDisplay();
            _txtHotkeyCapture.Text = $"✓ Captured: {FormatHotkey(_pendingModifier, _pendingVirtualKey)}";
        };

        this.FormClosing += (s, e) =>
        {
            if (this.DialogResult == DialogResult.OK)
            {
                // Save all form values back to properties
                DefaultSchema = _txtSchema?.Text.Trim() ?? "dbo";
                AutoCreateHistoryTable = _chkAutoHistory?.Checked ?? true;
                TemporalTableByDefault = _chkTemporal?.Checked ?? true;
                RunOnStartup = _chkRunOnStartup?.Checked ?? false;
                AutoAppendTemporalSuffix = _chkAutoAppendTemporalSuffix?.Checked ?? false;
                ShowFormOnStartup = _chkShowFormOnStartup?.Checked ?? false;
                HotKeyModifier = _pendingModifier;
                HotKeyVirtualKey = _pendingVirtualKey;
            }
        };

        this.ResumeLayout(false);
    }

    private void UpdateHotkeyDisplay()
    {
        if (_txtHotkey != null)
        {
            _txtHotkey.Text = FormatHotkey(_pendingModifier, _pendingVirtualKey);
        }
    }

    private string FormatHotkey(int modifiers, int vKey)
    {
        var keys = new List<string>();

        if ((modifiers & 0x0002) != 0) // MOD_CTRL
            keys.Add("Ctrl");
        if ((modifiers & 0x0001) != 0) // MOD_ALT
            keys.Add("Alt");
        if ((modifiers & 0x0004) != 0) // MOD_SHIFT
            keys.Add("Shift");

        // Convert virtual key to character (for printable characters)
        if (vKey >= 0x41 && vKey <= 0x5A) // A-Z
        {
            keys.Add(((char)vKey).ToString());
        }
        else if (vKey >= 0x30 && vKey <= 0x39) // 0-9
        {
            keys.Add(((char)vKey).ToString());
        }
        else if ((Keys)vKey == Keys.F1)
            keys.Add("F1");
        else if ((Keys)vKey == Keys.F12)
            keys.Add("F12");
        else
        {
            keys.Add($"0x{vKey:X}");
        }

        return string.Join("+", keys);
    }
}
