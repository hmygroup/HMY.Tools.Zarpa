using System.Runtime.InteropServices;

namespace CopyAsInsert.Services;

/// <summary>
/// Handles global hotkey registration and clipboard monitoring using Windows API
/// </summary>
public class ClipboardInterceptor : IDisposable
{
    private const int MOD_ALT = 0x0001;
    private const int MOD_SHIFT = 0x0004;
    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID = 9000;

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    private IntPtr _windowHandle;
    private bool _isRegistered = false;

    /// <summary>
    /// Event fired when Alt+Shift+I hotkey is pressed
    /// </summary>
    public event EventHandler<EventArgs>? HotKeyPressed;

    /// <summary>
    /// Initialize hotkey listener for the given window
    /// </summary>
    public void InitializeHotKey(IntPtr windowHandle)
    {
        _windowHandle = windowHandle;

        // Alt+Shift+I: I = 0x49
        _isRegistered = RegisterHotKey(windowHandle, HOTKEY_ID, MOD_ALT | MOD_SHIFT, 0x49);

        if (!_isRegistered)
        {
            throw new InvalidOperationException("Failed to register hotkey Alt+Shift+I. It may be in use by another application.");
        }
    }

    /// <summary>
    /// Call this from your form's WndProc to handle hotkey messages
    /// </summary>
    public void ProcessWindowMessage(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
        {
            HotKeyPressed?.Invoke(this, EventArgs.Empty);
        }
    }

    /// <summary>
    /// Get current clipboard content as text
    /// </summary>
    public static string? GetClipboardText()
    {
        try
        {
            return Clipboard.GetText();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Check if clipboard contains tabular data (TSV or CSV-like format)
    /// </summary>
    public static bool IsClipboardTabularData()
    {
        try
        {
            var text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
                return false;

            // Split by various newline formats
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2) // At least header + 1 data row
                return false;

            var firstLine = lines[0];
            // Check for delimiters (tab preferred over comma)
            bool hasDelimiter = firstLine.Contains('\t') || firstLine.Contains(',');
            
            return hasDelimiter;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set clipboard content to text
    /// </summary>
    public static void SetClipboardText(string text)
    {
        try
        {
            Clipboard.SetText(text);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set clipboard content", ex);
        }
    }

    /// <summary>
    /// Get clipboard as IDataObject (for file drops)
    /// </summary>
    public static IDataObject? GetClipboardData()
    {
        try
        {
            return Clipboard.GetDataObject();
        }
        catch
        {
            return null;
        }
    }

    public void Dispose()
    {
        if (_isRegistered && _windowHandle != IntPtr.Zero)
        {
            UnregisterHotKey(_windowHandle, HOTKEY_ID);
            _isRegistered = false;
        }
    }
}
