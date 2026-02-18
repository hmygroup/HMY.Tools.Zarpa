using System.Runtime.InteropServices;

namespace CopyAsInsert.Services;

/// <summary>
/// Handles global hotkey registration and clipboard monitoring using Windows API
/// </summary>
public class ClipboardInterceptor : IDisposable
{
    public const int MOD_ALT = 0x0001;
    public const int MOD_SHIFT = 0x0004;
    public const int MOD_CTRL = 0x0002;
    public const int WM_HOTKEY = 0x0312;
    public const int HOTKEY_ID = 9000;

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    private IntPtr _windowHandle;
    private int _currentModifier;
    private int _currentVirtualKey;
    private bool _isRegistered = false;

    /// <summary>
    /// Event fired when hotkey is pressed
    /// </summary>
    public event EventHandler<EventArgs>? HotKeyPressed;

    /// <summary>
    /// Initialize hotkey listener for the given window with default Alt+Shift+I
    /// </summary>
    public void InitializeHotKey(IntPtr windowHandle)
    {
        InitializeHotKey(windowHandle, MOD_ALT | MOD_SHIFT, 0x49);
    }

    /// <summary>
    /// Initialize hotkey listener with custom modifier and virtual key
    /// </summary>
    public void InitializeHotKey(IntPtr windowHandle, int modifiers, int virtualKey)
    {
        _windowHandle = windowHandle;
        _currentModifier = modifiers;
        _currentVirtualKey = virtualKey;

        _isRegistered = RegisterHotKey(windowHandle, HOTKEY_ID, modifiers, virtualKey);

        if (!_isRegistered)
        {
            throw new InvalidOperationException($"Failed to register hotkey. It may be in use by another application (Modifiers: 0x{modifiers:X}, VKey: 0x{virtualKey:X}).");
        }
    }

    /// <summary>
    /// Update hotkey to new configuration, unregistering the old one first
    /// </summary>
    public bool UpdateHotKey(int modifiers, int virtualKey)
    {
        // Unregister old hotkey if it was registered
        if (_isRegistered && _windowHandle != IntPtr.Zero)
        {
            UnregisterHotKey(_windowHandle, HOTKEY_ID);
            _isRegistered = false;
        }

        // Try to register with new hotkey
        if (_windowHandle != IntPtr.Zero)
        {
            _isRegistered = RegisterHotKey(_windowHandle, HOTKEY_ID, modifiers, virtualKey);

            if (_isRegistered)
            {
                _currentModifier = modifiers;
                _currentVirtualKey = virtualKey;
                return true;
            }
            else
            {
                // If registration failed, try to restore the old hotkey
                if (_currentModifier != modifiers || _currentVirtualKey != virtualKey)
                {
                    RegisterHotKey(_windowHandle, HOTKEY_ID, _currentModifier, _currentVirtualKey);
                    _isRegistered = true;
                }
                return false;
            }
        }

        return false;
    }

    /// <summary>
    /// Get current hotkey configuration as formatted string (e.g., "Ctrl+Alt+P")
    /// </summary>
    public string GetCurrentHotKeyString()
    {
        var keys = new List<string>();
        
        if ((_currentModifier & MOD_CTRL) != 0)
            keys.Add("Ctrl");
        if ((_currentModifier & MOD_ALT) != 0)
            keys.Add("Alt");
        if ((_currentModifier & MOD_SHIFT) != 0)
            keys.Add("Shift");

        // Convert virtual key to character
        string keyChar = ((char)_currentVirtualKey).ToString().ToUpper();
        keys.Add(keyChar);

        return string.Join("+", keys);
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
    /// Check if clipboard contains text data (any non-empty text is valid)
    /// </summary>
    public static bool IsClipboardTabularData()
    {
        try
        {
            var text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
                return false;

            // Accept any non-empty text (single values, multiple rows, etc.)
            return true;
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
