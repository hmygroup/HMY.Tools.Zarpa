namespace CopyAsInsert;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        
        var mainForm = new MainForm();
        mainForm.Load += (s, e) =>
        {
            // Hide after load so it's ready to receive hotkey messages
            mainForm.Hide();
        };
        Application.Run(mainForm);
    }    
}