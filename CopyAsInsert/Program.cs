namespace CopyAsInsert;

using System.Threading;

static class Program
{
    private static Mutex? _instanceMutex;
    private const string MutexName = "ZARPA_SingleInstance";

    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main(string[] args)
    {
        // Check if this is an update restart (new instance started by updater)
        bool isUpdateRestart = args.Contains("--update-restart");

        // Check if another instance is already running (skip check for update restarts)
        _instanceMutex = new Mutex(false, MutexName);
        
        if (!isUpdateRestart && !_instanceMutex.WaitOne(0))
        {
            // Another instance is already running
            MessageBox.Show(
                "Ya hay una instancia de ZARPA abierta.",
                "Aplicación en ejecución",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
            return;
        }

        // Acquire the mutex if this is an update restart
        if (isUpdateRestart)
        {
            _instanceMutex.WaitOne();
        }

        try
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
        finally
        {
            // Release the mutex when the application exits
            _instanceMutex?.ReleaseMutex();
            _instanceMutex?.Dispose();
        }
    }
}