using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace CopyAsInsert.Updater
{
    class Program
    {
        private static string? _logPath;

        static async Task Main(string[] args)
        {
            try
            {
                // Parse arguments: --version 1.0.9 --url https://... --app-path C:\path\
                var version = GetArgValue(args, "--version");
                var url = GetArgValue(args, "--url");
                var appPath = GetArgValue(args, "--app-path") ?? AppContext.BaseDirectory;

                SetupLogging(appPath);

                if (string.IsNullOrEmpty(version) || string.IsNullOrEmpty(url))
                {
                    Log("ERROR: Missing required arguments (--version, --url)");
                    return;
                }

                Log($"Starting update to version {version}");
                Log($"Download URL: {url}");
                Log($"App path: {appPath}");

                // Download the new executable
                var tempFile = Path.Combine(Path.GetTempPath(), $"CopyAsInsert-{version}.exe");
                if (!await DownloadFileAsync(url, tempFile))
                {
                    Log("ERROR: Failed to download update");
                    return;
                }

                Log($"Downloaded successfully to: {tempFile}");

                // Stage the new exe
                var stagedFile = Path.Combine(appPath, "CopyAsInsert.exe.new");
                try
                {
                    File.Copy(tempFile, stagedFile, overwrite: true);
                    Log($"Staged update to: {stagedFile}");

                    // Clean up temp file
                    try { File.Delete(tempFile); }
                    catch { /* Ignore */ }
                }
                catch (Exception ex)
                {
                    Log($"ERROR staging file: {ex.Message}");
                    return;
                }

                // Replace the running executable
                if (!ReplaceExecutable(appPath))
                {
                    Log("ERROR: Failed to replace executable");
                    return;
                }

                Log("Executable replaced successfully");

                // Launch the updated application
                var exePath = Path.Combine(appPath, "CopyAsInsert.exe");
                try
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = exePath,
                        UseShellExecute = true,
                        CreateNoWindow = false,
                        Arguments = "--update-restart"
                    };
                    Process.Start(startInfo);
                    Log("New process started with --update-restart flag");
                }
                catch (Exception ex)
                {
                    Log($"ERROR starting new process: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Log($"FATAL ERROR: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
            }
        }

        private static string? GetArgValue(string[] args, string argName)
        {
            for (int i = 0; i < args.Length - 1; i++)
            {
                if (args[i] == argName)
                    return args[i + 1];
            }
            return null;
        }

        private static async Task<bool> DownloadFileAsync(string url, string filePath)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "CopyAsInsert-Updater/1.0");
                    client.Timeout = TimeSpan.FromMinutes(10);

                    Log($"Downloading from: {url}");
                    var response = await client.GetAsync(url, HttpCompletionOption.ResponseContentRead);

                    if (!response.IsSuccessStatusCode)
                    {
                        Log($"HTTP Error {(int)response.StatusCode}: {response.StatusCode}");
                        return false;
                    }

                    var totalBytes = response.Content.Headers.ContentLength ?? -1;
                    Log($"File size: {totalBytes} bytes");

                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, useAsync: true))
                    {
                        var buffer = new byte[8192];
                        int bytesRead;
                        long totalRead = 0;

                        while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead);
                            totalRead += bytesRead;

                            if (totalBytes > 0)
                            {
                                var percent = (int)((totalRead * 100) / totalBytes);
                                Log($"Download progress: {percent}% ({totalRead}/{totalBytes} bytes)");
                            }
                        }
                    }

                    Log($"Download completed: {totalBytes} bytes written");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log($"Download error: {ex.Message}");
                return false;
            }
        }

        private static bool ReplaceExecutable(string appPath)
        {
            try
            {
                var exePath = Path.Combine(appPath, "CopyAsInsert.exe");
                var stagedPath = Path.Combine(appPath, "CopyAsInsert.exe.new");
                var backupPath = Path.Combine(appPath, "CopyAsInsert.exe.backup");

                if (!File.Exists(stagedPath))
                {
                    Log($"Staged file not found: {stagedPath}");
                    return false;
                }

                // Wait for file locks to release (main process exiting)
                int attempts = 0;
                while (attempts < 10)
                {
                    try
                    {
                        // Try to backup current exe
                        if (File.Exists(exePath))
                        {
                            if (File.Exists(backupPath))
                                File.Delete(backupPath);
                            File.Move(exePath, backupPath, overwrite: false);
                        }

                        // Move new exe to main location
                        File.Move(stagedPath, exePath, overwrite: false);
                        Log("File replacement successful");
                        return true;
                    }
                    catch (IOException ex)
                    {
                        attempts++;
                        if (attempts >= 10)
                        {
                            Log($"ERROR: Could not replace file after {attempts} attempts: {ex.Message}");
                            // Try to restore backup
                            if (File.Exists(backupPath))
                            {
                                try
                                {
                                    File.Move(backupPath, exePath, overwrite: true);
                                    Log("Restored backup after failed replacement");
                                }
                                catch { /* Ignore */ }
                            }
                            return false;
                        }
                        Log($"Retry {attempts}/10: File still locked");
                        System.Threading.Thread.Sleep(500);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Log($"ERROR in ReplaceExecutable: {ex.Message}");
                return false;
            }
        }

        private static void SetupLogging(string appPath)
        {
            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "CopyAsInsert"
                );
                Directory.CreateDirectory(logDir);
                _logPath = Path.Combine(logDir, "UpdaterLog.txt");
            }
            catch
            {
                _logPath = null;
            }
        }

        private static void Log(string message)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var logMessage = $"[{timestamp}] {message}";
            Console.WriteLine(logMessage);

            if (_logPath != null)
            {
                try
                {
                    File.AppendAllText(_logPath, logMessage + Environment.NewLine);
                }
                catch { /* Ignore logging errors */ }
            }
        }
    }
}
