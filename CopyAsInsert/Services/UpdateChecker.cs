using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;

namespace CopyAsInsert.Services
{
    /// <summary>
    /// Checks for and manages application updates
    /// </summary>
    public class UpdateChecker
    {
        private readonly string _currentVersion;
        private readonly string _updateCheckUrl;
        private readonly string _installationPath;
        private readonly string _logFilePath;

        public UpdateChecker(string? updateCheckUrl = null, string? installationPath = null)
        {
            // Get current version from assembly
            _currentVersion = GetCurrentVersion();
            
            // Default update check URL (can be overridden for testing)
            _updateCheckUrl = updateCheckUrl ?? "https://api.github.com/repos/hmygroup/HMY.Tools.Zarpa/releases/latest";
            
            // Get installation path
            _installationPath = installationPath ?? AppContext.BaseDirectory;

            // Setup logging
            _logFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CopyAsInsert",
                "UpdateChecker.log"
            );
            
            EnsureLogDirectoryExists();
        }

        private void EnsureLogDirectoryExists()
        {
            try
            {
                var logDir = Path.GetDirectoryName(_logFilePath);
                if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
            }
            catch { /* Ignore if we can't create log directory */ }
        }

        private void Log(string message)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var logMessage = $"[{timestamp}] {message}";
                System.Diagnostics.Debug.WriteLine(logMessage);
                
                File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
            }
            catch { /* Ignore logging errors */ }
        }

        /// <summary>
        /// Gets the current application version
        /// </summary>
        public string GetCurrentVersion()
        {
            var version = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
                .InformationalVersion ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
            
            // Strip git metadata (e.g., "1.0.9+7c693114c43c7bf874fce390d688ea15be6335b9" -> "1.0.9")
            version = version.Split('+')[0];
            
            return version;
        }

        /// <summary>
        /// Checks if a new version is available
        /// </summary>
        public async Task<UpdateCheckResult> CheckForUpdatesAsync()
        {
            try
            {
                Log($"Starting update check. Current version: {_currentVersion}");
                Log($"Update check URL: {_updateCheckUrl}");

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "CopyAsInsert-UpdateChecker/1.1.0");
                    client.Timeout = TimeSpan.FromSeconds(10);

                    Log("Sending HTTP request to GitHub API...");
                    var response = await client.GetAsync(_updateCheckUrl);
                    
                    Log($"HTTP Response Status: {(int)response.StatusCode} ({response.StatusCode})");

                    if (!response.IsSuccessStatusCode)
                    {
                        Log($"ERROR: HTTP {(int)response.StatusCode}: {response.StatusCode}");
                        return new UpdateCheckResult
                        {
                            ErrorMessage = $"HTTP Error {(int)response.StatusCode}"
                        };
                    }

                    var content = await response.Content.ReadAsStringAsync();
                    Log($"Response content length: {content.Length} bytes");
                    Log($"Raw response (first 500 chars): {content.Substring(0, Math.Min(500, content.Length))}");

                    using (JsonDocument doc = JsonDocument.Parse(content))
                    {
                        Log("JSON parsed successfully");
                        var root = doc.RootElement;

                        // Get tag name and convert to version (e.g., "v1.0.9" -> "1.0.9")
                        var tagName = root.GetProperty("tag_name").GetString() ?? "";
                        Log($"Tag name from response: {tagName}");

                        var version = CleanVersion(tagName);
                        Log($"Cleaned version: {version}");

                        // Compare versions
                        var versionComparison = CompareVersions(_currentVersion, version);
                        Log($"Version comparison: current={_currentVersion}, latest={version}, result={versionComparison}");

                        if (versionComparison >= 0)
                        {
                            // No update available
                            Log("No update available");
                            return new UpdateCheckResult
                            {
                                IsUpdateAvailable = false,
                                CurrentVersion = _currentVersion,
                                AvailableVersion = version
                            };
                        }

                        // Update is available - find the zip asset
                        var assets = root.GetProperty("assets");
                        Log($"Found {assets.GetArrayLength()} assets");

                        string? downloadUrl = null;
                        foreach (var asset in assets.EnumerateArray())
                        {
                            var assetName = asset.GetProperty("name").GetString() ?? "";
                            Log($"Checking asset: {assetName}");

                            // Look for the compressed release file (CopyAsInsert-v*.zip)
                            if (assetName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                            {
                                downloadUrl = asset.GetProperty("browser_download_url").GetString();
                                Log($"Found ZIP asset download URL: {downloadUrl}");
                                break;
                            }
                        }

                        // Get release notes
                        var body = root.GetProperty("body").GetString() ?? "";
                        Log($"Release notes retrieved: {body.Length} characters");

                        Log("Update available: True");
                        return new UpdateCheckResult
                        {
                            IsUpdateAvailable = true,
                            CurrentVersion = _currentVersion,
                            AvailableVersion = version,
                            DownloadUrl = downloadUrl,
                            ReleaseNotes = body
                        };
                    }
                }
            }
            catch (JsonException ex)
            {
                Log($"JSON parsing error: {ex.Message}");
                return new UpdateCheckResult { ErrorMessage = "Invalid response from server" };
            }
            catch (HttpRequestException ex)
            {
                Log($"Network error: {ex.Message}");
                return new UpdateCheckResult { ErrorMessage = $"Network error: {ex.Message}" };
            }
            catch (TaskCanceledException ex)
            {
                Log($"Request timeout: {ex.Message}");
                return new UpdateCheckResult { ErrorMessage = "Request timed out" };
            }
            catch (Exception ex)
            {
                Log($"ERROR: {ex.GetType().Name}: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                return new UpdateCheckResult { ErrorMessage = ex.Message };
            }
        }

        private string CleanVersion(string version)
        {
            // Remove leading 'v' if present
            if (version.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                version = version.Substring(1);
            return version;
        }

        /// <summary>
        /// Compares two semantic versions.
        /// Returns: negative if v1 < v2, zero if equal, positive if v1 > v2
        /// </summary>
        private int CompareVersions(string v1, string v2)
        {
            try
            {
                var parts1 = v1.Split('.');
                var parts2 = v2.Split('.');

                int maxParts = Math.Max(parts1.Length, parts2.Length);

                for (int i = 0; i < maxParts; i++)
                {
                    int num1 = i < parts1.Length && int.TryParse(parts1[i], out int n1) ? n1 : 0;
                    int num2 = i < parts2.Length && int.TryParse(parts2[i], out int n2) ? n2 : 0;

                    if (num1 < num2) return -1;
                    if (num1 > num2) return 1;
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Gets the path to the updater executable
        /// </summary>
        public string GetUpdaterPath()
        {
            return Path.Combine(_installationPath, "CopyAsInsert.Updater.exe");
        }
    }

    /// <summary>
    /// Result of an update check
    /// </summary>
    public class UpdateCheckResult
    {
        public bool IsUpdateAvailable { get; set; }
        public string CurrentVersion { get; set; } = string.Empty;
        public string? AvailableVersion { get; set; }
        public string? DownloadUrl { get; set; }
        public string? ReleaseNotes { get; set; }
        public string? ErrorMessage { get; set; }

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
                return $"Error: {ErrorMessage}";

            if (IsUpdateAvailable)
                return $"Update available: {CurrentVersion} → {AvailableVersion}";

            return $"You are on the latest version ({CurrentVersion})";
        }
    }
}
