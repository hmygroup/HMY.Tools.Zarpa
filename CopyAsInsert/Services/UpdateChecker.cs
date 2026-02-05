using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Reflection;

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

        public UpdateChecker(string? updateCheckUrl = null, string? installationPath = null)
        {
            // Get current version from assembly
            _currentVersion = GetCurrentVersion();
            
            // Default update check URL (can be overridden for testing)
            _updateCheckUrl = updateCheckUrl ?? "https://api.github.com/repos/hmygroup/HMY.Tools.Zarpa/releases/latest";
            
            // Get installation path
            _installationPath = installationPath ?? AppContext.BaseDirectory;
        }

        /// <summary>
        /// Gets the current application version
        /// </summary>
        public string GetCurrentVersion()
        {
            var version = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
                .InformationalVersion ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
            
            return version;
        }

        /// <summary>
        /// Checks if a new version is available
        /// </summary>
        public async Task<UpdateCheckResult> CheckForUpdatesAsync()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "CopyAsInsert-UpdateChecker");
                    client.Timeout = TimeSpan.FromSeconds(10);

                    var response = await client.GetAsync(_updateCheckUrl);
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        // 404 means no releases published yet (normal during initial development)
                        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                        {
                            return new UpdateCheckResult
                            {
                                IsUpdateAvailable = false,
                                CurrentVersion = _currentVersion,
                                ErrorMessage = null // No error - releases just haven't been published yet
                            };
                        }

                        return new UpdateCheckResult
                        {
                            IsUpdateAvailable = false,
                            CurrentVersion = _currentVersion,
                            ErrorMessage = $"Failed to check for updates: HTTP {(int)response.StatusCode}"
                        };
                    }

                    var content = await response.Content.ReadAsStringAsync();
                    var jsonDoc = JsonDocument.Parse(content);
                    var latestVersion = ExtractVersionFromGitHubResponse(jsonDoc);

                    if (latestVersion == null)
                    {
                        return new UpdateCheckResult
                        {
                            IsUpdateAvailable = false,
                            CurrentVersion = _currentVersion,
                            ErrorMessage = "Could not parse version from response"
                        };
                    }

                    var isNewer = CompareVersions(_currentVersion, latestVersion) < 0;

                    return new UpdateCheckResult
                    {
                        IsUpdateAvailable = isNewer,
                        CurrentVersion = _currentVersion,
                        AvailableVersion = latestVersion,
                        DownloadUrl = ExtractDownloadUrlFromGitHubResponse(jsonDoc),
                        ReleaseNotes = ExtractReleaseNotesFromGitHubResponse(jsonDoc)
                    };
                }
            }
            catch (Exception ex)
            {
                return new UpdateCheckResult
                {
                    IsUpdateAvailable = false,
                    CurrentVersion = _currentVersion,
                    ErrorMessage = $"Error checking for updates: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Compares two version strings. Returns negative if v1 is older, 0 if equal, positive if v1 is newer
        /// </summary>
        private int CompareVersions(string v1, string v2)
        {
            var parts1 = v1.Split('.');
            var parts2 = v2.Split('.');

            int maxLength = Math.Max(parts1.Length, parts2.Length);

            for (int i = 0; i < maxLength; i++)
            {
                int num1 = i < parts1.Length && int.TryParse(parts1[i], out var n1) ? n1 : 0;
                int num2 = i < parts2.Length && int.TryParse(parts2[i], out var n2) ? n2 : 0;

                if (num1 < num2) return -1;
                if (num1 > num2) return 1;
            }

            return 0;
        }

        private string? ExtractVersionFromGitHubResponse(JsonDocument doc)
        {
            try
            {
                var root = doc.RootElement;
                var tagName = root.GetProperty("tag_name").GetString();
                // Remove 'v' prefix if present
                return tagName?.TrimStart('v');
            }
            catch
            {
                return null;
            }
        }

        private string? ExtractDownloadUrlFromGitHubResponse(JsonDocument doc)
        {
            try
            {
                var root = doc.RootElement;
                var assets = root.GetProperty("assets");
                
                foreach (var asset in assets.EnumerateArray())
                {
                    var fileName = asset.GetProperty("name").GetString() ?? "";
                    if (fileName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    {
                        return asset.GetProperty("browser_download_url").GetString();
                    }
                }

                return root.GetProperty("html_url").GetString();
            }
            catch
            {
                return null;
            }
        }

        private string? ExtractReleaseNotesFromGitHubResponse(JsonDocument doc)
        {
            try
            {
                var root = doc.RootElement;
                return root.GetProperty("body").GetString();
            }
            catch
            {
                return null;
            }
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
