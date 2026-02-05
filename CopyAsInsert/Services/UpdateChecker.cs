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
                        // 404 means no releases published yet (normal during initial development)
                        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                        {
                            Log("No releases found (404) - this is normal for repositories without releases");
                            return new UpdateCheckResult
                            {
                                IsUpdateAvailable = false,
                                CurrentVersion = _currentVersion,
                                ErrorMessage = null
                            };
                        }

                        var errorMsg = $"Failed to check for updates: HTTP {(int)response.StatusCode}";
                        Log(errorMsg);
                        return new UpdateCheckResult
                        {
                            IsUpdateAvailable = false,
                            CurrentVersion = _currentVersion,
                            ErrorMessage = errorMsg
                        };
                    }

                    var content = await response.Content.ReadAsStringAsync();
                    Log($"Response content length: {content.Length} bytes");
                    Log($"Raw response (first 500 chars): {content.Substring(0, Math.Min(500, content.Length))}");

                    JsonDocument jsonDoc;
                    try
                    {
                        jsonDoc = JsonDocument.Parse(content);
                        Log("JSON parsed successfully");
                    }
                    catch (Exception jsonEx)
                    {
                        Log($"Failed to parse JSON: {jsonEx.Message}");
                        return new UpdateCheckResult
                        {
                            IsUpdateAvailable = false,
                            CurrentVersion = _currentVersion,
                            ErrorMessage = $"Failed to parse GitHub response: {jsonEx.Message}"
                        };
                    }

                    var latestVersion = ExtractVersionFromGitHubResponse(jsonDoc);
                    Log($"Extracted version from response: {latestVersion}");

                    if (latestVersion == null)
                    {
                        Log("Failed to extract version from GitHub response");
                        return new UpdateCheckResult
                        {
                            IsUpdateAvailable = false,
                            CurrentVersion = _currentVersion,
                            ErrorMessage = "Could not parse version from response"
                        };
                    }

                    var compareResult = CompareVersions(_currentVersion, latestVersion);
                    Log($"Version comparison: current={_currentVersion}, latest={latestVersion}, result={compareResult}");
                    var isNewer = compareResult < 0;

                    var downloadUrl = ExtractDownloadUrlFromGitHubResponse(jsonDoc);
                    var releaseNotes = ExtractReleaseNotesFromGitHubResponse(jsonDoc);
                    
                    Log($"Download URL: {downloadUrl}");
                    Log($"Update available: {isNewer}");

                    return new UpdateCheckResult
                    {
                        IsUpdateAvailable = isNewer,
                        CurrentVersion = _currentVersion,
                        AvailableVersion = latestVersion,
                        DownloadUrl = downloadUrl,
                        ReleaseNotes = releaseNotes
                    };
                }
            }
            catch (Exception ex)
            {
                Log($"Exception during update check: {ex.GetType().Name}: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                
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
                
                if (!root.TryGetProperty("tag_name", out var tagNameElement))
                {
                    Log("ERROR: 'tag_name' property not found in response");
                    return null;
                }

                var tagName = tagNameElement.GetString();
                Log($"Tag name from response: {tagName}");
                
                if (string.IsNullOrEmpty(tagName))
                {
                    Log("ERROR: 'tag_name' is empty or null");
                    return null;
                }

                var version = tagName.TrimStart('v');
                Log($"Cleaned version: {version}");
                return version;
            }
            catch (Exception ex)
            {
                Log($"ERROR in ExtractVersionFromGitHubResponse: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        private string? ExtractDownloadUrlFromGitHubResponse(JsonDocument doc)
        {
            try
            {
                var root = doc.RootElement;
                
                if (!root.TryGetProperty("assets", out var assetsElement))
                {
                    Log("WARNING: 'assets' property not found, trying html_url fallback");
                    if (root.TryGetProperty("html_url", out var htmlUrlElement))
                    {
                        return htmlUrlElement.GetString();
                    }
                    return null;
                }

                if (assetsElement.ValueKind != JsonValueKind.Array)
                {
                    Log($"WARNING: 'assets' is not an array (type: {assetsElement.ValueKind})");
                    return null;
                }

                Log($"Found {assetsElement.GetArrayLength()} assets");

                foreach (var asset in assetsElement.EnumerateArray())
                {
                    if (!asset.TryGetProperty("name", out var nameElement))
                    {
                        Log("WARNING: Asset without 'name' property");
                        continue;
                    }

                    var fileName = nameElement.GetString() ?? "";
                    Log($"Checking asset: {fileName}");

                    if (fileName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    {
                        if (asset.TryGetProperty("browser_download_url", out var urlElement))
                        {
                            var url = urlElement.GetString();
                            Log($"Found EXE asset download URL: {url}");
                            return url;
                        }
                    }
                }

                Log("No .exe asset found, falling back to html_url");
                if (root.TryGetProperty("html_url", out var htmlUrl))
                {
                    return htmlUrl.GetString();
                }

                return null;
            }
            catch (Exception ex)
            {
                Log($"ERROR in ExtractDownloadUrlFromGitHubResponse: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        private string? ExtractReleaseNotesFromGitHubResponse(JsonDocument doc)
        {
            try
            {
                var root = doc.RootElement;
                
                if (!root.TryGetProperty("body", out var bodyElement))
                {
                    Log("WARNING: 'body' property not found");
                    return null;
                }

                var body = bodyElement.GetString();
                Log($"Release notes retrieved: {body?.Length ?? 0} characters");
                return body;
            }
            catch (Exception ex)
            {
                Log($"ERROR in ExtractReleaseNotesFromGitHubResponse: {ex.GetType().Name}: {ex.Message}");
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
