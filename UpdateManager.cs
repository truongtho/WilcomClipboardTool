using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Toolkit.Uwp.Notifications;

namespace SMSDesignAgent
{
    public class ApplicationReleaseResponse
    {
        public string AppName { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
        public string DownloadUrl { get; set; } = string.Empty;
        public bool IsMandatory { get; set; }
        public string? ReleaseNotes { get; set; }
    }

    public static class UpdateManager
    {
        public static async Task CheckForUpdatesAsync(bool showToastIfUpToDate = false)
        {
            try
            {
                var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
                };

                using var client = new HttpClient(handler);
                client.Timeout = TimeSpan.FromSeconds(15);
                
                string apiUrl = $"{AppConfig.ApplicationReleaseApiUrl}/latest/SmsDesignAgent";
                
                var response = await client.GetAsync(apiUrl);
                if (response.IsSuccessStatusCode)
                {
                    var release = await response.Content.ReadFromJsonAsync<ApplicationReleaseResponse>();
                    if (release != null && !string.IsNullOrEmpty(release.Version) && !string.IsNullOrEmpty(release.DownloadUrl))
                    {
                        string currentVersionStr = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
                        if (Version.TryParse(currentVersionStr, out Version current) && Version.TryParse(release.Version, out Version latest))
                        {
                            if (latest > current)
                            {
                                // We have a newer version!
                                LogTrace($"New version found: {latest}. Current: {current}. Prompting user to update...");
                                new ToastContentBuilder()
                                    .AddText("Update Available")
                                    .AddText($"A new version ({latest}) of SMS Design Agent is available.")
                                    .AddButton(new ToastButton()
                                        .SetContent("Update")
                                        .AddArgument("action", "updateApp")
                                        .AddArgument("downloadUrl", release.DownloadUrl))
                                    .Show();
                            }
                            else if (showToastIfUpToDate)
                            {
                                new ToastContentBuilder()
                                    .AddText("Up to date")
                                    .AddText($"You are running the latest version: {currentVersionStr}")
                                    .Show();
                            }
                        }
                    }
                    else if (showToastIfUpToDate)
                    {
                        new ToastContentBuilder()
                            .AddText("No updates found")
                            .AddText("There is no valid update metadata on the server.")
                            .Show();
                    }
                }
                else if (showToastIfUpToDate)
                {
                    new ToastContentBuilder()
                        .AddText("Update check failed")
                        .AddText($"Server responded with status {response.StatusCode}.")
                        .Show();
                }
            }
            catch (Exception ex)
            {
                LogError("CheckForUpdatesAsync", ex);
            }
        }

        public static async Task DownloadAndInstallUpdateAsync(string downloadApiUrl, string originalBlobUrl)
        {
            try
            {
                new ToastContentBuilder()
                    .AddText("Downloading Update")
                    .AddText("The update is downloading in the background. The app will restart automatically.")
                    .Show();

                string tempDir = Path.Combine(Path.GetTempPath(), "SMSDesignAgent_Updates");
                Directory.CreateDirectory(tempDir);

                // Assuming the backend provides a direct link to the executable or MSI
                string fileName = "SMSDesignAgent_Update.exe";
                if (originalBlobUrl.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
                {
                    fileName = "SMSDesignAgent_Installer.msi";
                }
                
                string localFilePath = Path.Combine(tempDir, fileName);

                using var client = new HttpClient();
                using var response = await client.GetAsync(downloadApiUrl, HttpCompletionOption.ResponseHeadersRead);
                
                response.EnsureSuccessStatusCode();

                using (var fs = new FileStream(localFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await response.Content.CopyToAsync(fs);
                }

                LogTrace($"Downloaded update to {localFilePath}. About to restart.");

                // Run the newly downloaded installer package
                ProcessSetup(localFilePath);

                // Terminate current process so the installer can overwrite files
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                LogError("DownloadAndInstallUpdateAsync", ex);
            }
        }

        private static void ProcessSetup(string filePath)
        {
            if (filePath.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
            {
                var info = new ProcessStartInfo(filePath)
                {
                    UseShellExecute = true
                };
                Process.Start(info);
            }
            else
            {
                string currentExe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;
                if (string.IsNullOrEmpty(currentExe))
                {
                    LogError("ProcessSetup", new Exception("Could not determine current executable path."));
                    return;
                }

                string tempRoot = Path.GetTempPath();
                string psPath = Path.Combine(tempRoot, "sms_updater.ps1");
                string logPath = Path.Combine(tempRoot, "sms_updater_log.txt");

                string psContent = $@"
$logPath = '{logPath}'
$currentExe = '{currentExe}'
$filePath = '{filePath}'

function Log([string]$msg) {{
    ""[$(Get-Date -Format 'HH:mm:ss')] $msg"" | Out-File -FilePath $logPath -Append -Encoding utf8
}}

Log 'Starting PowerShell update script'
Log ""Target exe: $currentExe""
Log ""Source file: $filePath""

$retry = 0
while ($retry -lt 15) {{
    Start-Sleep -Seconds 2
    Log ""Attempting to move file (Attempt $retry)...""
    
    try {{
        Move-Item -Path $filePath -Destination $currentExe -Force -ErrorAction Stop
        Log 'Move successful! Launching...'
        Start-Process -FilePath $currentExe
        break
    }}
    catch {{
        Log ""Move failed: $_""
        $retry++
    }}
}}

Log 'PowerShell script finished.'
Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue
";
                File.WriteAllText(psPath, psContent, new System.Text.UTF8Encoding(true));
                LogTrace($"Generated updater ps1 at: {psPath}");

                var info = new ProcessStartInfo("powershell.exe", $"-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -NoProfile -File \"{psPath}\"")
                {
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true
                };

                try
                {
                    Process.Start(info);
                    LogTrace("Started powershell.exe to execute ps1 script.");
                }
                catch (Exception ex)
                {
                    LogError("ProcessSetup:StartPs", ex);
                }
            }
        }

        private static void LogTrace(string message)
        {
            string path = Path.Combine(AppContext.BaseDirectory, "update_trace.log");
            File.AppendAllText(path, $"{DateTime.Now}: UpdateManager - {message}\n");
        }

        private static void LogError(string context, Exception ex)
        {
            string path = Path.Combine(AppContext.BaseDirectory, "error.log");
            File.AppendAllText(path, $"{DateTime.Now}: Update Error in {context} - {ex.Message}\n");
        }
    }
}
