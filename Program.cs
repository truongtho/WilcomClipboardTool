using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using Microsoft.Toolkit.Uwp.Notifications;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SMSDesignAgent
{
    public static class AppConfig
    {
        public static string DesignLibraryApiUrl { get; set; } = "https://sms-api.mbczone.com/api/DesignLibrary/";
        public static string DeviceAuthApiUrl { get; set; } = "https://sms-api.mbczone.com/api/DeviceAuth";
        public static string DeviceAuthFrontendUrl { get; set; } = "https://sms.mbczone.com/device-auth";
        public static string ApplicationReleaseApiUrl => DeviceAuthApiUrl.Replace("/DeviceAuth", "/ApplicationRelease", StringComparison.OrdinalIgnoreCase);

        static AppConfig()
        {
            string configPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            if (File.Exists(configPath))
            {
                try
                {
                    string json = File.ReadAllText(configPath);
                    var doc = JsonDocument.Parse(json);
                    
                    if (doc.RootElement.TryGetProperty("DesignLibraryApiUrl", out var clipboardProp))
                        DesignLibraryApiUrl = clipboardProp.GetString() ?? DesignLibraryApiUrl;
                        
                    if (doc.RootElement.TryGetProperty("DeviceAuthApiUrl", out var authProp))
                        DeviceAuthApiUrl = authProp.GetString() ?? DeviceAuthApiUrl;
                        
                    if (doc.RootElement.TryGetProperty("DeviceAuthFrontendUrl", out var frontendProp))
                        DeviceAuthFrontendUrl = frontendProp.GetString() ?? DeviceAuthFrontendUrl;
                }
                catch { }
            }
            else
            {
                try
                {
                    var defaultConfig = new
                    {
                        DesignLibraryApiUrl = DesignLibraryApiUrl,
                        DeviceAuthApiUrl = DeviceAuthApiUrl,
                        DeviceAuthFrontendUrl = DeviceAuthFrontendUrl
                    };
                    File.WriteAllText(configPath, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
                }
                catch { }
            }
        }
    }

    class Program
    {


        [STAThread]
        static void Main(string[] args)
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool isFirstInstance;
            Mutex appMutex = new Mutex(true, "SMSDesignAgent_SingleInstance", out isFirstInstance);

            // Listen to toast popup button clicks
            ToastNotificationManagerCompat.OnActivated += toastArgs =>
            {
                ToastArguments argsDict = ToastArguments.Parse(toastArgs.Argument);
                if (argsDict.Contains("action") && argsDict["action"] == "copyCode" && argsDict.Contains("userCode"))
                {
                    Thread thread = new Thread(() =>
                    {
                        try { Clipboard.SetText(argsDict["userCode"]); } catch { }
                    });
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                }
                else if (argsDict.Contains("action") && argsDict["action"] == "updateApp" && argsDict.Contains("downloadUrl"))
                {
                    string downloadUrl = argsDict["downloadUrl"];
                    string downloadApiUrl = $"{AppConfig.ApplicationReleaseApiUrl}/latest/SmsDesignAgent/download";
                    Task.Run(async () => await UpdateManager.DownloadAndInstallUpdateAsync(downloadApiUrl, downloadUrl));
                }
            };

            RegisterUriScheme();

            // When run from Windows or Browser, the URI might be split into multiple args or wrapped in quotes.
            string fullArgs = string.Join(" ", args);
            bool handledUri = false;
            if (fullArgs.Contains("sms-design-agent://", StringComparison.OrdinalIgnoreCase))
            {
                var match = Regex.Match(fullArgs, @"sms-design-agent://\S+", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    HandleUriInvocation(match.Value).GetAwaiter().GetResult();
                    handledUri = true;
                    
                    // If a tray app is ALREADY running, this short-lived instance just exits now.
                    // If this is the FIRST instance, we want it to stay alive in the tray!
                    if (!isFirstInstance)
                    {
                        return; // Exit after handling URI as a secondary process
                    }
                }
            }

            // If we get here and it's NOT the first instance (e.g. user double-clicked the exe again manually)
            if (!isFirstInstance)
            {
                MessageBox.Show("SMS Design Agent is already running in the system tray.", "Already Running", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var oauthManager = new DesktopOAuthManager();
            var trayContext = new TrayApplicationContext(oauthManager);
            
            // Wire up event to refresh UI
            oauthManager.OnAuthStatusChanged += (s, e) =>
            {
                // Ensure UI updates happen on main thread
                if (System.Windows.Forms.Application.OpenForms.Count > 0)
                {
                    System.Windows.Forms.Application.OpenForms[0].Invoke(new Action(() => trayContext.UpdateStatus()));
                }
                else
                {
                    trayContext.UpdateStatus();
                }
            };

            // Start polling async
            _ = oauthManager.GetAccessTokenAsync();

            // Background async auto-update check
            _ = UpdateManager.CheckForUpdatesAsync();

            Application.Run(trayContext);

            // Keep Mutex alive until app exits explicitly
            GC.KeepAlive(appMutex);
        }

        public static async void HandleHotKey()
        {
            try
            {
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: HandleHotKey Triggered.\n");
                await UploadClipboardToApiAsync();
            }
            catch (Exception ex)
            {
                // In a background app, we swallow exceptions or log them to a file.
                File.AppendAllText(GetLogPath("error.log"), $"{DateTime.Now}: {ex.Message}{Environment.NewLine}");
            }
        }

        static async Task UploadClipboardToApiAsync()
        {
            var formats = ClipboardHelper.GetClipboardFormats();
            var customFormats = formats.Where(f => f.Id >= 0xC000).ToList();

            if (!formats.Any())
            {
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Clipboard is either empty or OpenClipboard failed immediately.\n");
                return;
            }

            if (!customFormats.Any())
            {
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: No custom >0xC000 formats found. Only {formats.Count} basic formats exist.\n");
                return;
            }

            File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Found {customFormats.Count} custom clipboard formats. Extracting...\n");

            List<ClipboardDataExport> exportList = new List<ClipboardDataExport>();

            foreach (var customFormat in customFormats)
            {
                byte[]? data = ClipboardHelper.GetClipboardDataBytes(customFormat.Id);
                
                if (data != null && data.Length > 0)
                {
                    exportList.Add(new ClipboardDataExport { FormatName = customFormat.Name, Data = data });
                }
            }

            if (exportList.Any())
            {
                string json = JsonSerializer.Serialize(exportList, new JsonSerializerOptions { WriteIndented = false });
                
                using var sha256 = System.Security.Cryptography.SHA256.Create();
                var hashBytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(json));
                string hashString = Convert.ToHexString(hashBytes).ToLowerInvariant();

                var payload = new { hash = hashString, wilcomObjectJson = json };

                var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
                };
                
                using HttpClient devClient = new HttpClient(handler);
                devClient.Timeout = TimeSpan.FromSeconds(30);

                var oauthManager = new DesktopOAuthManager();
                string? token = await oauthManager.GetAccessTokenAsync();
                
                if (string.IsNullOrEmpty(token))
                {
                    // Flow kicked off, abort current silent operation
                    return;
                }

                devClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                try 
                {
                    string familiesUrl = $"{AppConfig.DesignLibraryApiUrl.TrimEnd('/')}/families";
                    var familyPayload = new { name = "Clipboard Upload " + DateTime.Now.ToString("g") };
                    var familyResponse = await devClient.PostAsJsonAsync(familiesUrl, familyPayload);
                    
                    if (familyResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                    {
                        oauthManager.ClearTokens();
                        await oauthManager.GetAccessTokenAsync();
                        return;
                    }

                    if (familyResponse.IsSuccessStatusCode)
                    {
                        var familyResult = await familyResponse.Content.ReadFromJsonAsync<JsonElement>();
                        string familyId = "";
                        if (familyResult.TryGetProperty("id", out var idElement))
                        {
                            familyId = idElement.ToString();
                        }

                        if (!string.IsNullOrEmpty(familyId))
                        {
                            string variantsUrl = $"{familiesUrl}/{familyId}/variants";
                            var variantPayload = new { wilcomObjectJson = json, sizeLabel = "Default", colorwayName = "Default" };
                            var variantResponse = await devClient.PostAsJsonAsync(variantsUrl, variantPayload);

                            if (variantResponse.IsSuccessStatusCode)
                            {
                                new ToastContentBuilder()
                                    .AddText("Design uploaded!")
                                    .AddText("Successfully saved to cloud.")
                                    .Show();
                            }
                            else
                            {
                                string errorResponse = await variantResponse.Content.ReadAsStringAsync();
                                File.AppendAllText(GetLogPath("error.log"), $"{DateTime.Now}: API Variant Upload failed with response: {errorResponse}\n");
                                new ToastContentBuilder()
                                    .AddText("Upload Failed")
                                    .AddText($"Variant Status: {variantResponse.StatusCode}")
                                    .Show();
                            }
                        }
                    }
                    else
                    {
                        string errorResponse = await familyResponse.Content.ReadAsStringAsync();
                        File.AppendAllText(GetLogPath("error.log"), $"{DateTime.Now}: API Family Upload failed with response: {errorResponse}\n");
                        new ToastContentBuilder()
                            .AddText("Upload Failed")
                            .AddText($"Family Status: {familyResponse.StatusCode}")
                            .Show();
                    }
                }
                catch (Exception ex)
                {
                    File.AppendAllText(GetLogPath("error.log"), $"{DateTime.Now}: API Upload exception - {ex.Message}\n");
                    new ToastContentBuilder()
                        .AddText("Upload Failed")
                        .AddText("An error occurred. Check error.log for details.")
                        .Show();
                }
            }
        }

        static string GetLogPath(string filename)
        {
            return Path.Combine(AppContext.BaseDirectory, filename);
        }

        static void RegisterUriScheme()
        {
            try
            {
                string scheme = "sms-design-agent";
                // System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName often points to dotnet.exe during 'dotnet run'
                string executablePath = Environment.ProcessPath ?? System.Reflection.Assembly.GetExecutingAssembly().Location;
                
                // If running via 'dotnet run', we want to register the compiled .exe, not dotnet.exe
                if (executablePath.EndsWith("dotnet.exe", StringComparison.OrdinalIgnoreCase))
                {
                    // Fallback for dotnet run: assume the .dll path just needs .exe
                    executablePath = System.Reflection.Assembly.GetExecutingAssembly().Location.Replace(".dll", ".exe");
                }

                using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{scheme}"))
                {
                    key.SetValue("", $"URL:{scheme} Protocol");
                    key.SetValue("URL Protocol", "");

                    using (RegistryKey defaultIcon = key.CreateSubKey("DefaultIcon"))
                    {
                        defaultIcon.SetValue("", $"{executablePath},1");
                    }

                    using (RegistryKey commandKey = key.CreateSubKey(@"shell\open\command"))
                    {
                        commandKey.SetValue("", $"\"{executablePath}\" \"%1\"");
                    }
                }
                
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Successfully registered URI scheme '{scheme}' to point to {executablePath}\n");
            }
            catch (Exception ex)
            {
                File.AppendAllText(GetLogPath("error.log"), $"{DateTime.Now}: Failed to register URI scheme - {ex.Message}{Environment.NewLine}");
            }
        }

        static async Task HandleUriInvocation(string uri)
        {
            try
            {
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: HandleUriInvocation called with URI: {uri}\n");
                string payload = uri
                    .Replace("sms-design-agent://", "", StringComparison.OrdinalIgnoreCase)
                    .Replace("copy/?id=", "", StringComparison.OrdinalIgnoreCase)
                    .Replace("copy?id=", "", StringComparison.OrdinalIgnoreCase)
                    .Trim('/', '"', '\'', ' ');
                
                if (string.IsNullOrEmpty(payload)) 
                {
                    File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Payload was empty after processing.\n");
                    return;
                }

                string apiUrl = payload.StartsWith("http", StringComparison.OrdinalIgnoreCase) 
                                ? payload 
                                : $"{AppConfig.DesignLibraryApiUrl.TrimEnd('/')}/variants/{payload}";

                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Constructed API URL: {apiUrl}\n");

                using HttpClient client = new HttpClient();
                
                var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
                };
                client.CancelPendingRequests();
                client.Dispose();
                
                using HttpClient devClient = new HttpClient(handler);
                devClient.Timeout = TimeSpan.FromSeconds(15);
                
                var oauthManager = new DesktopOAuthManager();
                string? token = await oauthManager.GetAccessTokenAsync();
                
                if (string.IsNullOrEmpty(token))
                {
                    return;
                }
                
                devClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                
                var response = await devClient.GetAsync(apiUrl);
                
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: API Response Status: {(int)response.StatusCode} {response.StatusCode}\n");

                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    oauthManager.ClearTokens();
                    await oauthManager.GetAccessTokenAsync();
                    return;
                }

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    
                    using JsonDocument document = JsonDocument.Parse(jsonResponse);
                    string? wilcomObjectJson = null;

                    if (document.RootElement.TryGetProperty("blobPath", out JsonElement blobElement) && blobElement.ValueKind == JsonValueKind.String)
                    {
                        string blobUrl = blobElement.GetString() ?? "";
                        if (!string.IsNullOrEmpty(blobUrl))
                        {
                            using HttpClient simpleClient = new HttpClient();
                            var blobResponse = await simpleClient.GetAsync(blobUrl);
                            if (blobResponse.IsSuccessStatusCode)
                            {
                                wilcomObjectJson = await blobResponse.Content.ReadAsStringAsync();
                                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Fetched JSON from Azure Blob.\n");
                            }
                            else 
                            {
                                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Failed to download Blob JSON. Status: {blobResponse.StatusCode}\n");
                            }
                        }
                    }
                    else if (document.RootElement.TryGetProperty("wilcomObjectJson", out JsonElement legacyElement))
                    {
                        wilcomObjectJson = legacyElement.GetString() ?? "";
                    }

                    if (!string.IsNullOrEmpty(wilcomObjectJson))
                    {
                        var exportList = JsonSerializer.Deserialize<List<ClipboardDataExport>>(wilcomObjectJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                        
                        if (exportList != null && exportList.Any())
                        {
                            var clipboardDataList = new List<ClipboardHelper.ClipboardData>();
                            foreach (var export in exportList)
                            {
                                uint formatId = ClipboardHelper.RegisterClipboardFormat(export.FormatName);
                                if (formatId > 0 && export.Data != null && export.Data.Length > 0)
                                {
                                    clipboardDataList.Add(new ClipboardHelper.ClipboardData
                                    {
                                        FormatId = formatId,
                                        Data = export.Data
                                    });
                                }
                            }

                            if (ClipboardHelper.SetMultipleClipboardData(clipboardDataList))
                            {
                                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Clipboard data set successfully. Showing toast.\n");
                                new ToastContentBuilder()
                                    .AddText("Design copied! Paste in Wilcom with Ctrl+V.")
                                    .Show(toast =>
                                    {
                                        toast.ExpirationTime = DateTime.Now.AddSeconds(5);
                                    });
                                
                                // Keep the hidden win32 process alive for 6 seconds so the Toast Notification isn't killed
                                await Task.Delay(6000); 
                            }
                            else
                            {
                                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: SetMultipleClipboardData returned false.\n");
                            }
                        }
                    }
                    else
                    {
                         File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Valid 'wilcomObjectJson' missing in both API and Blob.\n");
                    }
                }
                else
                {
                    string errorResponse = await response.Content.ReadAsStringAsync();
                    File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: API call failed with response: {errorResponse}\n");
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Exception thrown inside HandleUriInvocation: {ex.Message}{Environment.NewLine}{ex.StackTrace}\n");
                File.AppendAllText(GetLogPath("error.log"), $"{DateTime.Now}: URI Invocation Error - {ex.Message}{Environment.NewLine}");
            }
        }
    }

    public class ClipboardDataExport
    {
        public string FormatName { get; set; } = "";
        public byte[] Data { get; set; } = Array.Empty<byte>();
    }

    public static class ClipboardHelper
    {
        public class ClipboardData
        {
            public uint FormatId { get; set; }
            public byte[] Data { get; set; } = Array.Empty<byte>();
        }

        public class ClipboardFormat
        {
            public uint Id { get; set; }
            public string Name { get; set; } = "";
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool EmptyClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint EnumClipboardFormats(uint format);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClipboardFormatName(uint format, [Out] StringBuilder lpszFormatName, int cchMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern uint RegisterClipboardFormat(string lpszFormat);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern UIntPtr GlobalSize(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalFree(IntPtr hMem);

        const uint GMEM_MOVEABLE = 0x0002;
        const uint GMEM_ZEROINIT = 0x0040;
        const uint GHND = (GMEM_MOVEABLE | GMEM_ZEROINIT);

        /// <summary>
        /// Retrieves all currently available clipboard formats.
        /// </summary>
        public static List<ClipboardFormat> GetClipboardFormats()
        {
            var formats = new List<ClipboardFormat>();
            
            bool opened = false;
            for (int i = 0; i < 20; i++) // Retry up to 2 seconds
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    opened = true;
                    break;
                }
                System.Threading.Thread.Sleep(100);
            }

            if (!opened)
            {
                return formats;
            }

            try
            {
                uint format = 0;
                while ((format = EnumClipboardFormats(format)) != 0)
                {
                    var sb = new StringBuilder(256);
                    int len = GetClipboardFormatName(format, sb, sb.Capacity);
                    string name = len > 0 ? sb.ToString() : GetStandardFormatName(format) ?? $"Unknown ({format})";
                    
                    formats.Add(new ClipboardFormat { Id = format, Name = name });
                }
            }
            finally
            {
                CloseClipboard();
            }

            return formats;
        }

        private static string? GetStandardFormatName(uint format)
        {
            return format switch
            {
                1 => "CF_TEXT",
                2 => "CF_BITMAP",
                3 => "CF_METAFILEPICT",
                4 => "CF_SYLK",
                5 => "CF_DIF",
                6 => "CF_TIFF",
                7 => "CF_OEMTEXT",
                8 => "CF_DIB",
                9 => "CF_PALETTE",
                10 => "CF_PENDATA",
                11 => "CF_RIFF",
                12 => "CF_WAVE",
                13 => "CF_UNICODETEXT",
                14 => "CF_ENHMETAFILE",
                15 => "CF_HDROP",
                16 => "CF_LOCALE",
                17 => "CF_DIBV5",
                _ => null
            };
        }

        /// <summary>
        /// Gets raw binary data for a specific clipboard format.
        /// </summary>
        public static byte[]? GetClipboardDataBytes(uint formatId)
        {
            bool opened = false;
            for (int i = 0; i < 20; i++)
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    opened = true;
                    break;
                }
                System.Threading.Thread.Sleep(100);
            }

            if (!opened)
                return null;

            try
            {
                IntPtr hMem = GetClipboardData(formatId);
                if (hMem == IntPtr.Zero)
                    return null;

                IntPtr pMem = GlobalLock(hMem);
                if (pMem == IntPtr.Zero)
                    return null;

                try
                {
                    ulong size = (ulong)GlobalSize(hMem);
                    if (size == 0) return null;

                    byte[] data = new byte[size];
                    Marshal.Copy(pMem, data, 0, (int)size);
                    return data;
                }
                finally
                {
                    GlobalUnlock(hMem);
                }
            }
            finally
            {
                CloseClipboard();
            }
        }

        /// <summary>
        /// Clears the clipboard and sets multiple custom formats.
        /// </summary>
        public static bool SetMultipleClipboardData(List<ClipboardData> dataList)
        {
            if (dataList == null || !dataList.Any())
                return false;

            if (!OpenClipboard(IntPtr.Zero))
                return false;

            try
            {
                EmptyClipboard();

                foreach (var item in dataList)
                {
                    if (item.Data == null || item.Data.Length == 0)
                        continue;

                    IntPtr hMem = GlobalAlloc(GHND, (UIntPtr)item.Data.Length);
                    if (hMem == IntPtr.Zero)
                        continue;

                    IntPtr pMem = GlobalLock(hMem);
                    if (pMem == IntPtr.Zero)
                    {
                        GlobalFree(hMem);
                        continue;
                    }

                    try
                    {
                        Marshal.Copy(item.Data, 0, pMem, item.Data.Length);
                    }
                    finally
                    {
                        GlobalUnlock(hMem);
                    }

                    IntPtr result = SetClipboardData(item.FormatId, hMem);
                    if (result == IntPtr.Zero)
                    {
                        GlobalFree(hMem);
                    }
                    // Do NOT GlobalFree hMem if SetClipboardData succeeds, 
                    // the system now owns the memory.
                }

                return true;
            }
            finally
            {
                CloseClipboard();
            }
        }
    }
}
