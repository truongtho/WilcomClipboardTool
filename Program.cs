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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Toolkit.Uwp.Notifications;

namespace WilcomClipboardTool
{
    class Program
    {


        [STAThread]
        static void Main(string[] args)
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool isFirstInstance;
            Mutex appMutex = new Mutex(true, "WilcomClipboardTool_SingleInstance", out isFirstInstance);

            // Listen to toast popup button clicks
            ToastNotificationManagerCompat.OnActivated += toastArgs =>
            {
                ToastArguments argsDict = ToastArguments.Parse(toastArgs.Argument);
                if (argsDict.Contains("action") && argsDict["action"] == "copyCode" && argsDict.Contains("userCode"))
                {
                    // Must be invoked on STA thread
                    if (System.Windows.Forms.Application.OpenForms.Count > 0)
                    {
                        System.Windows.Forms.Application.OpenForms[0].Invoke(new Action(() => 
                        {
                            Clipboard.SetText(argsDict["userCode"]);
                        }));
                    }
                }
            };

            RegisterUriScheme();

            // When run from Windows or Browser, the URI might be split into multiple args or wrapped in quotes.
            string fullArgs = string.Join(" ", args);
            bool handledUri = false;
            if (fullArgs.Contains("wilcom://", StringComparison.OrdinalIgnoreCase))
            {
                var match = Regex.Match(fullArgs, @"wilcom://\S+", RegexOptions.IgnoreCase);
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
                MessageBox.Show("Wilcom Clipboard Tool is already running in the system tray.", "Already Running", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            Application.Run(trayContext);

            // Keep Mutex alive until app exits explicitly
            GC.KeepAlive(appMutex);
        }

        public static void HandleHotKey()
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string defaultFilePath = $"WilcomClipboardData_{timestamp}.json";
                string hashFilePath = $"WilcomClipboardData_{timestamp}.sha256";

                SaveClipboardToFileAndHash(defaultFilePath, hashFilePath);
            }
            catch (Exception ex)
            {
                // In a background app, we swallow exceptions or log them to a file.
                File.AppendAllText("error.log", $"{DateTime.Now}: {ex.Message}{Environment.NewLine}");
            }
        }

        static void SaveClipboardToFileAndHash(string filePath, string hashFilePath)
        {
            var formats = ClipboardHelper.GetClipboardFormats();
            var customFormats = formats.Where(f => f.Id >= 0xC000).ToList();

            if (!customFormats.Any())
            {
                return;
            }

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
                string json = JsonSerializer.Serialize(exportList, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(filePath, json);

                // Compute SHA-256 hash of the JSON string
                byte[] jsonBytes = Encoding.UTF8.GetBytes(json);
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] hashBytes = sha256.ComputeHash(jsonBytes);
                    string hashString = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                    File.WriteAllText(hashFilePath, hashString);
                }

                new ToastContentBuilder()
                    .AddText("Design uploaded!")
                    .Show();
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
                string scheme = "wilcom";
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
                
                string payload = uri.Replace("wilcom://", "", StringComparison.OrdinalIgnoreCase).Trim('/');
                
                if (string.IsNullOrEmpty(payload)) 
                {
                    File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: Payload was empty after processing.\n");
                    return;
                }

                string apiUrl = payload.StartsWith("http", StringComparison.OrdinalIgnoreCase) 
                                ? payload 
                                : $"https://localhost:7178/api/WilcomClipboards/{payload}";

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
                string? token = oauthManager.GetStoredAccessToken();
                
                if (!string.IsNullOrEmpty(token))
                {
                    devClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                }
                
                var response = await devClient.GetAsync(apiUrl);
                
                File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: API Response Status: {(int)response.StatusCode} {response.StatusCode}\n");

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    
                    using JsonDocument document = JsonDocument.Parse(jsonResponse);
                    if (document.RootElement.TryGetProperty("wilcomObjectJson", out JsonElement jsonElement))
                    {
                        string wilcomObjectJson = jsonElement.GetString() ?? "";

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
                         File.AppendAllText(GetLogPath("dev_trace.log"), $"{DateTime.Now}: 'wilcomObjectJson' property missing in API response.\n");
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
            if (!OpenClipboard(IntPtr.Zero))
                return formats;

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
            if (!OpenClipboard(IntPtr.Zero))
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
