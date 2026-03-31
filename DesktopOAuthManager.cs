using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Win32;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.Toolkit.Uwp.Notifications;

namespace SMSDesignAgent
{
    public class DesktopOAuthManager
    {
        private string ApiBaseUrl => AppConfig.DeviceAuthApiUrl;
        private const string RegistryPath = @"Software\Desify\SMSDesignAgent";

        public class DeviceTokenResponse
        {
            public string? AccessToken { get; set; }
            public string? RefreshToken { get; set; }
            public string Status { get; set; } = string.Empty;
        }

        public class DeviceCodeResponse
        {
            public string DeviceCode { get; set; } = string.Empty;
            public string UserCode { get; set; } = string.Empty;
            public int ExpiresIn { get; set; }
        }

        public string? CurrentUserCode { get; private set; }
        public event EventHandler? OnAuthStatusChanged;

        public async Task<string?> GetAccessTokenAsync()
        {
            string? jwt = GetRegistryValue("AccessToken");
            
            if (!string.IsNullOrEmpty(jwt))
            {
                if (!IsTokenExpired(jwt))
                {
                    return jwt;
                }
                
                string? refreshToken = GetRegistryValue("RefreshToken");
                if (!string.IsNullOrEmpty(refreshToken))
                {
                    LogTrace("Access token expired. Attempting silent refresh.");
                    string? newJwt = await TrySilentRefreshAsync(refreshToken);
                    if (!string.IsNullOrEmpty(newJwt))
                    {
                        return newJwt;
                    }
                }
                
                LogTrace("Silent refresh unavailable or failed. Generating new Device Flow.");
                ClearTokens();
            }

            // No token or expired, start Auth Flow
            await StartAuthorizationFlow();
            return null; // The process is asynchronous and polling, token won't be ready immediately
        }

        public void ClearTokens()
        {
            DeleteRegistryValue("AccessToken");
            DeleteRegistryValue("RefreshToken");
        }

        private bool IsTokenExpired(string jwt)
        {
            try
            {
                var parts = jwt.Split('.');
                if (parts.Length != 3) return true;

                var payload = parts[1];
                payload = payload.Replace('-', '+').Replace('_', '/');
                switch (payload.Length % 4)
                {
                    case 2: payload += "=="; break;
                    case 3: payload += "="; break;
                }

                var jsonBytes = Convert.FromBase64String(payload);
                var json = Encoding.UTF8.GetString(jsonBytes);

                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("exp", out var expElement))
                {
                    long expSeconds = expElement.GetInt64();
                    var expDate = DateTimeOffset.FromUnixTimeSeconds(expSeconds).UtcDateTime;
                    
                    // Consider it expired 2 minutes before actual expiration to prevent edge cases
                    return expDate <= DateTime.UtcNow.AddMinutes(2);
                }
                return false;
            }
            catch
            {
                return true;
            }
        }

        private async Task<string?> TrySilentRefreshAsync(string refreshToken)
        {
            using HttpClient client = CreateHttpClient();

            try
            {
                var req = new { RefreshToken = refreshToken };
                var response = await client.PostAsJsonAsync($"{ApiBaseUrl}/refresh-token", req);
                
                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadFromJsonAsync<DeviceTokenResponse>();
                    if (result != null && result.Status == "granted" && !string.IsNullOrEmpty(result.AccessToken))
                    {
                        SaveRegistryValue("AccessToken", result.AccessToken);
                        if (!string.IsNullOrEmpty(result.RefreshToken))
                        {
                            SaveRegistryValue("RefreshToken", result.RefreshToken);
                        }

                        LogTrace("Silent refresh succeeded.");
                        return result.AccessToken;
                    }
                }
                else
                {
                    LogTrace($"Silent refresh failed: {response.StatusCode} {await response.Content.ReadAsStringAsync()}");
                }
            }
            catch (Exception ex)
            {
                LogError("TrySilentRefreshAsync", ex);
            }

            return null;
        }

        public string? GetStoredAccessToken()
        {
            return GetRegistryValue("AccessToken");
        }

        private async Task StartAuthorizationFlow()
        {
            string fingerprint = GetDeviceFingerprint();
            string machineName = Environment.MachineName;

            using HttpClient client = CreateHttpClient();

            try
            {
                var currentVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
                var appMetadata = new Dictionary<string, string>
                {
                    { "SmsDesignAgent", currentVersion }
                };

                var codeRequest = new { 
                    DeviceFingerprint = fingerprint, 
                    MachineName = machineName,
                    AppMetadata = appMetadata
                };
                var response = await client.PostAsJsonAsync($"{ApiBaseUrl}/device-code", codeRequest);
                
                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadFromJsonAsync<DeviceCodeResponse>();
                    if (result != null)
                    {
                        CurrentUserCode = result.UserCode;
                        OnAuthStatusChanged?.Invoke(this, EventArgs.Empty);

                        // Open Browser
                        OpenBrowser(AppConfig.DeviceAuthFrontendUrl);
                        
                        // Show Toast with Code
                        new ToastContentBuilder()
                            .AddText("Authorization Required")
                            .AddText($"Enter code {result.UserCode} in your browser to connect this device.")
                            .AddArgument("action", "copyCode")
                            .AddArgument("userCode", result.UserCode)
                            .AddButton(new ToastButton()
                                .SetContent("Copy Code")
                                .AddArgument("action", "copyCode")
                                .AddArgument("userCode", result.UserCode)
                            )
                            .Show(toast => { toast.ExpirationTime = DateTime.Now.AddMinutes(5); });

                        LogTrace($"Authorization flow started. UserCode: {result.UserCode}");
                        
                        // Start polling in background
                        _ = PollForTokenAsync(result.DeviceCode, result.ExpiresIn);
                    }
                }
                else
                {
                    LogTrace($"Failed to get device code: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                LogError("StartAuthorizationFlow", ex);
            }
        }

        private async Task PollForTokenAsync(string deviceCode, int expiresInSeconds)
        {
            using HttpClient client = CreateHttpClient();
            var startTime = DateTime.UtcNow;
            var timeout = TimeSpan.FromSeconds(expiresInSeconds);

            while (DateTime.UtcNow - startTime < timeout)
            {
                try
                {
                    var tokenReq = new { DeviceCode = deviceCode };
                    var response = await client.PostAsJsonAsync($"{ApiBaseUrl}/device-token", tokenReq);

                    if (response.IsSuccessStatusCode)
                    {
                        var result = await response.Content.ReadFromJsonAsync<DeviceTokenResponse>();
                        if (result != null && result.Status == "granted" && !string.IsNullOrEmpty(result.AccessToken))
                        {
                            SaveRegistryValue("AccessToken", result.AccessToken);
                            if (!string.IsNullOrEmpty(result.RefreshToken))
                            {
                                SaveRegistryValue("RefreshToken", result.RefreshToken);
                            }

                            new ToastContentBuilder()
                                .AddText("Device Connected!")
                                .AddText("You can now upload and retrieve designs from Wilcom.")
                                .Show();
                                
                            CurrentUserCode = null;
                            OnAuthStatusChanged?.Invoke(this, EventArgs.Empty);

                            LogTrace("Device authorized successfully.");
                            return;
                        }
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.BadRequest)
                    {
                        var errorJson = await response.Content.ReadAsStringAsync();
                        if (errorJson.Contains("authorization_pending"))
                        {
                            // Keep polling
                            await Task.Delay(5000);
                            continue;
                        }
                        else
                        {
                            LogTrace($"Polling stopped due to error: {errorJson}");
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError("PollForTokenAsync", ex);
                }

                await Task.Delay(5000);
            }

            LogTrace("Polling timed out.");
        }

        private string GetDeviceFingerprint()
        {
            try
            {
                // In a real application, combine Motherboard, CPU, HDD IDs using WMI.
                // For this implementation, we use MachineName + UserDomainName as a simple proxy.
                string rawId = $"{Environment.MachineName}_{Environment.UserDomainName}_{Environment.UserName}";
                
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawId));
                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                }
            }
            catch
            {
                return Guid.NewGuid().ToString("N");
            }
        }

        private HttpClient CreateHttpClient()
        {
            // Bypass SSL for localhost development
            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
            };
            return new HttpClient(handler);
        }

        private void OpenBrowser(string url)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                LogError("OpenBrowser", ex);
            }
        }

        private void SaveRegistryValue(string keyName, string value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
            {
                key.SetValue(keyName, value);
            }
        }

        private void DeleteRegistryValue(string keyName)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
            {
                key.DeleteValue(keyName, false);
            }
        }

        private string? GetRegistryValue(string keyName)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
            {
                return key?.GetValue(keyName) as string;
            }
        }

        private void LogTrace(string message)
        {
            string path = Path.Combine(AppContext.BaseDirectory, "dev_trace.log");
            File.AppendAllText(path, $"{DateTime.Now}: OAuth - {message}\n");
        }

        private void LogError(string context, Exception ex)
        {
            string path = Path.Combine(AppContext.BaseDirectory, "error.log");
            File.AppendAllText(path, $"{DateTime.Now}: OAuth Error in {context} - {ex.Message}\n");
        }
    }
}
