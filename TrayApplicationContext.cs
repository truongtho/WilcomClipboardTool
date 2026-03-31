using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace SMSDesignAgent
{
    public class TrayApplicationContext : ApplicationContext, IMessageFilter
    {
        private NotifyIcon _trayIcon;
        private DesktopOAuthManager _oauthManager;
        private ToolStripMenuItem _statusMenuItem;
        private ToolStripMenuItem _copyCodeMenuItem;
        private ToolStripMenuItem _checkUpdateMenuItem;
        
        // HotKey Win32 constants
        public const int WM_HOTKEY = 0x0312;
        public const int HOTKEY_ID = 1;
        public const int MOD_ALT = 0x0001;
        public const int MOD_CONTROL = 0x0002;
        public const int MOD_SHIFT = 0x0004;
        public const int MOD_WIN = 0x0008;
        public const int VK_K = 0x4B;

        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);

        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public TrayApplicationContext(DesktopOAuthManager oauthManager)
        {
            _oauthManager = oauthManager;

            _statusMenuItem = new ToolStripMenuItem("Status: Checking...");
            _statusMenuItem.Enabled = false;

            _copyCodeMenuItem = new ToolStripMenuItem("Copy Device Auth Code", null, CopyCode_Click);
            _copyCodeMenuItem.Visible = false;

            _checkUpdateMenuItem = new ToolStripMenuItem("Check for updates", null, CheckUpdate_Click);

            string versionStr = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0.0";
            var versionMenuItem = new ToolStripMenuItem($"Version {versionStr}");
            versionMenuItem.Enabled = false;
            InitializeAutoStartDefault();

            var exitMenuItem = new ToolStripMenuItem("Exit", null, Exit_Click);

            _trayIcon = new NotifyIcon()
            {
                Icon = GetEmbeddedIcon(),
                ContextMenuStrip = new ContextMenuStrip(),
                Visible = true,
                Text = "SMS Design Agent"
            };

            _trayIcon.ContextMenuStrip.Items.Add(_statusMenuItem);
            _trayIcon.ContextMenuStrip.Items.Add(versionMenuItem);
            _trayIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            _trayIcon.ContextMenuStrip.Items.Add(_copyCodeMenuItem);
            _trayIcon.ContextMenuStrip.Items.Add(_checkUpdateMenuItem);
            _trayIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            _trayIcon.ContextMenuStrip.Items.Add(exitMenuItem);

            // Hook application messages for HotKey
            Application.AddMessageFilter(this);
            RegisterHotKey(IntPtr.Zero, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, VK_K);

            UpdateStatus();
        }

        private Icon GetEmbeddedIcon()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                string resourceName = System.Linq.Enumerable.FirstOrDefault(assembly.GetManifestResourceNames(), x => x.EndsWith("SMSDesignAgent.ico", StringComparison.OrdinalIgnoreCase)) ?? string.Empty;
                if (!string.IsNullOrEmpty(resourceName))
                {
                    var stream = assembly.GetManifestResourceStream(resourceName);
                    if (stream != null) return new Icon(stream);
                }
            }
            catch { }
            return SystemIcons.Application;
        }

        public void UpdateStatus()
        {
            string? token = _oauthManager.GetStoredAccessToken();
            if (!string.IsNullOrEmpty(token))
            {
                _statusMenuItem.Text = "Status: Authorized";
                _copyCodeMenuItem.Visible = false;
            }
            else
            {
                _statusMenuItem.Text = "Status: Authorization Pending";
                if (!string.IsNullOrEmpty(_oauthManager.CurrentUserCode))
                {
                    _copyCodeMenuItem.Visible = true;
                }
            }
        }

        private void CopyCode_Click(object? sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_oauthManager.CurrentUserCode))
            {
                Clipboard.SetText(_oauthManager.CurrentUserCode);
                
                // Show standard Windows balloon tip via NotifyIcon instead of Toast for simplicity in Tray app
                _trayIcon.ShowBalloonTip(3000, "Code Copied", $"Authorization code {_oauthManager.CurrentUserCode} copied to clipboard.", ToolTipIcon.Info);
            }
        }

        private void Exit_Click(object? sender, EventArgs e)
        {
            _trayIcon.Visible = false;
            Application.Exit();
        }

        private async void CheckUpdate_Click(object? sender, EventArgs e)
        {
            _checkUpdateMenuItem.Enabled = false;
            _checkUpdateMenuItem.Text = "Checking...";
            await UpdateManager.CheckForUpdatesAsync(true);
            _checkUpdateMenuItem.Text = "Check for updates";
            _checkUpdateMenuItem.Enabled = true;
        }


        private const string RunKeyName = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "SMSDesignAgent";
        private const string AppSettingsKey = @"Software\SMSDesignAgent";

        private void InitializeAutoStartDefault()
        {
            using (RegistryKey? key = Registry.CurrentUser.CreateSubKey(AppSettingsKey))
            {
                if (key != null)
                {
                    object? initFlag = key.GetValue("AutoStartInitialized");
                    if (initFlag == null)
                    {
                        // Enable auto start on very first run
                        SetAutoStart(true);
                        key.SetValue("AutoStartInitialized", 1, RegistryValueKind.DWord);
                    }
                }
            }
        }

        private bool IsAutoStartEnabled()
        {
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(RunKeyName, false))
            {
                if (key != null)
                {
                    return key.GetValue(AppName) != null;
                }
            }
            return false;
        }

        private void SetAutoStart(bool enable)
        {
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(RunKeyName, true))
            {
                if (key != null)
                {
                    if (enable)
                    {
                        string exePath = Application.ExecutablePath;
                        if (!exePath.Contains("\""))
                        {
                            exePath = $"\"{exePath}\""; // Quote path to handle spaces
                        }
                        key.SetValue(AppName, exePath);
                    }
                    else
                    {
                        key.DeleteValue(AppName, false);
                    }
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Application.RemoveMessageFilter(this);
                UnregisterHotKey(IntPtr.Zero, HOTKEY_ID);
                _trayIcon?.Dispose();
            }
            base.Dispose(disposing);
        }

        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                // Call Program.HandleHotKey by making it internal or public
                Program.HandleHotKey();
                return true; // We handled it
            }
            return false;
        }
    }
}
