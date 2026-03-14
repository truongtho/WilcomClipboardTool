using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using Microsoft.Toolkit.Uwp.Notifications;

namespace WilcomClipboardTool
{
    class Program
    {
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);

        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        public static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        public static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage(ref MSG lpMsg);

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }

        const int MOD_ALT = 0x0001;
        const int MOD_CONTROL = 0x0002;
        const int MOD_SHIFT = 0x0004;
        const int MOD_WIN = 0x0008;
        const int WM_HOTKEY = 0x0312;
        const int HOTKEY_ID = 1;
        const int VK_K = 0x4B;

        [STAThread]
        static void Main(string[] args)
        {
            if (!RegisterHotKey(IntPtr.Zero, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, VK_K))
            {
                // Failed to register hotkey. App will exit.
                return;
            }

            MSG msg;
            while (GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                if (msg.message == WM_HOTKEY && msg.wParam.ToInt32() == HOTKEY_ID)
                {
                    HandleHotKey();
                }

                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }

            UnregisterHotKey(IntPtr.Zero, HOTKEY_ID);
        }

        static void HandleHotKey()
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
