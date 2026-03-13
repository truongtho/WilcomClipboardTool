using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace WilcomClipboardTool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=============================================");
            Console.WriteLine(" Wilcom Clipboard Transfer Tool");
            Console.WriteLine("=============================================\n");

            Console.WriteLine("Choose an action:");
            Console.WriteLine("1. Save current Wilcom clipboard data to file");
            Console.WriteLine("2. Load Wilcom clipboard data from file and paste to clipboard");
            Console.Write("Action (1/2): ");
            var choice = Console.ReadLine();

            string defaultFilePath = "WilcomClipboardData.json";

            if (choice == "1")
            {
                SaveClipboardToFile(defaultFilePath);
            }
            else if (choice == "2")
            {
                LoadClipboardFromFile(defaultFilePath);
            }
            else
            {
                Console.WriteLine("Invalid choice.");
            }
        }

        static void SaveClipboardToFile(string filePath)
        {
            // Scan clipboard formats
            var formats = ClipboardHelper.GetClipboardFormats();
            var customFormats = formats.Where(f => f.Id >= 0xC000).ToList();

            if (!customFormats.Any())
            {
                Console.WriteLine("\nNo custom clipboard formats found. Please copy something in Wilcom first.");
                return;
            }

            Console.WriteLine("\nFound Custom Formats (potentially Wilcom):");
            List<ClipboardDataExport> exportList = new List<ClipboardDataExport>();

            foreach (var customFormat in customFormats)
            {
                Console.WriteLine($"Reading Data for Format: {customFormat.Name} ({customFormat.Id})");
                byte[]? data = ClipboardHelper.GetClipboardDataBytes(customFormat.Id);
                
                if (data != null && data.Length > 0)
                {
                    Console.WriteLine($" => Successfully read {data.Length} bytes.");
                    exportList.Add(new ClipboardDataExport { FormatName = customFormat.Name, Data = data });
                }
            }

            if (exportList.Any())
            {
                string json = JsonSerializer.Serialize(exportList, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(filePath, json);
                Console.WriteLine($"\nSuccessfully saved {exportList.Count} formats to '{filePath}'");
                Console.WriteLine("You can now transfer this file to another machine.");
            }
            else
            {
                Console.WriteLine("\nNo valid data could be read from the clipboard.");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        static void LoadClipboardFromFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"\nFile not found: {filePath}");
                return;
            }

            Console.WriteLine($"\nLoading data from {filePath}...");
            string json = File.ReadAllText(filePath);
            var importList = JsonSerializer.Deserialize<List<ClipboardDataExport>>(json);

            if (importList == null || !importList.Any())
            {
                Console.WriteLine("No data found in the file.");
                return;
            }

            List<ClipboardHelper.ClipboardData> restoreList = new List<ClipboardHelper.ClipboardData>();

            foreach (var item in importList)
            {
                // Register the format by name to get the correct dynamic ID for THIS machine
                uint formatId = ClipboardHelper.RegisterClipboardFormat(item.FormatName);
                if (formatId == 0)
                {
                    Console.WriteLine($"Failed to register format: {item.FormatName}");
                    continue;
                }

                Console.WriteLine($"Registered format '{item.FormatName}' to local ID: {formatId}");
                restoreList.Add(new ClipboardHelper.ClipboardData 
                { 
                    FormatId = formatId, 
                    Data = item.Data 
                });
            }

            Console.WriteLine("\nClearing and restoring data to clipboard...");
            bool success = ClipboardHelper.SetMultipleClipboardData(restoreList);
            
            if (success)
            {
                Console.WriteLine("Data successfully restored to clipboard! You can now paste in Wilcom.");
            }
            else
            {
                Console.WriteLine("Failed to restore clipboard data.");
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
