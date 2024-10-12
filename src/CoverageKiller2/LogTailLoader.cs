using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace CoverageKiller2
{
    internal class LogTailLoader
    {
        private static string tempFilePath;
        public static string GetBareTailLog()
        {
            tempFilePath = @"C:\_LocalFiles\logs\log.txt";
            File.WriteAllText(tempFilePath, $"Log file created at {DateTime.Now} \n");
            return tempFilePath;
        }

        public static void StartBareTail()
        {
            if (!Process.GetProcessesByName("BareTail").Any())
            {
                string bareTailPath = Properties.Settings.Default.BareTailPath;
                Process.Start(bareTailPath, $"\"{tempFilePath}\"");
            }
        }

        public static void Cleanup()
        {
            try
            {
                if (File.Exists(tempFilePath) && !!Process.GetProcessesByName("BareTail").Any())
                {
                    File.Delete(tempFilePath);
                    Debug.WriteLine($"Temporary file {tempFilePath} deleted.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error deleting temporary file: {ex.Message}");
            }
        }
    }






}
