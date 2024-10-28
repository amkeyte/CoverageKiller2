using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace CoverageKiller2.Logging
{
    /// <summary>
    /// Handles loading and managing BareTail log files.
    /// </summary>
    internal class LogTailLoader
    {
        private static string tempFilePath;

        /// <summary>
        /// Creates a log file for BareTail and returns the file path.
        /// </summary>
        /// <returns>The file path of the created log file.</returns>
        public static string GetBareTailLog()
        {
            tempFilePath = @"C:\_LocalFiles\logs\log.txt";
            File.WriteAllText(tempFilePath, $"Log file created at {DateTime.Now} \n");
            return tempFilePath;
        }

        /// <summary>
        /// Starts BareTail with the log file if BareTail is not already running.
        /// </summary>
        public static void StartBareTail()
        {
            // Check if BareTail is already running
            if (!Process.GetProcessesByName("BareTail").Any())
            {
                string bareTailPath = Properties.Settings.Default.BareTailPath;
                Process.Start(bareTailPath, $"\"{tempFilePath}\"");
            }
        }

        /// <summary>
        /// Cleans up by deleting the temporary log file and closing BareTail if running.
        /// </summary>
        public static void Cleanup()
        {
            try
            {
                // Check if the file exists and BareTail is running
                //if (File.Exists(tempFilePath) && Process.GetProcessesByName("BareTail").Any())
                //{
                //    File.Delete(tempFilePath);
                //    Debug.WriteLine($"Temporary file {tempFilePath} deleted.");
                //}
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error deleting temporary file: {ex.Message}");
            }
        }
    }
}
