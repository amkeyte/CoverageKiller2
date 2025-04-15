using Serilog;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CoverageKiller2.Logging
{
    /// <summary>
    /// Handles loading and managing LogExpert log files.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    public class LogExpertLoader
    {
        /// <summary>
        /// Gets a value indicating whether the log viewer is currently open.
        /// </summary>
        public static bool LogOpen { get; private set; }

        /// <summary>
        /// Gets the name of the log file currently in use.
        /// </summary>
        public static string LogFileName { get; private set; }

        /// <summary>
        /// Creates a temporary log file and records the location in settings for reuse.
        /// </summary>
        /// <returns>The full file path of the created log file.</returns>
        public static string GetLogFile()
        {
            string filePath = Properties.Settings.Default.LastLogFile;

            filePath = Path.GetTempFileName();
            File.WriteAllText(filePath, $"Log file created at {DateTime.Now} \n");

            Properties.Settings.Default.LastLogFile = filePath;
            Properties.Settings.Default.Save();

            return filePath;
        }

        /// <summary>
        /// Starts LogExpert and opens the given log file, if it isn't already running.
        /// </summary>
        /// <param name="filePath">The file to open in LogExpert.</param>
        public static void StartLogExpert(string filePath, bool restartIfOpen)
        {
            var x = Process.GetProcessesByName("LogExpert").FirstOrDefault();
            if (x != null && restartIfOpen)
            {
                StopLogExpert();
            }

            string logExpertPath = Properties.Settings.Default.LogExpertPath;
            Process.Start(logExpertPath, $"\"{filePath}\"");
            LogOpen = true;
            LogFileName = filePath;

        }

        /// <summary>
        /// Attempts to close any running LogExpert process gracefully, falling back to force-kill if necessary.
        /// </summary>
        public static void StopLogExpert()
        {
            var processes = Process.GetProcessesByName("LogExpert");

            foreach (var process in processes)
            {
                try
                {
                    process.CloseMainWindow();
                    process.WaitForExit(5000);

                    if (!process.HasExited)
                    {
                        process.Kill();
                        process.WaitForExit();
                    }

                    LogOpen = false;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error stopping LogExpert process {process.Id}: {ex.Message}", ex);
                }
            }

            if (!processes.Any())
            {
                Console.WriteLine("No LogExpert process found.");
            }
        }

        /// <summary>
        /// Prompts the user to close LogExpert, and shuts it down if confirmed.
        /// </summary>
        public static void Cleanup()
        {
            try
            {
                if (LogOpen)
                {
                    Log.Information("Waiting for user input.");
                    DialogResult result = MessageBox.Show(
                        "Close logging viewer?",
                        "Close Logging Viewer",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        StopLogExpert();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error during cleanup: {ex.Message}");
            }
        }
    }
}
