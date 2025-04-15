using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CoverageKiller2.Logging
{




    /// <summary>
    /// Handles loading and managing BareTail log files.
    /// </summary>
    public class LogTailLoader
    {

        public static bool LogOpen { get; private set; }
        public static string LogFileName { get; private set; }
        /// <summary>
        /// Creates a log file for BareTail and returns the file path.
        /// </summary>
        /// <returns>The file path of the created log file.</returns>
        public static string GetLogFile()
        {
            string filePath = Properties.Settings.Default.LastLogFile;

            //// Check if previous log file exists
            //if (!string.IsNullOrEmpty(filePath))
            //{
            //    File.AppendAllText(filePath, $"\nLog file reused at {DateTime.Now} \n");
            //    //return tempFilePath;
            //}
            //else
            //{
            // Otherwise, create a new temporary file
            filePath = Path.GetTempFileName();
            File.WriteAllText(filePath, $"Log file created at {DateTime.Now} \n");
            //}

            // Save the new file path for reuse on next run
            Properties.Settings.Default.LastLogFile = filePath;
            Properties.Settings.Default.Save();

            return filePath;
        }

        public static void StartBareTail(string filePath)
        {
            // Check if BareTail is already running
            if (!Process.GetProcessesByName("BareTail").Any())
            {
                //// Ask the user if they want to open BareTail
                //var result = MessageBox.Show("Do you want to open the log viewer (BareTail)?", "Open Log Viewer", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                // If the user selects "Yes," start BareTail
                //if (result == DialogResult.Yes)
                //{
                string bareTailPath = Properties.Settings.Default.BareTailPath;
                Process.Start(bareTailPath, $"\"{filePath}\"");
                LogOpen = true;
                LogFileName = filePath;
                //}
            }
        }
        public static void StopBareTail()
        {
            var processes = Process.GetProcessesByName("BareTail");

            if (processes.Any())
            {
                foreach (var process in processes)
                {
                    try
                    {
                        // First, try to close gracefully
                        process.CloseMainWindow();
                        process.WaitForExit(5000); // Wait up to 5 seconds

                        if (!process.HasExited)
                        {
                            // Force kill if still running
                            process.Kill();
                            process.WaitForExit();
                        }
                        LogOpen = false;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error stopping BareTail process {process.Id}: {ex.Message}", ex);
                    }
                }
            }
            else
            {
                Console.WriteLine("No BareTail process found.");
            }
        }


        public static void Cleanup()
        {
            try
            {
                if (LogOpen)
                {
                    // Prompt the user to close the logging viewer.
                    DialogResult result = MessageBox.Show(
                        "Close logging viewer?",
                        "Close Logging Viewer",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    // If the user confirms by clicking yes, then stop the logging process.
                    if (result == DialogResult.Yes)
                    {
                        StopBareTail();
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
