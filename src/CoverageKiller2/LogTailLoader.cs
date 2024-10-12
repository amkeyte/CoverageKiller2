using CoverageKiller2.Properties;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace CoverageKiller2
{
    internal class LogTailLoader
    {
        private static string tempFilePath;
        public static string GetBareTailLog()
        {
            // Create a temporary text file
            //tempFilePath = Path.GetTempFileName() + ".txt";
            tempFilePath = @"C:\_LocalFiles\logs\log.txt";


            // Write some initial log data to the temp file
            File.WriteAllText(tempFilePath, $"Log file created at {DateTime.Now} \n");

            // Start BareTail to follow the log file
            //StartBareTail();

            //// Keep writing to the file for demonstration
            //for (int i = 1; i <= 10; i++)
            //{
            //    File.AppendAllText(tempFilePath, $"Log entry {i} at {DateTime.Now}\n");
            //    System.Threading.Thread.Sleep(1000); // Wait for 1 second
            //}
            return tempFilePath;
        }


        public static void StartBareTail()
        {
            // Path to the BareTail executable
            string bareTailPath = Properties.Settings.Default.BareTailPath;

            // Start BareTail process
            Process.Start(bareTailPath, $"\"{tempFilePath}\"");
        }

        public static void Cleanup()
        {



            try
            {
                if (Settings.Default.DEBUG)
                {
                    // Ask the user if they want to delete the temp file
                    DialogResult result = MessageBox.Show(
                                   "What would you like to do with the temporary log file?\n\n" +
                                   "Yes: Open in BareTail\n" +
                                   "No: Delete the file\n" +
                                   "Cancel: Do nothing",
                                   "Manage Temp File",
                                   MessageBoxButtons.YesNoCancel,
                                   MessageBoxIcon.Question
                                   );
                    if (result == DialogResult.Yes)
                    {
                        // Open the file in BareTail
                        StartBareTail();
                    }
                    else if (result == DialogResult.No)
                    {
                        // Delete the temporary file
                        if (File.Exists(tempFilePath))
                        {
                            File.Delete(tempFilePath);
                            Debug.WriteLine($"Temporary file {tempFilePath} deleted.");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("No action taken on the temporary file.");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error deleting temporary file: {ex.Message}");
            }
        }
    }






}
