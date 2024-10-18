using Serilog;

namespace CoverageKiller2
{
    /// <summary>
    /// The main class for the VSTO add-in. Handles initialization, shutdown,
    /// and custom ribbon functionality for the Word add-in.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Initializes logging and BareTail when the add-in starts.
        /// </summary>
        /// <param name="sender">The event source.</param>
        /// <param name="e">Event arguments.</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string logFile = LogTailLoader.GetBareTailLog();
            LoggingLoader.Configure(logFile, Serilog.Events.LogEventLevel.Debug);

            //debugging the big hangup.
            LogTailLoader.StartBareTail();

            Log.Debug("Logging started: Level {logEventLevel}", LoggingLoader.Level);
            Log.Information("ThisAddIn started.");
        }

        /// <summary>
        /// Cleans up logging and BareTail when the add-in is shut down.
        /// </summary>
        /// <param name="sender">The event source.</param>
        /// <param name="e">Event arguments.</param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Log.Information("ThisAddIn shutting down.");
            LoggingLoader.Cleanup();
            LogTailLoader.Cleanup();
        }

        /// <summary>
        /// Creates the custom ribbon for the add-in using Ribbon XML.
        /// </summary>
        /// <returns>An object that represents the custom ribbon.</returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CKRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support. Wires up the Startup and Shutdown events for the add-in.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
