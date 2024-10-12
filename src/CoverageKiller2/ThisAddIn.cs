//https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-word?view=vs-2022&tabs=csharp

//https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-a-custom-tab-by-using-ribbon-xml?view=vs-2022&tabs=csharp

using Serilog;

namespace CoverageKiller2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string logFile = LogTailLoader.GetBareTailLog();
            LoggingLoader.Configure(logFile, Serilog.Events.LogEventLevel.Debug);
            LogTailLoader.StartBareTail();

            Log.Debug("Logging started: Level {logEventLevel}", LoggingLoader.Level);
            Log.Information("ThisAddin started.");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Log.Information("ThisAddin shutting down.");
            LoggingLoader.Cleanup();
            LogTailLoader.Cleanup();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CKRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
