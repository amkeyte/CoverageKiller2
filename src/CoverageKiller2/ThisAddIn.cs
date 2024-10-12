//https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-word?view=vs-2022&tabs=csharp

//https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-a-custom-tab-by-using-ribbon-xml?view=vs-2022&tabs=csharp

using Serilog;
using Serilog.Events;

namespace CoverageKiller2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string logFile = LogTailLoader.GetBareTailLog();
            // Configure Serilog
            LogEventLevel logEventLevel = LogEventLevel.Debug;
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Is(logEventLevel)
                .WriteTo.Debug() // Log to Visual Studio's Output window
                                 //.WriteTo.File(logFile, rollingInterval: RollingInterval.Day) // Log to a file
                .WriteTo.Async(a => a.File(logFile)) // Log to a file asynchronously
                .CreateLogger();

            LogTailLoader.StartBareTail();

            // Example logs

            Log.Information("Logging started: Level {logEventLevel}", logEventLevel);
            Log.Information("ThisAddin started.");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Log.Information("ThisAddin shutting down.");
            // Ensure to flush and close the log at the end of the application
            Log.CloseAndFlush();
            LogTailLoader.Cleanup();
        }


        //void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        //{
        //    Doc.Paragraphs[1].Range.InsertParagraphBefore();
        //    Doc.Paragraphs[1].Range.Text = "This Addin has loaded./n";
        //}
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
